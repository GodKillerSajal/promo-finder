"""Promo Finder: finds ready-to-post promos in the scripts of a vertical microdrama.

A promo is 20-60 seconds long, grabs the viewer in the first 3 seconds (the hook) and ends on a
cliffhanger. It can join moments from one episode or from several.

How to use it:
  1. Put your Gemini API key(s) in the .env file next to this script.
  2. Put the show's zip file (or a folder of episode .txt scripts) next to this script.
  3. Run:  python promo_finder.py
     To try a few episodes first:  python promo_finder.py --episodes 1-3
  Results: output/<show name>/promos.csv (opens in Excel or Google Sheets), promos.json,
  raw_model_replies.jsonl (every AI reply, unedited) and run_log.txt.

How it works:
  1. Code reads every script into numbered lines with exact times.
  2. AI writes a short series bible: who is who, what is a spoiler.          (1 call)
  3. AI proposes up to 3 promos per episode as ranges of lines.              (1 call per episode)
  4. Code checks each cut (20-60 s, hook in the first 3 s, no cut inside a line, nothing said
     after the cliffhanger) and fixes it, or rebuilds it from the AI's hook and cliffhanger lines.
     Cuts that still fail go back to the AI once.
  5. AI picks moments for 1-2 trailers across episodes; code cuts them.      (1 call)
  6. AI judges all promos side by side, twice.                               (2 calls)
  7. Code scores and ranks them, groups near-duplicate cuts, and writes the files.

Needs Python 3.9 or newer. Nothing to install.
"""
import argparse
import csv
import hashlib
import json
import re
import sys
import tempfile
import time
import urllib.error
import urllib.request
import zipfile
from collections import Counter
from pathlib import Path

HERE = Path(__file__).resolve().parent  # .env, the zip, cache/ and output/ all sit next to this script


# ==================================================================================================
# 1. Settings
# ==================================================================================================

# The AI models, best first. If one is busy or out of quota, the next one answers.
# Any API that speaks the OpenAI chat format works: to swap models, change these three lines.
API_URL = 'https://generativelanguage.googleapis.com/v1beta/openai'
KEY_NAME = 'GEMINI_API_KEY'  # every key in .env whose name starts with this is used
MODELS = ['gemini-3.8-flash', 'gemini-3.6-flash', 'gemini-3.5-flash-lite', 'gemini-3.1-flash-lite']

PRICES = {'gemini-3.8-flash': (0.75, 3.75), 'gemini-3.6-flash': (0.75, 3.75),  # paid price in USD per 1M tokens
          'gemini-3.5-flash-lite': (0.30, 2.50), 'gemini-3.1-flash-lite': (0.25, 1.50)}  # (input, output)
MAX_TOKENS = 32768     # room for the longest answer (the judge)
TIMEOUT_S = 180        # a model that has not answered in 3 minutes counts as failed...
JUDGE_TIMEOUT_S = 600  # ...but the judge, which reads every promo at once, gets 10 minutes
SKIP_SLOW_MIN = 10     # a model that timed out, or a key out of quota, is skipped for 10 minutes
BUSY_WAIT_S = 60       # only if every model is busy at the same moment: wait this long once, then retry

# Promo rules
MIN_S, MAX_S = 20.0, 60.0  # promo length in seconds
HOOK_S = 3.0               # the hook must start within the first 3 seconds
MAX_HOLD_S = 10.0          # after the cliffhanger line: no new dialogue, and at most 10 s of silent shot
SNAP_S = 1.2               # a cut this close to a camera change is moved onto it
TARGET_S = 45.0            # when code has to shorten a stretch, it aims for this length

# Steps
PROMOS_PER_EPISODE = 3     # the AI proposes up to this many promos per episode
MAX_PER_EPISODE = 2        # the ranked list keeps at most 2 per episode; the rest become backups
TRAILER_MOMENTS = 4        # moments per episode offered for trailers across episodes
CAMERA_MAX_CHARS = 160     # long camera descriptions are shortened in prompts
JUDGE_PASSES = 2           # the judge ranks twice, the second time in reverse order (less position bias)
JUDGE_BATCH = 130          # promos judged side by side in one call
SELF_SCORE_WEIGHT = 0.5    # final scores = 50% judge + 50% the scores the AI gave when it picked the promo
SAME_OVERLAP = 0.5         # two cuts of one episode that overlap this much are the same promo...
SAME_QUESTION = 0.6        # ...and so are two that leave nearly the same open question
SLATE_SIZE = 10            # length of the "post these first" list


# ==================================================================================================
# 2. Reading the scripts (code only)
# Each script entry becomes a "block" with exact times: C = camera, A = action, S = sound,
# D = dialogue. Blocks are numbered b1, b2, ... in file order, so the AI can point at a line by its
# id and code turns the id back into exact times. The AI never types a timestamp.
# ==================================================================================================

TIMING = re.compile(r'Start:\s*\((\d+):(\d+(?:\.\d+)?)\)\s*\|\s*End:\s*\((\d+):(\d+(?:\.\d+)?)\)', re.I)
TAG = re.compile(r'^\*\*(.+?):\*\*\s*(.*)$')
SCENE = re.compile(r'^SCENE\s+(\d+)\s*$', re.I)
FIELD = re.compile(r'^(EPISODE NUMBER|EPISODE LOGLINE|CONTINUITY NOTES|SCENE HEADING|SCENE START|SCENE END)\s*:\s*(.*)$',
                   re.I)
CLOCK = re.compile(r'(\d+):(\d+(?:\.\d+)?)')
EXTENSIONS = {'O.S.', 'V.O.', "CONT'D", 'CONT’D', 'O.C.', 'CONT.'}
SKIP_FOLDERS = {'output', 'cache', '__pycache__', '__MACOSX'}


def secs(m, s):
    return round(int(m) * 60 + float(s), 3)


def clock(x):
    """87.21 -> '01:27.21' (the time format of the scripts)."""
    cs = int(round(x * 100))
    return f"{cs // 6000:02d}:{(cs % 6000) / 100:05.2f}"


def classify(tag):
    """'CAMERA' -> C, 'ACTION' -> A, 'SOUND FX' -> S; anything else is a speaker -> D."""
    t = tag.strip()
    up = t.upper()
    if up == 'CAMERA':
        return 'C', None, None, None
    if up == 'ACTION':
        return 'A', None, None, None
    if up in ('SOUND FX', 'SFX', 'SOUND', 'SOUND EFFECT'):
        return 'S', None, None, None
    speaker = t.split('(')[0].strip()
    parens = re.findall(r'\(([^)]*)\)', t)
    ext = [p.strip() for p in parens if p.strip().upper() in EXTENSIONS]
    emo = [p.strip() for p in parens if p.strip().upper() not in EXTENSIONS]
    return 'D', speaker, ' '.join(ext) or None, ', '.join(emo) or None


def parse_script(path):
    """Read one episode script into blocks with exact start and end times."""
    path = Path(path)
    lines = path.read_text(encoding='utf-8-sig').splitlines()
    has_markers = any('--- SCRIPT START ---' in ln.upper() for ln in lines)
    ep = {'file': path.name, 'episode': None, 'episode_label': None, 'logline': '', 'continuity': '',
          'scenes': [], 'blocks': [], 'warnings': []}
    m = re.search(r'episode[-_ ]?(\d+)', path.name, re.I)
    if m:
        ep['episode'] = int(m.group(1))
    scene, cur, in_script = None, None, not has_markers
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        i += 1
        up = line.upper()
        if '--- SCRIPT START ---' in up:
            in_script, cur = True, None
            continue
        if '--- SCRIPT END ---' in up:
            in_script, cur = False, None
            continue
        kv = FIELD.match(line)
        if kv:
            key, val = kv.group(1).upper(), kv.group(2).strip()
            if key == 'EPISODE NUMBER':
                ep['episode_label'] = val
                if ep['episode'] is None:
                    d = re.search(r'\d+', val)
                    ep['episode'] = int(d.group()) if d else None
            elif key == 'EPISODE LOGLINE':
                ep['logline'] = val
            elif key == 'CONTINUITY NOTES':
                ep['continuity'] = val
            elif scene is not None:
                if key == 'SCENE HEADING':
                    scene['heading'] = val
                else:
                    t = CLOCK.search(val)
                    scene['start' if key == 'SCENE START' else 'end'] = secs(*t.groups()) if t else None
            cur = None
            continue
        sm = SCENE.match(line)
        if sm and in_script:
            scene = {'scene': int(sm.group(1)), 'heading': None, 'start': None, 'end': None}
            ep['scenes'].append(scene)
            cur = None
            continue
        if up.startswith('ACTIONS AND DIALOGUE'):
            continue
        tm = TAG.match(line)
        if tm and in_script:
            kind, speaker, ext, emo = classify(tm.group(1))
            cur = {'id': f"b{len(ep['blocks']) + 1}", 'type': kind, 'tag': tm.group(1).strip(),
                   'speaker': speaker, 'ext': ext, 'emotion': emo,
                   'scene': scene['scene'] if scene else None, 'start': None, 'end': None, 'text': ''}
            ep['blocks'].append(cur)
            rest = tm.group(2).strip()
            t = TIMING.search(rest)
            if t:
                rest = TIMING.sub('', rest).strip()
            else:  # the timing is normally on the next non-empty line
                j = i
                while j < len(lines) and j < i + 3 and not lines[j].strip():
                    j += 1
                if j < len(lines):
                    t = TIMING.search(lines[j])
                    if t:
                        i = j + 1
            if t:
                cur['start'] = secs(t.group(1), t.group(2))
                cur['end'] = secs(t.group(3), t.group(4))
            else:
                ep['warnings'].append(f"{cur['id']}: no timing after tag {tm.group(1)!r}")
            if rest:
                cur['text'] = rest
            continue
        if not line:
            continue
        if cur is not None:
            cur['text'] = (cur['text'] + ' ' + line).strip()
        elif in_script:
            ep['warnings'].append(f"stray line {i}: {line[:80]!r}")
    for b in ep['blocks']:
        if b['start'] is not None and b['end'] is not None and b['end'] < b['start']:
            ep['warnings'].append(f"{b['id']}: end before start")
    ends = [b['end'] for b in ep['blocks'] if b['end'] is not None]
    ends += [s['end'] for s in ep['scenes'] if s.get('end')]
    ep['duration'] = max(ends) if ends else 0.0
    return ep


def scripts_in(folder):
    return [p for p in Path(folder).glob('*.txt') if not p.name.startswith('._')]


def episode_folder(root):
    """The folder at or under root that holds the most .txt scripts (None if there are none)."""
    best, best_n = None, 0
    for d in [Path(root)] + sorted(p for p in Path(root).rglob('*') if p.is_dir()):
        if any(part in SKIP_FOLDERS or part.startswith('.') for part in d.relative_to(root).parts):
            continue
        n = len(scripts_in(d))
        if n > best_n:
            best, best_n = d, n
    return best


def unzip(zip_path, dest):
    """Unpack a zip into dest (nothing may land outside it) and return the folder with the scripts."""
    dest = Path(dest).resolve()
    with zipfile.ZipFile(zip_path) as z:
        for m in z.infolist():
            target = (dest / m.filename).resolve()
            if not target.is_relative_to(dest) or '__MACOSX' in Path(m.filename).parts:
                continue
            if m.is_dir():
                target.mkdir(parents=True, exist_ok=True)
            else:
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(z.read(m))
    return episode_folder(dest)


def find_inputs(folder):
    """The zip files and script folders sitting in folder."""
    folder = Path(folder)
    zips = sorted(p for p in folder.glob('*.zip') if not p.name.startswith('.'))
    dirs = [d for d in sorted(folder.iterdir()) if d.is_dir() and d.name not in SKIP_FOLDERS
            and not d.name.startswith('.') and episode_folder(d) is not None]
    return zips + dirs


def parse_folder(folder):
    """Every .txt script in a folder, sorted by episode number."""
    eps = []
    for n, p in enumerate(sorted(scripts_in(folder)), 1):
        ep = parse_script(p)
        if not ep['blocks']:
            continue
        if ep['episode'] is None:
            ep['episode'] = n
        eps.append(ep)
    eps.sort(key=lambda e: e['episode'])
    return eps


def speaker_counts(ep):
    return Counter(b['speaker'] for b in ep['blocks'] if b['type'] == 'D' and b['speaker'])


def script_text(b):
    if b['type'] == 'D':
        who = b['speaker'] + (f" ({b['ext']})" if b['ext'] else '')
        return f"{who}: {b['text']}"
    kind = {'C': 'CAMERA', 'A': 'ACTION', 'S': 'SOUND FX'}.get(b['type'], b['type'])
    return f"[{kind}] {b['text']}"


def episode_label(nums):
    """[1, 2, 3, 7] -> '1-3, 7'"""
    nums = sorted(nums)
    if not nums:
        return 'none'
    runs, start, prev = [], nums[0], nums[0]
    for n in nums[1:] + [None]:
        if n is not None and n == prev + 1:
            prev = n
            continue
        runs.append(f'{start}' if start == prev else f'{start}-{prev}')
        if n is not None:
            start = prev = n
    return ', '.join(runs)


# ==================================================================================================
# 3. Checking and fixing cuts (code only)
# Code fixes what it safely can, and logs every fix:
#   a cut inside a line of dialogue        -> widen it to the whole line
#   a slow start before the hook           -> trim the start
#   new dialogue after the cliffhanger,
#   or a silent hold of more than 10 s     -> trim the end
#   a cut just off a camera change         -> move it onto the camera change
# Anything else (wrong length, unknown ids ...) goes back to the AI once, or the cut is dropped.
# ==================================================================================================

TOL = 0.25      # a cut may touch a line by this much without counting as "inside" it
CUT_EPS = 0.06  # a cut this close to a camera change counts as "on" it


class Episode:
    """Quick lookups over one parsed episode."""

    def __init__(self, ep):
        self.n = ep['episode']
        self.ep = ep
        self.order = [b for b in ep['blocks'] if b['start'] is not None and b['end'] is not None]
        self.pos = {b['id']: i for i, b in enumerate(self.order)}
        self.blocks = {b['id']: b for b in self.order}
        self.lines = sorted((b['start'], b['end'], b['id']) for b in self.order if b['type'] == 'D')
        cams = [b for b in self.order if b['type'] == 'C']
        self.cuts = sorted({round(b['start'], 3) for b in cams} | {round(b['end'], 3) for b in cams})
        self.headings = {s['scene']: s.get('heading') for s in ep['scenes']}

    def inside_line(self, t):
        for s, e, bid in self.lines:
            if s + TOL < t < e - TOL:
                return s, e, bid
        return None

    def on_cut(self, t):
        return any(abs(c - t) <= CUT_EPS for c in self.cuts)

    def cut_before(self, t, window):
        xs = [c for c in self.cuts if t - window <= c <= t + CUT_EPS]
        return max(xs) if xs else None

    def cut_after(self, t, window):
        xs = [c for c in self.cuts if t - CUT_EPS <= c <= t + window]
        return min(xs) if xs else None

    def edge_blocks(self, s, e):
        """The block ids that best describe where a segment starts and ends (for the output)."""
        starts = [b for b in self.order if abs(b['start'] - s) <= CUT_EPS]
        if starts:
            first = min(starts, key=lambda b: (b['type'] != 'C', self.pos[b['id']]))
        else:
            first = next((b for b in self.order if b['start'] <= s < b['end']), self.order[0])
        ends = [b for b in self.order if abs(b['end'] - e) <= CUT_EPS]
        if ends:
            last = min(ends, key=lambda b: (b['type'] != 'C', -self.pos[b['id']]))
        else:
            last = next((b for b in reversed(self.order) if b['start'] < e <= b['end']), self.order[-1])
        return first['id'], last['id']

    def scene_heading(self, t):
        for b in self.order:
            if b['start'] - CUT_EPS <= t < b['end']:
                return self.headings.get(b['scene'])
        return None

    def next_line_after(self, cb, end):
        """Start of the first dialogue line that begins after block cb and before end, or None."""
        xs = [s for s, e, bid in self.lines if bid != cb['id'] and cb['end'] - TOL <= s < end - TOL]
        return min(xs) if xs else None


def _resolve(eps, segs_spec, problems):
    segs = []
    for i, (n, f, t) in enumerate(segs_spec, 1):
        E = eps.get(n)
        if E is None:
            problems.append(f'segment {i}: episode {n} is not in this run')
            continue
        if isinstance(f, (int, float)) and isinstance(t, (int, float)):  # a cut built by code, in seconds
            if t - f <= 0.05:
                problems.append(f'segment {i}: empty')
                continue
            segs.append({'ep': n, 's': float(f), 'e': float(t)})
            continue
        b1, b2 = E.blocks.get(f), E.blocks.get(t)
        if b1 is None or b2 is None:
            problems.append(f'segment {i}: unknown block id {f if b1 is None else t} in episode {n}')
            continue
        if b2['end'] <= b1['start']:
            problems.append(f'segment {i}: {f}-{t} is empty or reversed')
            continue
        segs.append({'ep': n, 's': b1['start'], 'e': b2['end']})
    return segs


def check_promo(eps, segs_spec, hook_ref, cliff_ref):
    """Check one cut against the rules and fix what code can fix safely.

    segs_spec: [(episode, from_block, to_block)] (or seconds instead of block ids);
    hook_ref and cliff_ref: (episode, block_id).
    Returns the exact segments, hook, cliffhanger and checks, plus an audit trail: model_issues (the
    cut as the AI gave it), code_fixes, and problems (what is still wrong after fixing).
    """
    issues, fixes, problems, notes = [], [], [], []
    segs = _resolve(eps, segs_spec, problems)
    hb = eps[hook_ref[0]].blocks.get(hook_ref[1]) if hook_ref and hook_ref[0] in eps else None
    cb = eps[cliff_ref[0]].blocks.get(cliff_ref[1]) if cliff_ref and cliff_ref[0] in eps else None
    if hb is None:
        problems.append(f'unknown hook block {hook_ref[1] if hook_ref else None}')
    if cb is None:
        problems.append(f'unknown cliffhanger block {cliff_ref[1] if cliff_ref else None}')
    if problems or not segs:
        problems = problems or ['no usable segments']
        return {'ok': False, 'model_issues': list(problems), 'code_fixes': [], 'problems': problems}
    hook_ep, cliff_ep = hook_ref[0], cliff_ref[0]

    # the cut as the AI gave it (shows how often code had to step in)
    first, last = segs[0], segs[-1]
    for i, sg in enumerate(segs, 1):
        E = eps[sg['ep']]
        if E.inside_line(sg['s']):
            issues.append(f'segment {i} starts inside a line')
        if E.inside_line(sg['e']):
            issues.append(f'segment {i} ends inside a line')
    total0 = sum(sg['e'] - sg['s'] for sg in segs)
    if not MIN_S <= total0 <= MAX_S:
        issues.append(f'length {total0:.1f}s')
    if hook_ep != first['ep'] or hb['start'] >= first['e']:
        issues.append('hook is not in the first segment')
    elif hb['start'] - first['s'] > HOOK_S:
        issues.append(f"hook starts {hb['start'] - first['s']:.1f}s in")
    elif hb['start'] - first['s'] < -TOL:
        issues.append('hook starts before the promo')
    if cliff_ep != last['ep'] or not (cb['start'] >= last['s'] - TOL and cb['end'] <= last['e'] + TOL):
        issues.append('cliffhanger is not inside the last segment')
    elif eps[last['ep']].next_line_after(cb, last['e']) is not None:
        issues.append('keeps talking after the cliffhanger')
    elif last['e'] - cb['end'] > MAX_HOLD_S:
        issues.append(f"holds {last['e'] - cb['end']:.1f}s after the cliffhanger")

    # fix 1: whole lines only
    for i, sg in enumerate(segs, 1):
        E = eps[sg['ep']]
        line = E.inside_line(sg['s'])
        if line:
            sg['s'] = line[0]
            fixes.append(f'segment {i}: start moved back to the start of line {line[2]}')
        line = E.inside_line(sg['e'])
        if line:
            sg['e'] = line[1]
            fixes.append(f'segment {i}: end moved to the end of line {line[2]}')
    # fix 2: the hook must start within the first 3 s (keep up to 2.8 s of lead-in)
    if hook_ep == first['ep'] and hb['start'] < first['e']:
        off = hb['start'] - first['s']
        if off > HOOK_S:
            E = eps[first['ep']]
            new_s = hb['start'] - max(0.0, HOOK_S - 0.2)
            line = E.inside_line(new_s)
            if line:
                new_s = min(line[1], hb['start'])
            cuts = [c for c in E.cuts if new_s - CUT_EPS <= c <= hb['start'] + CUT_EPS and not E.inside_line(c)]
            if cuts:
                new_s = min(cuts)  # the earliest camera cut in the window keeps the most lead-in
            fixes.append(f"start trimmed by {new_s - first['s']:.1f}s so the hook lands in the first "
                         f"{HOOK_S:g}s")
            first['s'] = new_s
        elif off < -TOL and off >= -6.0:
            first['s'] = hb['start']
            fixes.append(f'start moved back {-off:.1f}s to include the whole hook')
    # fix 3: end on the cliffhanger (no new line after it, no endless hold)
    if cliff_ep == last['ep']:
        E = eps[last['ep']]
        if cb['end'] > last['e'] + TOL and cb['start'] >= last['s'] - TOL and cb['end'] - last['e'] <= 6.0:
            fixes.append(f"end extended {cb['end'] - last['e']:.1f}s to include the cliffhanger")
            last['e'] = cb['end']
        nxt = E.next_line_after(cb, last['e'])
        tail = last['e'] - cb['end']
        if cb['end'] > last['s'] and (nxt is not None or tail > MAX_HOLD_S):
            new_e = cb['end']
            room = SNAP_S if nxt is None else max(0.0, min(SNAP_S, nxt - cb['end']))
            c = E.cut_after(new_e, room)
            if c is not None and not E.inside_line(c):
                new_e = c
            why = 'a new line started after the cliffhanger' if nxt is not None else 'it held too long after the cliffhanger'
            fixes.append(f"end trimmed by {last['e'] - new_e:.1f}s: {why}")
            last['e'] = new_e
    # fix 4: move cuts onto nearby camera changes
    for i, sg in enumerate(segs):
        E = eps[sg['ep']]
        if not E.on_cut(sg['s']):
            c = E.cut_before(sg['s'], SNAP_S)
            pushes_hook = i == 0 and hook_ep == sg['ep'] and hb['start'] - (c or 0) > HOOK_S
            if c is not None and not E.inside_line(c) and not pushes_hook:
                fixes.append(f"segment {i + 1}: start snapped {sg['s'] - c:.2f}s back to a camera cut")
                sg['s'] = c
        if not E.on_cut(sg['e']):
            c = E.cut_after(sg['e'], SNAP_S)
            past_cliff = (i == len(segs) - 1 and cliff_ep == sg['ep'] and c is not None and
                          (E.next_line_after(cb, c) is not None or c - cb['end'] > MAX_HOLD_S))
            if c is not None and not E.inside_line(c) and not past_cliff:
                fixes.append(f"segment {i + 1}: end snapped {c - sg['e']:.2f}s forward to a camera cut")
                sg['e'] = c
    # fix 5: merge overlapping segments of the same episode
    merged = []
    for sg in segs:
        prev = merged[-1] if merged else None
        if prev and prev['ep'] == sg['ep'] and sg['s'] <= prev['e'] + 0.3 and sg['e'] >= prev['s'] - 0.3:
            prev['s'], prev['e'] = min(prev['s'], sg['s']), max(prev['e'], sg['e'])
            fixes.append('merged overlapping segments')
        else:
            merged.append(sg)
    segs = merged

    # final checks
    first, last = segs[0], segs[-1]
    total = round(sum(sg['e'] - sg['s'] for sg in segs), 2)
    mid = []
    for i, sg in enumerate(segs, 1):
        E = eps[sg['ep']]
        if sg['e'] - sg['s'] < 0.3:
            problems.append(f'segment {i} is empty after trimming')
        for side, t in (('starts', sg['s']), ('ends', sg['e'])):
            line = E.inside_line(t)
            if line:
                mid.append(f'segment {i} {side} inside line {line[2]}')
    hook_off = round(hb['start'] - first['s'], 2) if hook_ep == first['ep'] and hb['start'] < first['e'] else None
    hook_ok = hook_off is not None and -TOL <= hook_off <= HOOK_S
    cliff_in = cliff_ep == last['ep'] and cb['start'] >= last['s'] - TOL and cb['end'] <= last['e'] + TOL
    tail = round(last['e'] - cb['end'], 2) if cliff_ep == last['ep'] else None
    dur_ok = MIN_S <= total <= MAX_S
    if not dur_ok:
        problems.append(f'total length {total:.1f}s (must be {MIN_S:g}-{MAX_S:g}s)')
    if not hook_ok:
        problems.append(f'hook block {hook_ref[1]} starts {hook_off}s into the promo (must be within '
                        f'{HOOK_S:g}s)' if hook_off is not None else
                        f'hook block {hook_ref[1]} is not inside the first segment')
    if not cliff_in:
        problems.append(f'cliffhanger block {cliff_ref[1]} is not inside the last segment')
    problems += mid
    for i, sg in enumerate(segs, 1):
        E = eps[sg['ep']]
        if not (E.on_cut(sg['s']) and E.on_cut(sg['e'])):
            notes.append(f'segment {i} cuts mid-shot (no camera change there)')
        if sg['e'] - sg['s'] < 1.5:
            notes.append(f"segment {i} is very short ({sg['e'] - sg['s']:.1f}s)")
    if cliff_in and eps[last['ep']].next_line_after(cb, last['e']) is not None:
        notes.append('a line of dialogue still follows the cliffhanger')
    elif tail is not None and tail > MAX_HOLD_S:
        notes.append(f'holds {tail:.1f}s after the cliffhanger')
    for a, b in zip(segs, segs[1:]):
        if a['ep'] == b['ep'] and b['s'] < a['e']:
            notes.append('segments are out of story order')

    out_segs = []
    for i, sg in enumerate(segs, 1):
        E = eps[sg['ep']]
        fb, tb = E.edge_blocks(sg['s'], sg['e'])
        out_segs.append({'order': i, 'episode': sg['ep'], 'scene': E.scene_heading(sg['s']),
                         'start': clock(sg['s']), 'end': clock(sg['e']),
                         'start_s': round(sg['s'], 2), 'end_s': round(sg['e'], 2),
                         'duration_s': round(sg['e'] - sg['s'], 2), 'from_block': fb, 'to_block': tb})
    cliff_at = sum(x['duration_s'] for x in out_segs[:-1]) + (cb['start'] - last['s'])
    return {
        'ok': not problems, 'segments': out_segs, 'total_duration_s': total,
        'hook': {'episode': hook_ep, 'time': clock(hb['start']), 'seconds_into_promo': hook_off,
                 'block': hb['id'], 'script_text': script_text(hb)},
        'cliffhanger': {'episode': cliff_ep, 'time': clock(cb['start']), 'seconds_into_promo': round(cliff_at, 2),
                        'seconds_after_cliff_line_to_cut': tail, 'block': cb['id'], 'script_text': script_text(cb)},
        'checks': {'duration_20_60s': dur_ok, 'hook_within_3s': hook_ok, 'cliffhanger_in_last_segment': cliff_in,
                   'no_mid_line_cuts': not mid, 'notes': notes},
        'model_issues': issues, 'code_fixes': fixes, 'problems': problems,
    }


def transcript(eps, spans, max_chars=700):
    """The dialogue and action inside spans [{'episode', 'start_s', 'end_s'}], for prompts."""
    parts = []
    for k, sp in enumerate(spans):
        E = eps[sp['episode']]
        if k:
            parts.append('[CUT]')
        for b in E.order:
            if b['type'] in ('D', 'A') and b['start'] < sp['end_s'] - TOL and b['end'] > sp['start_s'] + TOL:
                parts.append(f"{b['speaker']}: {b['text']}" if b['type'] == 'D' else f"({b['text']})")
    txt = ' '.join(parts)
    if len(txt) > max_chars:
        txt = txt[:int(max_chars * 0.6)].rstrip() + ' ... ' + txt[-int(max_chars * 0.35):].lstrip()
    return txt


# ==================================================================================================
# 4. Building cuts from moments (code only)
# When the AI's own cut cannot be fixed, code builds one from the AI's hook line, cliffhanger line
# and the few lines a stranger needs. It cuts only on camera changes or line edges, never inside a
# line. Trailers across episodes are built the same way from the moments the AI picks.
# ==================================================================================================

TRANSITION = re.compile(r'\b(cut to black|fades? (?:to|out)|smash cut|black screen|title card|end credits|'
                        r'to be continued|end of episode)\b', re.I)


def is_transition(b):
    """A camera block that is only a transition (e.g. 'Cut to black.'), not a story moment."""
    text = b.get('text') or ''
    return b['type'] == 'C' and bool(TRANSITION.search(text)) and len(text) < 80


def _points(E):
    """Clean cut points: camera changes (preferred) and block edges, never inside a line."""
    cam = sorted({round(t, 3) for b in E.order if b['type'] == 'C' for t in (b['start'], b['end'])})
    edge = sorted({round(t, 3) for b in E.order if b['type'] != 'C' for t in (b['start'], b['end'])})
    return [p for p in cam if not E.inside_line(p)], [p for p in edge if not E.inside_line(p)]


def _start(E, H, window, cam, edge):
    lo = H['start'] - max(0.0, window - 0.2)  # keep up to (window - 0.2 s) of lead-in
    for pts in (cam, edge):
        xs = [p for p in pts if lo - CUT_EPS <= p <= H['start'] + CUT_EPS]
        if xs:
            return min(xs)
    line = E.inside_line(H['start'])
    return line[0] if line and H['start'] - line[0] <= window else H['start']


def _end(E, C, max_hold, cam, edge, longest=False):
    hi = C['end'] + max_hold
    nxt = E.next_line_after(C, hi + 1.0)  # no new line of dialogue after the cliffhanger
    if nxt is not None:
        hi = min(hi, nxt)
    for pts in (cam, edge):
        xs = [p for p in pts if C['end'] - CUT_EPS <= p <= hi + CUT_EPS]
        if xs:
            return max(xs) if longest else min(xs)
    line = E.inside_line(C['end'])
    return line[1] if line else C['end']


def _earliest(cam, edge, lo, hi):
    for pts in (cam, edge):
        xs = [p for p in pts if lo - CUT_EPS <= p <= hi + CUT_EPS]
        if xs:
            return min(xs)
    return hi


def _latest(cam, edge, lo, hi):
    for pts in (cam, edge):
        xs = [p for p in pts if lo - CUT_EPS <= p <= hi + CUT_EPS]
        if xs:
            return max(xs)
    return lo


def trailer(eps, picks, max_ctx=12.0):
    """Build a trailer across episodes from moments [(episode, block_id)], in the given order.

    The first clip opens on its moment (the hook), middle clips carry a few seconds of context, and
    the last clip builds up to its moment (the cliffhanger). The context then grows or shrinks until
    the total is 20-60 s. Returns (segments [(episode, start, end)], problems).
    """
    info = []
    for n, bid in picks:
        E = eps.get(n)
        b = E.blocks.get(bid) if E else None
        if b is None:
            return None, [f'unknown moment {bid} in episode {n}']
        info.append((n, E, b, *_points(E)))
    last_i = len(info) - 1

    def build(ctx_mid, ctx_last, tail_first):
        segs = []
        for i, (n, E, b, cam, edge) in enumerate(info):
            if i == 0:
                s = _start(E, b, HOOK_S, cam, edge)
            else:
                s = _earliest(cam, edge, b['start'] - (ctx_last if i == last_i else ctx_mid), b['start'])
            if i == last_i:
                e = _end(E, b, MAX_HOLD_S, cam, edge)
            else:
                e = _latest(cam, edge, b['end'], b['end'] + (tail_first if i == 0 else 2.0))
            line = E.inside_line(e)
            e = line[1] if line else e
            segs.append((n, s, max(e, b['end'])))
        return segs

    ctx_mid, ctx_last, tail_first = 3.0, 5.0, 4.0
    segs = build(ctx_mid, ctx_last, tail_first)
    total = sum(e - s for _, s, e in segs)
    while total < MIN_S and (ctx_mid < max_ctx or ctx_last < max_ctx or tail_first < max_ctx):
        ctx_mid, ctx_last, tail_first = min(max_ctx, ctx_mid + 2), min(max_ctx, ctx_last + 3), min(max_ctx, tail_first + 2)
        segs = build(ctx_mid, ctx_last, tail_first)
        total = sum(e - s for _, s, e in segs)
    while total > MAX_S and (ctx_mid > 0 or ctx_last > 0 or tail_first > 0):
        ctx_mid, ctx_last, tail_first = max(0.0, ctx_mid - 2), max(0.0, ctx_last - 2), max(0.0, tail_first - 2)
        segs = build(ctx_mid, ctx_last, tail_first)
        total = sum(e - s for _, s, e in segs)
    problems = [] if MIN_S <= total <= MAX_S else [
        f'trailer runs {total:.1f}s (must be {MIN_S:g}-{MAX_S:g}s): use {"more" if total < MIN_S else "fewer"} moments']
    return segs, problems


def beat(E, block_id, max_len=12.0):
    """A short clip around one moment, for trailers: (start, end) or None."""
    b = E.blocks.get(block_id)
    if b is None:
        return None
    cam, edge = _points(E)
    s = max([p for p in cam + edge if b['start'] - 2.5 <= p <= b['start'] + CUT_EPS], default=b['start'])
    e = min([p for p in cam + edge if b['end'] - CUT_EPS <= p <= b['end'] + 2.5], default=b['end'])
    line = E.inside_line(s)
    s = line[0] if line else s
    line = E.inside_line(e)
    e = line[1] if line else e
    if e - s > max_len:
        s, e = b['start'], min(b['end'], b['start'] + max_len)
    return (round(s, 3), round(e, 3)) if e - s >= 1.5 else None


def assemble(E, hook_id, cliff_id, keep_ids):
    """Build a cut from moments. Returns (segments [(start, end)], hook, cliff, notes, problems).

    notes: what code decided (shown in the output); problems: why no valid cut exists.
    """
    notes, problems = [], []
    H, C = E.blocks.get(hook_id), E.blocks.get(cliff_id)
    if H is None or C is None:
        return None, H, C, notes, [f'unknown block id {hook_id if H is None else cliff_id}']
    if is_transition(C):
        i = E.pos[C['id']]
        alt = next((b for b in reversed(E.order[:i]) if not is_transition(b)), None)
        if alt is not None:
            notes.append(f"cliffhanger was a transition shot ({C['id']}); used {alt['id']} instead")
            C = alt
    if C['end'] <= H['start'] + 1.0:
        return None, H, C, notes, [f"the cliffhanger ({C['id']}) does not come after the hook ({H['id']})"]
    cam, edge = _points(E)
    s = _start(E, H, HOOK_S, cam, edge)
    e = _end(E, C, MAX_HOLD_S, cam, edge)
    if e - s < MIN_S:
        e = _end(E, C, MAX_HOLD_S, cam, edge, longest=True)
    if e - s <= MAX_S:
        if e - s < MIN_S:
            problems.append(f'hook to cliffhanger is only {e - s:.1f}s (needs {MIN_S:g}s): choose an earlier '
                            'hook or a later cliffhanger')
        return [(s, e)], H, C, notes, problems

    # too long: split the stretch into units at camera changes (long units also at line edges)
    pts = sorted({s, e} | {p for p in cam if s < p < e})
    units = []
    for a, b in zip(pts, pts[1:]):
        inner = [p for p in edge if a + 2.0 <= p <= b - 2.0] if b - a > 8.0 else []
        cur = a
        for p in inner:
            if p - cur >= 2.0:
                units.append((cur, p))
                cur = p
        units.append((cur, b))
    units = [u for u in units if u[1] - u[0] > 0.05]

    def hits(u, x0, x1):
        return u[0] < x1 - TOL and u[1] > x0 + TOL

    def length(idx):
        return sum(units[i][1] - units[i][0] for i in idx)

    keep = {i for i, u in enumerate(units) if hits(u, s, H['end']) or hits(u, C['start'], e)}
    if not keep:
        keep = {0, len(units) - 1}
    ignored = []
    for kid in keep_ids:
        K = E.blocks.get(kid)
        if K is None or K['start'] < s - TOL or K['end'] > e + TOL:
            ignored.append(kid)
            continue
        keep |= {i for i, u in enumerate(units) if hits(u, K['start'], K['end'])}
    if ignored:
        notes.append(f"essential lines outside the hook-to-cliffhanger stretch ignored: {', '.join(map(str, ignored))}")
    if length(keep) > MAX_S:
        return None, H, C, notes, [f'the hook, cliffhanger and essential lines alone run {length(keep):.0f}s '
                                   f'(max {MAX_S:g}s): mark fewer essential lines or pick a closer hook']
    last_hook = max(i for i in keep if units[i][0] < H['end'] + TOL)
    first_cliff = min(i for i in keep if hits(units[i], C['start'], e))
    before_cliff = [i for i in range(first_cliff - 1, last_hook, -1)]
    after_hook = [i for i in range(last_hook + 1, first_cliff)]
    order, a, b = [], 0, 0
    while a < len(before_cliff) or b < len(after_hook):  # two units of build-up per unit after the hook
        for _ in range(2):
            if a < len(before_cliff):
                order.append(before_cliff[a])
                a += 1
        if b < len(after_hook):
            order.append(after_hook[b])
            b += 1
    for i in order:
        if length(keep) >= TARGET_S:
            break
        if i not in keep and length(keep | {i}) <= MAX_S:
            keep.add(i)
    runs = []
    for i in sorted(keep):
        if runs and i == runs[-1][-1] + 1:
            runs[-1].append(i)
        else:
            runs.append([i])
    segs = [(units[r[0]][0], units[r[-1]][1]) for r in runs]
    while len(segs) > 1:  # bridge tiny gaps, and keep at most 4 segments
        g, j = min((segs[k + 1][0] - segs[k][1], k) for k in range(len(segs) - 1))
        total = sum(y - x for x, y in segs)
        if (g < 2.0 or len(segs) > 4) and total + g <= MAX_S:
            segs[j] = (segs[j][0], segs[j + 1][1])
            del segs[j + 1]
        else:
            break
    notes.append(f'stretch from hook to cliffhanger was {e - s:.0f}s; kept {len(segs)} part(s), '
                 f'{sum(y - x for x, y in segs):.0f}s')
    if sum(y - x for x, y in segs) < MIN_S:
        problems.append('could not keep enough material between hook and cliffhanger')
    return segs, H, C, notes, problems


# ==================================================================================================
# 5. Asking the AI
# Each call tries the best model first. If it is busy, refuses or gives a broken answer, the next
# model answers that call at once, and the next call starts again from the best model. A model
# that hangs is skipped for SKIP_SLOW_MIN. A key that runs out of quota is skipped too, and the
# next key is used at once. Only if every model is busy at the same moment does the tool wait
# (BUSY_WAIT_S, once). Every answer is saved in cache/, so running the same show again costs nothing.
# ==================================================================================================

class ModelError(Exception):
    """A model could not give a usable answer."""


class Busy(ModelError):
    """The model is overloaded right now (HTTP 5xx)."""


class TooSlow(ModelError):
    """No answer within the time limit."""


def load_keys(path):
    """Every key in .env whose name starts with KEY_NAME, e.g. GEMINI_API_KEY_1=..., GEMINI_API_KEY_2=...
    ('NAME: key' lines and a bare key on its own line also work)."""
    keys = []
    if not path.exists():
        return keys
    for line in path.read_text(encoding='utf-8-sig').splitlines():
        line = line.strip()
        if not line or line.startswith('#'):
            continue
        m = re.match(r'^(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)[^=:]*[=:]\s*(.*)$', line)
        if m:
            if not m.group(1).upper().startswith(KEY_NAME):
                continue
            value = m.group(2).strip().strip('"').strip("'")
        elif ' ' not in line and len(line) >= 30:
            value = line
        else:
            continue
        if value and value not in keys:
            keys.append(value)
    return keys


def extract_json(text):
    """The JSON object in a reply (code fences and <think> blocks are ignored)."""
    t = re.sub(r'<think>.*?</think>', '', text or '', flags=re.S).strip()
    t = re.sub(r'^```(?:json)?\s*', '', t)
    t = re.sub(r'\s*```\s*$', '', t)
    try:
        return json.loads(t)
    except json.JSONDecodeError:
        pass
    a, b = t.find('{'), t.rfind('}')
    if a != -1 and b > a:
        return json.loads(t[a:b + 1])
    raise ValueError('the reply contains no JSON object')


class Models:
    def __init__(self, keys, cache_dir, raw_file):
        self.keys = keys
        self.cache_dir = cache_dir
        self.raw_file = raw_file
        self.rest = {}  # a model, a key number, or (model, key number) -> time until which it is left alone
        self.log = []   # one record per answer, for the cost line

    def resting(self, thing):
        return time.time() < self.rest.get(thing, 0)

    def keys_for(self, model):
        return [i for i in range(len(self.keys)) if not self.resting(i) and not self.resting((model, i))]

    def cache_file(self, model, messages):
        # (this shape matches the saved answers of earlier runs, so they are reused)
        ident = {'url': API_URL, 'model': model, 'temperature': None, 'extra': {}}
        blob = json.dumps({'p': ident, 'm': messages}, sort_keys=True, ensure_ascii=False)
        h = hashlib.sha256(blob.encode('utf-8')).hexdigest()
        return self.cache_dir / h[:2] / f'{h}.json'

    def ask(self, name, messages, check, timeout=TIMEOUT_S):
        """The answer of the best model that gives a usable one: {'data', 'text', 'model'}."""
        for model in MODELS:  # asked before? reuse the saved answer (delete cache/ to ask again)
            saved = self.cache_file(model, messages)
            if saved.exists():
                try:
                    rec = json.loads(saved.read_text(encoding='utf-8'))
                    data = extract_json(rec['text'])
                    check(data)
                    return self._done(name, model, rec, data, True)
                except (OSError, KeyError, ValueError):
                    pass
        errors = []
        for attempt in (1, 2):
            busy = False
            for model in MODELS:
                if not self.keys_for(model):
                    errors.append(f'{model}: no usable API key left')
                    continue
                if self.resting(model):
                    errors.append(f'{model}: skipped, it timed out recently')
                    continue
                try:
                    rec = self._call(model, messages, timeout)
                    data = extract_json(rec['text'])
                    check(data)
                except (ModelError, ValueError) as e:
                    if isinstance(e, TooSlow):
                        self.rest[model] = time.time() + SKIP_SLOW_MIN * 60
                    busy = busy or isinstance(e, Busy)
                    print(f'      {model} failed ({e}) -> trying the next model')
                    errors.append(f'{model}: {e}')
                    continue
                saved = self.cache_file(model, messages)
                saved.parent.mkdir(parents=True, exist_ok=True)
                saved.write_text(json.dumps(rec, ensure_ascii=False, indent=1), encoding='utf-8')
                return self._done(name, model, rec, data, False)
            if not busy or attempt == 2:
                break
            print(f'      every model is busy right now: waiting {BUSY_WAIT_S}s, then trying once more')
            time.sleep(BUSY_WAIT_S)
        raise ModelError(f'{name}: no model could answer ({"; ".join(errors)})')

    def _call(self, model, messages, timeout):
        """One request to one model, trying the keys in turn."""
        body = json.dumps({'model': model, 'messages': messages, 'max_tokens': MAX_TOKENS,
                           'response_format': {'type': 'json_object'}}).encode('utf-8')
        for i in self.keys_for(model):
            req = urllib.request.Request(API_URL + '/chat/completions', data=body, method='POST', headers={
                'Content-Type': 'application/json', 'Authorization': f'Bearer {self.keys[i]}',
                'User-Agent': 'promo-finder'})
            t0 = time.time()
            try:
                with urllib.request.urlopen(req, timeout=timeout) as resp:
                    reply = json.loads(resp.read().decode('utf-8'))
            except urllib.error.HTTPError as e:
                text = e.read().decode('utf-8', 'replace')
                m = re.search(r'"message"\s*:\s*"([^"]{0,160})', text)
                why = f'HTTP {e.code}' + (f': {m.group(1)}' if m else '')
                if e.code == 429:  # this key has used up its quota for this model
                    self.rest[(model, i)] = time.time() + SKIP_SLOW_MIN * 60
                    print(f'      {model}: key {i + 1}/{len(self.keys)} is out of quota -> trying the next key')
                    continue
                if e.code in (401, 403) or (e.code == 400 and 'api key' in text.lower()):  # e.g. "API key not valid"
                    self.rest[i] = float('inf')  # a rejected key is not used again
                    print(f'      key {i + 1}/{len(self.keys)} was rejected ({why}) -> trying the next key')
                    continue
                if e.code >= 500:  # e.g. 503 "high demand"
                    raise Busy(why) from None
                raise ModelError(why) from None
            except Exception as e:  # no connection, no answer in time, or a broken response
                if 'timed out' in str(e) or isinstance(e, TimeoutError):
                    raise TooSlow(f'no answer within {timeout}s') from None
                raise ModelError(f'{type(e).__name__}: {str(e)[:160]}') from None
            try:
                choice = reply['choices'][0]
                text = (choice.get('message') or {}).get('content') or ''
            except (KeyError, IndexError, TypeError, AttributeError):
                raise ModelError(f'unexpected response: {str(reply)[:160]}') from None
            if not text:  # e.g. Gemini's content filter refused the request
                raise ModelError(f"no answer (finish reason: {choice.get('finish_reason')})")
            usage = reply.get('usage') or {}
            return {'text': text, 'finish_reason': choice.get('finish_reason'),
                    'prompt_tokens': usage.get('prompt_tokens') or 0,
                    'completion_tokens': usage.get('completion_tokens') or 0,
                    'latency_s': round(time.time() - t0, 2), 'model': model,
                    'saved_at': time.strftime('%Y-%m-%d %H:%M:%S')}
        raise ModelError('no usable API key left')

    def _done(self, name, model, rec, data, cached):
        r = {'call': name, 'model': model, 'from_cache': cached, 'prompt_tokens': rec.get('prompt_tokens', 0),
             'completion_tokens': rec.get('completion_tokens', 0), 'latency_s': rec.get('latency_s', 0.0)}
        self.log.append(r)
        with open(self.raw_file, 'a', encoding='utf-8') as f:  # every reply, unedited
            f.write(json.dumps(dict(r, reply=rec['text']), ensure_ascii=False) + '\n')
        return {'data': data, 'text': rec['text'], 'model': model}

    def usage(self):
        cost = 0.0
        for r in self.log:  # what the same run would cost from scratch at paid prices
            p_in, p_out = PRICES.get(r['model'], (0.0, 0.0))
            cost += r['prompt_tokens'] / 1e6 * p_in + r['completion_tokens'] / 1e6 * p_out
        return {'calls': len(self.log), 'live_calls': sum(1 for r in self.log if not r['from_cache']),
                'cached_calls': sum(1 for r in self.log if r['from_cache']),
                'prompt_tokens': sum(r['prompt_tokens'] for r in self.log),
                'completion_tokens': sum(r['completion_tokens'] for r in self.log),
                'model_seconds': round(sum(r['latency_s'] for r in self.log), 1),
                'paid_cost_usd': round(cost, 4), 'calls_by_model': dict(Counter(r['model'] for r in self.log))}


# ==================================================================================================
# 6. Prompts
# The AI is asked for judgment only (which moments, why, how good) and answers with block ids.
# Nothing here is about one show: show knowledge comes from the series bible the AI writes itself.
# ==================================================================================================

RUBRIC = """- hook: 5 = a shocking line, act or image in the first 3 s (betrayal, slap, whip, fire, public humiliation, a forbidden kiss, an accusation); 3 = intriguing; 1 = flat.
- cliffhanger: 5 = cuts on a high-stakes open question; 3 = some tension, partly resolved; 1 = resolved.
- standalone: 5 = a cold viewer gets who, what and why; 3 = unclear who or why; 1 = needs earlier episodes. A promo without either lead sells a subplot: at most 3.
- spoiler_safety: 5 = setup only; 3 = reveals a mid-season twist; 1 = gives away the climax or the ending. The premise stated in the show's title is not a spoiler."""

BIBLE_SYSTEM = """You are a story editor for vertical microdrama series. From the loglines, continuity notes and speaker labels of every episode, write a compact series bible. Promo editors will use it to decide what a cold viewer understands and what counts as a spoiler. Output only a JSON object."""

BIBLE_SHAPE = """{
  "premise": "<2-3 sentences: who, what conflict, what promise>",
  "title_promise": "<what the show's title already tells viewers; this is not a spoiler>",
  "leads": [{"name": "...", "role": "...", "speaker_labels": ["..."]}],
  "characters": [{"name": "...", "role": "...", "speaker_labels": ["..."]}],
  "label_warnings": ["<speaker labels that change meaning, or that cover several people, with episode numbers>"],
  "arc": [{"episodes": "1-4", "summary": "..."}],
  "twists": [{"episode": 19, "twist": "...", "level": "premise | mid | endgame"}]
}"""


def bible_messages(title, episodes, counts):
    rows = []
    for ep in episodes:
        sp = ', '.join(f'{k}({v})' for k, v in counts[ep['episode']].most_common(12))
        rows.append(f"E{ep['episode']:02d}: {ep.get('logline') or '-'} | notes: {ep.get('continuity') or '-'} | speakers: {sp}")
    user = (f"SERIES TITLE: {title}\n\nEpisodes ({len(episodes)}):\n" + '\n'.join(rows) +
            "\n\nReturn JSON in exactly this shape:\n" + BIBLE_SHAPE +
            "\nMap descriptive labels (for example 'A STRANGER' or 'LADY IN RED') to named characters when the loglines "
            "make it clear. Keep the whole bible under 600 words.")
    return [{'role': 'system', 'content': BIBLE_SYSTEM}, {'role': 'user', 'content': user}]


def _people(xs):
    return '; '.join(f"{x.get('name')} ({x.get('role', '')}; labels: {', '.join(map(str, x.get('speaker_labels') or []))})"
                     for x in xs)


def bible_header(b):
    """The parts of the bible every prompt needs: premise, title promise, leads, twists."""
    if not isinstance(b, dict) or not b:
        return '(no series bible available)'
    out = []
    if b.get('premise'):
        out.append('PREMISE: ' + str(b['premise']))
    if b.get('title_promise'):
        out.append('TITLE PROMISE (not a spoiler): ' + str(b['title_promise']))
    leads = [x for x in (b.get('leads') or []) if isinstance(x, dict)]
    if leads:
        out.append('LEADS: ' + _people(leads))
    tw = [t for t in (b.get('twists') or []) if isinstance(t, dict)]
    if tw:
        out.append('TWISTS (spoiler level): ' + ' | '.join(f"ep {t.get('episode')}: {t.get('twist')} [{t.get('level')}]"
                                                          for t in tw))
    return '\n'.join(out)


def _in_range(spec, n):
    nums = [int(x) for x in re.findall(r'\d+', str(spec))]
    return (nums[0] <= n <= nums[1]) if len(nums) >= 2 else (bool(nums) and nums[0] == n)


def bible_notes(b, n, labels):
    """Only what matters for episode n: where it sits in the story, who speaks in it, label warnings."""
    if not isinstance(b, dict) or not b:
        return ''
    up = {str(lab).upper() for lab in labels}
    out = []
    arc = [a for a in (b.get('arc') or []) if isinstance(a, dict) and _in_range(a.get('episodes'), n)]
    if arc:
        out.append('Where this sits in the story: ' + ' | '.join(str(a.get('summary')) for a in arc))
    people = [x for x in (b.get('leads') or []) + (b.get('characters') or [])
              if isinstance(x, dict) and up & {str(lab).upper() for lab in (x.get('speaker_labels') or [])}]
    if people:
        out.append('Who speaks here: ' + _people(people))
    warns = [str(w) for w in (b.get('label_warnings') or []) if any(lab in str(w).upper() for lab in up if len(lab) > 2)]
    if warns:
        out.append('Label warnings: ' + ' | '.join(warns))
    return '\n'.join(out)


def _minutes(x):
    return f"{int(x) // 60}:{int(x) % 60:02d}"


def script_listing(ep):
    """The script as the AI reads it: one line per block with a rough start time (no end times,
    so the AI never does arithmetic)."""
    out, last = [], object()
    scenes = {s['scene']: s for s in ep['scenes']}
    kind = {'C': 'CAM', 'A': 'ACT', 'S': 'SFX'}
    for b in ep['blocks']:
        if b['start'] is None or b['end'] is None:
            continue
        if b['scene'] != last:
            s = scenes.get(b['scene'])
            if s:
                out.append(f"-- SCENE {s['scene']}: {s.get('heading') or ''}")
            last = b['scene']
        t = f"[{_minutes(b['start'])}]"
        if b['type'] == 'D':
            who = b['speaker'] + (f" ({b['ext']})" if b['ext'] else '')
            emo = f" [{b['emotion']}]" if b['emotion'] else ''
            out.append(f"{t} {b['id']} {who}{emo}: {b['text']}")
        else:
            txt = b['text']
            if b['type'] == 'C' and CAMERA_MAX_CHARS and len(txt) > CAMERA_MAX_CHARS:
                txt = txt[:CAMERA_MAX_CHARS].rstrip() + '...'
            out.append(f"{t} {b['id']} {kind.get(b['type'], b['type'])} {txt}")
    out.append(f"== END OF EPISODE ({_minutes(ep['duration'])}) ==")
    return '\n'.join(out)


EXAMPLE = """=== EPISODE 7 of 40 ===
Logline: At her stepsister's gala, a maid is humiliated, until a guest reveals who she married.
-- SCENE 1: INT. BALLROOM - NIGHT
[0:00] b1 CAM Wide shot of a glittering ballroom; guests turn toward the stairs.
[0:01] b2 HOST [Cheerful]: "Welcome to the Lin family gala!"
[0:04] b3 CAM Close-up of MIA in a plain apron, carrying a tray.
[0:04] b4 VIVIAN [Mocking]: "Look, the maid thinks she's a guest."
[0:08] b5 ACT Vivian throws a glass of red wine in Mia's face.
[0:08] b6 SFX Glass shatters on the marble floor.
[0:09] b7 CAM Medium shot: the crowd laughs as Mia wipes her face.
[0:10] b8 MIA [Shaking]: "You'll regret this."
[0:13] b9 VIVIAN [Laughing]: "Who would ever marry you?"
[0:16] b10 CAM Wide shot: the doors burst open; ETHAN, the CEO, walks in with bodyguards.
[0:18] b11 ETHAN [Ice cold]: "Who touched my wife?"
[0:22] b12 CAM Close-up of Vivian's face draining of color.
[0:25] b13 CAM Cut to black.
== END OF EPISODE (0:26) ==
A good answer:
{"episodes": [{"episode": 7,
  "moments": [{"block": "b5", "kind": "hook", "what": "wine thrown in the maid's face"},
              {"block": "b4", "kind": "hook", "what": "'the maid thinks she is a guest'"},
              {"block": "b11", "kind": "cliffhanger", "what": "the CEO: 'Who touched my wife?'"},
              {"block": "b12", "kind": "turn", "what": "the bully goes pale"}],
  "promos": [{"title": "Who touched my wife?", "segments": [{"from": "b3", "to": "b12"}], "hook_block": "b4", "cliffhanger_block": "b11",
              "essential_blocks": ["b9"],
              "why": "Public humiliation in the first second, then a reversal that leaves 'what happens to her now?' open.",
              "scores": {"hook": 5, "cliffhanger": 5, "standalone": 5, "spoiler_safety": 4}}],
  "no_promo_reason": ""}]}
One continuous range: it opens on the camera shot of the hook line b4 and ends on the silent reaction b12, right after the cliffhanger line b11, not on "Cut to black" (b13)."""

CANDIDATE_SYSTEM = """You are a senior promo editor at a vertical microdrama app (shows like ReelShort or DramaBox). You choose the moments in a screenplay that become short ads on TikTok, Reels and Shorts. An ad works when someone who has never seen the show stops scrolling, understands the conflict within seconds, and has to know what happens next.

YOU PROPOSE THE CUT; SOFTWARE CHECKS IT
Give each promo as 1-3 ranges of block ids, in story order (a range runs from the start of its first block to the end of its last block). Pick by story, not arithmetic: the software computes exact times, enforces the rules (20-60 seconds, hook in the first 3 seconds, no cut inside a line, nothing said after the cliffhanger) and fixes small timing errors. If your ranges cannot be fixed, it rebuilds the promo from your hook, cliffhanger and essential lines.

STEP 1 - FIND THE MOMENTS. For each episode, list its 3-8 strongest moments: lines or actions that could open an ad (hook) or end one (cliffhanger), plus big turns.
- Hooks that stop a scroll: a betrayal revealed, public humiliation, a slap, shove or other violence, a cruel or shocking line, an accusation, an ultimatum, an identity reveal, a forbidden kiss, a villain gloating, a last-second rescue.
- Cliffhangers that make people click: a threat or ultimatum not yet answered, a reveal cut off, someone arriving or a door opening, danger about to land, an accusation hanging in the air, a choice that must be made now.
- Weak moments: exposition, travel, small talk, calm romance with nothing at stake, a problem already solved, the end of a story arc, a scene transition such as "Cut to black".
- An episode's ending is usually written as a cliffhanger, and its opening often picks up the previous one at full intensity.

STEP 2 - BUILD PROMOS. The best promo is usually ONE continuous stretch of 20-60 seconds in which the hook and the cliffhanger sit close together, often in the same scene. Start the first range on the hook (or at most 2 seconds before it) and end the last range on the cliffhanger; a silent reaction shot after it is fine, a new line of dialogue is not. Use a second or third range only to skip a slow middle: every jump makes a promo harder to follow. Name the hook_block, the cliffhanger_block and up to 3 essential_blocks a stranger needs (who is who, what is at stake). Then check as a viewer who knows nothing:
- Does the first line or action grab me with no context?
- By the end, do I understand who is against whom?
- Is a question left open, with nothing after it that answers it?
Prefer moments with the leads. Each promo must use a different hook or cliffhanger. Give up to MAXP promos per episode, best first. If nothing would stop a scroll (a quiet, talky or transitional episode), return no promos and say why.

SCORES (1-5, be strict; 5 is rare)
""" + RUBRIC + """

OUTPUT: only a JSON object in exactly this shape, one entry per episode you were given:
{"episodes": [{"episode": <number>,
  "moments": [{"block": "<id>", "kind": "hook | cliffhanger | turn", "what": "<a few words>"}],
  "promos": [{"title": "<short working title>", "segments": [{"from": "<id>", "to": "<id>"}],
              "hook_block": "<id>", "cliffhanger_block": "<id>", "essential_blocks": ["<id>"],
              "why": "<one line on why it works>",
              "scores": {"hook": 1-5, "cliffhanger": 1-5, "standalone": 1-5, "spoiler_safety": 1-5}}],
  "no_promo_reason": "<only when promos is empty>"}]}

EXAMPLE (a made-up show, only to show the format; never reuse its names or lines):
""" + EXAMPLE


def candidate_messages(title, bible, ep, total, counts):
    n = ep['episode']
    notes = bible_notes(bible, n, counts[n].keys())
    user = (f"SERIES: {title}\n\nSERIES BIBLE\n{bible_header(bible)}\n\n"
            f"Below: episode(s) {n} of {total}. [m:ss] is roughly when a block starts. CAM = camera shot, "
            "ACT = action, SFX = sound; other lines are dialogue (SPEAKER [emotion]: line).\n"
            f"\n=== EPISODE {n} of {total} ===\nLogline: {ep.get('logline') or '-'}\n"
            f"Picks up from: {ep.get('continuity') or '-'}\n" + (notes + '\n' if notes else '') + script_listing(ep))
    return [{'role': 'system', 'content': CANDIDATE_SYSTEM.replace('MAXP', str(PROMOS_PER_EPISODE))},
            {'role': 'user', 'content': user}]


def repair_message(problems):
    return ("The software could not build these promos:\n" + '\n'.join(problems) +
            "\n\nReturn corrected versions in the same JSON shape (the \"episodes\" list), with only the promos "
            "you fixed: choose a different hook or cliffhanger, or fewer essential lines. Drop a promo if it "
            "cannot be fixed. moments may be empty.")


CROSS_SYSTEM = """You cut multi-episode trailers for a vertical microdrama series. A trailer joins 2-5 moments from different episodes; the software turns each moment into a short clip and checks the total length (20-60 s, so 3-5 moments is typical). The first moment must work as a hook and the last must be a cliffhanger.
Build up to two trailers:
1. PREMISE TRAILER (always try): sell the TITLE PROMISE to someone who knows nothing. Open on the setup's most shocking moment from the early episodes, add the turn that sets the premise in motion, and end on the moment that reveals or teases the title promise, cut before anything is explained. The title promise is not a spoiler.
2. STORY TRAILER (optional): a threat, ultimatum or danger in one episode and its consequence in a later one, ending unresolved.
Never use moments that resolve the main story (endgame twists). Each moment must be understandable next to the others: prefer moments with the leads.
Output only JSON: {"trailers": [{"kind": "premise | story", "title": "...", "moments": ["m03", "m11", "m20"], "why": "<one line>"}]}
Use only moment ids from the list."""


def cross_messages(title, bible, moments_text):
    user = (f"SERIES: {title}\n\nSERIES BIBLE\n{bible_header(bible)}\n\n"
            f"MOMENTS (id | episode | kind | clip length | what is said or shown):\n{moments_text}")
    return [{'role': 'system', 'content': CROSS_SYSTEM}, {'role': 'user', 'content': user}]


JUDGE_SYSTEM = """You are the head of performance marketing at a microdrama app. Your editors proposed promo cuts for one series. Work in three steps:
1. For each candidate, write the open question its ending leaves a viewer with: a specific question with stakes ("Will she escape in time?", "Who sent the letter?", "What will he do to them now?"). If the ending closes the scene's conflict, write "none", however good the line is. Closing endings: someone wins the argument with a last word, a calm goodbye, a hug or kiss that settles things, a quiet decision to endure. Not closing: a threat, a promise of revenge, or a vow of what someone will do next leaves "will they, and what happens then?" open, and a cruel reveal leaves "how will she react?" open.
2. Order all candidates from best to worst as ads that win new viewers.
3. Score each one 1-5 on four criteria, consistent with that order (a better-ranked candidate never gets a lower total). Use the full scale; give 5 only to the best few. A candidate whose open question is "none" scores 1-2 on cliffhanger.
""" + RUBRIC + """
Use the series bible and the episode number for spoilers: resolutions from late episodes are spoilers.
If a candidate sells the same moment as a better candidate, set same_as to the better one's id.
Output only JSON: {"ranking": ["c03", "c01", "..."], "scores": [{"id": "c01", "open_question": "<the question, or none>", "hook": 4, "cliffhanger": 5, "standalone": 4, "spoiler_safety": 5, "same_as": null, "note": "<one line>"}]}
Rank and score every id exactly once."""

NONE_LIKE = {'', 'none', 'n/a', 'na', 'no', 'nothing', '-', 'no question', 'no open question', 'null'}


def judge_messages(title, bible, cands_text):
    user = f"SERIES: {title}\n\nSERIES BIBLE\n{bible_header(bible)}\n\nCANDIDATES\n{cands_text}"
    return [{'role': 'system', 'content': JUDGE_SYSTEM}, {'role': 'user', 'content': user}]


# ==================================================================================================
# 7. The steps, in order
# ==================================================================================================

SCORE_KEYS = ('hook', 'cliffhanger', 'standalone', 'spoiler_safety')


def _scores(sc, as_float=False):
    sc = sc if isinstance(sc, dict) else {}
    out = {}
    for k in SCORE_KEYS:
        v = sc.get(k, sc.get('cliff') if k == 'cliffhanger' else None)
        try:
            x = max(1.0, min(5.0, float(v)))
            out[k] = round(x, 2) if as_float else int(round(x))
        except (TypeError, ValueError):
            out[k] = 3
    return out


def _bid(x):
    """'b12', 'B12', '12', 12 -> 'b12'"""
    s = str(x if x is not None else '').strip().lower()
    s = s[1:] if s.startswith('b') else s
    return f'b{int(s)}' if s.isdigit() else str(x)


def _object(d):
    if not isinstance(d, dict):
        raise ValueError('expected a JSON object')


def _check_candidates(d):
    _object(d)
    if 'episodes' not in d and isinstance(d.get('promos'), list):
        return  # the one-episode shape; handled in _episode_entries
    if not isinstance(d.get('episodes'), list):
        raise ValueError('missing the "episodes" list')
    for e in d['episodes']:
        if not isinstance(e, dict) or not isinstance(e.get('promos', []), list):
            raise ValueError('every episode needs a "promos" list')


def _check_trailers(d):
    _object(d)
    if not isinstance(d.get('trailers', []), list):
        raise ValueError('"trailers" must be a list')


def _check_judge(d):
    _object(d)
    if not isinstance(d.get('scores'), list):
        raise ValueError('missing the "scores" list')


def _episode_entries(data, n):
    """The parts of a reply that are about episode n."""
    if 'episodes' not in data and isinstance(data.get('promos'), list):
        data = {'episodes': [dict(data, episode=data.get('episode', n))]}
    out = []
    for e in data.get('episodes') or []:
        if not isinstance(e, dict):
            continue
        try:
            k = int(str(e.get('episode')).strip().lstrip('Ee'))
        except (TypeError, ValueError):
            k = n
        if k == n:
            out.append(e)
    return out


def _kind(segments):
    eps = {s['episode'] for s in segments}
    return 'cross-episode' if len(eps) > 1 else ('multi-segment' if len(segments) > 1 else 'single-segment')


def make_candidate(eps, n, p, model, words, repaired=False):
    """Check the AI's own cut. If code cannot fix it, rebuild the cut from the AI's hook,
    cliffhanger and essential lines."""
    E = eps[n]
    segs = [(n, _bid(s.get('from')), _bid(s.get('to'))) for s in (p.get('segments') or []) if isinstance(s, dict)]
    hook, cliff = _bid(p.get('hook_block')), _bid(p.get('cliffhanger_block'))
    keep = [_bid(x) for x in (p.get('essential_blocks') or []) if x is not None][:5]

    def describe(bid, blk):
        return words.get(bid) or (script_text(blk)[:140] if blk else '')

    c = {'title': str(p.get('title') or 'untitled').strip(), 'episodes': [n],
         'picked_segments': [list(s) for s in segs], 'picked_hook': hook, 'picked_cliffhanger': cliff,
         'essential': keep, 'why': str(p.get('why') or ''), 'self_scores': _scores(p.get('scores')),
         'model': model, 'repaired': repaired}
    first = None
    if segs:
        chk = check_promo(eps, segs, (n, hook), (n, cliff))
        if chk['ok']:
            return dict(c, hook_note=describe(hook, E.blocks.get(hook)),
                        cliff_note=describe(cliff, E.blocks.get(cliff)), **chk)
        first = chk
    # rebuild from the moments
    built, H, C, notes, problems = assemble(E, hook, cliff, keep)
    issues = list(problems) + [x for x in notes if 'transition' in x]
    if built is not None and not problems:
        chk = check_promo(eps, [(n, a, b) for a, b in built], (n, H['id']), (n, C['id']))
        chk['code_fixes'] = notes + chk['code_fixes']
        chk['model_issues'] = issues
    else:
        chk = {'ok': False, 'model_issues': issues, 'code_fixes': notes, 'problems': problems or ['no cut could be built']}
    if first is not None:
        chk['model_issues'] = list(first.get('model_issues') or []) + list(chk.get('model_issues') or [])
        if chk['ok']:
            chk['code_fixes'] = [f"the model's own cut could not be fixed ({'; '.join(first['problems'])}); rebuilt "
                                 "from its hook and cliffhanger"] + chk['code_fixes']
    return dict(c, hook_note=describe(hook, H), cliff_note=describe(C['id'] if C else cliff, C), **chk)


def _lead_in_promo(eps, c, leads):
    """True if a lead (from the series bible) speaks, or is named, inside the promo."""
    labels = {str(lab).upper() for x in leads for lab in (x.get('speaker_labels') or [])}
    names = {w.upper() for x in leads for w in str(x.get('name') or '').replace(',', ' ').split() if len(w) > 2}
    for sg in c['segments']:
        for b in eps[sg['episode']].order:
            if b['start'] < sg['end_s'] - TOL and b['end'] > sg['start_s'] + TOL:
                if b['type'] == 'D' and (b['speaker'] or '').upper() in labels:
                    return True
                if any(nm in (b['text'] or '').upper() for nm in names):
                    return True
    return False


def _judge_line(eps, c, total):
    last = c['segments'][-1]
    gap = eps[last['episode']].ep['duration'] - last['end_s']
    facts = [f"{len(c['segments'])} part(s)"]
    if c['type'] != 'cross-episode' and gap <= 8.0:
        facts.append("ends on the episode's final beat")
    return (f"{c['_jid']} | episode(s) {', '.join(map(str, c['episodes']))} of {total} | {c['total_duration_s']}s | "
            f"{c['type']}, {', '.join(facts)}\n  hook ({c['hook']['seconds_into_promo']}s in): {c['hook']['script_text']}\n"
            f"  cliffhanger: {c['cliffhanger']['script_text']}\n"
            f"  transcript: {transcript(eps, c['segments'], 700)}\n  editor's note: {c['why']}")


def _public(c):
    """The final promo record."""
    hook = dict(c['hook'], description=c['hook_note'])
    cliff = dict(c['cliffhanger'], description=c['cliff_note'])
    return {'promo_id': c['promo_id'], 'rank': None, 'tier': None, 'title': c['title'], 'type': c['type'],
            'episodes': c['episodes'], 'total_duration_s': c['total_duration_s'], 'segments': c['segments'],
            'hook': hook, 'cliffhanger': cliff, 'why_it_works': c['why'], 'scores': c['scores'],
            'total_score': None, 'self_scores': c['self_scores'], 'judge_note': c.get('judge_note', ''),
            'judge_position': c.get('judge_position'), 'open_question': c.get('open_question'),
            'same_as': c.get('same_as'), 'checks': c['checks'],
            'model_issues': c['model_issues'], 'code_fixes': c['code_fixes'], 'repaired': c['repaired'],
            'model': c['model'],
            'picked': {'segments': c['picked_segments'], 'hook_block': c['picked_hook'],
                       'cliffhanger_block': c['picked_cliffhanger'], 'essential_blocks': c.get('essential', [])}}


def _rejected(c, why=None):
    return {'title': c['title'], 'episodes': c['episodes'], 'picked_segments': c['picked_segments'],
            'hook_block': c['picked_hook'], 'cliffhanger_block': c['picked_cliffhanger'],
            'essential_blocks': c.get('essential', []),
            'problems': c['problems'], 'model_issues': c['model_issues'], 'repaired': c['repaired'],
            'status': why or 'dropped', 'model': c['model']}


def find_promos(folder, out, title, keys, select=None):
    """Run every step on one show. Returns how many episodes were processed."""
    t0 = time.time()
    eps_all = parse_folder(folder)
    if not eps_all:
        raise SystemExit(f'No screenplay .txt files with timed lines found in {folder}')
    total = max(e['episode'] for e in eps_all)
    sel = [e for e in eps_all if select is None or e['episode'] in select]
    if not sel:
        raise SystemExit('None of the requested episodes exist in the scripts.')
    ai = Models(keys, HERE / 'cache', out / 'raw_model_replies.jsonl')
    ai.raw_file.write_text('', encoding='utf-8')
    eps = {e['episode']: Episode(e) for e in eps_all}
    warnings = [f"E{e['episode']:02d}: {w}" for e in eps_all for w in e['warnings']]
    failures, done = [], []
    print(f"Show: {title} | {len(eps_all)} episodes found, running on {len(sel)}: {episode_label([e['episode'] for e in sel])}")
    print(f"Models, best first: {', '.join(MODELS)} | {len(keys)} API key(s)")

    # 1. series bible: one call that reads every episode's logline, so spoiler and standalone
    #    judgments know the whole story
    print('[1/5] series bible')
    counts = {e['episode']: speaker_counts(e) for e in eps_all}
    bible = {}
    try:
        bible = ai.ask('series_bible', bible_messages(title, eps_all, counts), _object)['data']
    except ModelError as e:
        failures.append(str(e))
        print('   failed, continuing without a bible:', e)

    # 2. promos per episode: checked and fixed by code; cuts that still fail go back to the AI once
    good, rejected, moments = [], [], []
    what = {}  # (episode, block) -> the AI's few words about that moment

    def collect(data, model, n, repaired):
        ok, bad = [], []
        for e in _episode_entries(data, n):
            for m in e.get('moments') or []:
                if isinstance(m, dict) and m.get('block') is not None:
                    bid = _bid(m.get('block'))
                    what[(n, bid)] = str(m.get('what') or '')
                    if not repaired:
                        moments.append({'episode': n, 'block': bid, 'kind': str(m.get('kind') or '').lower(),
                                        'what': str(m.get('what') or '')})
            for p in e.get('promos') or []:
                if isinstance(p, dict):
                    words = {b: w for (nn, b), w in what.items() if nn == n}
                    c = make_candidate(eps, n, p, model, words, repaired)
                    (ok if c['ok'] else bad).append(c)
        return ok, bad

    for i, ep in enumerate(sel, 1):
        n = ep['episode']
        tag = f'E{n:02d}'
        print(f'[2/5] candidates {tag} ({i}/{len(sel)})')
        msgs = candidate_messages(title, bible, ep, total, counts)
        try:
            res = ai.ask('candidates_' + tag, msgs, _check_candidates)
        except ModelError as e:
            failures.append(str(e))
            print('   failed:', e)
            continue
        done.append(n)
        ok, to_fix = collect(res['data'], res['model'], n, False)
        good += ok
        if not to_fix:
            continue
        problems = [f"- episode {c['episodes'][0]}, \"{c['title']}\" (hook {c['picked_hook']}, cliffhanger "
                    f"{c['picked_cliffhanger']}, essential {', '.join(c.get('essential') or []) or 'none'}): "
                    f"{'; '.join(c['problems'])}" for c in to_fix]
        msgs2 = msgs + [{'role': 'assistant', 'content': res['text']},
                        {'role': 'user', 'content': repair_message(problems)}]
        print(f'      {len(to_fix)} pick(s) failed the checks, asking for a repair')
        try:
            res2 = ai.ask('repair_' + tag, msgs2, _check_candidates)
            ok2, bad2 = collect(res2['data'], res2['model'], n, True)
            good += ok2
            rejected += [_rejected(c, 'repair also failed the checks') for c in bad2]
        except ModelError as e:
            failures.append(str(e))
        rejected += [_rejected(c, 'failed the checks; sent back for one repair') for c in to_fix]

    # 3. trailers across episodes: the AI picks moments from its own lists; code cuts the clips
    trailers = []
    clips, seen_m, per = [], set(), {}
    prio = {'cliffhanger': 0, 'hook': 1}
    for m in sorted(moments, key=lambda m: (m['episode'], prio.get(m['kind'], 2))):
        if (m['episode'], m['block']) in seen_m or per.get(m['episode'], 0) >= TRAILER_MOMENTS:
            continue
        span = beat(eps[m['episode']], m['block'])
        if span:
            seen_m.add((m['episode'], m['block']))
            per[m['episode']] = per.get(m['episode'], 0) + 1
            clips.append(dict(m, start_s=span[0], end_s=span[1]))
    clips.sort(key=lambda m: (m['episode'], m['start_s']))
    if len({c['episode'] for c in clips}) >= 2:
        print(f'[3/5] multi-episode trailers from {len(clips)} moments')
        for i, cl in enumerate(clips, 1):
            cl['id'] = f'm{i:02d}'
        lines = [f"{cl['id']} | E{cl['episode']:02d} | {cl['kind']} | {cl['end_s'] - cl['start_s']:.1f}s | "
                 f"{transcript(eps, [cl], 220)}" for cl in clips]
        try:
            res = ai.ask('cross_episode', cross_messages(title, bible, '\n'.join(lines)), _check_trailers)
            by_id = {cl['id']: cl for cl in clips}
            for t in res['data'].get('trailers') or []:
                if not isinstance(t, dict):
                    continue
                ids = [x for x in (t.get('moments') or t.get('beats') or []) if x in by_id]
                if len(ids) < 2 or len({by_id[x]['episode'] for x in ids}) < 2:
                    continue
                first, last = by_id[ids[0]], by_id[ids[-1]]
                segs, tprob = trailer(eps, [(by_id[x]['episode'], by_id[x]['block']) for x in ids])
                c = {'title': str(t.get('title') or 'trailer'), 'episodes': [], 'picked_segments': [],
                     'picked_hook': first['block'], 'picked_cliffhanger': last['block'], 'essential': [],
                     'hook_note': first['what'], 'cliff_note': last['what'], 'why': str(t.get('why') or ''),
                     'self_scores': _scores(None), 'model': res['model'], 'repaired': False}
                if segs is None:
                    rejected.append(_rejected(dict(c, problems=tprob, model_issues=tprob)))
                    continue
                chk = check_promo(eps, segs, (first['episode'], first['block']), (last['episode'], last['block']))
                if tprob:
                    chk['problems'] = tprob + chk['problems']
                    chk['ok'] = False
                c = dict(c, episodes=sorted({s[0] for s in segs}),
                         picked_segments=[[s[0], round(s[1], 2), round(s[2], 2)] for s in segs], **chk)
                if c['ok']:
                    trailers.append(c)
                else:
                    rejected.append(_rejected(c))
        except ModelError as e:
            failures.append(str(e))
    else:
        print('[3/5] multi-episode trailers: skipped (moments from fewer than 2 episodes)')

    # promo ids: E05-1, E05-2 ... in the AI's own order; trailers T1, T2
    per_ep, seen, final = {}, set(), []
    for c in good + trailers:
        sig = tuple((s['episode'], s['start_s'], s['end_s']) for s in c['segments'])
        if sig in seen:
            continue  # the same cut twice (e.g. the repair repeated a good pick)
        seen.add(sig)
        c['type'] = _kind(c['segments'])
        if c['type'] == 'cross-episode':
            per_ep['T'] = per_ep.get('T', 0) + 1
            c['promo_id'] = f"T{per_ep['T']}"
        else:
            n = c['episodes'][0]
            per_ep[n] = per_ep.get(n, 0) + 1
            c['promo_id'] = f'E{n:02d}-{per_ep[n]}'
        final.append(c)

    # 4. judge: all promos side by side, twice (the second time in reverse order), then averaged
    if final:
        print(f'[4/5] judging {len(final)} promos')
        for i, c in enumerate(final, 1):
            c['_jid'] = f'c{i:02d}'
        jid = {c['_jid']: c['promo_id'] for c in final}
        votes, pos_votes = {}, {}  # judge id -> the scores / ranking positions from each pass
        for b0 in range(0, len(final), JUDGE_BATCH):
            chunk = final[b0:b0 + JUDGE_BATCH]
            for ps in range(JUDGE_PASSES):
                order = chunk if ps % 2 == 0 else list(reversed(chunk))
                text = '\n\n'.join(_judge_line(eps, c, total) for c in order)
                label = f'batch{b0 // JUDGE_BATCH + 1}' + (f'-pass{ps + 1}' if JUDGE_PASSES > 1 else '')
                try:
                    res = ai.ask('judge_' + label, judge_messages(title, bible, text), _check_judge, JUDGE_TIMEOUT_S)
                    for s in res['data'].get('scores') or []:
                        if isinstance(s, dict) and s.get('id') is not None:
                            votes.setdefault(str(s['id']).strip(), []).append(s)
                    for pos, x in enumerate(res['data'].get('ranking') or []):
                        pos_votes.setdefault(str(x).strip(), []).append(b0 + pos)
                except ModelError as e:
                    failures.append(str(e))
        scores, positions = {}, {}
        for jid_, vs in votes.items():  # average the passes
            fl = [_scores(v, as_float=True) for v in vs]
            scores[jid_] = {**vs[0], **{k: sum(f[k] for f in fl) / len(fl) for k in SCORE_KEYS}}
            qs = [str(v.get('open_question') if v.get('open_question') is not None else '-').strip() for v in vs]
            none = sum(1 for q in qs if q.lower().strip(' .!') in NONE_LIKE)
            scores[jid_]['open_question'] = next((q for q in qs if q.lower().strip(' .!') not in NONE_LIKE), 'none')
            scores[jid_]['_no_question'] = none * 2 >= len(qs)  # most passes found nothing left open
        for jid_, ps_ in pos_votes.items():
            positions[jid_] = sum(ps_) / len(ps_)
        leads = [x for x in (bible.get('leads') or []) if isinstance(x, dict)] if isinstance(bible, dict) else []
        for c in final:
            s = scores.get(c['_jid'])
            if s:
                js = _scores(s, as_float=True)
                w = 0.0 if c['type'] == 'cross-episode' else SELF_SCORE_WEIGHT  # trailers have no self-scores
                c['scores'] = {k: round((1 - w) * js[k] + w * c['self_scores'][k], 2) for k in SCORE_KEYS}
                c['judge_note'] = str(s.get('note') or '')
                c['same_as'] = jid.get(str(s.get('same_as') or '').strip())
                c['open_question'] = s.get('open_question')
                if s.get('_no_question') and c['scores']['cliffhanger'] > 2:  # nothing left open: not a cliffhanger
                    c['scores']['cliffhanger'] = 2
                    c['judge_note'] = (c['judge_note'] + ' ').lstrip() + '[ending leaves no open question: cliffhanger capped at 2]'
            else:
                c['scores'] = dict(c['self_scores'])
                c['judge_note'] = '(no judge score: self-scores used)'
                c['same_as'] = None
            c['judge_position'] = positions.get(c['_jid'])
            if leads and c['scores']['standalone'] > 3 and not _lead_in_promo(eps, c, leads):
                c['scores']['standalone'] = 3
                c['judge_note'] = (c['judge_note'] + ' ').lstrip() + '[no lead appears: standalone capped at 3]'

    # 5. rank, group other cuts of the same moment, cap 2 per episode, pick the slate, write the files
    promos = rank([_public(c) for c in final])
    promos, alternates = group_variants(promos)
    promos, backups = cap_per_episode(promos, MAX_PER_EPISODE)
    for p in backups:
        p['backup'] = True
    alternates = backups + alternates
    slate_list = slate(promos, SLATE_SIZE)
    usage = ai.usage()
    elapsed = round(time.time() - t0, 1)
    meta = {'generated_at': time.strftime('%Y-%m-%d %H:%M'),
            'models': sorted({r['model'] for r in ai.log}) or ['(no model calls)'],
            'episodes_processed': sorted(done), 'episodes_label': episode_label(done),
            'episodes_failed': sorted(set(e['episode'] for e in sel) - set(done)),
            'settings': {'models': MODELS, 'min_s': MIN_S, 'max_s': MAX_S, 'hook_window_s': HOOK_S,
                         'max_hold_s': MAX_HOLD_S, 'snap_window_s': SNAP_S, 'target_s': TARGET_S,
                         'promos_per_episode': PROMOS_PER_EPISODE, 'max_per_episode': MAX_PER_EPISODE,
                         'judge_passes': JUDGE_PASSES, 'self_score_weight': SELF_SCORE_WEIGHT},
            'usage': usage, 'series_bible': bible,
            'run': {'elapsed_s': elapsed, 'failures': failures, 'parse_warnings': warnings, 'calls': ai.log}}
    write_outputs(out, title, meta, promos, rejected, slate_list, alternates)

    n = {t: sum(1 for p in promos if p['tier'] == t) for t in 'ABC'}
    print(f"[5/5] {len(promos)} promos ({n['A']} A, {n['B']} B, {n['C']} C) + {len(alternates)} alternative cuts, "
          f"{len(rejected)} rejected, {len(failures)} failed calls")
    for p in promos[:10]:
        print(f"   {p['rank']:>2}. {p['tier']} {p['total_score']:>5} {p['promo_id']:<7} {p['total_duration_s']:>5}s  "
              f"{segline(p)}  {p['title']}")
    print(f"   calls: {usage['calls']} ({usage['cached_calls']} reused from cache) | tokens in/out: "
          f"{usage['prompt_tokens']}/{usage['completion_tokens']} | cost at paid prices: ${usage['paid_cost_usd']:.2f} | "
          f"{elapsed}s")
    if usage['calls_by_model']:
        print('   answers by model: ' + ', '.join(f'{m} ({k})' for m, k in usage['calls_by_model'].items()))
    for f in failures:
        print('   FAILED:', f)
    return len(done)


# ==================================================================================================
# 8. Ranking (code only)
# ==================================================================================================

WEIGHTS = {'hook': 0.35, 'cliffhanger': 0.30, 'standalone': 0.25, 'spoiler_safety': 0.10}
FORMULA = 'total_score = 20 * (0.35*hook + 0.30*cliffhanger + 0.25*standalone + 0.10*spoiler_safety); sub-scores 1-5, total 20-100'
TIERS = 'A = post first (score >= 90), B = strong (85-89.9), C = usable / backup (< 85)'
STOP_WORDS = set('the a an and or of to in on at for with will would be is are was were her his him she he they them it '
                 'its this that what who how does do did from by as into after before their there than then can could '
                 'whether'.split())


def _f(x):
    try:
        return min(5.0, max(1.0, float(x)))
    except (TypeError, ValueError):
        return 3.0


def total_score(scores):
    return round(20 * sum(w * _f(scores.get(k)) for k, w in WEIGHTS.items()), 1)


def tier(score):
    return 'A' if score >= 90 else ('B' if score >= 85 else 'C')


def rank(promos):
    for p in promos:
        p['total_score'] = total_score(p['scores'])
    promos.sort(key=lambda p: (-p['total_score'], -_f(p['scores'].get('cliffhanger')), -_f(p['scores'].get('hook')),
                               p.get('judge_position') if p.get('judge_position') is not None else 10 ** 6,
                               min(p['episodes']), p['promo_id']))
    for i, p in enumerate(promos, 1):
        p['rank'] = i
        p['tier'] = tier(p['total_score'])
    return promos


def _spans(segs):
    """Segments -> {episode: [(start, end), ...]}, with overlapping spans merged."""
    d = {}
    for s in segs:
        d.setdefault(s['episode'], []).append((s['start_s'], s['end_s']))
    for ep, v in d.items():
        v.sort()
        merged = []
        for a, b in v:
            if merged and a <= merged[-1][1]:
                merged[-1] = (merged[-1][0], max(merged[-1][1], b))
            else:
                merged.append((a, b))
        d[ep] = merged
    return d


def overlap(segs1, segs2):
    """Shared time / combined time of two cuts (0 = nothing shared, 1 = the same cut)."""
    x, y = _spans(segs1), _spans(segs2)
    shared = sum(max(0.0, min(b, d) - max(a, c)) for ep in set(x) & set(y) for a, b in x[ep] for c, d in y[ep])
    union = sum(b - a for v in x.values() for a, b in v) + sum(b - a for v in y.values() for a, b in v) - shared
    return shared / union if union > 0 else 0.0


def question_similarity(a, b):
    """Word overlap of two open questions (0 when either ending leaves nothing open)."""
    if str(a or '').strip(' .!').lower() in NONE_LIKE or str(b or '').strip(' .!').lower() in NONE_LIKE:
        return 0.0
    wa = {w for w in re.findall(r"[a-z']+", str(a).lower()) if w not in STOP_WORDS and len(w) > 2}
    wb = {w for w in re.findall(r"[a-z']+", str(b).lower()) if w not in STOP_WORDS and len(w) > 2}
    return len(wa & wb) / len(wa | wb) if wa and wb else 0.0


def group_variants(promos):
    """Keep the best-ranked cut of each moment; other cuts of it become alternates.

    Two promos of the same episode are the same moment when their cuts overlap by half or more,
    when they share both the hook and the cliffhanger, or when they leave nearly the same open
    question. Input must be ranked. Returns (main list re-ranked 1..n, alternates).
    """
    main, alternates = [], []

    def same(m, p):
        return ((m['cliffhanger']['block'] == p['cliffhanger']['block'] and m['hook']['block'] == p['hook']['block'])
                or overlap(m['segments'], p['segments']) >= SAME_OVERLAP
                or question_similarity(m.get('open_question'), p.get('open_question')) >= SAME_QUESTION)

    for p in promos:
        host = None
        if p['type'] != 'cross-episode':
            host = next((m for m in main if m['type'] != 'cross-episode' and m['episodes'] == p['episodes'] and
                         same(m, p)), None)
        if host is None:
            main.append(p)
        else:
            p['variant_of'] = host['promo_id']
            host.setdefault('alternates', []).append(p['promo_id'])
            alternates.append(p)
    for i, p in enumerate(main, 1):
        p['rank'] = i
    for p in alternates:
        p['rank'] = None
    return main, alternates


def cap_per_episode(promos, cap):
    """Keep at most cap promos per episode (trailers are not capped); the rest become backups."""
    main, backups, seen = [], [], {}
    for p in promos:
        if p['type'] != 'cross-episode':
            n = p['episodes'][0]
            seen[n] = seen.get(n, 0) + 1
            if seen[n] > cap:
                p['rank'] = None
                backups.append(p)
                continue
        main.append(p)
    for i, p in enumerate(main, 1):
        p['rank'] = i
    return main, backups


def slate(promos, size=10):
    """"Post these first": the best promo per episode (a trailer is its own pick), skipping any the
    judge called the same moment as one already picked. Order: tier, trailers first, then the
    earliest episode, because a viewer who clicks starts at episode 1."""
    chosen, used_eps, ids = [], set(), set()
    for p in promos:  # already ranked
        if len(chosen) >= size:
            break
        if p.get('same_as') in ids:
            continue
        if p['type'] != 'cross-episode' and p['episodes'][0] in used_eps:
            continue
        chosen.append(p)
        ids.add(p['promo_id'])
        if p['type'] != 'cross-episode':
            used_eps.add(p['episodes'][0])
    chosen.sort(key=lambda p: (p['tier'], p['type'] != 'cross-episode', min(p['episodes'])))
    return chosen


# ==================================================================================================
# 9. Output files: promos.csv (one row per promo) and promos.json (everything)
# ==================================================================================================

CSV_COLUMNS = ['rank', 'slate_position', 'list', 'tier', 'total_score', 'promo_id', 'title', 'type', 'episodes',
               'segments', 'total_duration_s', 'hook_time', 'hook_line', 'hook', 'cliffhanger_time',
               'cliffhanger_line', 'cliffhanger', 'open_question', 'why_it_works', 'hook_score',
               'cliffhanger_score', 'standalone_score', 'spoiler_safety', 'judge_note', 'variant_of', 'code_fixes',
               'check_notes', 'model', 'reviewer_verdict', 'reviewer_notes']


def segline(p):
    return ' + '.join(f"E{s['episode']:02d} {s['start']}–{s['end']}" for s in p['segments'])


def _row(p, slate_pos):
    kind = 'ranked' if p.get('rank') else ('alternative cut' if p.get('variant_of') else 'backup')
    sc = p['scores']
    return [p.get('rank') or '', slate_pos.get(p['promo_id'], ''), kind, p['tier'], p['total_score'], p['promo_id'],
            p['title'], p['type'], ' '.join(map(str, p['episodes'])), segline(p).replace('–', '-'),
            p['total_duration_s'], f"E{p['hook']['episode']:02d} {p['hook']['time']}", p['hook']['script_text'],
            p['hook']['description'], f"E{p['cliffhanger']['episode']:02d} {p['cliffhanger']['time']}",
            p['cliffhanger']['script_text'], p['cliffhanger']['description'], p.get('open_question') or '',
            p['why_it_works'], sc['hook'], sc['cliffhanger'], sc['standalone'], sc['spoiler_safety'],
            p.get('judge_note', ''), p.get('variant_of') or '', '; '.join(p['code_fixes']),
            '; '.join(p['checks']['notes']), p['model'], '', '']


def write_outputs(out, show, meta, promos, rejected, slate_list, alternates):
    slate_pos = {p['promo_id']: i for i, p in enumerate(slate_list, 1)}
    doc = {'show': show, **meta,
           'timestamps': 'MM:SS.cs from the start of each episode, as in the scripts; *_s fields are seconds.',
           'blocks': 'bN = the Nth tagged block in the episode file. A segment runs from the start of from_block '
                     'to the end of to_block.',
           'score_formula': FORMULA, 'tiers': TIERS,
           'lists': 'promos = the ranked list (at most max_per_episode_final per episode); alternates = other cuts '
                    'of a ranked promo (variant_of names it) and backups beyond the per-episode limit',
           'promo_count': len(promos),
           'slate': [{'position': i, 'promo_id': p['promo_id'], 'rank': p['rank'], 'title': p['title']}
                     for i, p in enumerate(slate_list, 1)],
           'promos': promos, 'alternates': alternates, 'rejected_candidates': rejected}
    (out / 'promos.json').write_text(json.dumps(doc, indent=2, ensure_ascii=False), encoding='utf-8')
    with open(out / 'promos.csv', 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(CSV_COLUMNS)
        for p in promos + alternates:
            w.writerow(_row(p, slate_pos))


# ==================================================================================================
# 10. Command line
# ==================================================================================================

class Tee:
    """Shows everything printed on screen and also writes it to run_log.txt."""

    def __init__(self, path, screen):
        self.file = open(path, 'w', encoding='utf-8')
        self.screen = screen

    def write(self, s):
        self.screen.write(s)
        self.file.write(s)

    def flush(self):
        self.screen.flush()
        self.file.flush()


def episode_numbers(text):
    """'1-5,12' -> {1, 2, 3, 4, 5, 12}"""
    out = set()
    for part in text.replace(' ', '').split(','):
        if '-' in part:
            a, b = part.split('-', 1)
            out.update(range(int(a), int(b) + 1))
        elif part:
            out.add(int(part))
    return out


def main():
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')  # so Windows consoles show every character
    except (AttributeError, ValueError):
        pass
    ap = argparse.ArgumentParser(description='Finds ready-to-post promos in the scripts of a microdrama. Put the '
                                             "show's zip file and your .env next to this script, then run it.")
    ap.add_argument('inputs', nargs='*', help='zip file(s) or script folder(s) (default: every zip or script folder '
                                              'next to this script)')
    ap.add_argument('--episodes', help='only these episodes, e.g. 1-3 or 1,5,9')
    args = ap.parse_args()

    keys = load_keys(HERE / '.env')
    if not keys and not (HERE / 'cache').exists():
        sys.exit('No API key found. Open the .env file next to promo_finder.py and paste your Gemini key after '
                 'GEMINI_API_KEY_1= (get one free at https://aistudio.google.com/apikey).')
    inputs = [Path(x) for x in args.inputs] or find_inputs(HERE)
    if not inputs:
        sys.exit(f"No scripts found. Put the show's zip file (or a folder of episode .txt scripts) in {HERE} "
                 'and run again.')
    select = episode_numbers(args.episodes) if args.episodes else None
    for src in inputs:
        if not src.exists():
            sys.exit(f'Not found: {src}')
        with tempfile.TemporaryDirectory() as tmp:
            folder = episode_folder(src) if src.is_dir() else unzip(src, tmp)
            if folder is None:
                sys.exit(f'No episode .txt scripts found in {src.name}')
            # the show's name: the folder inside the zip, or the zip's own name when that folder has a plain name
            plain = folder.resolve() == Path(tmp).resolve() or (
                not src.is_dir() and folder.name.lower() in ('scripts', 'episodes'))
            title = src.stem if plain else folder.name
            out = HERE / 'output' / (re.sub(r'[^A-Za-z0-9]+', '_', title).strip('_')[:80] or 'show')
            out.mkdir(parents=True, exist_ok=True)
            screen = sys.stdout
            sys.stdout = Tee(out / 'run_log.txt', screen)
            try:
                print(f'Input: {src.name}')
                if not keys:
                    print('No API key in .env: only answers saved in cache/ can be used.')
                processed = find_promos(folder, out, title, keys, select)
                if not processed:
                    print('No episode could be processed: every model call failed (see above). Check your internet '
                          'connection and your API key in .env, then run again.')
            finally:
                sys.stdout.file.close()
                sys.stdout = screen
        if not processed:
            sys.exit(1)
        print(f'Done. Results are in output/{out.name}/ next to promo_finder.py')


if __name__ == '__main__':
    main()
