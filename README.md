# Promo Finder

Promo Finder reads the screenplays of a vertical microdrama (short, episodic, phone-first drama) and finds ready-to-post promos for TikTok, Reels and Shorts. Each promo:
- is 20–60 seconds long
- grabs a stranger in the first 3 seconds (the **hook**)
- ends on an unresolved, high-stakes moment (the **cliffhanger**)
- makes sense to someone who has never seen the show

A promo can be one continuous clip, 2–4 clips from one episode, or a trailer that joins moments from several episodes. For every promo you get the exact start and end time of each clip, the hook line, the cliffhanger line, why it works, scores, a tier, and a "post these first" list.

It is one Python file with no dependencies. The AI model (Google Gemini by default) does the judging. Plain code does everything that must be exact: timestamps, rules, scoring and ranking.

## Contents

```
promo_finder.py   the whole tool (Python 3.9+, standard library only)
.env              put your API key(s) here
README.md         this file
```

## Quick start

1. **Python 3.9 or newer.** Nothing else to install.
2. **API keys.** Get a free Gemini key at https://aistudio.google.com/apikey and paste it after `GEMINI_API_KEY_1=` in `.env`. Add more keys on the next lines (`GEMINI_API_KEY_2=`, `GEMINI_API_KEY_3=` ...). The free tier gives each key a small daily quota, and a 50-episode show needs about 55–60 calls. When a key runs out, the tool moves to the next one by itself. Make each key in a different Google Cloud project, because keys in the same project share one quota.
3. **Scripts.** Put the show's zip file (or a folder of episode `.txt` files) next to `promo_finder.py`. The format is described below.
4. **Run:**

   ```bash
   python promo_finder.py
   ```

   - Try a few episodes first: `python promo_finder.py --episodes 1-3`
   - Read a zip or folder from somewhere else: `python promo_finder.py path/to/show.zip`
   - On Windows, `py promo_finder.py` also works. The script always looks for `.env`, the zip, `cache/` and `output/` next to itself, so you can run it from any folder.
5. **Results** are in `output/<show name>/`: `promos.csv` (opens in Excel or Google Sheets), `promos.json`, `raw_model_replies.jsonl` and `run_log.txt`.

**Do not commit** your real keys, your scripts, `cache/` or `output/` to a public repo. The `.env` in this repo is an empty template.

**Google Colab:** upload `promo_finder.py` and your zip, then run:

```python
open('.env', 'w').write('GEMINI_API_KEY_1=your-key\nGEMINI_API_KEY_2=another-key\n')
!python promo_finder.py
```

The results are in the `output` folder in the Colab file browser.

## Input format

One `.txt` file per episode, in a zip or a folder. The parser expects this layout (the content below is made up):

```
EPISODE NUMBER: 7
EPISODE LOGLINE: At her stepsister's gala, a maid is humiliated, until a guest reveals who she married.
CONTINUITY NOTES: Picks up the morning after her secret wedding.

--- SCRIPT START ---

SCENE 01
Scene heading: INT. BALLROOM - NIGHT
Scene start: (00:00.00)
Scene end: (00:26.00)

Actions and dialogue:

**CAMERA:**
Start: (00:00.00) | End: (00:04.00)
Wide shot of a glittering ballroom; guests turn toward the stairs.

**VIVIAN (Mocking):**
Start: (00:04.20) | End: (00:07.50)
"Look, the maid thinks she's a guest."

**ACTION:**
Start: (00:08.00) | End: (00:09.00)
Vivian throws a glass of red wine in Mia's face.

**SOUND FX:**
Start: (00:08.10) | End: (00:08.60)
Glass shatters on the marble floor.

--- SCRIPT END ---
```

- Each block starts with a tag: `**CAMERA:**`, `**ACTION:**`, `**SOUND FX:**`, or a speaker such as `**VIVIAN (Mocking):**`. Text in brackets after a speaker is read as the emotion; `(O.S.)`, `(V.O.)` and `(CONT'D)` are read as extensions.
- The line after a tag holds the timing, `Start: (MM:SS.cc) | End: (MM:SS.cc)`, counted from the start of the episode. The text follows on the next lines (or on the tag line itself).
- The episode number comes from the file name (`episode-007...txt`) or from the `EPISODE NUMBER:` line.
- The logline and continuity notes are optional but help the AI write the series bible.
- `--- SCRIPT START ---` / `--- SCRIPT END ---` are optional. If they are present, only the text between them is read.

Each tagged block becomes a numbered block (`b1`, `b2` ...) with exact times. The AI picks promos by block number and never types a timestamp. Code turns the block numbers back into exact times.

## How it works

| Step | Who | What happens |
|---|---|---|
| 1. Read | code | Every script block becomes a numbered block with exact start and end times |
| 2. Series bible | AI, 1 call | From every episode's logline, continuity notes and speaker list: the premise, what the title already promises, the leads and their speaker labels, label warnings, the story arc, and the twists with their spoiler level. This is the only show knowledge; nothing in the code is tied to one show |
| 3. Promos per episode | AI, 1 call each | The 3–8 strongest moments of the episode, then up to 3 promos as 1–3 ranges of block numbers, with the hook block, the cliffhanger block, up to 3 essential blocks, a title, why it works and four 1–5 scores |
| 4. Check and fix | code | Turns block numbers into times and enforces the rules below. Fixes small errors. If the AI's cut can't be fixed, rebuilds it from the AI's hook, cliffhanger and essential lines. Cuts that still fail go back to the AI once (one repair round); the rest are rejected |
| 5. Trailers | AI picks, code cuts | From the moments listed in step 3, the AI picks up to two trailers: a premise trailer that sells the title promise, and an optional story trailer (a threat in one episode, its consequence in a later one). Code turns each moment into a short clip and fits the total to 20–60 s |
| 6. Judge | AI, 2 calls | Sees every promo side by side (hook line, cliffhanger line, a short transcript, length, number of parts). It writes the open question each ending leaves, ranks all promos, and scores them. The second pass sees them in reverse order to cancel position bias |
| 7. Rank and write | code | Final scores, tiers, grouping of near-duplicate cuts, at most 2 promos per episode, the "post these first" list, and the output files |

### The rules code enforces

- Total length 20–60 seconds.
- The hook must start within the first 3 seconds.
- The promo ends on the cliffhanger: no new line of dialogue after it, and at most 10 seconds of silent reaction shot.
- No cut falls inside a line of dialogue.

What code fixes on its own (each fix is logged in the `code_fixes` column):

- A cut inside a line is widened to the whole line.
- A slow start is trimmed so the hook lands in the first 3 seconds (keeping up to 2.8 s of lead-in, starting on a camera cut if possible).
- New dialogue after the cliffhanger, or a hold longer than 10 s, is trimmed away.
- A cut within 1.2 s of a camera change is moved onto it.
- Overlapping clips are merged.

**Rebuilding a cut.** If the AI's ranges can't be fixed, code builds the cut from the hook block to the cliffhanger block. A stretch over 60 s is split at camera changes, and code keeps the hook, the cliffhanger, the essential lines and the build-up to the cliffhanger (two units of build-up for each unit after the hook) until the cut is about 45 s. A cliffhanger that is only a transition shot ("Cut to black") is replaced by the line before it.

### The prompts

All prompts are in section 6 of `promo_finder.py`. In short:

- **Promos per episode:** the AI plays a senior promo editor. The prompt lists hooks that stop a scroll (a betrayal revealed, public humiliation, a slap, an accusation, an ultimatum, an identity reveal...), cliffhangers that make people click (a threat not yet answered, a reveal cut off, someone arriving, danger about to land...), and weak moments (exposition, small talk, a problem already solved, a scene transition). It works in two steps: find the moments, then build promos, preferring one continuous stretch where hook and cliffhanger sit close together. Then it checks as a viewer who knows nothing. Scores use a strict 1–5 rubric, the answer is JSON, and there is one worked example from a made-up show. The script is shown as one line per block with a rough start time (`[0:04] b4 VIVIAN [Mocking]: ...`) and no end times, so the model never does arithmetic.
- **Judge:** step 1, write the open question each ending leaves ("none" if the ending closes the scene's conflict); step 2, order all candidates from best to worst; step 3, score them consistently with that order. An ending with no open question scores 1–2 on cliffhanger.
- **Trailers:** a premise trailer that opens on the setup's most shocking early moment and ends on the moment that teases the title promise, plus an optional story trailer. No moments that resolve the main story.

### Scores, tiers and the ranked list

- Each promo gets four 1–5 scores: **hook**, **cliffhanger**, **standalone** (does a stranger understand it?) and **spoiler safety**.
- The final scores are 50% the judge's (averaged over its two passes) and 50% the scores the AI gave when it picked the promo. Trailers use the judge's scores only.
- Two caps: if most judge passes found no open question, cliffhanger is capped at 2; if no lead character (from the series bible) speaks or is named in the promo, standalone is capped at 3.
- `total_score = 20 × (0.35 hook + 0.30 cliffhanger + 0.25 standalone + 0.10 spoiler safety)`, from 20 to 100. Ties are broken by cliffhanger, then hook, then the judge's ranking position.
- **Tiers:** A = post first (90 or more), B = strong (85 to 89.9), C = usable or backup (below 85).
- **Alternative cuts:** two promos from the same episode count as the same moment when their clips overlap by half or more (shared time ÷ combined time), when they share both the hook and the cliffhanger, or when their open questions share at least 60% of their words. The best-ranked cut stays in the list; the others become `alternative cut` rows, good for A/B tests.
- **At most 2 promos per episode** stay in the ranked list; the rest become `backup` rows. Trailers are not capped.
- **"Post these first" list (slate):** going down the ranked list, take the best promo from each episode (a trailer counts as its own pick), skip any the judge called the same moment as one already picked, and stop at 10. Then order them by tier, trailers first, then by episode number, because a new viewer starts at episode 1.

### Model switching (works on the free tier)

- Every call tries the models in order: `gemini-3.8-flash`, `gemini-3.6-flash`, `gemini-3.5-flash-lite`, `gemini-3.1-flash-lite`.
- If a model is busy (HTTP 5xx), refuses (content filter) or gives a broken answer, the next model answers that call at once. The next call starts again from the best model.
- A model that doesn't answer within 3 minutes (10 minutes for the judge) is skipped for 10 minutes.
- A key that is out of quota (HTTP 429) is skipped for that model for 10 minutes, and the next key is used at once. A rejected key is never used again.
- The tool waits only if every model is busy at the same moment, and then only once, for one minute.
- Every answer is saved in `cache/`, keyed by model and prompt, so running the same show again is instant and free. Delete `cache/` to get fresh answers.
- **Other models:** every call goes to an OpenAI-compatible `/chat/completions` endpoint, with JSON output requested. To use another provider, change `API_URL`, `KEY_NAME` and `MODELS` at the top of the script (and `PRICES` for the cost line).

## Output

`output/<show name>/` holds four files:

| File | What it is |
|---|---|
| `promos.csv` | One row per promo: the ranked promos first, then backups and alternative cuts |
| `promos.json` | Everything: the ranked promos, alternatives, rejected picks (with the reason), the series bible, the slate, token usage, and every call |
| `raw_model_replies.jsonl` | Every AI reply, unedited, one per line |
| `run_log.txt` | The screen output of the run |

The columns of `promos.csv`, from left to right:

| Column | Meaning |
|---|---|
| `rank` | Position in the ranked list, 1 = best. Empty for backups and alternative cuts |
| `slate_position` | The suggested posting order for the first 10 promos (1 = post first). Empty for all other promos |
| `list` | `ranked`, `alternative cut` (another cut of a ranked promo) or `backup` (a good promo from an episode that already has 2 in the list) |
| `tier` | A, B or C (see above) |
| `total_score` | 20–100, computed by code from the four scores |
| `promo_id` | `E05-1` = the first promo found in episode 5; `T1` = the first trailer across episodes |
| `title` | A short working title written by the AI |
| `type` | `single-segment` (one clip), `multi-segment` (2–4 clips from one episode) or `cross-episode` (a trailer) |
| `episodes` | The episodes the promo uses, e.g. `3 9 14` for a trailer |
| `segments` | The exact cut: play these clips in this order, e.g. `E07 00:04.20-00:22.40 + E07 00:31.00-00:52.10` (minutes:seconds from the start of each episode) |
| `total_duration_s` | Length in seconds (always 20–60) |
| `hook_time`, `hook_line`, `hook` | When the hook starts, the script line itself, and the AI's short description of it |
| `cliffhanger_time`, `cliffhanger_line`, `cliffhanger` | The same for the line the promo ends on |
| `open_question` | The question the ending leaves the viewer with, written by the judge (`none` = the ending closes the scene) |
| `why_it_works` | One line from the AI |
| `hook_score`, `cliffhanger_score`, `standalone_score`, `spoiler_safety` | The four 1–5 scores |
| `judge_note` | The judge's one-line comment, plus a note when code capped a score |
| `variant_of` | For an alternative cut: the promo it is another cut of |
| `code_fixes` | What code changed in the AI's cut to meet the rules |
| `check_notes` | Small warnings that are not rule breaks (e.g. a cut not on a camera change) |
| `model` | The model that picked the promo |
| `reviewer_verdict`, `reviewer_notes` | Empty, for an editor to fill in |

## Settings

All settings are plain values at the top of `promo_finder.py`:

| Setting | Default | What it does |
|---|---|---|
| `API_URL`, `KEY_NAME`, `MODELS` | Gemini (OpenAI-compatible endpoint) | Which models to call, best first; every `.env` key whose name starts with `KEY_NAME` is used |
| `PRICES` | Gemini paid prices | Used only for the cost line at the end of a run |
| `MAX_TOKENS` | 32768 | Room for the longest answer (the judge) |
| `TIMEOUT_S`, `JUDGE_TIMEOUT_S` | 180, 600 | When a call counts as failed |
| `SKIP_SLOW_MIN` | 10 | How long a model that timed out (or a key out of quota) is skipped |
| `BUSY_WAIT_S` | 60 | The one wait when every model is busy |
| `MIN_S`, `MAX_S` | 20, 60 | Promo length in seconds |
| `HOOK_S` | 3 | The hook must start within this many seconds |
| `MAX_HOLD_S` | 10 | The longest silent hold after the cliffhanger line |
| `SNAP_S` | 1.2 | A cut this close to a camera change is moved onto it |
| `TARGET_S` | 45 | The length code aims for when it rebuilds a cut |
| `PROMOS_PER_EPISODE` | 3 | How many promos the AI may propose per episode |
| `MAX_PER_EPISODE` | 2 | How many promos per episode stay in the ranked list |
| `TRAILER_MOMENTS` | 4 | Moments per episode offered for trailers |
| `CAMERA_MAX_CHARS` | 160 | Long camera descriptions are shortened in prompts |
| `JUDGE_PASSES`, `JUDGE_BATCH` | 2, 130 | Judge passes, and how many promos one judge call compares |
| `SELF_SCORE_WEIGHT` | 0.5 | The share of the AI's own scores in the final scores |
| `SAME_OVERLAP`, `SAME_QUESTION` | 0.5, 0.6 | When two cuts count as the same moment |
| `SLATE_SIZE` | 10 | Length of the "post these first" list |

## How I tested it

**Reference set.** To measure quality I needed answers to compare against, so I built a reference set for a 50-episode test show with Claude Opus 5.5, the strongest model I could use. It has 73 ranked promos (68 when alternative cuts of the same moment count once, 27 of them strong, tier A or B), plus 19 "near misses": moments that look like promos but aren't, such as a triumphant speech that leaves nothing open. Code checked every pick against the rules. It is one AI's opinion, not ground truth.

**Matching.** A promo counts as found when it overlaps a reference promo by at least half of their combined time (the same `overlap()` function the tool uses for grouping). A looser 30% overlap counts as "the same moment, cut differently".

**Measures.** Strong reference promos found, all found, precision, how many of the top 5 and top 10 match, rank agreement (Spearman), picks that hit a near miss, promos in quiet episodes (episodes with no good promo), and rule breaks.

**Method.** I tuned on 12 episodes and checked the main changes on 12 other episodes, so I wasn't only fitting the episodes I tuned on.

**Choosing the model.** Finding promos needs story sense (what shocks, what leaves a question open, what a stranger understands), so I used an LLM rather than rules or a small classifier. On the 12 tuning episodes, Gemini 3.6 Flash found 8 of 9 strong promos, Gemini 3.5 Flash-Lite found 5, and gpt-oss-120b on Groq found 3. Other free services asked for payment or refused my first requests. Gemini 3.8 Flash is newer, but on the free tier it was often overloaded.

**What I tried, from simple to smarter prompting** (Gemini 3.6 Flash, held-out episodes):

| Version | What the AI did | Result |
|---|---|---|
| 1. Simple prompt | Picked exact block ranges for each promo | Good picks, but 21–45% broke a timing rule (mostly late hooks). 6 of 9 strong promos found |
| 2. Less work for the AI | Named only the moments; code built every cut | No rule breaks and 3× fewer calls. Helped the weak Flash-Lite, but hurt 3.6 Flash: precision fell from 71% to 32% (choppy cuts) |
| 3. Hybrid (final) | Proposes a tight cut; code checks and fixes it, or rebuilds it from the hook and cliffhanger | Best of both: 8 of 9 strong promos, 12 of 16 overall |
| Better judging | Second judge pass in reverse order; the open-question test; judge scores blended 50/50 with the picker's own | Rank agreement with the reference set rose from 0.11 to 0.56 |

Tried and dropped: letting code pair extra hook and cliffhanger moments into new promos (more found on Flash-Lite, less precise on 3.6 Flash), grouping by the judge's own "same moment" flags (it merged different promos), and keeping 1 or 3 promos per episode (2 was the best balance).

**Results on the full 50-episode test show** (two complete runs):

| | Run 1 | Run 2 |
|---|---|---|
| Promos found | 85 ranked (13 A, 15 B, 57 C) + 16 alternative cuts and 11 backups | 76 ranked (12 A, 16 B, 48 C) + 5 alternative cuts and 6 backups |
| Strong reference promos found | 23 of 27 (26 counting the same moment cut differently) | 14 of 27 (19) |
| Top 5 / top 10 that match | 4 of 5, 8 of 10 | 2 of 5, 6 of 10 |
| Rule breaks | 0 | 0 |
| Which model answered | 3.6 Flash 50, 3.8 Flash 7 | 3.6 Flash 36, Flash-Lite 23 |
| Calls, cost at paid prices, time on the free tier | 57, $0.43, 57 minutes | 59, $0.34, 31 minutes |

In run 2, Gemini 3.6 Flash was overloaded, so the weaker Flash-Lite answered 21 of the 50 episodes. That shows how much the model matters. In run 1, 43% of the AI's first-draft cuts needed a code fix (mostly hooks that started late), and the two quiet episodes got no promo.

## Where it fails

- **It depends on the model.** When Flash-Lite answers instead of 3.6 Flash, far fewer strong promos are found and quiet episodes get weak promos.
- **Resolved endings** that look dramatic (a big speech, a vow) can still rank too high.
- **Same moment, different cut:** about a third of the misses are the right scene cut differently, which is a matter of taste.
- **Ranking is weaker than finding:** rank agreement with the reference set was 0.49 on the full show.
- **Text only:** a line can read well and still play badly on screen.
- **Content filter:** Gemini sometimes refuses a violent episode; the tool then skips that call and logs it.
- **The reference set is AI-made**, so the numbers show how often two AIs agree, not the truth.

## Cost and scale

- A 50-episode show takes about 57 calls: 1 bible, 50 episodes, a few repairs, 1 trailer call and 2 judge passes. That's about 330K input and 49K output tokens, or about $0.43 at paid Gemini Flash prices (under 1 cent per episode, so roughly $430 for 1,000 shows).
- On the free tier the bottleneck is quota and overloaded models: 30–60 minutes per show. On a paid key with episodes sent in parallel it would take a few minutes.
- A new show needs nothing but its zip; the series bible is the only show knowledge.
- Next steps: an editor reviews the top of the CSV (it has reviewer columns) and A/B tests the alternative cuts. Their verdicts and real numbers (hook rate, click-through) are the best way to tune the score weights, and later to fine-tune a smaller, cheaper model. Reading the video as well as the text would catch lines that play badly on screen.

**Privacy:** prompts sent to Gemini's free tier may be used by Google to improve its products. Use a paid key for confidential scripts.

## How I built it

I used Claude (Claude Code with Claude Opus 5.5) to write the code, run the tests, do the repetitive work, and build the reference set. The decisions were mine: the problem, the model, which ideas to try, which results to trust, and what went into the final tool. Gemini runs inside Promo Finder.

Where the AI tools got it wrong:
- Claude's version-2 prompt told the model that hooks and cliffhangers could be "far apart", because code would trim the middle. Gemini 3.6 Flash then built choppy cuts, and precision fell from 71% to 32%. The reference set caught it, and version 3 fixed it.
- The first draft of the reference set broke the timing rules in 11 of 73 picks until code checked it.
- Gemini's series bible mixed up one speaker.
- The first "open question" rule also demoted real revenge-vow cliffhangers, until I changed its wording.
