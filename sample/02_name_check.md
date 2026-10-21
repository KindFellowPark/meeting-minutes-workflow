# Name check for the sample meeting

Step 2 of the skill. Claude compares every name in `01_transcript.txt` against
`skill/minutes-builder/references/known_names.md` and stops with one line per suspect name.
The person preparing the minutes answers each line before any minutes are drafted.

## What Claude flagged

```
[NAME CHECK] Transcript says "Priya Ramen". Possible match: Priya Ramanathan (Development Director). Correct?
[NAME CHECK] Transcript says "Marcus Oil Aaron". Possible match: Marcus Oyelaran (Events Coordinator). Correct?
[NAME CHECK] Transcript says "Castle Lane gallery". Possible match: Castellane Gallery (event venue). Correct?
[NAME CHECK] Transcript says "Lumen Hall". Possible match: Lumenhall Catering (caterer). Correct?
[NAME CHECK] Transcript says "Teodora". Not on the list; no surname given. How should she appear in the minutes?
```

## What the person preparing the minutes answered

| Flag | Answer |
|---|---|
| Priya Ramen | Yes, Priya Ramanathan |
| Marcus Oil Aaron | Yes, Marcus Oyelaran |
| Castle Lane gallery | Yes, Castellane Gallery |
| Lumen Hall | Yes, Lumenhall Catering |
| Teodora | Use "a scholarship student (name to be confirmed)" until Priya confirms her |

Only after these answers does Claude write `03_minutes_input.txt`. The last flag shows why the
check waits for a person: Claude has no way to know a student's surname, and guessing one into
a document that goes to an executive is worse than leaving it open.
