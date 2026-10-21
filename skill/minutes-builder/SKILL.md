---
name: minutes-builder
description: >
  Turn a meeting transcript into finished, branded meeting minutes as a Word document.
  Use when the user says "write up the minutes", "format these minutes", "run the minutes
  builder", "create minutes from this transcript", or has a meeting transcript ready to turn
  into a document. Checks every name against a known-names list and waits for a person to
  confirm each uncertain one before drafting.
---

# minutes-builder

This skill turns a meeting transcript into a minutes document in a fixed house style. The
styling (heading colours, the two-column details table, the three-column action-items table)
lives in `scripts/generate_minutes_docx.py`, so the skill never re-creates it by hand. The
skill's job is to (1) check names, (2) shape the transcript into the script's input format,
and (3) run the script.

The transcript can come from anywhere: a Teams or Zoom transcript that already has speaker
labels, or a recording transcribed on the user's own machine. Recordings with sensitive
content should be transcribed locally, so the audio never leaves the computer.

## Step 1: Check that python-docx loads

```bash
python3 -c "import docx" || { echo "python-docx missing: pip install python-docx"; exit 1; }
```

Stop if it fails. Do not install packages on every run.

## Step 2: Check names against the known-names list

Read `references/known_names.md`. Then read the raw transcript and look for every name of a
person, vendor, place or program.

Speech-to-text gets uncommon names wrong more often than anything else ("Oyelaran" comes out
as "Oil Aaron"; "Lumenhall" comes out as "Lumen Hall"). For each name that is a near-miss for
an entry on the list, or is not on the list at all, write one line:

```
[NAME CHECK] Transcript says "Oil Aaron". Possible match: Marcus Oyelaran (Events Coordinator). Correct?
[NAME CHECK] Transcript says "Teodora". Not on the list; no surname given. How should she appear?
```

Then **stop and wait** for the user to confirm or correct every line. This is a review gate,
not an auto-correct: never substitute a name silently, and never guess a name that is not in
the transcript or the list. If there are no suspect names, say so in one line and continue.

When the user confirms a new name, offer to add it to `references/known_names.md` so the next
meeting is checked against it.

## Step 3: Shape the transcript into the input format

The script reads a structured plain-text file, not a raw transcript. Write the minutes as a
`.txt` file using these markers, each at the start of a line:

```
##TITLE## MEETING MINUTES
##SUBTITLE## Organization, Topic
##META##
Date: Tuesday, March 3, 2026
Time: 10:00 AM to 11:00 AM
Location: Zoom
Participants: Name One, Name Two, Name Three
##SUMMARY## One-paragraph summary. Omit the line if there is no summary.
##SECTION## First Section Header
Prose for the section. A blank line separates paragraphs.
- A line starting with "- " becomes a bullet.
##TABLE## Optional caption
Header one | Header two
Row label | Value
##ACTIONITEMS##
- Action description — Owner — Deadline
##BYLINE## Prepared by Name, Title, Date
```

Rules that matter:

- Write what was decided and who owns what, not a word-for-word record. Use confirmed names
  from Step 2 only.
- Never invent a deadline. If the meeting gave none, write `TBD` or the words that were said
  ("This week").
- `**bold**` and `_italic_` work in prose and bullets only. Use bold sparingly, for the figure,
  the deadline or the constraint a reader must not miss.
- Action items use space, em dash, space as the field separator: `Action — Owner — Deadline`.
  The parser splits on it.
- A `##TABLE##` block runs until the next `##` marker, so put it last in its section.

## Step 4: Run the generator

```bash
python3 scripts/generate_minutes_docx.py --input "<structured .txt>" --output "<minutes .docx>"
```

A successful run prints `Saved: <path>`.

## Step 5: Hand off

Tell the user where the document is. The user reads it and makes any final edits in Word
before it goes out; the skill produces a draft for a person to finish, not a document that
is sent without being read.
