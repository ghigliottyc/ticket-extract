# Ticket Extractor

This script turns a folder of baseball ticket photos into a review-ready Excel workbook. Each image is sent to Claude's vision API, the response is validated against a schema, cross-checked with deterministic rules, and written to a spreadsheet with color-coded status so you know exactly which rows to double-check by hand.

## Features

- **Structured, validated output** — every field comes back as `{value, confidence, evidence}` and is validated against a Pydantic schema before it's trusted. Malformed model output is retried, not silently accepted.
- **Anti-hallucination extraction prompt** — the model is explicitly instructed to transcribe only what's visible, never infer from baseball knowledge or "what the ticket should say," and to return `null` rather than guess.
- **Deterministic cross-field validation** — independent of the model, the script:
  - Re-derives the day of week from the parsed date and flags it if it disagrees with what's printed on the ticket
  - Flags home team == away team
  - Flags a price field with no digits in it
  - Flags fields that are empty strings instead of `null`
  - Flags rows where the model's overall confidence is "high" but an important field isn't
- **Explicit REVIEW status** — Each row gets `OK`, `REVIEW` (validation flagged something), or `ERROR` (extraction failed), so you can filter the sheet and only look at what needs attention.
- **Image handling** — corrects EXIF orientation, handles HEIC (iPhone) images via `pillow-heif`, resizes and attempts to sharpen images before sending them to the API.
- **Concurrent processing with retries** — configurable worker pool, exponential backoff on transient API errors (rate limits, timeouts, connection errors), separate retry path for invalid/malformed model output.
- **Full audit trail** — the raw model transcription, validation issues, model name, and timestamp are written to the sheet for every row, so any result can be traced back and independently checked.

## Requirements

```bash
pip install anthropic openpyxl pillow pydantic pillow-heif
```

You'll also need an Anthropic API key, available via the `ANTHROPIC_API_KEY` environment variable or the `--api-key` flag.

## Usage

```bash
# Basic run — defaults to the haiku model
python extract_tickets_v2.py --tickets ./tickets --output tickets.xlsx

# Use a stronger model
python extract_tickets_v2.py --tickets ./tickets --output tickets.xlsx --model sonnet

# More concurrent requests
python extract_tickets_v2.py --tickets ./tickets --output tickets.xlsx --workers 6
```

### CLI options

| Flag | Default | Description |
|---|---|---|
| `--tickets` | *(required)* | Folder containing ticket images |
| `--output` | `baseball_tickets_v2.xlsx` | Output Excel path |
| `--model` | `haiku` | Model alias — `haiku` or `sonnet` |
| `--api-key` | `$ANTHROPIC_API_KEY` | Anthropic API key |
| `--workers` | `4` | Max concurrent API requests |
| `--retries` | `3` | Retries for transient/invalid responses |
| `--timeout` | `120.0` | Per-request timeout, in seconds |
| `--max-image-dimension` | `2400` | Resize the image's longest side to this many pixels before sending |

Supported image formats: `.jpg`, `.jpeg`, `.png`, `.heic`, `.webp`, `.bmp`, `.tif`, `.tiff`.

## Output

The workbook has two sheets:

**`Tickets`** — one row per image, with a column per extracted field (date, day printed vs. day calculated, time, home/away team, stadium, game #, section, row, seat, seat type, gate, price, season ticket), plus:

- **Confidence** — the model's self-reported overall confidence
- **Validation Issues** — any deterministic checks that fired, newline-separated
- **Status** — `OK`, `REVIEW`, or `ERROR`, color-coded (green/yellow/red) for quick scanning
- **Raw Text** — the model's transcription of the ticket, for auditing any field against the source
- **Model** / **Processed At** — which model produced the row and when

The sheet has frozen header/first column, an autofilter, and conditional formatting on the confidence column.

**`Summary`** — totals: tickets processed, successfully extracted, errors, rows needing review, and a confidence breakdown.

### A note on the REVIEW status

`Validation Issues` are deterministic sanity checks. A flagged row often just means "something here is worth a second look" — e.g. a date format the parser doesn't recognize yet, not necessarily a misread ticket.

## How extraction works

1. **Image normalization** (`normalize_image`) — corrects orientation via EXIF, converts to RGB, resizes if needed, applies mild contrast/sharpness enhancement, and encodes as base64 JPEG.
2. **Extraction** (`extract_ticket`) — sends the image + prompt to Claude, strips markdown code fences from the response, parses it as JSON, and validates it against the `TicketExtraction` Pydantic model. Invalid JSON or schema violations are retried (with backoff) up to `--retries` times; so are transient API errors (rate limits, timeouts, connection issues, 5xxs) — these two failure classes are handled separately since one is a model-output problem and the other is a transport problem.
3. **Validation** (`validate_extraction`) — runs the deterministic checks described above and returns a list of human-readable issue strings.
4. **Excel output** (`build_excel`) — writes both sheets, with per-row coloring based on status.
