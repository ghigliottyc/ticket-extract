#!/usr/bin/env python3
"""
Baseball Ticket Extractor
Extracts structured data from ticket images using Claude's vision API
and saves results to an Excel spreadsheet.

Requirements:
    pip install anthropic openpyxl pillow

Usage:
    python extract_tickets.py --tickets /path/to/ticket/images --output tickets.xlsx
    python extract_tickets.py --tickets ./tickets --output my_games.xlsx --model haiku
"""

import anthropic
import base64
import json
import argparse
import sys
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Configuration ────────────────────────────────────────────────────────────

SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".heic", ".webp", ".bmp"}

MODELS = {
    "haiku":  "claude-haiku-4-5-20251001",   # Fastest & cheapest (~$0.01–0.02 per 200 images)
    "sonnet": "claude-sonnet-4-6",            # More accurate, still affordable
}

FIELDS = [
    "date",           # e.g. "July 13, 2008"
    "day_of_week",    # e.g. "Sunday"
    "game_time",      # e.g. "8:05 PM"
    "home_team",      # e.g. "New York Mets"
    "away_team",      # e.g. "Colorado Rockies"
    "stadium",        # e.g. "Shea Stadium"
    "game_number",    # e.g. "47"
    "section",        # e.g. "27"
    "row",            # e.g. "R"
    "seat",           # e.g. "8"
    "seat_type",      # e.g. "Upper", "Loge Reserved", "Field Box"
    "gate",           # e.g. "D"
    "price",          # e.g. "$37.00"
    "season_ticket",  # "Yes" / "No"
    "notes",          # Any unusual info or flags for manual review
    "confidence",     # "high" / "medium" / "low"
]

EXTRACTION_PROMPT = """You are extracting data from a baseball ticket image. 
Return ONLY a valid JSON object with these exact keys. Use null for any field you cannot read.

Keys to extract:
- date: Full date as written on ticket (e.g. "July 13, 2008")
- day_of_week: Day of week (e.g. "Sunday")  
- game_time: Game start time (e.g. "8:05 PM")
- home_team: Full home team name (e.g. "New York Mets")
- away_team: Full visiting/away team name (e.g. "Colorado Rockies")
- stadium: Stadium or venue name
- game_number: Game number shown on ticket (just the number, e.g. "47")
- section: Section number or code
- row: Row letter or number
- seat: Seat number
- seat_type: Seating area type (e.g. "Upper Deck", "Loge Reserved", "Field Box", "Diamond Club")
- gate: Entry gate (e.g. "D")
- price: Ticket face value including $ symbol (e.g. "$37.00"). Use null if not shown.
- season_ticket: "Yes" if this is a season ticket, "No" otherwise
- notes: Any flags for manual review (e.g. "rain check stub only", "date partially obscured")
- confidence: Your overall confidence in the extraction — "high", "medium", or "low"

Return ONLY the JSON object. No explanation, no markdown, no backticks."""

# ── Core Extraction ──────────────────────────────────────────────────────────

def encode_image(path: Path) -> tuple[str, str]:
    """Return (base64_data, media_type) for an image file."""
    suffix = path.suffix.lower()
    media_types = {
        ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        ".png": "image/png",  ".webp": "image/webp",
        ".gif": "image/gif",  ".bmp": "image/png",   # bmp → send as png after convert
        ".heic": "image/jpeg",                        # HEIC treated as jpeg
    }
    media_type = media_types.get(suffix, "image/jpeg")

    # Convert HEIC / BMP if Pillow is available
    if suffix in {".heic", ".bmp"}:
        try:
            from PIL import Image
            import io
            from pillow_heif import register_heif_opener
            register_heif_opener()
            img = Image.open(path).convert("RGB")
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=90)
            return base64.standard_b64encode(buf.getvalue()).decode("utf-8"), "image/jpeg"
        except ImportError:
            pass  # Fall through to raw read

    with open(path, "rb") as f:
        return base64.standard_b64encode(f.read()).decode("utf-8"), media_type


def extract_ticket(client: anthropic.Anthropic, image_path: Path, model: str) -> dict:
    """Send one ticket image to Claude and return parsed JSON fields."""
    img_data, media_type = encode_image(image_path)

    response = client.messages.create(
        model=model,
        max_tokens=1000,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": media_type,
                        "data": img_data,
                    },
                },
                {"type": "text", "text": EXTRACTION_PROMPT},
            ],
        }],
    )

    raw = response.content[0].text.strip()

    # Strip markdown fences if model adds them
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()

    result = json.loads(raw)

    # Ensure all expected fields are present
    for field in FIELDS:
        result.setdefault(field, None)

    return result


# ── Excel Output ─────────────────────────────────────────────────────────────

HEADERS = [
    "File Name", "Date", "Day", "Time", "Home Team", "Away Team",
    "Stadium", "Game #", "Section", "Row", "Seat", "Seat Type",
    "Gate", "Price", "Season Ticket", "Notes", "Confidence",
]

FIELD_TO_HEADER = {
    "date": "Date", "day_of_week": "Day", "game_time": "Time",
    "home_team": "Home Team", "away_team": "Away Team", "stadium": "Stadium",
    "game_number": "Game #", "section": "Section", "row": "Row",
    "seat": "Seat", "seat_type": "Seat Type", "gate": "Gate",
    "price": "Price", "season_ticket": "Season Ticket",
    "notes": "Notes", "confidence": "Confidence",
}

CONFIDENCE_COLORS = {
    "high":   "C6EFCE",  # green
    "medium": "FFEB9C",  # yellow
    "low":    "FFC7CE",  # red
}

HEADER_FILL   = PatternFill("solid", start_color="1F4E79")
HEADER_FONT   = Font(bold=True, color="FFFFFF", name="Arial", size=10)
DATA_FONT     = Font(name="Arial", size=10)
CENTER        = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT          = Alignment(horizontal="left",   vertical="center", wrap_text=True)
THIN          = Side(style="thin", color="CCCCCC")
BORDER        = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

COL_WIDTHS = {
    "File Name": 28, "Date": 18, "Day": 10, "Time": 10,
    "Home Team": 22, "Away Team": 22, "Stadium": 22,
    "Game #": 8, "Section": 9, "Row": 6, "Seat": 6,
    "Seat Type": 18, "Gate": 6, "Price": 10,
    "Season Ticket": 13, "Notes": 35, "Confidence": 12,
}


def build_excel(records: list[dict], output_path: Path):
    wb = Workbook()

    # ── Data sheet ──
    ws = wb.active
    ws.title = "Tickets"
    ws.freeze_panes = "B2"
    ws.row_dimensions[1].height = 30

    # Headers
    for col_idx, header in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font   = HEADER_FONT
        cell.fill   = HEADER_FILL
        cell.alignment = CENTER
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = COL_WIDTHS.get(header, 14)

    # Data rows
    for row_idx, rec in enumerate(records, 2):
        ws.row_dimensions[row_idx].height = 18
        data = rec.get("data", {})
        err  = rec.get("error")
        conf = (data.get("confidence") or "").lower()
        fill_color = CONFIDENCE_COLORS.get(conf, "FFFFFF")
        row_fill   = PatternFill("solid", start_color=fill_color)

        row_values = [
            rec["filename"],
            data.get("date"),
            data.get("day_of_week"),
            data.get("game_time"),
            data.get("home_team"),
            data.get("away_team"),
            data.get("stadium"),
            data.get("game_number"),
            data.get("section"),
            data.get("row"),
            data.get("seat"),
            data.get("seat_type"),
            data.get("gate"),
            data.get("price"),
            data.get("season_ticket"),
            err if err else data.get("notes"),
            data.get("confidence") if not err else "ERROR",
        ]

        for col_idx, value in enumerate(row_values, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font      = DATA_FONT
            cell.border    = BORDER
            cell.fill      = row_fill
            cell.alignment = CENTER if col_idx != len(HEADERS) - 1 else LEFT

    # Auto-filter
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}1"

    # ── Summary sheet ──
    ws2 = wb.create_sheet("Summary")
    total     = len(records)
    succeeded = sum(1 for r in records if not r.get("error"))
    failed    = total - succeeded
    low_conf  = sum(1 for r in records
                    if not r.get("error") and
                    (r["data"].get("confidence") or "").lower() == "low")

    summary_rows = [
        ("Total tickets processed", total),
        ("Successfully extracted",  succeeded),
        ("Errors (needs manual review)", failed),
        ("Low confidence extractions",   low_conf),
        ("Run date", datetime.now().strftime("%B %d, %Y %I:%M %p")),
    ]

    ws2["A1"] = "Ticket Extraction Summary"
    ws2["A1"].font = Font(bold=True, size=14, name="Arial", color="1F4E79")
    ws2.column_dimensions["A"].width = 35
    ws2.column_dimensions["B"].width = 20

    for i, (label, value) in enumerate(summary_rows, 3):
        ws2.cell(row=i, column=1, value=label).font = Font(bold=True, name="Arial")
        ws2.cell(row=i, column=2, value=value).font = Font(name="Arial")

    ws2["A9"] = "Color Legend"
    ws2["A9"].font = Font(bold=True, name="Arial")
    for i, (conf, color, label) in enumerate([
        ("high",   "C6EFCE", "High confidence"),
        ("medium", "FFEB9C", "Medium confidence — review recommended"),
        ("low",    "FFC7CE", "Low confidence — manual review required"),
    ], 10):
        c = ws2.cell(row=i, column=1, value=label)
        c.fill = PatternFill("solid", start_color=color)
        c.font = Font(name="Arial")

    wb.save(output_path)


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Extract baseball ticket data to Excel")
    parser.add_argument("--tickets", required=True,
                        help="Path to folder containing ticket images")
    parser.add_argument("--output", default="baseball_tickets.xlsx",
                        help="Output Excel file path (default: baseball_tickets.xlsx)")
    parser.add_argument("--model", choices=["haiku", "sonnet"], default="haiku",
                        help="Claude model to use (default: haiku — fastest/cheapest)")
    parser.add_argument("--api-key", default=None,
                        help="Anthropic API key (or set ANTHROPIC_API_KEY env var)")
    args = parser.parse_args()

    tickets_dir = Path(args.tickets)
    if not tickets_dir.is_dir():
        print(f"Error: '{tickets_dir}' is not a valid directory.")
        sys.exit(1)

    image_files = sorted([
        p for p in tickets_dir.iterdir()
        if p.suffix.lower() in SUPPORTED_EXTENSIONS
    ])

    if not image_files:
        print(f"No supported image files found in '{tickets_dir}'.")
        print(f"Supported formats: {', '.join(SUPPORTED_EXTENSIONS)}")
        sys.exit(1)

    model_id = MODELS[args.model]
    client   = anthropic.Anthropic(api_key=args.api_key) if args.api_key else anthropic.Anthropic()

    print(f"\n🎟️  Baseball Ticket Extractor")
    print(f"   Found {len(image_files)} image(s) in '{tickets_dir}'")
    print(f"   Model : {args.model} ({model_id})")
    print(f"   Output: {args.output}\n")

    records = []
    errors  = 0

    for i, img_path in enumerate(image_files, 1):
        print(f"[{i:>3}/{len(image_files)}] {img_path.name} ... ", end="", flush=True)
        try:
            data = extract_ticket(client, img_path, model_id)
            conf = (data.get("confidence") or "?").upper()
            print(f"✓  ({conf})")
            records.append({"filename": img_path.name, "data": data})
        except json.JSONDecodeError as e:
            print(f"⚠  JSON parse error: {e}")
            records.append({"filename": img_path.name, "data": {}, "error": f"JSON parse error: {e}"})
            errors += 1
        except Exception as e:
            print(f"✗  {e}")
            records.append({"filename": img_path.name, "data": {}, "error": str(e)})
            errors += 1

    output_path = Path(args.output)
    build_excel(records, output_path)

    print(f"\n✅ Done! {len(records) - errors}/{len(records)} tickets extracted successfully.")
    if errors:
        print(f"⚠️  {errors} ticket(s) had errors — see the 'Notes' column for details.")
    print(f"📊 Saved to: {output_path.resolve()}\n")


if __name__ == "__main__":
    main()
