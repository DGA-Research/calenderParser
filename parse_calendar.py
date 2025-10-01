from __future__ import annotations

import csv
import re
from datetime import datetime
from pathlib import Path

import pdfplumber
from pdfplumber.utils import cluster_objects

PDF_NAME = 'AG Taylor Calendar - September 2023.pdf'
OUTPUT_NAME = 'ag_taylor_calendar_sept_2023.csv'

DATE_PATTERN = re.compile(r'([A-Za-z]+ \d{1,2}, \d{4})')
DAY_NAMES = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
DAY_NAMES_UPPER = {name.upper() for name in DAY_NAMES}
DIGITS_ONLY_PATTERN = re.compile(r'^[0-9]+$')

TIME_COLUMN_X_THRESHOLD = 70
LINE_CLUSTER_TOLERANCE = 3
INDENT_NEW_EVENT_THRESHOLD = 25
NEW_EVENT_KEYWORDS = {
    'MEETING',
    'MTG',
    'WORK',
    'PRESS',
    'PH:',
    'CALL',
    'BRIEF',
    'CONF',
    'CHECK',
    'LUNCH',
    'DINNER',
    'TOUR',
    'HEARING',
    'TRAVEL',
    'INTERVIEW',
}
NOISE_SEGMENTS = {'BOI BOI'}

SPECIAL_SPLIT_MARKERS = [
    "CHIEF'S CONFERENCE",
    "CHIEFS CONFERENCE",
]

LOCATION_HINTS = {
    'MICROSOFT TEAMS',
    'IN PERSON',
    'TELECONFERENCE',
    'ZOOM',
    'BOI',
    'ANC',
    'JNU',
    'ROOM',
    'CONF',
    'SUITE',
    'FLOOR',
    'OFFICE',
    'BUILDING',
    'CAPITOL',
    'ANCHORAGE',
    'JUNEAU',
    'PALMER',
    'FAIRBANKS',
    'TEAMS',
}

FORCED_EVENT_PREFIXES = (
    'CHECK IN',
    'STATEHOOD',
    'OP-ED',
    'OP ED',
    'WORK ON LIST',
    'UPDATE',
    "CHIEF'S CONFERENCE",
)

PERSON_HINTS = {
    'GOV)',
    'LAW)',
    'DNR)',
    'DOL)',
    'DOC)',
    'DEC)',
    'DOT)',
    'EDU)',
    'JUD)',
    'LEG)',
}


def build_lines(words):
    clusters = cluster_objects(words, 'top', tolerance=LINE_CLUSTER_TOLERANCE)
    lines = []
    for cluster in clusters:
        cluster_sorted = sorted(cluster, key=lambda w: w['x0'])
        text = ' '.join(word['text'] for word in cluster_sorted).strip()
        if not text:
            continue
        x0 = min(word['x0'] for word in cluster_sorted)
        lines.append({'text': text, 'x0': x0, 'words': cluster_sorted})
    return lines


def extract_date(lines):
    for line in lines:
        match = DATE_PATTERN.search(line['text'])
        if match:
            try:
                return datetime.strptime(match.group(1), '%B %d, %Y').date()
            except ValueError:
                continue
    return None


def find_day_line_index(lines):
    for idx, line in enumerate(lines):
        if line['text'].strip().upper() in DAY_NAMES_UPPER:
            return idx
    return None


def split_event_text(event_text: str) -> list[str]:
    if not event_text:
        return []
    normalized = event_text.replace(' PH:', ';PH:')
    for marker in SPECIAL_SPLIT_MARKERS:
        pattern = re.compile(r'\s+' + re.escape(marker), flags=re.IGNORECASE)
        normalized = pattern.sub(lambda m: '; ' + m.group(0).strip(), normalized)
    parts = [part.strip() for part in normalized.split(';')]
    return [part for part in parts if part and part.upper() not in NOISE_SEGMENTS]


def segment_has_digits(text: str) -> bool:
    return any(ch.isdigit() for ch in text)


def should_force_new_event(segment: str, current_segments: list[str]) -> bool:
    cleaned = segment.strip()
    if not cleaned:
        return False
    upper = cleaned.upper()
    if upper.startswith('PH:'):
        return False
    if 'TELECONFERENCE' in upper:
        return False
    if is_location_segment(segment):
        return False
    if re.fullmatch(r'[A-Z][a-z]+(?: [A-Z][a-z]+){0,2}(?: [A-Z]\.)?(?: \(.+\))?', cleaned):
        return True
    if any(hint in upper for hint in PERSON_HINTS):
        return False
    if re.match(r'^\d{1,2}[:.]\d{2}', cleaned):
        return True
    if re.match(r'^\d{1,2}\s?(AM|PM)', upper):
        return True
    if any(upper.startswith(prefix) for prefix in FORCED_EVENT_PREFIXES):
        return True
    if any(keyword in upper for keyword in ('MTG:', 'HEARING', 'BRIEFING', 'INTERVIEW', 'DEFENSE', 'REVIEW', 'PRESS CONFERENCE')):
        if not any(hint in upper for hint in LOCATION_HINTS):
            return True
        words = [w for w in cleaned.split() if w]
        capitalized = sum(1 for w in words if w[0].isalpha() and w[0].isupper())
        if 1 < len(words) <= 8 and capitalized >= len(words) - 1:
            if not any(hint in upper for hint in LOCATION_HINTS):
                return True
    return False




def is_bold_word(word: dict) -> bool:
    font_name = (word.get('fontname') or '').upper()
    return 'BOLD' in font_name or 'BD' in font_name or 'BLACK' in font_name





def merge_events(lines, start_index):
    events: list[list[str]] = []
    current_segments: list[str] = []
    pending_new_event = True
    current_event_x0: float | None = None
    pending_location_segments: list[str] = []

    def flush_current() -> None:
        nonlocal current_segments, pending_new_event, current_event_x0
        if current_segments:
            events.append(current_segments.copy())
        current_segments = []
        pending_new_event = True
        current_event_x0 = None

    for line in lines[start_index:]:
        text = line['text'].strip()
        if not text:
            continue
        upper_text = text.upper()
        if upper_text in DAY_NAMES_UPPER:
            break
        if DATE_PATTERN.search(text):
            break
        if DIGITS_ONLY_PATTERN.fullmatch(text) and line['x0'] > 200:
            break

        time_words = [word for word in line['words'] if word['x0'] <= TIME_COLUMN_X_THRESHOLD]
        event_words = [word for word in line['words'] if word['x0'] > TIME_COLUMN_X_THRESHOLD]
        event_text = ' '.join(word['text'] for word in event_words).strip()
        segments = split_event_text(event_text)

        event_x0 = min((word['x0'] for word in event_words), default=None)
        line_is_bold = any(is_bold_word(word) for word in event_words[:3])

        if time_words:
            if segments and not pending_new_event and current_segments and all(is_person_segment(seg) for seg in segments):
                current_segments.extend(segments)
                if event_x0 is not None and current_event_x0 is None:
                    current_event_x0 = event_x0
                continue
            flush_current()
            if pending_location_segments:
                current_segments.extend(pending_location_segments)
                pending_location_segments.clear()
            if segments:
                current_segments.extend(segments)
                pending_new_event = False
                current_event_x0 = event_x0
            else:
                pending_new_event = True
                current_event_x0 = None
            continue

        if not segments:
            continue

        all_person = all(is_person_segment(seg) for seg in segments)
        all_location = all(is_location_segment(seg) for seg in segments)

        if all_location and current_segments:
            current_segments.extend(segments)
            if event_x0 is not None and current_event_x0 is None:
                current_event_x0 = event_x0
            continue

        if all_person and current_segments:
            current_segments.extend(segments)
            continue

        if all_location and not current_segments:
            pending_location_segments.extend(segments)
            if event_x0 is not None and current_event_x0 is None:
                current_event_x0 = event_x0
            continue

        indent_trigger = (
            event_x0 is not None
            and current_event_x0 is not None
            and event_x0 - current_event_x0 > INDENT_NEW_EVENT_THRESHOLD
        )

        primary_segment = segments[0]
        forced_new = False
        if current_segments:
            if indent_trigger and not all_person:
                forced_new = True
            elif not all_person and should_force_new_event(primary_segment, current_segments):
                forced_new = True
            elif line_is_bold and not all_person and not is_location_segment(primary_segment):
                forced_new = True

        if forced_new:
            flush_current()
            if pending_location_segments:
                current_segments.extend(pending_location_segments)
                pending_location_segments.clear()

        if pending_new_event and not current_segments:
            pending_new_event = False

        if not current_segments and pending_location_segments:
            current_segments.extend(pending_location_segments)
            pending_location_segments.clear()

        current_segments.extend(segments)
        split_index = None
        for idx, seg in enumerate(current_segments):
            if idx == 0:
                continue
            upper_seg = seg.upper()
            if upper_seg.startswith('PH:') or any(upper_seg.startswith(marker) for marker in SPECIAL_SPLIT_MARKERS):
                split_index = idx
                break
        if split_index is not None:
            trailing = current_segments[split_index:]
            current_segments = current_segments[:split_index]
            flush_current()
            current_segments.extend(trailing)
            pending_new_event = False
            if event_x0 is not None:
                current_event_x0 = event_x0

        if event_x0 is not None:
            if current_event_x0 is None:
                current_event_x0 = event_x0
            else:
                current_event_x0 = min(current_event_x0, event_x0)

    if pending_location_segments:
        if current_segments:
            current_segments.extend(pending_location_segments)
        elif events:
            events[-1].extend(pending_location_segments)
        pending_location_segments.clear()

    flush_current()
    return events



def clean_meeting_place(text: str) -> str:
    cleaned = text.strip()
    if cleaned.upper().startswith('PH:'):
        cleaned = cleaned.split('PH:', 1)[1].strip()
    return cleaned


def is_location_segment(segment: str) -> bool:
    cleaned = segment.strip()
    if not cleaned:
        return False
    upper = cleaned.upper()
    if 'TELECONFERENCE' in upper:
        return True
    if "CHIEF'S CONFERENCE" in upper:
        return False
    if 'CONFERENCE' in upper and 'ROOM' not in upper and 'CENTER' not in upper and 'CALL' not in upper and 'LINE' not in upper:
        return False
    return any(hint in upper for hint in LOCATION_HINTS)


def is_person_segment(segment: str) -> bool:
    cleaned = segment.strip()
    if not cleaned:
        return False
    upper = cleaned.upper()
    if upper.startswith('PH:'):
        return False
    if 'TELECONFERENCE' in upper:
        return False
    if is_location_segment(segment):
        return False
    if upper.startswith('MTG:'):
        return False
    if any(keyword in upper for keyword in ('STATEHOOD', 'DEFENSE', 'MEETING', 'CONFERENCE', 'CABINET', 'BRIEFING')):
        return False
    if ',' in cleaned and not segment_has_digits(cleaned):
        parts = [part.strip() for part in cleaned.split(',') if part.strip()]
        if parts and any(part[0].isalpha() for part in parts):
            return True
    if re.fullmatch(r'[A-Z][a-z]+(?: [A-Z][a-z]+){0,2}(?: [A-Z]\.)?(?: \(.+\))?', cleaned):
        return True
    return any(hint in upper for hint in PERSON_HINTS)



def format_event_segments(segments: list[str]) -> tuple[str, str, str]:
    if not segments:
        return '', '', ''

    field1 = segments[0].strip()
    rest_raw = [seg.strip() for seg in segments[1:] if seg and seg.strip()]

    def normalize_for_compare(value: str) -> str:
        return re.sub(r'\s+', ' ', value.upper()).strip(' ;')

    def should_skip_segment(value: str) -> bool:
        upper = value.upper().strip(' ;')
        field_norm = normalize_for_compare(field1)
        return upper == field_norm or field_norm.endswith(upper)

    processed_rest: list[str] = []
    for segment in rest_raw:
        if not segment:
            continue
        if segment.upper().startswith('PH:'):
            field1 = f"{field1} {segment}".strip()
            continue
        if should_skip_segment(segment):
            continue
        match = re.search(r'[A-Z][a-z]+,\s+[A-Z]', segment)
        if match and match.start() > 0:
            before = segment[: match.start()].strip(' ;')
            after = segment[match.start():].strip()
            if before:
                processed_rest.append(before)
            if after:
                processed_rest.append(after)
        else:
            processed_rest.append(segment)

    meeting_parts: list[str] = []
    person_parts: list[str] = []

    for segment in processed_rest:
        if not segment:
            continue
        if is_person_segment(segment):
            person_parts.append(segment)
        else:
            meeting_parts.append(segment)

    field_upper = field1.upper()
    if 'PH: INTERVIEW' in field_upper:
        if meeting_parts:
            person_parts = meeting_parts + person_parts
            meeting_parts = []
    elif field_upper.startswith('PH:') and not meeting_parts and person_parts:
        meeting_parts = person_parts
        person_parts = []
    elif field_upper.startswith('MTG:') and not meeting_parts and person_parts:
        meeting_parts = person_parts
        person_parts = []

    meeting_place = '; '.join(part.strip() for part in meeting_parts if part.strip())
    meeting_place = clean_meeting_place(meeting_place)

    person = '; '.join(part.strip() for part in person_parts if part.strip())

    return field1, meeting_place, person


def extract_events(pdf_path: Path):
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            words = page.extract_words(use_text_flow=True, keep_blank_chars=False, extra_attrs=['fontname'])
            if not words:
                continue
            lines = build_lines(words)

            date_obj = extract_date(lines)
            if not date_obj:
                continue

            day_idx = find_day_line_index(lines)
            if day_idx is None:
                continue

            schedule_start = day_idx + 2
            events = merge_events(lines, schedule_start)

            if not events:
                yield date_obj.isoformat(), '', '', ''
                continue

            for event_segments in events:
                event_name, meeting_place, person = format_event_segments(event_segments)
                yield date_obj.isoformat(), event_name, meeting_place, person


def main():
    pdf_path = Path(PDF_NAME)
    if not pdf_path.exists():
        raise SystemExit(f'PDF not found: {pdf_path}')

    rows = list(extract_events(pdf_path))

    output_path = Path(OUTPUT_NAME)
    with output_path.open('w', newline='', encoding='utf-8') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(['date', 'Event Name', 'Meeting Place', 'Person'])
        writer.writerows(rows)

    print(f'Wrote {len(rows)} rows to {output_path}')


if __name__ == '__main__':
    main()
