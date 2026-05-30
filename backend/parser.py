from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable

import pdfplumber

from .city_district_map import get_district

STATE_CATEGORIES = [
    "GOPENS",
    "GSCS",
    "GSTS",
    "GVJS",
    "GNT1S",
    "GNT2S",
    "GNT3S",
    "GOBCS",
    "GSEBCS",
    "LOPENS",
    "LSCS",
    "LSTS",
    "LVJS",
    "LNT1S",
    "LNT2S",
    "LNT3S",
    "LOBCS",
    "LSEBCS",
]

STAGE_CATEGORIES = [
    "PWDOPENS",
    "PWDOBCS",
    "PWDRSCS",
    "PWDROBCS",
    "DEFOPENS",
    "DEFOBCS",
    "DEFROBCS",
    "DEFRNT3S",
    "TFWS",
    "EWS",
    "ORPHAN",
]

EXPORTED_CATEGORY_CODES = {
    "GOPENS",
    "LOPENS",
    "GSCS",
    "GSTS",
    "GVJS",
    "GNT1S",
    "GNT2S",
    "GNT3S",
    "GOBCS",
    "GSEBCS",
    "LSCS",
    "LSTS",
    "LVJS",
    "LNT1S",
    "LNT2S",
    "LNT3S",
    "LOBCS",
    "LSEBCS",
    "TFWS",
    "EWS",
    "PWDOPENS",
    "PWDOBCS",
    "PWDRSCS",
    "DEFOPENS",
    "DEFOBCS",
    "DEFROBCS",
    "ORPHAN",
}

ALL_KNOWN_CATEGORIES = tuple(dict.fromkeys(STATE_CATEGORIES + STAGE_CATEGORIES))
ALL_KNOWN_CATEGORY_SET = set(ALL_KNOWN_CATEGORIES)
NOISE_LINES = (
    "government of maharashtra",
    "state common entrance test cell",
    "cut off list for maharashtra",
    "degree courses in engineering",
    "legends: starting character g-general",
    "maharashtra state seats",
    "d i r",
    "figures in bracket",
    "starting character",
)
COLLEGE_HEADER_RE = re.compile(r"^(?P<college_code>\d{5})\s*-\s*(?P<college_name>.+)$")
BRANCH_HEADER_RE = re.compile(r"^(?P<branch_code>\d{10})\s*-\s*(?P<branch_name>.+)$")
HOME_UNIVERSITY_RE = re.compile(r"Home University\s*:\s*(.+)$", re.IGNORECASE)
PERCENTILE_RE = re.compile(r"\(\s*([\d.]+)\s*\)")
YEAR_RE = re.compile(r"\b(20\d{2})\s*[-/]\s*(\d{2,4})\b")
ROUND_RE = re.compile(r"CAP\s*Round\s*([IVX]+)", re.IGNORECASE)
ROMAN_TOKEN_RE = re.compile(r"^[IVX]+$", re.IGNORECASE)
TOKEN_RE = re.compile(r"\S+")
CATEGORY_FRAGMENT_RE = re.compile(r"^[A-Z][A-Z0-9]*$")
COMPLETE_EXTRA_CATEGORIES = {"PWDROBCS", "DEFRNT3S", "DEFRSEBCS"}
COMPLETE_CATEGORY_CODES = ALL_KNOWN_CATEGORY_SET | COMPLETE_EXTRA_CATEGORIES
SECTION_SUFFIX_PRIORITY = {"H": 3, "S": 2, "O": 1, "": 0}

ProgressCallback = Callable[[dict[str, Any]], None]


@dataclass(slots=True)
class ParsedRow:
    college_code: str
    college_name: str
    city: str
    district: str
    college_type: str
    minority_status: str
    home_university: str
    branch_code: str
    branch_name: str
    data: dict[str, dict[str, float | int]]


@dataclass(slots=True)
class ParseResult:
    round_label: str
    academic_year: str | None
    sheet_name: str
    output_filename: str
    rows: list[ParsedRow]
    warnings: list[str] = field(default_factory=list)
    ignored_categories: list[str] = field(default_factory=list)

    @property
    def colleges_found(self) -> int:
        return len({row.college_code for row in self.rows})

    @property
    def branches_found(self) -> int:
        return len(self.rows)

    @property
    def rows_found(self) -> int:
        return len(self.rows)


@dataclass(slots=True)
class TableColumn:
    code: str
    start: int


@dataclass(slots=True)
class ParserState:
    round_label: str
    academic_year: str | None
    sheet_name: str
    output_filename: str
    rows: list[ParsedRow] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    ignored_categories: set[str] = field(default_factory=set)
    colleges_seen: set[str] = field(default_factory=set)
    current_college_code: str = ""
    current_college_name: str = ""
    current_city: str = ""
    current_branch_code: str = ""
    current_branch_name: str = ""
    current_college_type: str = "Unknown"
    current_minority_status: str = "General"
    current_home_university: str = ""
    current_data: dict[str, dict[str, float | int]] = field(default_factory=dict)
    current_data_priority: dict[str, int] = field(default_factory=dict)
    mode: str = "seeking"
    table_columns: list[TableColumn] = field(default_factory=list)
    pending_rank_pairs: list[tuple[str, int]] = field(default_factory=list)
    stage_split_pending: bool = False

    def reset_branch_state(self) -> None:
        self.current_branch_code = ""
        self.current_branch_name = ""
        self.current_college_type = "Unknown"
        self.current_minority_status = "General"
        self.current_home_university = ""
        self.current_data = {}
        self.current_data_priority = {}
        self.mode = "seeking"
        self.table_columns = []
        self.pending_rank_pairs = []
        self.stage_split_pending = False

    def start_college(self, code: str, name: str) -> None:
        self.current_college_code = code
        self.current_college_name = name.strip()
        self.current_city = extract_city(self.current_college_name)
        self.colleges_seen.add(code)

    def start_branch(self, branch_code: str, branch_name: str) -> None:
        self.reset_branch_state()
        self.current_branch_code = branch_code
        self.current_branch_name = branch_name.strip()

    def start_block(self) -> None:
        self.mode = "reading_table_headers"
        self.table_columns = []
        self.pending_rank_pairs = []
        self.stage_split_pending = False

    def commit_pending_cutoffs(self, percentile_tokens: list[str]) -> None:
        total_pairs = min(len(self.pending_rank_pairs), len(percentile_tokens))
        if total_pairs < len(self.pending_rank_pairs):
            self.warnings.append(
                f"Incomplete cutoff row for {self.current_branch_code or 'unknown branch'} "
                f"({len(self.pending_rank_pairs)} mapped ranks, {len(percentile_tokens)} percentiles)."
            )

        for index in range(total_pairs):
            category, rank_value = self.pending_rank_pairs[index]
            normalized_category, priority = normalize_category_code(category)
            percentile_text = percentile_tokens[index]
            if priority < self.current_data_priority.get(normalized_category, -1):
                continue
            self.current_data[normalized_category] = {
                "rank": int(rank_value),
                "pct": float(percentile_text),
            }
            self.current_data_priority[normalized_category] = priority
            if normalized_category not in EXPORTED_CATEGORY_CODES:
                self.ignored_categories.add(category)

        self.mode = "reading_table_headers"
        self.pending_rank_pairs = []

    def save_current_row(self) -> None:
        if not self.current_branch_code:
            return

        self.rows.append(
            ParsedRow(
                college_code=self.current_college_code,
                college_name=self.current_college_name,
                city=self.current_city,
                district=get_district(self.current_city),
                college_type=self.current_college_type,
                minority_status=self.current_minority_status,
                home_university=self.current_home_university,
                branch_code=self.current_branch_code,
                branch_name=self.current_branch_name,
                data=dict(self.current_data),
            )
        )
        self.reset_branch_state()

    def to_result(self) -> ParseResult:
        return ParseResult(
            round_label=self.round_label,
            academic_year=self.academic_year,
            sheet_name=self.sheet_name,
            output_filename=self.output_filename,
            rows=list(self.rows),
            warnings=list(self.warnings),
            ignored_categories=sorted(self.ignored_categories),
        )


class PartialParseError(RuntimeError):
    def __init__(self, message: str, partial_result: ParseResult | None = None):
        super().__init__(message)
        self.partial_result = partial_result


def normalize_line(line: str) -> str:
    return re.sub(r"\s+", " ", line.replace("\xa0", " ")).strip()


def is_noise(line: str) -> bool:
    lowered = normalize_line(line).lower()
    if not lowered:
        return True
    if any(noise in lowered for noise in NOISE_LINES):
        return True
    return bool(re.fullmatch(r"\d{1,3}", lowered))


def extract_city(college_name: str) -> str:
    parts = college_name.rsplit(",", 1)
    return parts[1].strip() if len(parts) == 2 else ""


def extract_status_details(status_line: str) -> tuple[str, str, str]:
    status_text = normalize_line(status_line.replace("Status:", "", 1))

    if "Government-Aided" in status_text:
        college_type = "Government-Aided"
    elif "Deemed University" in status_text:
        college_type = "Deemed University"
    elif "University Department" in status_text:
        college_type = "University Dept"
    elif "Un-Aided" in status_text:
        college_type = "Private (Unaided)"
    elif "Government" in status_text:
        college_type = "Government"
    else:
        college_type = "Unknown"

    minority_status = "General"
    linguistic_match = re.search(r"Linguistic Minority - ([A-Za-z]+)", status_text, re.IGNORECASE)
    religious_match = re.search(r"Religious Minority - ([A-Za-z]+)", status_text, re.IGNORECASE)
    if linguistic_match:
        minority_status = f"Linguistic ({linguistic_match.group(1)})"
    elif religious_match:
        minority_status = f"Religious ({religious_match.group(1)})"

    home_university_match = HOME_UNIVERSITY_RE.search(status_line)
    home_university = normalize_line(home_university_match.group(1)) if home_university_match else ""
    return college_type, minority_status, home_university


def normalize_category_code(category: str) -> tuple[str, int]:
    category = category.upper()
    match = re.fullmatch(r"([GL])(.+)([HSO])", category)
    if not match:
        return category, SECTION_SUFFIX_PRIORITY[""]

    prefix, middle, suffix = match.groups()
    normalized_middle = {
        "OPEN": "OPEN",
        "SC": "SC",
        "ST": "ST",
        "VJ": "VJ",
        "NT1": "NT1",
        "NT2": "NT2",
        "NT3": "NT3",
        "OBC": "OBC",
        "SEBC": "SEBC",
    }.get(middle)
    if normalized_middle is None:
        return category, SECTION_SUFFIX_PRIORITY[""]

    return f"{prefix}{normalized_middle}S", SECTION_SUFFIX_PRIORITY[suffix]


def detect_round_and_year(text: str, source_name: str | None = None) -> tuple[str, str | None, str, str]:
    probe_text = " ".join(part for part in [text, source_name or ""] if part)
    round_match = ROUND_RE.search(probe_text)
    year_match = YEAR_RE.search(probe_text)

    round_label = f"CAP Round {round_match.group(1).upper()}" if round_match else "CAP Round"
    academic_year = None
    if year_match:
        start_year = year_match.group(1)
        end_fragment = year_match.group(2)
        academic_year = f"{start_year}-{end_fragment if len(end_fragment) == 2 else end_fragment[-2:]}"

    sheet_name = f"{round_label} Cutoffs"
    if round_match and academic_year:
        output_filename = f"MHT_CET_{round_label.replace(' ', '_')}_Cutoffs_{academic_year}.xlsx"
    else:
        output_filename = "MHT_CET_CAP_Cutoffs.xlsx"
    return round_label, academic_year, sheet_name[:31], output_filename


def _extract_category_tokens(line: str) -> list[str]:
    return [token for token in line.upper().split() if token in ALL_KNOWN_CATEGORY_SET]


def _line_contains_category_header(raw_line: str) -> bool:
    for token, _, _ in _tokenize_with_positions(raw_line):
        if token.upper() == "STAGE":
            continue
        if normalize_category_code(token)[0] in EXPORTED_CATEGORY_CODES:
            return True
    return False


def _is_rank_line(line: str) -> bool:
    tokens = line.split()
    return bool(tokens) and all(re.fullmatch(r"\d+", token) for token in tokens)


def _tokenize_with_positions(raw_line: str) -> list[tuple[str, int, int]]:
    return [(match.group(), match.start(), match.end()) for match in TOKEN_RE.finditer(raw_line)]


def _is_stage_value_line(raw_line: str) -> bool:
    tokens = raw_line.split()
    if not tokens:
        return False
    if ROMAN_TOKEN_RE.fullmatch(tokens[0]):
        return any(re.fullmatch(r"\d+", token) for token in tokens[1:])
    return all(re.fullmatch(r"\d+", token) for token in tokens)


def _merge_table_header_tokens(columns: list[TableColumn], raw_line: str) -> list[TableColumn]:
    merged = list(columns)
    suffix_tokens: list[tuple[str, int]] = []

    for token, start, _ in _tokenize_with_positions(raw_line):
        if token != token.upper():
            continue
        upper = token.upper()
        if upper == "STAGE" or ROMAN_TOKEN_RE.fullmatch(upper):
            continue
        if re.fullmatch(r"\d+", upper) or PERCENTILE_RE.fullmatch(token):
            continue
        if not CATEGORY_FRAGMENT_RE.fullmatch(upper):
            continue
        if len(upper) == 1 and merged:
            suffix_tokens.append((upper, start))
            continue
        if any(column.start == start and column.code == upper for column in merged):
            continue
        merged.append(TableColumn(code=upper, start=start))

    if suffix_tokens:
        used_indexes: set[int] = set()
        for suffix, start in suffix_tokens:
            candidates = [
                (index, abs(column.start - start))
                for index, column in enumerate(merged)
                if index not in used_indexes and column.code not in COMPLETE_CATEGORY_CODES
            ]
            if not candidates:
                continue
            index = min(candidates, key=lambda item: item[1])[0]
            merged[index].code = f"{merged[index].code}{suffix}"
            used_indexes.add(index)

    merged.sort(key=lambda column: column.start)
    return merged


def _extract_rank_tokens(raw_line: str) -> list[tuple[int, int]]:
    rank_tokens: list[tuple[int, int]] = []
    for token, start, _ in _tokenize_with_positions(raw_line):
        if ROMAN_TOKEN_RE.fullmatch(token):
            continue
        if re.fullmatch(r"\d+", token):
            rank_tokens.append((int(token), start))
    return rank_tokens


def _align_rank_tokens_to_columns(
    columns: list[TableColumn],
    rank_tokens: list[tuple[int, int]],
) -> list[tuple[str, int]]:
    if not columns or not rank_tokens:
        return []

    ordered_columns = sorted(columns, key=lambda column: column.start)
    if len(rank_tokens) >= len(ordered_columns):
        return [
            (ordered_columns[index].code, rank_tokens[index][0])
            for index in range(min(len(ordered_columns), len(rank_tokens)))
        ]

    column_count = len(ordered_columns)
    token_count = len(rank_tokens)
    infinity = float("inf")
    dp = [[infinity] * (token_count + 1) for _ in range(column_count + 1)]
    decision: list[list[tuple[str, int, int] | None]] = [
        [None] * (token_count + 1) for _ in range(column_count + 1)
    ]
    dp[0][0] = 0.0

    for column_index in range(column_count):
        remaining_columns = column_count - column_index
        for token_index in range(token_count + 1):
            current_cost = dp[column_index][token_index]
            if current_cost == infinity:
                continue

            remaining_tokens = token_count - token_index
            if remaining_columns > remaining_tokens:
                skip_cost = current_cost + 0.25
                if skip_cost < dp[column_index + 1][token_index]:
                    dp[column_index + 1][token_index] = skip_cost
                    decision[column_index + 1][token_index] = ("skip", column_index, token_index)

            if token_index < token_count:
                match_cost = current_cost + abs(
                    ordered_columns[column_index].start - rank_tokens[token_index][1]
                )
                if match_cost < dp[column_index + 1][token_index + 1]:
                    dp[column_index + 1][token_index + 1] = match_cost
                    decision[column_index + 1][token_index + 1] = ("match", column_index, token_index)

    assignments: list[tuple[str, int]] = []
    column_index = column_count
    token_index = token_count
    while column_index > 0 or token_index > 0:
        step = decision[column_index][token_index]
        if step is None:
            break
        action, previous_column, previous_token = step
        if action == "match":
            assignments.append(
                (ordered_columns[previous_column].code, rank_tokens[previous_token][0])
            )
        column_index = previous_column
        token_index = previous_token

    assignments.reverse()
    return assignments


def _emit_progress(
    callback: ProgressCallback | None,
    state: ParserState,
    pages_processed: int,
    total_pages: int,
) -> None:
    if callback is None:
        return

    callback(
        {
            "pages_processed": pages_processed,
            "total_pages": total_pages,
            "colleges_found": len(state.colleges_seen),
            "branches_found": len(state.rows) + (1 if state.current_branch_code else 0),
            "rows_found": len(state.rows) + (1 if state.current_branch_code else 0),
            "current_college": state.current_college_name,
            "current_branch": state.current_branch_name,
            "message": (
                f"Extracting branch cutoffs... ({len(state.rows) + (1 if state.current_branch_code else 0)} "
                f"branches processed)"
            ),
        }
    )


def parse_text_pages(
    page_texts: Iterable[str],
    progress_callback: ProgressCallback | None = None,
    source_name: str | None = None,
) -> ParseResult:
    texts = list(page_texts)
    preview_text = "\n".join(texts[:3])
    round_label, academic_year, sheet_name, output_filename = detect_round_and_year(
        preview_text,
        source_name=source_name,
    )
    state = ParserState(
        round_label=round_label,
        academic_year=academic_year,
        sheet_name=sheet_name,
        output_filename=output_filename,
    )
    total_pages = len(texts)

    try:
        for page_index, text in enumerate(texts, start=1):
            for raw_line in (text or "").splitlines():
                raw_line = raw_line.replace("\xa0", " ").rstrip()
                line = normalize_line(raw_line)
                if not line or is_noise(line):
                    continue

                if state.stage_split_pending and re.fullmatch(r"[IVX]+", line, re.IGNORECASE):
                    state.start_block()
                    continue

                if state.stage_split_pending and not re.fullmatch(r"[IVX]+", line, re.IGNORECASE):
                    state.start_block()

                branch_match = BRANCH_HEADER_RE.match(line)
                if branch_match:
                    state.save_current_row()
                    state.start_branch(
                        branch_code=branch_match.group("branch_code"),
                        branch_name=branch_match.group("branch_name"),
                    )
                    continue

                college_match = COLLEGE_HEADER_RE.match(line)
                if college_match and len(college_match.group("college_code")) == 5:
                    state.save_current_row()
                    state.start_college(
                        code=college_match.group("college_code"),
                        name=college_match.group("college_name"),
                    )
                    continue

                if line.startswith("Status:"):
                    college_type, minority_status, home_university = extract_status_details(line)
                    state.current_college_type = college_type
                    state.current_minority_status = minority_status
                    state.current_home_university = home_university
                    continue

                if re.match(r"^State\s+Level\b", line, re.IGNORECASE):
                    state.start_block()
                    continue

                if re.match(r"^Stage\s+I\b", line, re.IGNORECASE):
                    state.start_block()
                    continue

                if re.match(r"^Stage\b", line, re.IGNORECASE) and _line_contains_category_header(raw_line):
                    state.start_block()
                    state.table_columns = _merge_table_header_tokens(state.table_columns, raw_line)
                    continue

                if re.fullmatch(r"Stage", line, re.IGNORECASE):
                    state.stage_split_pending = True
                    continue

                if state.mode == "reading_table_headers":
                    if PERCENTILE_RE.search(line) and state.pending_rank_pairs:
                        percentile_tokens = PERCENTILE_RE.findall(line)
                        state.commit_pending_cutoffs(percentile_tokens)
                        continue

                    if _is_stage_value_line(line):
                        rank_tokens = _extract_rank_tokens(raw_line)
                        state.pending_rank_pairs = _align_rank_tokens_to_columns(
                            state.table_columns,
                            rank_tokens,
                        )
                        state.mode = "reading_table_percentiles"
                        continue

                    updated_columns = _merge_table_header_tokens(state.table_columns, raw_line)
                    if updated_columns:
                        state.table_columns = updated_columns
                        continue

                if state.mode == "reading_table_percentiles":
                    percentile_tokens = PERCENTILE_RE.findall(line)
                    if percentile_tokens:
                        state.commit_pending_cutoffs(percentile_tokens)
                        continue

                    if _is_stage_value_line(line):
                        rank_tokens = _extract_rank_tokens(raw_line)
                        state.pending_rank_pairs = _align_rank_tokens_to_columns(
                            state.table_columns,
                            rank_tokens,
                        )
                        continue

            _emit_progress(progress_callback, state, page_index, total_pages)

        state.save_current_row()
        return state.to_result()
    except Exception as exc:
        state.save_current_row()
        partial = state.to_result()
        if partial.rows:
            raise PartialParseError(str(exc), partial_result=partial) from exc
        raise


def parse_pdf(
    pdf_path: str | Path,
    progress_callback: ProgressCallback | None = None,
    max_pages: int | None = None,
    source_name: str | None = None,
) -> ParseResult:
    pdf_path = Path(pdf_path)
    with pdfplumber.open(str(pdf_path)) as pdf:
        pages = pdf.pages[: max_pages or len(pdf.pages)]
        page_texts = [
            page.extract_text(layout=True, x_tolerance=1, y_tolerance=3) or ""
            for page in pages
        ]
    return parse_text_pages(
        page_texts,
        progress_callback=progress_callback,
        source_name=source_name or pdf_path.name,
    )
