from __future__ import annotations

import itertools
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable

import pdfplumber

from .city_district_map import get_district

# =============================================================================
# CATEGORY CODES
# =============================================================================
# MHT-CET CAP rounds have two allocation levels:
#   STATE  – Maharashtra state-level seats (General + Ladies seats)
#   STAGE  – Special categories: PWD, Defence, TFWS, EWS, Orphan

STATE_CATEGORIES = [
    "GOPENS",               # General Open State
    "GSCS",                 # SC State
    "GSTS",                 # ST State
    "GVJS",                 # VJ/DT State
    "GNT1S",                # NT1 (Matang) State
    "GNT2S",                # NT2 State
    "GNT3S",                # NT3 State
    "GOBCS",                # OBC State
    "GSEBCS",               # SEBC (Maratha reservation) State
    "LOPENS",               # Ladies Open State
    "LSCS",                 # Ladies SC State
    "LSTS",                 # Ladies ST State
    "LVJS",                 # Ladies VJ State
    "LNT1S",                # Ladies NT1 State
    "LNT2S",                # Ladies NT2 State
    "LNT3S",                # Ladies NT3 State
    "LOBCS",                # Ladies OBC State
    "LSEBCS",               # Ladies SEBC State
]

STAGE_CATEGORIES = [
    "PWDOPENS",             # Persons with Disability – Open
    "PWDOBCS",              # Persons with Disability – OBC
    "PWDRSCS",              # Persons with Disability – SC
    "PWDROBCS",             # Persons with Disability – OBC (Rural) — parsed but not exported
    "DEFOPENS",             # Defence – Open
    "DEFOBCS",              # Defence – OBC
    "DEFROBCS",             # Defence – OBC (Rural)
    "DEFRNT3S",             # Defence – NT3 Rural — parsed but not exported
    "TFWS",                 # Tuition Fee Waiver Scheme
    "EWS",                  # Economically Weaker Section
    "ORPHAN",               # Orphan
]

# Categories that appear as columns in the final Excel workbook.
# All other parsed categories are tracked under ignored_categories.
EXPORTED_CATEGORY_CODES = {
    "GOPENS", "LOPENS",
    "GSCS",   "GSTS",   "GVJS",
    "GNT1S",  "GNT2S",  "GNT3S",
    "GOBCS",  "GSEBCS",
    "LSCS",   "LSTS",   "LVJS",
    "LNT1S",  "LNT2S",  "LNT3S",
    "LOBCS",  "LSEBCS",
    "TFWS",   "EWS",
    "PWDOPENS", "PWDOBCS", "PWDRSCS",
    "DEFOPENS", "DEFOBCS", "DEFROBCS",
    "ORPHAN",
}

ALL_KNOWN_CATEGORIES = tuple(dict.fromkeys(STATE_CATEGORIES + STAGE_CATEGORIES))
ALL_KNOWN_CATEGORY_SET = set(ALL_KNOWN_CATEGORIES)

# Category codes that are already fully spelled out — no suffix character to append.
COMPLETE_EXTRA_CATEGORIES = {"PWDROBCS", "DEFRNT3S", "DEFRSEBCS"}
COMPLETE_CATEGORY_CODES = ALL_KNOWN_CATEGORY_SET | COMPLETE_EXTRA_CATEGORIES

# Section suffix priority (used when the same base category appears in multiple
# sections of the same branch page, e.g. Home University vs State vs Other).
# Higher number = wins when merging duplicate keys.
#   H = Home University seats  (highest – most specific)
#   S = State-level seats
#   O = Other Than Home University seats
#   "" = bare code with no suffix (lowest)
SECTION_SUFFIX_PRIORITY: dict[str, int] = {"H": 3, "S": 2, "O": 1, "": 0}

# How many pages to peek at for round/year detection before starting the parse.
_PREVIEW_PAGE_COUNT = 3

# =============================================================================
# NOISE FILTER
# Lines that contain no useful data are skipped before anything else.
# =============================================================================

_NOISE_SUBSTRINGS = (
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

# =============================================================================
# COMPILED REGEX  (compiled once at import time for speed)
# =============================================================================

# "06006 - College of Engineering Pune, Pune"
_COLLEGE_HEADER_RE = re.compile(r"^(?P<college_code>\d{5})\s*-\s*(?P<college_name>.+)$")

# "0600624210 - Computer Science and Engineering"
_BRANCH_HEADER_RE = re.compile(r"^(?P<branch_code>\d{10})\s*-\s*(?P<branch_name>.+)$")

# "Home University : Savitribai Phule Pune University"
_HOME_UNIVERSITY_RE = re.compile(r"Home University\s*:\s*(.+)$", re.IGNORECASE)

# "(97.3737374)" – percentile values are always in parentheses in the PDF
_PERCENTILE_RE = re.compile(r"\(\s*([\d.]+)\s*\)")

# Academic year: "2025-26" or "2025/26"
_YEAR_RE = re.compile(r"\b(20\d{2})\s*[-/]\s*(\d{2,4})\b")

# CAP Round I / II / III in any mix of spacing
_ROUND_RE = re.compile(r"CAP\s*Round\s*([IVX]+)", re.IGNORECASE)

# Roman numerals used as Stage markers (I, II, III, IV, …)
_ROMAN_RE = re.compile(r"^[IVX]+$", re.IGNORECASE)

# Non-whitespace token scanner
_TOKEN_RE = re.compile(r"\S+")

# Valid category fragment: only uppercase letters and digits, starts with a letter
_CATEGORY_FRAGMENT_RE = re.compile(r"^[A-Z][A-Z0-9]*$")

# Type alias for the progress callback function
ProgressCallback = Callable[[dict[str, Any]], None]


# =============================================================================
# DATA CLASSES  (public API)
# =============================================================================

@dataclass(slots=True)
class ParsedRow:
    """
    One complete college–branch entry with all category cutoff data.

    The `data` dict is keyed by category code (e.g. "GOPENS") and contains:
        {"rank": int, "pct": float}
    Missing categories are simply absent from the dict (not None).
    """

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
    """
    Final output of a successful (or partial) parse.

    Use colleges_found / branches_found / rows_found for progress stats.
    All three branch-level properties derive from len(rows) — rows is the
    single source of truth.
    """

    round_label: str            # e.g. "CAP Round I"
    academic_year: str | None   # e.g. "2025-26"
    sheet_name: str             # Excel sheet title (max 31 chars)
    output_filename: str        # Suggested .xlsx filename
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
    """
    One category header column identified in the PDF table.

    `start` is the character offset in the raw (layout-preserved) line —
    used for positional alignment of rank values to category codes.
    """

    code: str
    start: int


class PartialParseError(RuntimeError):
    """
    Raised when parsing crashes mid-way but at least one row was already collected.

    The caller (app.py) can offer the partial result as a downloadable file
    rather than showing a blank error page to the user.
    """

    def __init__(self, message: str, partial_result: ParseResult | None = None) -> None:
        super().__init__(message)
        self.partial_result = partial_result


# =============================================================================
# PARSER STATE MACHINE
# =============================================================================

@dataclass(slots=True)
class _ParserState:
    """
    Mutable state updated line-by-line as pages are streamed through.

    Parser operates in one of three modes:
      "seeking"                  – scanning for a State Level / Stage block header
      "reading_table_headers"    – accumulating category column codes from header rows
      "reading_table_percentiles"– waiting for the percentile row after a rank row

    A new college resets college-level context.
    A new branch resets all branch-level context (type, minority, data, mode).
    """

    # Set at construction from the PDF preview (round label, year)
    round_label: str
    academic_year: str | None
    sheet_name: str
    output_filename: str

    # Accumulated output
    rows: list[ParsedRow] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    ignored_categories: set[str] = field(default_factory=set)
    colleges_seen: set[str] = field(default_factory=set)

    # Current college (persists across branches of the same college)
    current_college_code: str = ""
    current_college_name: str = ""
    current_city: str = ""

    # Current branch (reset on every new branch header)
    current_branch_code: str = ""
    current_branch_name: str = ""
    current_college_type: str = "Unknown"
    current_minority_status: str = "General"
    current_home_university: str = ""
    current_data: dict[str, dict[str, float | int]] = field(default_factory=dict)
    current_data_priority: dict[str, int] = field(default_factory=dict)

    # Table parsing bookkeeping
    mode: str = "seeking"
    table_columns: list[TableColumn] = field(default_factory=list)
    pending_rank_pairs: list[tuple[str, int]] = field(default_factory=list)
    stage_split_pending: bool = False   # True when "Stage" appears alone on a line

    # ------------------------------------------------------------------
    # Transition helpers
    # ------------------------------------------------------------------

    def reset_branch_state(self) -> None:
        """Clear everything that belongs to the current branch."""
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
        self.current_city = _extract_city(self.current_college_name)
        self.colleges_seen.add(code)

    def start_branch(self, branch_code: str, branch_name: str) -> None:
        self.reset_branch_state()
        self.current_branch_code = branch_code
        self.current_branch_name = branch_name.strip()

    def start_block(self) -> None:
        """Enter table-header-reading mode for a new State Level / Stage block."""
        self.mode = "reading_table_headers"
        self.table_columns = []
        self.pending_rank_pairs = []
        self.stage_split_pending = False

    def commit_pending_cutoffs(self, percentile_tokens: list[str]) -> None:
        """
        Pair buffered rank values with the incoming percentile strings, store
        them in current_data (respecting section priority), then reset for the
        next data block.
        """
        total_pairs = min(len(self.pending_rank_pairs), len(percentile_tokens))
        if total_pairs < len(self.pending_rank_pairs):
            self.warnings.append(
                f"Incomplete cutoff row for branch '{self.current_branch_code or 'unknown'}': "
                f"got {len(percentile_tokens)} percentile(s) for "
                f"{len(self.pending_rank_pairs)} rank(s)."
            )

        for i in range(total_pairs):
            raw_category, rank_value = self.pending_rank_pairs[i]
            normalized, priority = _normalize_category_code(raw_category)

            # Keep the entry with the highest section priority (H > S > O > bare).
            if priority < self.current_data_priority.get(normalized, -1):
                continue

            self.current_data[normalized] = {
                "rank": int(rank_value),
                "pct": float(percentile_tokens[i]),
            }
            self.current_data_priority[normalized] = priority

            if normalized not in EXPORTED_CATEGORY_CODES:
                self.ignored_categories.add(raw_category)

        self.mode = "reading_table_headers"
        self.pending_rank_pairs = []

    def save_current_row(self) -> None:
        """
        Flush the in-progress branch entry into rows and reset branch state.
        No-op if no branch is currently active.
        """
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


# =============================================================================
# LINE-LEVEL HELPERS
# =============================================================================

def _normalize_line(line: str) -> str:
    """Collapse all whitespace (including non-breaking spaces) to single spaces."""
    return re.sub(r"\s+", " ", line.replace("\xa0", " ")).strip()


def _is_noise(line: str) -> bool:
    """
    Return True if the line carries no parseable data.

    Filters out:
    - Empty lines
    - Known PDF header/footer boilerplate
    - Bare page numbers (1–3 digits)
    """
    lowered = _normalize_line(line).lower()
    if not lowered:
        return True
    if any(fragment in lowered for fragment in _NOISE_SUBSTRINGS):
        return True
    return bool(re.fullmatch(r"\d{1,3}", lowered))


def _extract_city(college_name: str) -> str:
    """
    Extract city from a college name of the form "College Name, City".
    Returns empty string if no comma is found.
    """
    parts = college_name.rsplit(",", 1)
    return parts[1].strip() if len(parts) == 2 else ""


def _extract_status_details(status_line: str) -> tuple[str, str, str]:
    """
    Parse a 'Status:' line and return (college_type, minority_status, home_university).

    Example input:
        Status: Un-Aided Linguistic Minority - Marathi Home University : Pune University
    Returns:
        ("Private (Unaided)", "Linguistic (Marathi)", "Pune University")
    """
    text = _normalize_line(status_line.replace("Status:", "", 1))

    # College type
    if "Government-Aided" in text:
        college_type = "Government-Aided"
    elif "Deemed University" in text:
        college_type = "Deemed University"
    elif "University Department" in text:
        college_type = "University Dept"
    elif "Un-Aided" in text:
        college_type = "Private (Unaided)"
    elif "Government" in text:
        college_type = "Government"
    else:
        college_type = "Unknown"

    # Minority status
    minority_status = "General"
    ling = re.search(r"Linguistic Minority - ([A-Za-z]+)", text, re.IGNORECASE)
    rel = re.search(r"Religious Minority - ([A-Za-z]+)", text, re.IGNORECASE)
    if ling:
        minority_status = f"Linguistic ({ling.group(1)})"
    elif rel:
        minority_status = f"Religious ({rel.group(1)})"

    # Home university
    hu_match = _HOME_UNIVERSITY_RE.search(status_line)
    home_university = _normalize_line(hu_match.group(1)) if hu_match else ""

    return college_type, minority_status, home_university


def _normalize_category_code(category: str) -> tuple[str, int]:
    """
    Map a raw PDF category token to its canonical export code and section priority.

    The PDF uses section suffixes H/S/O to distinguish allocation pools:
        GOPENO  → GOPENS, priority 1  (Other Than Home University)
        GOPENH  → GOPENS, priority 3  (Home University)
        GOPENS  → GOPENS, priority 2  (State Level)

    All three map to the single export column "GOPENS"; the highest-priority
    value wins when duplicates exist for the same branch.
    """
    category = category.upper()
    match = re.fullmatch(r"([GL])(.+)([HSO])", category)
    if not match:
        return category, SECTION_SUFFIX_PRIORITY[""]

    prefix, middle, suffix = match.groups()
    middle_map = {
        "OPEN": "OPEN", "SC": "SC", "ST": "ST", "VJ": "VJ",
        "NT1": "NT1", "NT2": "NT2", "NT3": "NT3",
        "OBC": "OBC", "SEBC": "SEBC",
    }
    normalized_middle = middle_map.get(middle)
    if normalized_middle is None:
        return category, SECTION_SUFFIX_PRIORITY[""]

    return f"{prefix}{normalized_middle}S", SECTION_SUFFIX_PRIORITY[suffix]


def _detect_round_and_year(
    text: str, source_name: str | None = None
) -> tuple[str, str | None, str, str]:
    """
    Detect the CAP round label and academic year from the first few pages.

    Returns (round_label, academic_year, sheet_name, output_filename).
    Falls back to generic strings if no match is found.
    """
    probe = " ".join(p for p in [text, source_name or ""] if p)
    round_match = _ROUND_RE.search(probe)
    year_match = _YEAR_RE.search(probe)

    round_label = f"CAP Round {round_match.group(1).upper()}" if round_match else "CAP Round"

    academic_year: str | None = None
    if year_match:
        start = year_match.group(1)
        end = year_match.group(2)
        academic_year = f"{start}-{end if len(end) == 2 else end[-2:]}"

    sheet_name = f"{round_label} Cutoffs"
    if round_match and academic_year:
        output_filename = (
            f"MHT_CET_{round_label.replace(' ', '_')}_Cutoffs_{academic_year}.xlsx"
        )
    else:
        output_filename = "MHT_CET_CAP_Cutoffs.xlsx"

    return round_label, academic_year, sheet_name[:31], output_filename


# =============================================================================
# TABLE PARSING HELPERS
# =============================================================================

def _tokenize_with_positions(raw_line: str) -> list[tuple[str, int, int]]:
    """Return (token, start, end) for every non-whitespace run in the line."""
    return [(m.group(), m.start(), m.end()) for m in _TOKEN_RE.finditer(raw_line)]


def _line_contains_category_header(raw_line: str) -> bool:
    """True if the line has at least one recognizable exported category token."""
    for token, _, _ in _tokenize_with_positions(raw_line):
        if token.upper() == "STAGE":
            continue
        if _normalize_category_code(token)[0] in EXPORTED_CATEGORY_CODES:
            return True
    return False


def _is_rank_data_line(line: str) -> bool:
    """
    True if the line looks like a rank data row.

    Accepted formats:
        "9196 16679 25163"          (pure digit tokens)
        "I 9196 16679 25163"        (Roman numeral followed by digits)
    """
    tokens = line.split()
    if not tokens:
        return False
    if _ROMAN_RE.fullmatch(tokens[0]):
        return any(re.fullmatch(r"\d+", t) for t in tokens[1:])
    return all(re.fullmatch(r"\d+", t) for t in tokens)


def _merge_table_header_tokens(
    existing: list[TableColumn], raw_line: str
) -> list[TableColumn]:
    """
    Absorb new category tokens from raw_line into the column list.

    Two-pass logic:
    Pass 1 – Add multi-character tokens as new columns (skipping Roman numerals,
              pure digits, and duplicates).
    Pass 2 – Attach single-character suffix tokens (H, S, O) to the nearest
              incomplete column by horizontal character distance.
    """
    merged = list(existing)
    pending_suffixes: list[tuple[str, int]] = []

    for token, start, _ in _tokenize_with_positions(raw_line):
        upper = token.upper()
        # Only process uppercase tokens
        if token != upper:
            continue
        # Skip structural tokens
        if upper == "STAGE" or _ROMAN_RE.fullmatch(upper):
            continue
        # Skip numeric or parenthesised values
        if re.fullmatch(r"\d+", upper) or _PERCENTILE_RE.fullmatch(token):
            continue
        # Must look like a category fragment
        if not _CATEGORY_FRAGMENT_RE.fullmatch(upper):
            continue
        # Single character: treat as a suffix to attach in Pass 2
        if len(upper) == 1 and merged:
            pending_suffixes.append((upper, start))
            continue
        # Avoid duplicates
        if any(c.start == start and c.code == upper for c in merged):
            continue
        merged.append(TableColumn(code=upper, start=start))

    # Pass 2: attach suffixes to nearest incomplete column
    if pending_suffixes:
        used: set[int] = set()
        for suffix, suf_start in pending_suffixes:
            candidates = [
                (i, abs(col.start - suf_start))
                for i, col in enumerate(merged)
                if i not in used and col.code not in COMPLETE_CATEGORY_CODES
            ]
            if not candidates:
                continue
            best_i = min(candidates, key=lambda x: x[1])[0]
            merged[best_i].code = f"{merged[best_i].code}{suffix}"
            used.add(best_i)

    merged.sort(key=lambda c: c.start)
    return merged


def _extract_rank_tokens(raw_line: str) -> list[tuple[int, int]]:
    """
    Return (rank_value, char_position) for every purely-numeric token
    in the line, skipping Roman numerals (Stage indicators).
    """
    return [
        (int(m.group()), m.start())
        for m in _TOKEN_RE.finditer(raw_line)
        if re.fullmatch(r"\d+", m.group()) and not _ROMAN_RE.fullmatch(m.group())
    ]


def _align_ranks_to_columns(
    columns: list[TableColumn],
    rank_tokens: list[tuple[int, int]],
) -> list[tuple[str, int]]:
    """
    Assign rank values to category columns using positional alignment.

    When rank_count == column_count  → simple 1-to-1 zip by order.
    When rank_count <  column_count  → dynamic-programming least-cost
        assignment that skips columns (categories with no admits) rather
        than misaligning values to wrong categories.

    Returns list of (category_code, rank_value) pairs.
    """
    if not columns or not rank_tokens:
        return []

    ordered = sorted(columns, key=lambda c: c.start)

    # Fast path: enough tokens to fill all columns
    if len(rank_tokens) >= len(ordered):
        return [(ordered[i].code, rank_tokens[i][0]) for i in range(len(ordered))]

    # DP alignment (fewer tokens than columns)
    n_cols = len(ordered)
    n_toks = len(rank_tokens)
    INF = float("inf")

    # dp[ci][ti] = minimum cost to assign the first ti tokens to the first ci columns
    dp = [[INF] * (n_toks + 1) for _ in range(n_cols + 1)]
    # decision[ci][ti] = ("match"|"skip", prev_ci, prev_ti)
    decision: list[list[tuple[str, int, int] | None]] = [
        [None] * (n_toks + 1) for _ in range(n_cols + 1)
    ]
    dp[0][0] = 0.0

    for ci in range(n_cols):
        remaining_cols = n_cols - ci
        for ti in range(n_toks + 1):
            cost = dp[ci][ti]
            if cost == INF:
                continue

            remaining_toks = n_toks - ti

            # Option A: skip this column (only allowed when we can still fill the rest)
            if remaining_cols > remaining_toks:
                skip_cost = cost + 0.25
                if skip_cost < dp[ci + 1][ti]:
                    dp[ci + 1][ti] = skip_cost
                    decision[ci + 1][ti] = ("skip", ci, ti)

            # Option B: assign next token to this column
            if ti < n_toks:
                match_cost = cost + abs(ordered[ci].start - rank_tokens[ti][1])
                if match_cost < dp[ci + 1][ti + 1]:
                    dp[ci + 1][ti + 1] = match_cost
                    decision[ci + 1][ti + 1] = ("match", ci, ti)

    # Traceback
    assignments: list[tuple[str, int]] = []
    ci, ti = n_cols, n_toks
    while ci > 0 or ti > 0:
        step = decision[ci][ti]
        if step is None:
            break
        action, prev_ci, prev_ti = step
        if action == "match":
            assignments.append((ordered[prev_ci].code, rank_tokens[prev_ti][0]))
        ci, ti = prev_ci, prev_ti

    assignments.reverse()
    return assignments


# =============================================================================
# PROGRESS REPORTING
# =============================================================================

def _emit_progress(
    callback: ProgressCallback | None,
    state: _ParserState,
    pages_processed: int,
    total_pages: int,
) -> None:
    """Push a progress update to the callback if one was provided."""
    if callback is None:
        return

    in_progress = 1 if state.current_branch_code else 0
    total_branches = len(state.rows) + in_progress

    callback({
        "pages_processed": pages_processed,
        "total_pages": total_pages,
        "colleges_found": len(state.colleges_seen),
        "branches_found": total_branches,
        "rows_found": total_branches,
        "current_college": state.current_college_name,
        "current_branch": state.current_branch_name,
        "message": f"Extracting branch cutoffs... ({total_branches} branches processed)",
    })


# =============================================================================
# LINE PROCESSOR  (the heart of the state machine)
# =============================================================================

def _process_line(raw_line: str, state: _ParserState) -> None:
    """
    Feed one raw PDF line into the parser state machine.

    This function is intentionally side-effect only — it mutates `state`
    and returns nothing.  The caller (parse_text_pages) is responsible for
    iterating pages and emitting progress.
    """
    # Normalise whitespace variants before anything else
    raw_line = raw_line.replace("\xa0", " ").rstrip()
    line = _normalize_line(raw_line)

    if not line or _is_noise(line):
        return

    # ------------------------------------------------------------------
    # Stage page-break handling
    # The PDF sometimes puts "Stage" alone on the last line of a page and
    # the Roman numeral (I / II / III) at the top of the next page.
    # ------------------------------------------------------------------
    if state.stage_split_pending:
        if _ROMAN_RE.fullmatch(line):
            # This is the continuation Roman numeral → start a fresh block
            state.start_block()
            return
        else:
            # Something else came after "Stage" — start the block anyway
            # then fall through so this line is also processed normally.
            state.start_block()

    # ------------------------------------------------------------------
    # Branch header  (10-digit code)
    # "0600624210 - Computer Science and Engineering"
    # ------------------------------------------------------------------
    branch_match = _BRANCH_HEADER_RE.match(line)
    if branch_match:
        state.save_current_row()
        state.start_branch(
            branch_code=branch_match.group("branch_code"),
            branch_name=branch_match.group("branch_name"),
        )
        return

    # ------------------------------------------------------------------
    # College header  (5-digit code)
    # "06006 - College of Engineering Pune, Pune"
    # ------------------------------------------------------------------
    college_match = _COLLEGE_HEADER_RE.match(line)
    if college_match and len(college_match.group("college_code")) == 5:
        state.save_current_row()
        state.start_college(
            code=college_match.group("college_code"),
            name=college_match.group("college_name"),
        )
        return

    # ------------------------------------------------------------------
    # Status line  (college type, minority, home university)
    # "Status: Government Home University : Mumbai University"
    # ------------------------------------------------------------------
    if line.startswith("Status:"):
        college_type, minority, home_uni = _extract_status_details(line)
        state.current_college_type = college_type
        state.current_minority_status = minority
        state.current_home_university = home_uni
        return

    # ------------------------------------------------------------------
    # Table block starters
    # ------------------------------------------------------------------

    # "State Level" — opens the general category table
    if re.match(r"^State\s+Level\b", line, re.IGNORECASE):
        state.start_block()
        return

    # "Stage I" — opens a stage table with Roman numeral on the same line
    if re.match(r"^Stage\s+I\b", line, re.IGNORECASE):
        state.start_block()
        return

    # "Stage GOPENS GSCS ..." — stage keyword + categories on one line
    if re.match(r"^Stage\b", line, re.IGNORECASE) and _line_contains_category_header(raw_line):
        state.start_block()
        state.table_columns = _merge_table_header_tokens(state.table_columns, raw_line)
        return

    # "Stage" alone on a line — Roman numeral follows on the next line
    if re.fullmatch(r"Stage", line, re.IGNORECASE):
        state.stage_split_pending = True
        return

    # ------------------------------------------------------------------
    # Table data parsing (mode-driven)
    # ------------------------------------------------------------------

    if state.mode == "reading_table_headers":
        # If percentiles arrive while we still have pending ranks, commit them
        if _PERCENTILE_RE.search(line) and state.pending_rank_pairs:
            state.commit_pending_cutoffs(_PERCENTILE_RE.findall(line))
            return

        # A rank row transitions us to "waiting for percentiles"
        if _is_rank_data_line(line):
            rank_tokens = _extract_rank_tokens(raw_line)
            state.pending_rank_pairs = _align_ranks_to_columns(
                state.table_columns, rank_tokens
            )
            state.mode = "reading_table_percentiles"
            return

        # Still reading headers — absorb any new category tokens
        updated = _merge_table_header_tokens(state.table_columns, raw_line)
        if updated:
            state.table_columns = updated
        return

    if state.mode == "reading_table_percentiles":
        percentile_tokens = _PERCENTILE_RE.findall(line)
        if percentile_tokens:
            state.commit_pending_cutoffs(percentile_tokens)
            return

        # Another rank row before we received percentiles — buffer it
        if _is_rank_data_line(line):
            rank_tokens = _extract_rank_tokens(raw_line)
            state.pending_rank_pairs = _align_ranks_to_columns(
                state.table_columns, rank_tokens
            )
            return


# =============================================================================
# PUBLIC PARSING FUNCTIONS
# =============================================================================

def parse_text_pages(
    page_texts: Iterable[str],
    total_pages: int = 0,
    progress_callback: ProgressCallback | None = None,
    source_name: str | None = None,
) -> ParseResult:
    """
    Parse an iterable of page-text strings into a ParseResult.

    Designed to work with both lists (for tests and small files) and generators
    (for large PDFs where you don't want all text in RAM at once).

    Args:
        page_texts:        Iterable of per-page text strings (one string per page).
        total_pages:       Total page count for progress percentage.
                           Pass 0 if unknown (progress still works but % is approximate).
        progress_callback: Called after every page with a dict of current stats.
        source_name:       PDF filename — used as fallback for round/year detection.

    Returns:
        ParseResult with all rows.

    Raises:
        PartialParseError: Crash mid-parse with ≥1 rows → caller can offer partial download.
        Any other exception if zero rows were collected before the crash.
    """
    page_iterator = iter(page_texts)

    # Peek at the first few pages for round/year detection.
    # We then chain them back so no pages are skipped during the main parse.
    preview_buffer: list[str] = []
    for _ in range(_PREVIEW_PAGE_COUNT):
        try:
            preview_buffer.append(next(page_iterator))
        except StopIteration:
            break

    round_label, academic_year, sheet_name, output_filename = _detect_round_and_year(
        "\n".join(preview_buffer), source_name=source_name
    )

    state = _ParserState(
        round_label=round_label,
        academic_year=academic_year,
        sheet_name=sheet_name,
        output_filename=output_filename,
    )

    # Chain preview pages back with the rest — nothing is re-extracted
    all_pages = itertools.chain(preview_buffer, page_iterator)

    page_index = 0
    try:
        for text in all_pages:
            page_index += 1
            for raw_line in (text or "").splitlines():
                _process_line(raw_line, state)
            # Use page_index as total fallback so percentage is never 0
            _emit_progress(progress_callback, state, page_index, total_pages or page_index)

        state.save_current_row()
        return state.to_result()

    except Exception as exc:
        # Flush whatever branch is in progress, then surface a partial result
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
    """
    Open a MHT-CET CAP round cutoff PDF and parse it into a ParseResult.

    Memory strategy — pages are extracted one at a time via a generator.
    Only one page worth of text exists in RAM at any moment.  This is critical
    for large PDFs (1000+ pages) which would otherwise exhaust the container's
    memory and crash the server.

    Args:
        pdf_path:          Path to the PDF file.
        progress_callback: Called after every page with a status dict.
        max_pages:         Limit parsing to the first N pages (useful for previews).
        source_name:       Override filename used for round/year detection.

    Returns:
        ParseResult with all parsed rows.
    """
    pdf_path = Path(pdf_path)

    with pdfplumber.open(str(pdf_path)) as pdf:
        page_slice = pdf.pages[:max_pages] if max_pages else pdf.pages
        total_pages = len(page_slice)

        def _page_generator():
            """
            Yield one page's extracted text at a time.

            layout=True preserves horizontal character positions, which is
            essential for the positional alignment of rank values to category
            column headers.  x_tolerance=1 keeps columns tight; y_tolerance=3
            groups characters on the same visual line.
            """
            for page in page_slice:
                yield page.extract_text(layout=True, x_tolerance=1, y_tolerance=3) or ""

        # IMPORTANT: parse_text_pages is called INSIDE the `with` block.
        # The generator holds a reference to the open pdfplumber PDF object.
        # Closing the PDF before the generator is fully consumed would cause
        # a crash on the second page.
        return parse_text_pages(
            _page_generator(),
            total_pages=total_pages,
            progress_callback=progress_callback,
            source_name=source_name or pdf_path.name,
        )
