import io
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st


# =====================================================
# Little League Scheduler
# =====================================================

st.set_page_config(
    page_title="Little League Scheduler",
    layout="wide",
    initial_sidebar_state="collapsed",
)

DIVISION_SHEETS = ["T-ball", "5-6", "7-8", "9-10", "11-12"]

ALL_FIELDS = ["Field 1", "Field 2", "Field 3", "Field 4", "Field 5"]

DEFAULT_FIELDS_BY_DIVISION = {
    "T-ball": ["Field 5"],
    "5-6": ["Field 4", "Field 5"],
    "7-8": ["Field 3"],
    "9-10": ["Field 2", "Field 1"],
    "11-12": ["Field 1"],
}

WEEKDAY_NAME_TO_INT = {
    "Monday": 0,
    "Tuesday": 1,
    "Wednesday": 2,
    "Thursday": 3,
    "Friday": 4,
    "Saturday": 5,
    "Sunday": 6,
}
INT_TO_WEEKDAY_NAME = {v: k for k, v in WEEKDAY_NAME_TO_INT.items()}

TEAM_GRID_COLUMNS = ["Division", "Team", "TeamColor", "CoachName"]

# No Friday games
WEEKDAY_OPTIONS = ["Monday", "Tuesday", "Wednesday", "Thursday"]

# Fixed slot rules
WEEKDAY_SLOTS = ["17:00", "18:30"]  # 1.5 hours
SATURDAY_BASE_SLOTS = ["09:30", "11:30", "13:30"]  # 2 hours
SATURDAY_OVERFLOW_SLOT = "17:30"  # use only if needed

WEEKDAY_GAME_LENGTH_MINUTES = 90
SATURDAY_GAME_LENGTH_MINUTES = 120

DIVISION_NIGHT_GROUP_OPTIONS = ["None", "M/W", "T/Th"]
DIVISION_NIGHT_GROUP_MAP = {
    "None": {0, 1, 2, 3},
    "M/W": {0, 2},
    "T/Th": {1, 3},
}

# Brand palette
VEGAS_GOLD = "#C5B358"
DEEP_GOLD = "#A48F3A"
SOFT_GOLD = "#E9DFB0"
CHARCOAL = "#111111"
OFF_BLACK = "#1C1C1C"
BORDER = "#D9D4C2"
PAPER = "#FAF8F2"
WHITE = "#FFFFFF"
MUTED = "#6B7280"


@dataclass
class TeamRecord:
    division: str
    team: str
    team_color: str
    coach_name: str
    preferred_weekdays: List[int]  # Saturday added automatically


@dataclass
class GameRecord:
    division: str
    home_team: str
    away_team: str
    game_date: date
    start_time: time
    end_time: time
    field: str
    slot_label: str
    duration_minutes: int

    def to_dict(self):
        return {
            "Division": self.division,
            "Home Team": self.home_team,
            "Away Team": self.away_team,
            "Date": self.game_date.isoformat(),
            "Start": self.start_time.strftime("%H:%M"),
            "End": self.end_time.strftime("%H:%M"),
            "Field": self.field,
            "Slot": self.slot_label,
            "DurationMinutes": self.duration_minutes,
            "Matchup": f"{self.away_team} @ {self.home_team}",
        }


# -------------------------
# Styling
# -------------------------
def inject_app_css():
    st.markdown(
        f"""
        <style>
        .stApp {{
            background: linear-gradient(180deg, {PAPER} 0%, #F5F3EB 100%);
        }}

        .block-container {{
            padding-top: 1.3rem;
            padding-bottom: 2rem;
            max-width: 1400px;
        }}

        h1, h2, h3, h4 {{
            color: {OFF_BLACK};
            letter-spacing: -0.02em;
        }}

        .hero-card {{
            background: linear-gradient(135deg, {OFF_BLACK} 0%, #262626 100%);
            border: 1px solid {DEEP_GOLD};
            border-radius: 20px;
            padding: 28px 28px 22px 28px;
            margin-bottom: 18px;
            box-shadow: 0 14px 30px rgba(0,0,0,0.12);
        }}

        .hero-title {{
            color: {WHITE} !important;
            font-size: 2rem;
            font-weight: 800;
            margin: 0 0 6px 0;
        }}

        .hero-sub {{
            color: #E8E5D8 !important;
            font-size: 1rem;
            margin: 0;
        }}

        .gold-chip {{
            display: inline-block;
            background: {SOFT_GOLD};
            color: {OFF_BLACK} !important;
            border: 1px solid {DEEP_GOLD};
            border-radius: 999px;
            padding: 4px 10px;
            font-size: 0.78rem;
            font-weight: 700;
            margin-right: 6px;
            margin-bottom: 6px;
        }}

        .section-card {{
            background: {WHITE};
            border: 1px solid {BORDER};
            border-radius: 18px;
            padding: 18px 18px 14px 18px;
            box-shadow: 0 8px 24px rgba(17,17,17,0.06);
            margin-bottom: 16px;
        }}

        .section-title {{
            font-size: 1.05rem;
            font-weight: 800;
            color: {OFF_BLACK} !important;
            margin-bottom: 6px;
        }}

        .section-sub {{
            color: {OFF_BLACK} !important;
            font-size: 0.92rem;
            margin-bottom: 10px;
        }}

        .mini-metric {{
            background: {WHITE};
            border: 1px solid {BORDER};
            border-left: 6px solid {VEGAS_GOLD};
            border-radius: 14px;
            padding: 14px 16px;
            box-shadow: 0 6px 18px rgba(17,17,17,0.05);
            margin-bottom: 10px;
        }}

        .mini-metric-label {{
            color: {OFF_BLACK} !important;
            font-size: 0.82rem;
            text-transform: uppercase;
            letter-spacing: 0.04em;
            font-weight: 700;
            margin-bottom: 4px;
        }}

        .mini-metric-value {{
            color: {OFF_BLACK} !important;
            font-size: 1.15rem;
            font-weight: 800;
            margin: 0;
        }}

        .scheduler-note {{
            background: #FFF9E8;
            border: 1px solid #EAD9A0;
            border-left: 6px solid {VEGAS_GOLD};
            border-radius: 14px;
            padding: 12px 14px;
            color: {OFF_BLACK} !important;
            font-size: 0.92rem;
            margin-top: 6px;
            margin-bottom: 12px;
        }}

        .scheduler-note * {{
            color: {OFF_BLACK} !important;
        }}

        div.stButton > button {{
            border-radius: 12px !important;
            border: 1px solid {DEEP_GOLD} !important;
            background: linear-gradient(180deg, {VEGAS_GOLD} 0%, {DEEP_GOLD} 100%) !important;
            color: {WHITE} !important;
            font-weight: 700 !important;
            box-shadow: 0 6px 14px rgba(0,0,0,0.10);
        }}

        div.stDownloadButton > button {{
            border-radius: 12px !important;
            border: 1px solid {DEEP_GOLD} !important;
            background: {WHITE} !important;
            color: {OFF_BLACK} !important;
            font-weight: 700 !important;
        }}

        .stTabs [data-baseweb="tab-list"] {{
            gap: 8px;
        }}

        .stTabs [data-baseweb="tab"] {{
            background: {WHITE};
            border-radius: 12px 12px 0 0;
            border: 1px solid {BORDER};
            padding-left: 16px;
            padding-right: 16px;
            color: {OFF_BLACK} !important;
        }}

        .stTabs [data-baseweb="tab"] * {{
            color: {OFF_BLACK} !important;
        }}

        .stTabs [aria-selected="true"] {{
            background: {SOFT_GOLD} !important;
            color: {OFF_BLACK} !important;
            border-bottom-color: {SOFT_GOLD} !important;
        }}

        .stTabs [aria-selected="true"] * {{
            color: {OFF_BLACK} !important;
        }}

        .stExpander {{
            border-radius: 14px !important;
            border: 1px solid {BORDER} !important;
            background: {WHITE} !important;
        }}

        .stExpander summary,
        .stExpander summary * {{
            color: {OFF_BLACK} !important;
        }}

        .stMarkdown,
        .stText,
        .stCaption,
        .stSelectbox label,
        .stMultiSelect label,
        .stDateInput label,
        .stTextArea label,
        .stNumberInput label,
        .stRadio label,
        .stCheckbox label,
        .stFileUploader label {{
            color: {OFF_BLACK} !important;
        }}

        div[data-testid="stWidgetLabel"] *,
        label[data-testid="stWidgetLabel"] * {{
            color: {OFF_BLACK} !important;
        }}

        div[data-baseweb="select"] > div,
        div[data-baseweb="input"] > div,
        .stDateInput > div > div,
        .stTextArea textarea {{
            border-radius: 12px !important;
        }}

        .stDataFrame, .stTable {{
            border-radius: 16px;
            overflow: hidden;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_hero():
    st.markdown(
        f"""
        <div class="hero-card">
            <div class="hero-title">⚾ Little League Schedule Builder</div>
            <p class="hero-sub">
                Build division schedules, control fields and weekday rules, balance home/away,
                review conflicts, and manage overflow games with a polished calendar view.
            </p>
            <div style="margin-top:14px;">
                <span class="gold-chip">Vegas Gold Theme</span>
                <span class="gold-chip">Saturday-First Scheduling</span>
                <span class="gold-chip">Night Group Rules</span>
                <span class="gold-chip">Manual Placement</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_open(title: str, subtitle: str = ""):
    st.markdown(
        f"""
        <div class="section-card">
            <div class="section-title">{title}</div>
            {"<div class='section-sub'>" + subtitle + "</div>" if subtitle else ""}
        """,
        unsafe_allow_html=True,
    )


def section_close():
    st.markdown("</div>", unsafe_allow_html=True)


def mini_metric(label: str, value: str):
    st.markdown(
        f"""
        <div class="mini-metric">
            <div class="mini-metric-label">{label}</div>
            <p class="mini-metric-value">{value}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


inject_app_css()


# -------------------------
# Session state init
# -------------------------
if "master_team_df" not in st.session_state:
    st.session_state.master_team_df = pd.DataFrame(columns=TEAM_GRID_COLUMNS)
if "team_config_df" not in st.session_state:
    st.session_state.team_config_df = pd.DataFrame()
if "schedule_df" not in st.session_state:
    st.session_state.schedule_df = pd.DataFrame()
if "unscheduled_df" not in st.session_state:
    st.session_state.unscheduled_df = pd.DataFrame()
if "blockout_dates" not in st.session_state:
    st.session_state.blockout_dates = []
if "loaded_file_name" not in st.session_state:
    st.session_state.loaded_file_name = None
if "loaded_file_size" not in st.session_state:
    st.session_state.loaded_file_size = None
if "max_games_per_week_by_division" not in st.session_state:
    st.session_state.max_games_per_week_by_division = {}
if "global_allowed_weekdays" not in st.session_state:
    st.session_state.global_allowed_weekdays = [0, 1, 2, 3, 5]
if "season_start" not in st.session_state:
    st.session_state.season_start = None
if "season_end" not in st.session_state:
    st.session_state.season_end = None
if "division_night_groups" not in st.session_state:
    st.session_state.division_night_groups = {division: "None" for division in DIVISION_SHEETS}
if "division_fields" not in st.session_state:
    st.session_state.division_fields = DEFAULT_FIELDS_BY_DIVISION.copy()


# -------------------------
# Utility helpers
# -------------------------
def safe_strip(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def serialize_list(values: List[str]) -> str:
    return ", ".join(values)


def deserialize_list(value) -> List[str]:
    if value is None:
        return []
    try:
        if pd.isna(value):
            return []
    except Exception:
        pass
    value = str(value).strip()
    if not value:
        return []
    return [x.strip() for x in value.split(",") if x.strip()]


def get_division_weekday_set(division: str, division_night_groups: Dict[str, str]) -> set:
    group = division_night_groups.get(division, "None")
    return set(DIVISION_NIGHT_GROUP_MAP.get(group, DIVISION_NIGHT_GROUP_MAP["None"]))


@st.cache_data(show_spinner=False)
def load_division_workbook(file_bytes: bytes) -> Tuple[Dict[str, pd.DataFrame], Dict[str, List[str]], List[str]]:
    workbook = pd.ExcelFile(io.BytesIO(file_bytes))
    found_sheets = workbook.sheet_names
    division_frames: Dict[str, pd.DataFrame] = {}
    detected_columns: Dict[str, List[str]] = {}
    warnings: List[str] = []

    for sheet in found_sheets:
        df = pd.read_excel(workbook, sheet_name=sheet)
        detected_columns[sheet] = df.columns.tolist()
        if sheet in DIVISION_SHEETS:
            division_frames[sheet] = df

    missing = [sheet for sheet in DIVISION_SHEETS if sheet not in division_frames]
    if missing:
        warnings.append(f"These division sheets were not found: {', '.join(missing)}")

    return division_frames, detected_columns, warnings


def validate_division_frames(division_frames: Dict[str, pd.DataFrame]) -> List[str]:
    errors = []
    required = ["Team", "TeamColor", "CoachName"]

    if not division_frames:
        errors.append("No valid division sheets were found in the workbook.")
        return errors

    for division, df in division_frames.items():
        for col in required:
            if col not in df.columns:
                errors.append(f"Sheet '{division}' is missing required column: {col}")
    return errors


@st.cache_data(show_spinner=False)
def create_sample_workbook_bytes() -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for division in DIVISION_SHEETS:
            sample_df = pd.DataFrame(
                [
                    {"Team": f"{division} Team 1", "TeamColor": "Blue", "CoachName": "Coach A"},
                    {"Team": f"{division} Team 2", "TeamColor": "Red", "CoachName": "Coach B"},
                    {"Team": f"{division} Team 3", "TeamColor": "Green", "CoachName": "Coach C"},
                ]
            )
            sample_df.to_excel(writer, sheet_name=division, index=False)
    output.seek(0)
    return output.getvalue()


def build_master_team_df(division_frames: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    rows = []
    for division, df in division_frames.items():
        for _, row in df.iterrows():
            team = safe_strip(row.get("Team"))
            team_color = safe_strip(row.get("TeamColor"))
            coach_name = safe_strip(row.get("CoachName"))
            if team:
                rows.append(
                    {
                        "Division": division,
                        "Team": team,
                        "TeamColor": team_color,
                        "CoachName": coach_name,
                    }
                )
    return pd.DataFrame(rows, columns=TEAM_GRID_COLUMNS)


def normalize_team_config_df(master_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in master_df.iterrows():
        division = row["Division"]
        rows.append(
            {
                "Division": division,
                "Team": row["Team"],
                "TeamColor": row["TeamColor"],
                "CoachName": row["CoachName"],
                "PreferredWeekdays": serialize_list(WEEKDAY_OPTIONS),
            }
        )
    return pd.DataFrame(rows)


def parse_blockout_tokens(tokens: List[str], year: int) -> Tuple[set, List[str]]:
    blocked = set()
    bad_tokens = []
    for token in tokens:
        value = token.strip()
        if not value:
            continue
        parsed = None
        for fmt in ["%m-%d", "%m/%d", "%B %d", "%b %d"]:
            try:
                parsed = datetime.strptime(value, fmt)
                blocked.add(date(year, parsed.month, parsed.day))
                break
            except ValueError:
                continue
        if parsed is None:
            bad_tokens.append(value)
    return blocked, bad_tokens


def daterange(start_date: date, end_date: date):
    current = start_date
    while current <= end_date:
        yield current
        current += timedelta(days=1)


def build_slot(start_label: str, duration_minutes: int) -> Tuple[str, time, time, int]:
    start_dt = datetime.strptime(start_label, "%H:%M")
    end_dt = start_dt + timedelta(minutes=duration_minutes)
    return start_label, start_dt.time(), end_dt.time(), duration_minutes


def get_day_slot_defs(current_date: date, use_overflow_saturday_slot: bool) -> List[Tuple[str, time, time, int]]:
    if current_date.weekday() in [0, 1, 2, 3]:
        return [build_slot(slot, WEEKDAY_GAME_LENGTH_MINUTES) for slot in WEEKDAY_SLOTS]
    if current_date.weekday() == 5:
        slots = [build_slot(slot, SATURDAY_GAME_LENGTH_MINUTES) for slot in SATURDAY_BASE_SLOTS]
        if use_overflow_saturday_slot:
            slots.append(build_slot(SATURDAY_OVERFLOW_SLOT, SATURDAY_GAME_LENGTH_MINUTES))
        return slots
    return []


def is_weekday(current_date: date) -> bool:
    return current_date.weekday() in [0, 1, 2, 3]


def is_saturday(current_date: date) -> bool:
    return current_date.weekday() == 5


def build_team_lookup(teams: List[TeamRecord]) -> Dict[str, TeamRecord]:
    return {f"{team.division}::{team.team}": team for team in teams}


def allowed_weekdays_for_matchup(
    team_lookup: Dict[str, TeamRecord],
    division: str,
    team_a: str,
    team_b: str,
    division_night_groups: Dict[str, str],
) -> set:
    team_a_days = set(team_lookup[f"{division}::{team_a}"].preferred_weekdays)
    team_b_days = set(team_lookup[f"{division}::{team_b}"].preferred_weekdays)
    shared = team_a_days.intersection(team_b_days)

    division_weekday_set = get_division_weekday_set(division, division_night_groups)
    shared = shared.intersection(division_weekday_set)

    shared.add(WEEKDAY_NAME_TO_INT["Saturday"])
    return shared


def eligible_fields_for_matchup(division: str, division_fields: Dict[str, List[str]]) -> List[str]:
    return division_fields.get(division, [])


def generate_round_robin_pairings(team_names: List[str], games_per_team: int) -> List[Tuple[str, str]]:
    teams = list(team_names)
    if len(teams) < 2 or games_per_team <= 0:
        return []

    pairings = []
    game_count = {team: 0 for team in teams}
    matchup_counts = {}
    safety_cap = max(2000, len(teams) * games_per_team * 10)
    loops = 0

    while min(game_count.values()) < games_per_team and loops < safety_cap:
        loops += 1
        sorted_teams = sorted(teams, key=lambda x: (game_count[x], x))
        used_this_pass = set()
        made_pair = False

        for i, team_a in enumerate(sorted_teams):
            if team_a in used_this_pass or game_count[team_a] >= games_per_team:
                continue

            candidates = []
            for team_b in sorted_teams[i + 1:]:
                if team_b in used_this_pass or team_b == team_a or game_count[team_b] >= games_per_team:
                    continue
                matchup_key = tuple(sorted((team_a, team_b)))
                candidates.append((matchup_counts.get(matchup_key, 0), game_count[team_b], team_b))

            if not candidates:
                continue

            candidates.sort(key=lambda x: (x[0], x[1], x[2]))
            _, _, team_b = candidates[0]
            pairings.append((team_a, team_b))
            game_count[team_a] += 1
            game_count[team_b] += 1
            matchup_key = tuple(sorted((team_a, team_b)))
            matchup_counts[matchup_key] = matchup_counts.get(matchup_key, 0) + 1
            used_this_pass.add(team_a)
            used_this_pass.add(team_b)
            made_pair = True

        if not made_pair:
            break

    return pairings


def parse_team_records_from_config(config_df: pd.DataFrame) -> List[TeamRecord]:
    team_records = []
    for _, row in config_df.iterrows():
        division = safe_strip(row.get("Division"))
        team = safe_strip(row.get("Team"))
        if not division or not team:
            continue

        preferred_weekday_names = deserialize_list(row.get("PreferredWeekdays"))
        preferred_weekdays = []
        for day_name in preferred_weekday_names:
            if day_name in WEEKDAY_NAME_TO_INT and day_name != "Saturday":
                preferred_weekdays.append(WEEKDAY_NAME_TO_INT[day_name])

        team_records.append(
            TeamRecord(
                division=division,
                team=team,
                team_color=safe_strip(row.get("TeamColor")),
                coach_name=safe_strip(row.get("CoachName")),
                preferred_weekdays=sorted(set(preferred_weekdays)),
            )
        )
    return team_records


def monthly_calendar_html(schedule_df: pd.DataFrame, year: int, month: int, division_filter: str = "All") -> str:
    from calendar import monthrange
    import html

    division_colors = {
        "T-ball": {"bg": "#F7F1D6", "border": VEGAS_GOLD, "text": OFF_BLACK},
        "5-6": {"bg": "#F3F4F6", "border": "#4B5563", "text": OFF_BLACK},
        "7-8": {"bg": "#FDF8E7", "border": DEEP_GOLD, "text": OFF_BLACK},
        "9-10": {"bg": "#ECECEC", "border": CHARCOAL, "text": OFF_BLACK},
        "11-12": {"bg": "#FFFDF7", "border": "#8C7A2C", "text": OFF_BLACK},
    }

    df = schedule_df.copy()
    if df.empty:
        return "<p style='font-family:Arial,sans-serif;'>No scheduled games yet.</p>"

    df["DateOnly"] = df["Date"].dt.date
    if division_filter != "All":
        df = df[df["Division"] == division_filter]

    days_in_month = monthrange(year, month)[1]
    first_day_weekday = date(year, month, 1).weekday()
    game_map = {}

    for _, row in df.iterrows():
        d = row["DateOnly"]
        if d.year == year and d.month == month:
            division = row["Division"]
            palette = division_colors.get(
                division,
                {"bg": "#F5F5F5", "border": "#888888", "text": OFF_BLACK},
            )

            matchup = html.escape(f"{row['Away Team']} @ {row['Home Team']}")
            field = html.escape(str(row["Field"]))
            start = html.escape(str(row["Start"]))
            division_label = html.escape(str(division))

            card = f"""
            <div style="
                font-size:12px;
                margin-top:6px;
                padding:7px 8px;
                border-radius:10px;
                background:{palette['bg']};
                border-left:5px solid {palette['border']};
                color:{palette['text']};
                line-height:1.35;
                box-shadow:0 2px 8px rgba(0,0,0,0.06);
                font-family:Arial,sans-serif;
            ">
                <div style="font-weight:700;">{start} · {division_label}</div>
                <div>{matchup}</div>
                <div style="opacity:0.8;">{field}</div>
            </div>
            """
            game_map.setdefault(d.day, []).append(card)

    legend_html = ""
    if division_filter == "All":
        legend_items = []
        for division, palette in division_colors.items():
            legend_items.append(
                f"""
                <div style="display:flex;align-items:center;gap:6px;margin-right:14px;margin-bottom:6px;">
                    <span style="
                        width:14px;
                        height:14px;
                        display:inline-block;
                        border-radius:3px;
                        background:{palette['bg']};
                        border-left:5px solid {palette['border']};
                    "></span>
                    <span style="font-size:13px;font-family:Arial,sans-serif;">{html.escape(division)}</span>
                </div>
                """
            )
        legend_html = f"""
        <div style="display:flex;flex-wrap:wrap;align-items:center;margin-bottom:12px;">
            {''.join(legend_items)}
        </div>
        """

    html_out = f"""
    <style>
    .cal-grid {{
        display:grid;
        grid-template-columns:repeat(7, 1fr);
        gap:10px;
        font-family:Arial,sans-serif;
    }}
    .cal-header {{
        font-weight:700;
        text-align:center;
        padding:10px 8px;
        background:{OFF_BLACK};
        color:{WHITE};
        border-radius:10px;
    }}
    .cal-day {{
        min-height:160px;
        border:1px solid {BORDER};
        border-radius:12px;
        padding:8px;
        background:{WHITE};
        overflow:hidden;
        box-shadow:0 4px 12px rgba(0,0,0,0.04);
    }}
    .cal-muted {{
        background:#F8F8F8;
        color:#C0C0C0;
    }}
    .cal-day-number {{
        font-weight:800;
        font-size:14px;
        margin-bottom:4px;
        color:{OFF_BLACK};
    }}
    </style>

    {legend_html}

    <div class='cal-grid'>
    """

    for wd in ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]:
        html_out += f"<div class='cal-header'>{wd}</div>"

    total_cells = ((first_day_weekday + days_in_month + 6) // 7) * 7
    for cell in range(total_cells):
        day_num = cell - first_day_weekday + 1
        if day_num < 1 or day_num > days_in_month:
            html_out += "<div class='cal-day cal-muted'></div>"
        else:
            games_html = "".join(game_map.get(day_num, []))
            html_out += f"<div class='cal-day'><div class='cal-day-number'>{day_num}</div>{games_html}</div>"

    html_out += "</div>"
    return html_out


def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def get_all_blocked_dates(current_year: int, extra_text: str) -> set:
    blocked_dates = set(st.session_state.blockout_dates)
    tokens = [x.strip() for x in extra_text.split(",") if x.strip()]
    parsed_from_text, bad_tokens = parse_blockout_tokens(tokens, current_year)
    blocked_dates.update(parsed_from_text)
    if bad_tokens:
        st.warning(f"These blackout values could not be parsed: {', '.join(bad_tokens)}")
    return blocked_dates


def division_games_on_date(scheduled_games: List[GameRecord], division: str, current_date: date) -> int:
    return sum(1 for g in scheduled_games if g.division == division and g.game_date == current_date)


def division_games_in_week(scheduled_games: List[GameRecord], division: str, current_date: date) -> int:
    iso_year, iso_week, _ = current_date.isocalendar()
    return sum(
        1
        for g in scheduled_games
        if g.division == division and g.game_date.isocalendar()[:2] == (iso_year, iso_week)
    )


def choose_home_away(
    division: str,
    team_a: str,
    team_b: str,
    home_counts: Dict[Tuple[str, str], int],
    away_counts: Dict[Tuple[str, str], int],
) -> Tuple[str, str]:
    a_home = home_counts.get((division, team_a), 0)
    b_home = home_counts.get((division, team_b), 0)
    a_away = away_counts.get((division, team_a), 0)
    b_away = away_counts.get((division, team_b), 0)

    if a_home < b_home:
        return team_a, team_b
    if b_home < a_home:
        return team_b, team_a

    if a_away > b_away:
        return team_a, team_b
    if b_away > a_away:
        return team_b, team_a

    return (team_a, team_b) if team_a <= team_b else (team_b, team_a)


def get_candidate_dates_sorted(
    start_date: date,
    end_date: date,
    blocked_dates: set,
    shared_allowed_days: set,
    global_allowed_weekdays: List[int],
    division: str,
    scheduled_games: List[GameRecord],
) -> List[date]:
    candidates = []
    for current_date in daterange(start_date, end_date):
        if current_date in blocked_dates:
            continue
        if current_date.weekday() not in global_allowed_weekdays:
            continue
        if current_date.weekday() not in shared_allowed_days:
            continue
        candidates.append(current_date)

    def sort_key(d: date):
        if is_saturday(d):
            return (0, d)
        return (
            1,
            division_games_on_date(scheduled_games, division, d),
            division_games_in_week(scheduled_games, division, d),
            d.weekday(),
            d,
        )

    return sorted(candidates, key=sort_key)


def try_generate_schedule(
    teams: List[TeamRecord],
    games_per_team_by_division: Dict[str, int],
    max_games_per_week_by_division: Dict[str, int],
    division_night_groups: Dict[str, str],
    division_fields: Dict[str, List[str]],
    start_date: date,
    end_date: date,
    global_allowed_weekdays: List[int],
    blocked_dates: set,
    use_overflow_saturday_slot: bool,
) -> Tuple[pd.DataFrame, pd.DataFrame, List[str]]:
    warnings: List[str] = []
    scheduled_games: List[GameRecord] = []
    unscheduled_rows: List[dict] = []

    if not teams:
        return pd.DataFrame(), pd.DataFrame(), ["No teams found to schedule."]

    team_lookup = build_team_lookup(teams)
    divisions = sorted({t.division for t in teams})
    teams_by_division = {division: [t for t in teams if t.division == division] for division in divisions}

    inventory = {}
    for current_date in daterange(start_date, end_date):
        if current_date in blocked_dates:
            continue
        if current_date.weekday() not in global_allowed_weekdays:
            continue
        for slot_label, _, _, _ in get_day_slot_defs(current_date, use_overflow_saturday_slot):
            for field in ALL_FIELDS:
                inventory[(current_date, slot_label, field)] = None

    team_busy = set()
    team_week_counts = {}
    home_counts: Dict[Tuple[str, str], int] = {}
    away_counts: Dict[Tuple[str, str], int] = {}

    for division in divisions:
        division_teams = teams_by_division[division]
        team_names = [t.team for t in division_teams]
        games_per_team = games_per_team_by_division.get(division, 0)
        max_games_per_week = max_games_per_week_by_division.get(division, 99)

        if len(team_names) < 2 or games_per_team <= 0:
            continue

        pairings = generate_round_robin_pairings(team_names, games_per_team)

        for team_a, team_b in pairings:
            shared_allowed_days = allowed_weekdays_for_matchup(
                team_lookup,
                division,
                team_a,
                team_b,
                division_night_groups,
            ).intersection(set(global_allowed_weekdays))

            allowed_fields = eligible_fields_for_matchup(division, division_fields)
            placed = False

            if not allowed_fields:
                unscheduled_rows.append(
                    {
                        "Division": division,
                        "Home Team": team_a,
                        "Away Team": team_b,
                        "Reason": "No fields selected for this division.",
                    }
                )
                continue

            candidate_dates = get_candidate_dates_sorted(
                start_date=start_date,
                end_date=end_date,
                blocked_dates=blocked_dates,
                shared_allowed_days=shared_allowed_days,
                global_allowed_weekdays=global_allowed_weekdays,
                division=division,
                scheduled_games=scheduled_games,
            )

            for current_date in candidate_dates:
                iso_year, iso_week, _ = current_date.isocalendar()
                a_week_key = (division, team_a, iso_year, iso_week)
                b_week_key = (division, team_b, iso_year, iso_week)

                if team_week_counts.get(a_week_key, 0) >= max_games_per_week:
                    continue
                if team_week_counts.get(b_week_key, 0) >= max_games_per_week:
                    continue

                same_day_conflict = any(
                    g.game_date == current_date
                    and g.division == division
                    and (g.home_team in (team_a, team_b) or g.away_team in (team_a, team_b))
                    for g in scheduled_games
                )
                if same_day_conflict:
                    continue

                slot_defs = get_day_slot_defs(current_date, use_overflow_saturday_slot)

                for slot_label, start_t, end_t, duration_minutes in slot_defs:
                    if (current_date, slot_label, f"{division}::{team_a}") in team_busy or (
                        current_date, slot_label, f"{division}::{team_b}"
                    ) in team_busy:
                        continue

                    for field in allowed_fields:
                        key = (current_date, slot_label, field)
                        if key not in inventory or inventory[key] is not None:
                            continue

                        home_team, away_team = choose_home_away(
                            division=division,
                            team_a=team_a,
                            team_b=team_b,
                            home_counts=home_counts,
                            away_counts=away_counts,
                        )

                        scheduled_games.append(
                            GameRecord(
                                division=division,
                                home_team=home_team,
                                away_team=away_team,
                                game_date=current_date,
                                start_time=start_t,
                                end_time=end_t,
                                field=field,
                                slot_label=slot_label,
                                duration_minutes=duration_minutes,
                            )
                        )

                        inventory[key] = f"{away_team} @ {home_team}"
                        team_busy.add((current_date, slot_label, f"{division}::{team_a}"))
                        team_busy.add((current_date, slot_label, f"{division}::{team_b}"))
                        team_week_counts[a_week_key] = team_week_counts.get(a_week_key, 0) + 1
                        team_week_counts[b_week_key] = team_week_counts.get(b_week_key, 0) + 1
                        home_counts[(division, home_team)] = home_counts.get((division, home_team), 0) + 1
                        away_counts[(division, away_team)] = away_counts.get((division, away_team), 0) + 1

                        placed = True
                        break

                    if placed:
                        break
                if placed:
                    break

            if not placed and division in division_fields and division_fields[division]:
                unscheduled_rows.append(
                    {
                        "Division": division,
                        "Home Team": team_a,
                        "Away Team": team_b,
                        "Reason": "No open slot matched date, field, weekday, night-group, or weekly-limit rules.",
                    }
                )

    schedule_df = pd.DataFrame([game.to_dict() for game in scheduled_games])
    if not schedule_df.empty:
        schedule_df["Date"] = pd.to_datetime(schedule_df["Date"])
        schedule_df = schedule_df.sort_values(["Date", "Start", "Field", "Division"]).reset_index(drop=True)

    unscheduled_df = pd.DataFrame(unscheduled_rows)
    return schedule_df, unscheduled_df, warnings


def generate_schedule(
    teams: List[TeamRecord],
    games_per_team_by_division: Dict[str, int],
    max_games_per_week_by_division: Dict[str, int],
    division_night_groups: Dict[str, str],
    division_fields: Dict[str, List[str]],
    start_date: date,
    end_date: date,
    global_allowed_weekdays: List[int],
    blocked_dates: set,
) -> Tuple[pd.DataFrame, pd.DataFrame, List[str]]:
    warnings: List[str] = []

    schedule_df, unscheduled_df, _ = try_generate_schedule(
        teams=teams,
        games_per_team_by_division=games_per_team_by_division,
        max_games_per_week_by_division=max_games_per_week_by_division,
        division_night_groups=division_night_groups,
        division_fields=division_fields,
        start_date=start_date,
        end_date=end_date,
        global_allowed_weekdays=global_allowed_weekdays,
        blocked_dates=blocked_dates,
        use_overflow_saturday_slot=False,
    )

    if unscheduled_df.empty:
        return schedule_df, unscheduled_df, warnings

    overflow_schedule_df, overflow_unscheduled_df, _ = try_generate_schedule(
        teams=teams,
        games_per_team_by_division=games_per_team_by_division,
        max_games_per_week_by_division=max_games_per_week_by_division,
        division_night_groups=division_night_groups,
        division_fields=division_fields,
        start_date=start_date,
        end_date=end_date,
        global_allowed_weekdays=global_allowed_weekdays,
        blocked_dates=blocked_dates,
        use_overflow_saturday_slot=True,
    )

    if len(overflow_unscheduled_df) < len(unscheduled_df):
        warnings.append("Saturday 5:30 PM overflow slots were used because the base Saturday-first schedule could not fit all games.")
        if not overflow_unscheduled_df.empty:
            warnings.append(
                f"{len(overflow_unscheduled_df)} game(s) still could not be scheduled automatically and were saved to the unscheduled list."
            )
        return overflow_schedule_df, overflow_unscheduled_df, warnings

    warnings.append(
        f"{len(unscheduled_df)} game(s) could not be scheduled automatically and were saved to the unscheduled list."
    )
    return schedule_df, unscheduled_df, warnings


def get_open_slots(
    schedule_df: pd.DataFrame,
    selected_date: date,
    selected_division: str,
    blocked_dates: set,
    global_allowed_weekdays: List[int],
    division_night_groups: Dict[str, str],
    division_fields: Dict[str, List[str]],
    include_saturday_overflow: bool = True,
) -> List[Tuple[str, str]]:
    if selected_date in blocked_dates:
        return []
    if selected_date.weekday() not in global_allowed_weekdays:
        return []

    if is_weekday(selected_date):
        division_weekdays = get_division_weekday_set(selected_division, division_night_groups)
        if selected_date.weekday() not in division_weekdays:
            return []

    open_slots = []
    schedule_copy = schedule_df.copy()
    if not schedule_copy.empty:
        schedule_copy["Date"] = pd.to_datetime(schedule_copy["Date"])

    slot_defs = get_day_slot_defs(selected_date, include_saturday_overflow)
    allowed_fields = division_fields.get(selected_division, [])

    for slot_label, _, _, _ in slot_defs:
        for field in allowed_fields:
            occupied = False
            if not schedule_copy.empty:
                occupied = (
                    (schedule_copy["Date"].dt.date == selected_date)
                    & (schedule_copy["Slot"] == slot_label)
                    & (schedule_copy["Field"] == field)
                ).any()
            if not occupied:
                open_slots.append((slot_label, field))
    return open_slots


def get_slot_duration_for_date_and_time(selected_date: date, slot_label: str) -> int:
    if is_weekday(selected_date):
        return WEEKDAY_GAME_LENGTH_MINUTES
    if is_saturday(selected_date):
        return SATURDAY_GAME_LENGTH_MINUTES
    return WEEKDAY_GAME_LENGTH_MINUTES


def can_place_manual_game(
    schedule_df: pd.DataFrame,
    game_row: pd.Series,
    selected_date: date,
    slot_label: str,
    field: str,
    team_config_df: pd.DataFrame,
    blocked_dates: set,
    global_allowed_weekdays: List[int],
    max_games_per_week_by_division: Dict[str, int],
    division_night_groups: Dict[str, str],
    division_fields: Dict[str, List[str]],
) -> Tuple[bool, str]:
    if selected_date in blocked_dates:
        return False, "That date is blocked out."

    if selected_date.weekday() not in global_allowed_weekdays:
        return False, "That day is not allowed league-wide."

    division = game_row["Division"]
    home_team = game_row["Home Team"]
    away_team = game_row["Away Team"]

    if field not in division_fields.get(division, []):
        return False, "Selected field is not allowed for this division."

    if is_weekday(selected_date):
        division_weekdays = get_division_weekday_set(division, division_night_groups)
        if selected_date.weekday() not in division_weekdays:
            return False, "That division is not assigned to that weekday night group."

    home_team_row = team_config_df[
        (team_config_df["Division"] == division) & (team_config_df["Team"] == home_team)
    ]
    away_team_row = team_config_df[
        (team_config_df["Division"] == division) & (team_config_df["Team"] == away_team)
    ]

    if home_team_row.empty or away_team_row.empty:
        return False, "One or both teams are missing team config."

    home_team_row = home_team_row.iloc[0]
    away_team_row = away_team_row.iloc[0]

    weekday_name = INT_TO_WEEKDAY_NAME[selected_date.weekday()]
    if weekday_name != "Saturday":
        home_days = deserialize_list(home_team_row["PreferredWeekdays"])
        away_days = deserialize_list(away_team_row["PreferredWeekdays"])
        if weekday_name not in home_days or weekday_name not in away_days:
            return False, "Selected weekday is not allowed for both teams."

    if not schedule_df.empty:
        df = schedule_df.copy()
        df["Date"] = pd.to_datetime(df["Date"])

        same_slot_conflict = (
            (df["Date"].dt.date == selected_date)
            & (df["Slot"] == slot_label)
            & (df["Field"] == field)
        ).any()
        if same_slot_conflict:
            return False, "That field/slot is already occupied."

        same_day_conflict = (
            (df["Date"].dt.date == selected_date)
            & (df["Division"] == division)
            & (
                (df["Home Team"].isin([home_team, away_team]))
                | (df["Away Team"].isin([home_team, away_team]))
            )
        ).any()
        if same_day_conflict:
            return False, "One of those teams already has a game that day."

        iso_year, iso_week, _ = selected_date.isocalendar()
        division_games = df[df["Division"] == division].copy()
        iso = division_games["Date"].dt.isocalendar()
        division_games["IsoYear"] = iso.year
        division_games["IsoWeek"] = iso.week

        max_games = max_games_per_week_by_division.get(division, 99)
        for team_name in [home_team, away_team]:
            team_week_count = division_games[
                ((division_games["Home Team"] == team_name) | (division_games["Away Team"] == team_name))
                & (division_games["IsoYear"] == iso_year)
                & (division_games["IsoWeek"] == iso_week)
            ].shape[0]
            if team_week_count >= max_games:
                return False, f"{team_name} is already at the weekly max."

    return True, "OK"


def add_manual_game_to_schedule(
    schedule_df: pd.DataFrame,
    unscheduled_df: pd.DataFrame,
    unscheduled_index: int,
    selected_date: date,
    slot_label: str,
    field: str,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    row = unscheduled_df.loc[unscheduled_index]
    duration_minutes = get_slot_duration_for_date_and_time(selected_date, slot_label)
    start_dt = datetime.strptime(slot_label, "%H:%M")
    end_dt = start_dt + timedelta(minutes=duration_minutes)

    new_row = {
        "Division": row["Division"],
        "Home Team": row["Home Team"],
        "Away Team": row["Away Team"],
        "Date": pd.to_datetime(selected_date),
        "Start": start_dt.strftime("%H:%M"),
        "End": end_dt.strftime("%H:%M"),
        "Field": field,
        "Slot": slot_label,
        "DurationMinutes": duration_minutes,
        "Matchup": f"{row['Away Team']} @ {row['Home Team']}",
    }

    updated_schedule = pd.concat([schedule_df, pd.DataFrame([new_row])], ignore_index=True)
    updated_schedule["Date"] = pd.to_datetime(updated_schedule["Date"])
    updated_schedule = updated_schedule.sort_values(["Date", "Start", "Field", "Division"]).reset_index(drop=True)

    updated_unscheduled = unscheduled_df.drop(index=unscheduled_index).reset_index(drop=True)
    return updated_schedule, updated_unscheduled


# -------------------------
# UI
# -------------------------
render_hero()

top_a, top_b, top_c = st.columns([1.8, 1.2, 1.2])
with top_a:
    mini_metric("Scheduling Mode", "Saturday First")
with top_b:
    mini_metric("Weekday Format", "Mon–Thu / 5:00 & 6:30")
with top_c:
    mini_metric("Weekend Format", "Saturday / 9:30, 11:30, 1:30")


section_open(
    "Workbook Setup",
    "Upload one worksheet per division using the expected tab names and columns.",
)
left_info, right_info = st.columns([2.2, 1])
with left_info:
    st.markdown(
        """
        **Workbook format**
        - Tabs: T-ball, 5-6, 7-8, 9-10, 11-12
        - Columns: Team, TeamColor, CoachName
        """
    )
with right_info:
    st.download_button(
        "Download sample workbook",
        data=create_sample_workbook_bytes(),
        file_name="little_league_scheduler_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

uploaded_file = st.file_uploader("Upload division workbook (.xlsx)", type=["xlsx"])
section_close()


# -------------------------
# Workbook ingest
# -------------------------
if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    file_size = len(file_bytes)

    if (
        st.session_state.loaded_file_name != uploaded_file.name
        or st.session_state.loaded_file_size != file_size
    ):
        try:
            division_frames, detected_columns, load_warnings = load_division_workbook(file_bytes)
            errors = validate_division_frames(division_frames)

            if errors:
                for err in errors:
                    st.error(err)
                st.stop()

            for warning in load_warnings:
                st.warning(warning)

            master_team_df = build_master_team_df(division_frames)
            st.session_state.master_team_df = master_team_df.copy()
            st.session_state.team_config_df = normalize_team_config_df(master_team_df)
            st.session_state.schedule_df = pd.DataFrame()
            st.session_state.unscheduled_df = pd.DataFrame()
            st.session_state.loaded_file_name = uploaded_file.name
            st.session_state.loaded_file_size = file_size

            st.session_state.division_night_groups = {
                division: st.session_state.division_night_groups.get(division, "None")
                for division in master_team_df["Division"].unique().tolist()
            }
            st.session_state.division_fields = {
                division: st.session_state.division_fields.get(
                    division, DEFAULT_FIELDS_BY_DIVISION.get(division, [])
                )
                for division in master_team_df["Division"].unique().tolist()
            }

            st.success("Workbook loaded successfully.")
            with st.expander("Workbook preview", expanded=False):
                st.write("Detected sheets and columns")
                st.json(detected_columns)
                st.dataframe(master_team_df, use_container_width=True, hide_index=True)

        except Exception as exc:
            st.error(f"Unable to load workbook: {exc}")

master_team_df = st.session_state.master_team_df.copy()
team_config_df = st.session_state.team_config_df.copy()
schedule_df = st.session_state.schedule_df.copy()
unscheduled_df = st.session_state.unscheduled_df.copy()


# -------------------------
# Main app once workbook is loaded
# -------------------------
if not master_team_df.empty:
    section_open(
        "League Scheduling Rules",
        "Set season dates, allowed days, blackout dates, division night groups, fields, and weekly limits.",
    )

    current_year = date.today().year
    rules_col1, rules_col2, rules_col3 = st.columns(3)

    with rules_col1:
        season_start = st.date_input("Season start date", value=date(current_year, 4, 1))
        season_end = st.date_input("Season end date", value=date(current_year, 6, 15))
        st.session_state.season_start = season_start
        st.session_state.season_end = season_end

        st.markdown(
            f"""
            <div class="scheduler-note">
                <strong>Weekday games</strong><br>
                Monday–Thursday only<br>
                5:00 PM and 6:30 PM<br><br>
                <strong>Saturday games</strong><br>
                9:30 AM, 11:30 AM, 1:30 PM<br>
                5:30 PM overflow only if needed
            </div>
            """,
            unsafe_allow_html=True,
        )

    with rules_col2:
        global_allowed_weekdays_names = st.multiselect(
            "League-wide allowed days",
            options=["Monday", "Tuesday", "Wednesday", "Thursday", "Saturday"],
            default=["Monday", "Tuesday", "Wednesday", "Thursday", "Saturday"],
            help="No Friday games in this version.",
        )

    with rules_col3:
        st.markdown("**Blackout dates**")
        new_blackout_date = st.date_input("Pick blackout date", value=None, key="new_blackout_picker")
        add_blackout = st.button("Add blackout date")
        if add_blackout and new_blackout_date:
            if new_blackout_date not in st.session_state.blockout_dates:
                st.session_state.blockout_dates.append(new_blackout_date)
                st.session_state.blockout_dates = sorted(st.session_state.blockout_dates)

        manual_blockout_text = st.text_area(
            "Also add blackout dates by text",
            value="",
            help="Examples: 04-20, 05-25, May 30",
        )

    if st.session_state.blockout_dates:
        st.write("Current blackout dates")
        blockout_df = pd.DataFrame({"BlackoutDate": st.session_state.blockout_dates})
        st.dataframe(blockout_df, use_container_width=True, hide_index=True)

        remove_choice = st.selectbox(
            "Remove blackout date",
            options=["None"] + [str(d) for d in st.session_state.blockout_dates],
            key="remove_blackout_choice",
        )
        if st.button("Remove selected blackout date"):
            if remove_choice != "None":
                st.session_state.blockout_dates = [
                    d for d in st.session_state.blockout_dates if str(d) != remove_choice
                ]
                st.rerun()

    st.markdown("### Division Setup")
    division_names = sorted(master_team_df["Division"].unique().tolist())
    games_per_team_by_division: Dict[str, int] = {}
    max_games_per_week_by_division: Dict[str, int] = {}
    division_night_groups: Dict[str, str] = {}
    division_fields: Dict[str, List[str]] = {}

    setup_cols = st.columns(len(division_names)) if division_names else []
    for idx, division in enumerate(division_names):
        with setup_cols[idx]:
            games_per_team_by_division[division] = st.number_input(
                f"{division} total games/team",
                min_value=0,
                max_value=30,
                value=8,
                step=1,
                key=f"games_per_team_{division}",
            )
            max_games_per_week_by_division[division] = st.number_input(
                f"{division} max games/week",
                min_value=1,
                max_value=7,
                value=2,
                step=1,
                key=f"max_games_per_week_{division}",
            )
            division_night_groups[division] = st.selectbox(
                f"{division} night group",
                options=DIVISION_NIGHT_GROUP_OPTIONS,
                index=DIVISION_NIGHT_GROUP_OPTIONS.index(
                    st.session_state.division_night_groups.get(division, "None")
                ),
                key=f"night_group_{division}",
                help="None = not assigned to a special weekday group. Saturday is still allowed.",
            )
            division_fields[division] = st.multiselect(
                f"{division} allowed fields",
                options=ALL_FIELDS,
                default=st.session_state.division_fields.get(division, DEFAULT_FIELDS_BY_DIVISION.get(division, [])),
                key=f"division_fields_{division}",
                help="These fields are available for this entire division.",
            )

    st.session_state.max_games_per_week_by_division = max_games_per_week_by_division.copy()
    st.session_state.division_night_groups = division_night_groups.copy()
    st.session_state.division_fields = division_fields.copy()

    st.markdown("### Division Team Configuration")
    division_tabs = st.tabs(division_names)
    updated_rows = []

    current_team_config_lookup = {}
    if not team_config_df.empty:
        for _, row in team_config_df.iterrows():
            current_team_config_lookup[(row["Division"], row["Team"])] = row.to_dict()

    for idx, division in enumerate(division_names):
        with division_tabs[idx]:
            st.markdown(f"**{division}**")
            st.caption(f"Night group: {division_night_groups[division]} | Fields: {', '.join(division_fields[division]) if division_fields[division] else 'None selected'}")
            division_teams = master_team_df[master_team_df["Division"] == division].copy().reset_index(drop=True)

            for _, team_row in division_teams.iterrows():
                team_name = team_row["Team"]
                team_color = team_row["TeamColor"]
                coach_name = team_row["CoachName"]
                config_key = (division, team_name)

                existing = current_team_config_lookup.get(
                    config_key,
                    {
                        "PreferredWeekdays": serialize_list(WEEKDAY_OPTIONS),
                    },
                )

                default_days = set(deserialize_list(existing.get("PreferredWeekdays", "")))

                with st.expander(f"{team_name} · {team_color} · {coach_name}", expanded=False):
                    st.markdown("**Allowed weekdays**")
                    day_cols = st.columns(len(WEEKDAY_OPTIONS))
                    selected_days = []
                    for i, weekday in enumerate(WEEKDAY_OPTIONS):
                        with day_cols[i]:
                            checked = st.checkbox(
                                weekday[:3],
                                value=weekday in default_days,
                                key=f"day_{division}_{team_name}_{weekday}",
                            )
                            if checked:
                                selected_days.append(weekday)

                    st.caption("Saturday is always allowed automatically. Friday is not used. Fields are now controlled at the division level.")

                    updated_rows.append(
                        {
                            "Division": division,
                            "Team": team_name,
                            "TeamColor": team_color,
                            "CoachName": coach_name,
                            "PreferredWeekdays": serialize_list(selected_days),
                        }
                    )

    if updated_rows:
        team_config_df = pd.DataFrame(updated_rows)
        st.session_state.team_config_df = team_config_df.copy()

    global_allowed_weekdays = [WEEKDAY_NAME_TO_INT[name] for name in global_allowed_weekdays_names]
    st.session_state.global_allowed_weekdays = global_allowed_weekdays.copy()

    blocked_dates = get_all_blocked_dates(current_year, manual_blockout_text)

    action_col1, action_col2 = st.columns([1, 2])
    with action_col1:
        generate_clicked = st.button("Generate Schedule", type="primary", use_container_width=True)
    with action_col2:
        st.markdown(
            """
            <div class="scheduler-note">
                Scheduling priority: fill Saturdays first, then spread remaining games across weekdays by division.
                Home/away is balanced automatically.
            </div>
            """,
            unsafe_allow_html=True,
        )

    section_close()

    if generate_clicked:
        team_records = parse_team_records_from_config(team_config_df)
        schedule_df, unscheduled_df, warnings = generate_schedule(
            teams=team_records,
            games_per_team_by_division=games_per_team_by_division,
            max_games_per_week_by_division=max_games_per_week_by_division,
            division_night_groups=division_night_groups,
            division_fields=division_fields,
            start_date=season_start,
            end_date=season_end,
            global_allowed_weekdays=global_allowed_weekdays,
            blocked_dates=blocked_dates,
        )
        st.session_state.schedule_df = schedule_df.copy()
        st.session_state.unscheduled_df = unscheduled_df.copy()

        for warning in warnings:
            st.warning(warning)

        if schedule_df.empty:
            st.error("No games were scheduled. Try expanding the date range, selecting fields, or loosening restrictions.")
        else:
            st.success(f"Generated {len(schedule_df)} scheduled games.")

schedule_df = st.session_state.schedule_df.copy()
unscheduled_df = st.session_state.unscheduled_df.copy()


# -------------------------
# Manual placement area
# -------------------------
if not unscheduled_df.empty:
    section_open(
        "Manual Placement for Unscheduled Games",
        "Pick an unscheduled game and place it into an open, valid slot.",
    )

    current_year = date.today().year
    blocked_dates = get_all_blocked_dates(current_year, "")
    global_allowed_weekdays = st.session_state.global_allowed_weekdays
    max_games_per_week_by_division = st.session_state.max_games_per_week_by_division
    division_night_groups = st.session_state.division_night_groups
    division_fields = st.session_state.division_fields

    default_manual_date = st.session_state.season_start if st.session_state.season_start else date.today()

    unscheduled_display = unscheduled_df.copy()
    unscheduled_display["Label"] = (
        unscheduled_display["Division"]
        + " | "
        + unscheduled_display["Away Team"]
        + " @ "
        + unscheduled_display["Home Team"]
    )

    selected_label = st.selectbox("Choose an unscheduled game", options=unscheduled_display["Label"].tolist())
    selected_idx = unscheduled_display[unscheduled_display["Label"] == selected_label].index[0]
    selected_game = unscheduled_df.loc[selected_idx]
    selected_division = selected_game["Division"]

    mp_col1, mp_col2, mp_col3 = st.columns(3)
    with mp_col1:
        manual_date = st.date_input("Manual game date", value=default_manual_date, key="manual_date")
    with mp_col2:
        open_slots = get_open_slots(
            schedule_df=schedule_df,
            selected_date=manual_date,
            selected_division=selected_division,
            blocked_dates=blocked_dates,
            global_allowed_weekdays=global_allowed_weekdays,
            division_night_groups=division_night_groups,
            division_fields=division_fields,
            include_saturday_overflow=True,
        )
        open_slot_labels = [f"{slot} | {field}" for slot, field in open_slots]
        selected_open_slot = st.selectbox(
            "Open slot",
            options=open_slot_labels if open_slot_labels else ["No open slots"],
        )
    with mp_col3:
        mini_metric("Selected Matchup", f"{selected_game['Away Team']} @ {selected_game['Home Team']}")

    if open_slots and selected_open_slot != "No open slots":
        chosen_slot, chosen_field = selected_open_slot.split(" | ")

        is_valid, validation_message = can_place_manual_game(
            schedule_df=schedule_df,
            game_row=selected_game,
            selected_date=manual_date,
            slot_label=chosen_slot,
            field=chosen_field,
            team_config_df=team_config_df,
            blocked_dates=blocked_dates,
            global_allowed_weekdays=global_allowed_weekdays,
            max_games_per_week_by_division=max_games_per_week_by_division,
            division_night_groups=division_night_groups,
            division_fields=division_fields,
        )

        if is_valid:
            st.success("Slot is valid for manual placement.")
            if st.button("Place game into schedule", use_container_width=True):
                updated_schedule, updated_unscheduled = add_manual_game_to_schedule(
                    schedule_df=schedule_df,
                    unscheduled_df=unscheduled_df,
                    unscheduled_index=selected_idx,
                    selected_date=manual_date,
                    slot_label=chosen_slot,
                    field=chosen_field,
                )
                st.session_state.schedule_df = updated_schedule
                st.session_state.unscheduled_df = updated_unscheduled
                st.success("Game placed successfully.")
                st.rerun()
        else:
            st.error(validation_message)

    st.markdown("**Remaining Unscheduled Games**")
    st.dataframe(unscheduled_df, use_container_width=True, hide_index=True)
    st.download_button(
        "Download unscheduled games CSV",
        data=dataframe_to_csv_bytes(unscheduled_df),
        file_name="unscheduled_games.csv",
        mime="text/csv",
        use_container_width=True,
    )
    section_close()


# -------------------------
# Schedule Views
# -------------------------
if not schedule_df.empty:
    section_open(
        "Schedule Views",
        "Filter the schedule by division, team, and field, then view it as a calendar or as a game list.",
    )

    schedule_df["Date"] = pd.to_datetime(schedule_df["Date"])
    schedule_df["Year"] = schedule_df["Date"].dt.year
    schedule_df["Month"] = schedule_df["Date"].dt.month

    filter_col1, filter_col2, filter_col3, filter_col4 = st.columns(4)
    with filter_col1:
        division_filter = st.selectbox("Division", ["All"] + sorted(schedule_df["Division"].unique().tolist()))
    with filter_col2:
        team_options = sorted(set(schedule_df["Home Team"]).union(set(schedule_df["Away Team"])))
        team_filter = st.selectbox("Team", ["All"] + team_options)
    with filter_col3:
        field_filter = st.selectbox("Field", ["All"] + sorted(schedule_df["Field"].unique().tolist()))
    with filter_col4:
        view_mode = st.radio("View", ["Calendar", "Games List"], horizontal=True)

    filtered_df = schedule_df.copy()
    if division_filter != "All":
        filtered_df = filtered_df[filtered_df["Division"] == division_filter]
    if team_filter != "All":
        filtered_df = filtered_df[
            (filtered_df["Home Team"] == team_filter) | (filtered_df["Away Team"] == team_filter)
        ]
    if field_filter != "All":
        filtered_df = filtered_df[filtered_df["Field"] == field_filter]

    filtered_df = filtered_df.sort_values(["Date", "Start", "Field", "Division"]).reset_index(drop=True)

    export_col1, export_col2, export_col3 = st.columns(3)
    with export_col1:
        st.download_button(
            "Download filtered schedule CSV",
            data=dataframe_to_csv_bytes(
                filtered_df.drop(columns=[c for c in ["Year", "Month"] if c in filtered_df.columns])
            ),
            file_name="filtered_schedule.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with export_col2:
        st.download_button(
            "Download full schedule CSV",
            data=dataframe_to_csv_bytes(
                schedule_df.drop(columns=[c for c in ["Year", "Month"] if c in schedule_df.columns])
            ),
            file_name="full_schedule.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with export_col3:
        st.download_button(
            "Download team config CSV",
            data=dataframe_to_csv_bytes(st.session_state.team_config_df),
            file_name="team_config.csv",
            mime="text/csv",
            use_container_width=True,
        )

    if view_mode == "Games List":
        display_df = filtered_df.copy()
        display_df["Date"] = display_df["Date"].dt.strftime("%Y-%m-%d")
        st.dataframe(display_df, use_container_width=True, hide_index=True)
    else:
        cal_col1, cal_col2 = st.columns(2)
        with cal_col1:
            year_choice = st.selectbox("Calendar year", sorted(filtered_df["Year"].unique().tolist()))
        with cal_col2:
            month_choice = st.selectbox(
                "Calendar month",
                options=list(range(1, 13)),
                format_func=lambda m: datetime(2000, m, 1).strftime("%B"),
                index=max(0, min(11, int(filtered_df["Month"].min()) - 1)),
            )
        calendar_html = monthly_calendar_html(filtered_df, year_choice, month_choice, division_filter=division_filter)
        st.components.v1.html(calendar_html, height=1000, scrolling=True)

    st.markdown("### Home/Away Summary")
    summary_df = pd.concat(
        [
            schedule_df[["Division", "Home Team"]].rename(columns={"Home Team": "Team"}).assign(HomeGames=1, AwayGames=0),
            schedule_df[["Division", "Away Team"]].rename(columns={"Away Team": "Team"}).assign(HomeGames=0, AwayGames=1),
        ],
        ignore_index=True,
    )
    summary_df = (
        summary_df.groupby(["Division", "Team"], as_index=False)[["HomeGames", "AwayGames"]]
        .sum()
        .sort_values(["Division", "Team"])
    )
    summary_df["Diff"] = summary_df["HomeGames"] - summary_df["AwayGames"]
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    section_close()


# -------------------------
# Empty state
# -------------------------
if master_team_df.empty and uploaded_file is None:
    section_open("Get Started", "Upload your workbook to begin building the season schedule.")
    st.info("Upload your workbook to get started. Use the sample workbook if you want a template.")
    section_close()


# -------------------------
# Notes / Next steps
# -------------------------
with st.expander("What this version supports / next upgrades", expanded=False):
    st.markdown(
        """
        **Included now**
        - One workbook tab per division
        - Team name, team color, coach name ingestion
        - Division-level field selection
        - Checkbox-based weekday selection
        - No Friday games
        - Weekday schedule: 5:00 and 6:30, 1.5 hours
        - Saturday schedule: 9:30, 11:30, 1:30, 2 hours
        - Saturday 5:30 overflow only if needed
        - Saturday-first scheduling priority
        - Remaining weekday games spread more evenly by division
        - Division night groups: None, M/W, T/Th
        - Home/away balancing
        - League blackout dates with quick add/remove
        - Games per team by division
        - Max games per week by division
        - Unscheduled games list
        - Manual placement of unscheduled games into open slots
        - Filterable list and calendar views
        - Division-colored calendar
        - Home/away summary table

        **Still good next upgrades**
        - Conflict reporting explaining exactly why a game could not be placed
        - Opponent-spacing rules
        - Saved schedules / persistence
        - True drag/drop calendar with a custom frontend component
        """
    )