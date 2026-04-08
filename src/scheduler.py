from pathlib import Path
import math
import pandas as pd
from math import radians, sin, cos, asin, sqrt



# ===========================================
# CONFIG
# ===========================================

DEFAULT_MONTHLY_CAP = 45  # max jobs per engineer per month (adjust if needed)

BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_DIR = BASE_DIR / "data" / "input"
OUTPUT_DIR = BASE_DIR / "data" / "output"

JOB_PATH = INPUT_DIR / "job_data.xlsx"                # From ServicePro
ENG_PATH = INPUT_DIR / "engineers.xlsx"               # Engineer info
SCHED_PATH = INPUT_DIR / "SERVICE_SCHEDULE_MASTER.xlsx"  # Hayley's schedule
OUTCODE_PATH = INPUT_DIR / "postcode-outcodes.csv"    # UK postcodes/outcodes + lat/lon


# ===========================================
# HELPER FUNCTIONS: POSTCODES
# ===========================================

def detect_postcode_column(df: pd.DataFrame) -> str:
    """
    Detect which column in a DataFrame is the postcode column.
    Uses known names first, then falls back to anything containing 'post'.
    """
    possible_names = [
        "postcode",
        "site postal code",
        "site_postal_code",
        "post code",
        "sitepostcode",
        "site_post_code"
    ]

    lower_map = {c.lower(): c for c in df.columns}

    # 1) Try exact known names (case-insensitive)
    for known in possible_names:
        if known in lower_map:
            return lower_map[known]

    # 2) Fallback: any column containing "post"
    for col in df.columns:
        if "post" in col.lower():
            return col

    # 3) If we get here, we failed
    raise ValueError(
        f"Could not detect a postcode column. Columns found: {df.columns.tolist()}"
    )


def add_postal_area_to_jobs(jobs: pd.DataFrame) -> pd.DataFrame:
    """
    Add Outcode (e.g. 'B24') and postalArea (e.g. 'B') columns to the jobs dataframe.
    Uses whatever postcode column is found (e.g. 'Site Postal Code').
    """
    jobs = jobs.copy()

    postcode_col = detect_postcode_column(jobs)
    print(f"[jobs] Using '{postcode_col}' as the postcode column in job_data.xlsx")

    jobs[postcode_col] = jobs[postcode_col].astype(str).str.strip()

    # Outcode = e.g. 'B24' from 'B24 8AB'
    jobs["Outcode"] = jobs[postcode_col].str.extract(r"^([A-Z]{1,2}\d{1,2})", expand=False)

    # postalArea = just letters, e.g. 'B' or 'AB'
    jobs["postalArea"] = jobs[postcode_col].str.extract(r"^([A-Z]{1,2})", expand=False)

    return jobs


def add_postal_area_to_engineers(engineers: pd.DataFrame) -> pd.DataFrame:
    """
    Add Outcode and postalArea for engineers based on their home 'Post Code' (or equivalent).
    """
    engineers = engineers.copy()

    postcode_col = detect_postcode_column(engineers)
    print(f"[engineers] Using '{postcode_col}' as the postcode column in engineers.xlsx")

    engineers[postcode_col] = engineers[postcode_col].astype(str).str.strip().str.upper()

    # Outcode = e.g. 'B36' from 'B36 0JU'
    engineers["Outcode"] = engineers[postcode_col].str.extract(r"^([A-Z]{1,2}\d{1,2})", expand=False)

    # postalArea = just letters, e.g. 'B' or 'AB'
    engineers["postalArea"] = engineers[postcode_col].str.extract(r"^([A-Z]{1,2})", expand=False)

    return engineers


# ===========================================
# GEO HELPERS (DISTANCE / CENTROIDS)
# ===========================================

def haversine_km(lat1, lon1, lat2, lon2) -> float:
    """
    Compute great-circle distance between two points on Earth (km).
    """
    # Handle missing coords
    if any(pd.isna(x) for x in [lat1, lon1, lat2, lon2]):
        return float("inf")

    # Convert decimal degrees to radians
    lat1, lon1, lat2, lon2 = map(radians, [lat1, lon1, lat2, lon2])

    dlat = lat2 - lat1
    dlon = lon2 - lon1

    a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlon / 2) ** 2
    c = 2 * asin(sqrt(a))
    r = 6371  # Radius of earth in km
    return c * r


def build_geo_maps(outcodes: pd.DataFrame):
    """
    Build:
    - outcode_coords: mapping Outcode -> (lat, lon)
    - area_centroids: mapping postalArea -> (lat, lon)
    from postcode-outcodes.csv.
    """

    oc = outcodes.copy()
    oc["postcode"] = oc["postcode"].astype(str).str.strip().str.upper()

    # Extract Outcode and postalArea from outcodes file
    oc["Outcode"] = oc["postcode"].str.extract(r"^([A-Z]{1,2}\d{1,2})", expand=False)
    oc["postalArea"] = oc["postcode"].str.extract(r"^([A-Z]{1,2})", expand=False)

    # Outcode-level coordinates (mean if multiple rows per outcode)
    outcode_coords = (
        oc.dropna(subset=["Outcode"])
          .groupby("Outcode")[["latitude", "longitude"]]
          .mean()
          .dropna()
          .to_dict("index")
    )

    # Postal area centroids (e.g. AB, B, NG)
    area_centroids = (
        oc.dropna(subset=["postalArea"])
          .groupby("postalArea")[["latitude", "longitude"]]
          .mean()
          .dropna()
          .to_dict("index")
    )

    return outcode_coords, area_centroids


# ===========================================
# SCHEDULE MASTER NORMALISATION (HAYLEY FILE)
# ===========================================

MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTH_TO_NUM = {m: i+1 for i, m in enumerate(MONTH_ORDER)}
NUM_TO_MONTH = {v: k for k, v in MONTH_TO_NUM.items()}


def build_schedule_cycles(schedule: pd.DataFrame) -> pd.DataFrame:
    """
    Take Hayley's SERVICE_SCHEDULE_MASTER sheet and turn it into a tidy
    table of (Postcode, postalArea, CycleID, Month1, Month2 or single month).

    Assumptions:
    - 'Postcode' column exists.
    - Month columns are named: Jan, Feb, ..., Dec.
    - Non-empty / non-zero => serviced that month.
    - You already enforced 6-month patterns manually.
    """
    sched = schedule.copy()

    if "Postcode" not in sched.columns:
        raise ValueError("Expected a 'Postcode' column in SERVICE_SCHEDULE_MASTER.xlsx")

    sched["postalArea"] = sched["Postcode"].astype(str).str.extract(r"^([A-Z]{1,2})", expand=False)

    records = []
    for _, row in sched.iterrows():
        postcode = str(row["Postcode"]).strip()
        postal_area = row["postalArea"]

        active_months = []
        for m in MONTH_ORDER:
            if m in row.index:
                val = row[m]
                if pd.notna(val) and str(val).strip() not in ("", "0"):
                    active_months.append(MONTH_TO_NUM[m])

        if not active_months:
            continue

        active_months = sorted(active_months)

        # Once-a-year (e.g. ZE = just Mar)
        if len(active_months) == 1:
            m = active_months[0]
            records.append({
                "Postcode": postcode,
                "postalArea": postal_area,
                "CycleID": 1,
                "Month1": NUM_TO_MONTH[m],
                "Month2": None,
                "IsAnnualOnly": True
            })
            continue

        # General case: 6-month pairs
        used = set()
        cycle_id = 1

        for i, m1 in enumerate(active_months):
            if m1 in used:
                continue
            partner = None
            for m2 in active_months:
                if m2 in used or m2 == m1:
                    continue
                diff = (m2 - m1) % 12
                if diff == 6:
                    partner = m2
                    break

            if partner is not None:
                records.append({
                    "Postcode": postcode,
                    "postalArea": postal_area,
                    "CycleID": cycle_id,
                    "Month1": NUM_TO_MONTH[m1],
                    "Month2": NUM_TO_MONTH[partner],
                    "IsAnnualOnly": False
                })
                used.add(m1)
                used.add(partner)
                cycle_id += 1
            else:
                # Single-month cycle if no exact partner
                records.append({
                    "Postcode": postcode,
                    "postalArea": postal_area,
                    "CycleID": cycle_id,
                    "Month1": NUM_TO_MONTH[m1],
                    "Month2": None,
                    "IsAnnualOnly": False
                })
                used.add(m1)
                cycle_id += 1

    cycles_df = pd.DataFrame(records)
    return cycles_df


# ===========================================
# DISTANCE-BASED ENGINEER ORDERING
# ===========================================

def build_area_engineer_distance_map(
    engineers_with_area: pd.DataFrame,
    outcode_coords: dict,
    area_centroids: dict,
    cycles_df: pd.DataFrame,
):
    """
    For each postalArea in cycles_df, get a list of engineers sorted by
    distance (km) from that area's centroid.
    Returns:
    - area_eng_map: postalArea -> [EngineerName sorted by distance]
    - all_engineers: list of all EngineerName (for fallback)
    """

    eng = engineers_with_area.copy()

    # Build readable name
    first_col = next((c for c in eng.columns if c.lower() in ("first name", "firstname", "first_name")), None)
    last_col = next((c for c in eng.columns if c.lower() in ("last name", "lastname", "last_name")), None)

    if first_col and last_col:
        eng["EngineerName"] = eng[first_col].astype(str).str.strip() + " " + eng[last_col].astype(str).str.strip()
    else:
        if "Employee Id" in eng.columns:
            eng["EngineerName"] = eng["Employee Id"].astype(str)
        else:
            eng["EngineerName"] = eng.index.astype(str)

    # Attach lat/lon to engineers from Outcode map or area centroid as fallback
    eng["lat"] = None
    eng["lon"] = None

    for idx, row in eng.iterrows():
        outcode = row.get("Outcode")
        area = row.get("postalArea")

        lat = lon = None
        if isinstance(outcode, str) and outcode in outcode_coords:
            lat = outcode_coords[outcode]["latitude"]
            lon = outcode_coords[outcode]["longitude"]
        elif isinstance(area, str) and area in area_centroids:
            lat = area_centroids[area]["latitude"]
            lon = area_centroids[area]["longitude"]

        eng.at[idx, "lat"] = lat
        eng.at[idx, "lon"] = lon

    # Drop engineers with no coordinates at all
    eng_valid = eng.dropna(subset=["lat", "lon"])

    all_engineers = eng_valid["EngineerName"].tolist()
    area_eng_map = {}

    unique_areas = cycles_df["postalArea"].dropna().unique()

    for area in unique_areas:
        if area not in area_centroids:
            # No centroid for this area, fallback to all engineers as-is
            area_eng_map[area] = all_engineers[:]
            continue

        area_lat = area_centroids[area]["latitude"]
        area_lon = area_centroids[area]["longitude"]

        distances = []
        for _, erow in eng_valid.iterrows():
            d = haversine_km(area_lat, area_lon, erow["lat"], erow["lon"])
            distances.append((erow["EngineerName"], d))

        distances_sorted = sorted(distances, key=lambda x: x[1])
        area_eng_map[area] = [name for name, _ in distances_sorted]

    return area_eng_map, all_engineers


# ===========================================
# JOB DISTRIBUTION + ENGINEER ASSIGNMENT
# ===========================================

def assign_jobs_to_engineers(cycles_df: pd.DataFrame,
                             jobs_with_area: pd.DataFrame,
                             area_eng_map: dict,
                             all_engineers: list,
                             monthly_cap: int) ->   tuple[pd.DataFrame, pd.DataFrame]:

    """
    Use:
    - cycles_df: schedule cycles from Hayley's master
    - jobs_with_area: job_data with postalArea
    - area_eng_map: postalArea -> [EngineerName sorted by distance]
    - all_engineers: global list of engineers (fallback)

    Steps:
    - Count total jobs per postalArea from job_data.
    - For each postalArea, work out how many 'visits' per year.
    - Distribute total jobs evenly across visits.
    - For each visit (postalArea + Month1/Month2), assign jobs:
        * primary = nearest engineer for that area
        * enforce MONTHLY_CAP
        * spillover = next-nearest engineer, etc.
    """
    # Count jobs per postalArea
    jobs_counts = jobs_with_area.groupby("postalArea").size().to_dict()

    # Compute total visits per postalArea from cycles
    visits_per_area = {}
    for area, group in cycles_df.groupby("postalArea"):
        visits = 0
        for _, row in group.iterrows():
            if pd.isna(row["Month2"]) or row["Month2"] is None:
                visits += 1
            else:
                visits += 2
        visits_per_area[area] = visits if visits > 0 else 1

    # Track monthly load per engineer
    engineer_load = {}

    assignment_records = []
    base_year = 2025  # reference planning year

    for area, group in cycles_df.groupby("postalArea"):
        total_jobs = jobs_counts.get(area, 0)
        if total_jobs == 0:
            continue

        total_visits = visits_per_area.get(area, 1)
        jobs_per_visit = math.ceil(total_jobs / total_visits)

        # Nearest engineers for this area
        candidate_engineers = area_eng_map.get(area)
        if not candidate_engineers:
            candidate_engineers = all_engineers  # fallback

        if not candidate_engineers:
            # no engineers at all (should not happen)
            continue

        primary_engineer = candidate_engineers[0]

        for _, row in group.iterrows():
            months = []
            m1 = row["Month1"]
            m2 = row["Month2"]
            if isinstance(m1, str):
                months.append(MONTH_TO_NUM[m1])
            elif isinstance(m1, int):
                months.append(m1)

            if pd.notna(m2):
                if isinstance(m2, str):
                    months.append(MONTH_TO_NUM[m2])
                elif isinstance(m2, int):
                    months.append(m2)

            for m in months:
                year_month = f"{base_year}-{m:02d}"

                assigned_engineer = None
                spillover = False
                over_cap = False

                # Try engineers in order of distance
                for idx, eng_name in enumerate(candidate_engineers):
                    key = (eng_name, year_month)
                    current_load = engineer_load.get(key, 0)

                    if current_load + jobs_per_visit <= monthly_cap:
                        assigned_engineer = eng_name
                        engineer_load[key] = current_load + jobs_per_visit
                        if idx > 0:
                            spillover = True
                        break

                if assigned_engineer is None:
                    # Everyone at/over cap -> force onto primary, mark over-cap
                    key = (primary_engineer, year_month)
                    current_load = engineer_load.get(key, 0)
                    engineer_load[key] = current_load + jobs_per_visit
                    assigned_engineer = primary_engineer
                    over_cap = True

                assignment_records.append({
                    "YearMonth": year_month,
                    "postalArea": area,
                    "PostcodeExample": row["Postcode"],
                    "CycleID": row["CycleID"],
                    "PlannedMonthName": NUM_TO_MONTH[m],
                    "Jobs_Assigned": jobs_per_visit,
                    "Engineer": assigned_engineer,
                    "IsSpillover": spillover,
                    "IsOverCap": over_cap
                })

    assignments_df = pd.DataFrame(assignment_records)

    # Build summary per engineer per month
    summary_records = []
    for (eng_name, year_month), count in engineer_load.items():
        summary_records.append({
            "Engineer": eng_name,
            "YearMonth": year_month,
            "Jobs_Assigned": count,
            "Capacity": monthly_cap,
            "UtilisationPct": round(100 * count / monthly_cap, 1),
            "Status": (
                "Over cap" if count > monthly_cap
                else "At cap" if count == monthly_cap
                else "Under cap"
            )
        })

    summary_df = pd.DataFrame(summary_records)

    return assignments_df, summary_df


# ===========================================
# REBUILD HAYLEY-STYLE SCHEDULE (STEP 7)
# ===========================================

def rebuild_hayley_schedule(schedule: pd.DataFrame,
                            assignments_df: pd.DataFrame) -> pd.DataFrame:
    """
    Take the original Hayley SERVICE_SCHEDULE_MASTER layout and overwrite
    the Jan–Dec columns with the new job counts based on assignments_df.
    """
    sched = schedule.copy()

    if "Postcode" not in sched.columns:
        raise ValueError("Expected 'Postcode' column in schedule.")

    if assignments_df.empty:
        return sched

    # Pivot assignments: PostcodeExample x PlannedMonthName -> total jobs
    pivot = assignments_df.pivot_table(
        index="PostcodeExample",
        columns="PlannedMonthName",
        values="Jobs_Assigned",
        aggfunc="sum",
        fill_value=0,
    )

    # Ensure all month columns exist
    for m in MONTH_ORDER:
        if m not in pivot.columns:
            pivot[m] = 0

    pivot = pivot[MONTH_ORDER]

    def get_new_month_value(row, month_name):
        pc = str(row["Postcode"]).strip()
        if pc in pivot.index:
            return pivot.loc[pc, month_name]
        if month_name in row.index:
            return row[month_name]
        return 0

    for m in MONTH_ORDER:
        if m in sched.columns:
            sched[m] = sched.apply(lambda r, mm=m: get_new_month_value(r, mm), axis=1)

    return sched


# ===========================================
# MAIN
# ===========================================

from typing import Optional

def main(monthly_cap: int = DEFAULT_MONTHLY_CAP,
         engineers_override_path: Optional[Path] = None):

    print("=== Service Scheduler – Preview (Steps 1–8) ===\n")

    # 1. Load all input files
    print(f"Loading job data from:        {JOB_PATH}")
    jobs = pd.read_excel(JOB_PATH)

    eng_path = engineers_override_path or ENG_PATH
    print(f"Loading engineers from:       {eng_path}")
    engineers = pd.read_excel(eng_path)

    print(f"Loading schedule master from: {SCHED_PATH}")
    schedule = pd.read_excel(SCHED_PATH)

    print(f"Loading postcode outcodes from: {OUTCODE_PATH}")
    outcodes = pd.read_csv(OUTCODE_PATH)

    # 2. Show basic info
    print("\n--- Columns detected ---")
    print("Jobs columns:      ", jobs.columns.tolist())
    print("Engineers columns: ", engineers.columns.tolist())
    print("Schedule columns:  ", schedule.columns.tolist())
    print("Outcodes columns:  ", outcodes.columns.tolist())

    # 3. Derive Outcode/postalArea for jobs
    print("\n--- Sample job rows with Outcode/postalArea ---")
    jobs_with_area = add_postal_area_to_jobs(jobs)
    print(jobs_with_area[["Outcode", "postalArea"]].head())

    # 4. Derive Outcode/postalArea for engineers
    print("\n--- Sample engineer rows with Outcode/postalArea ---")
    engineers_with_area = add_postal_area_to_engineers(engineers)
    engineer_name_cols = [c for c in engineers_with_area.columns if c.lower() in ("first name", "firstname", "first_name")]
    last_name_cols = [c for c in engineers_with_area.columns if c.lower() in ("last name", "lastname", "last_name")]

    cols_to_show = []
    if engineer_name_cols:
        cols_to_show.append(engineer_name_cols[0])
    if last_name_cols:
        cols_to_show.append(last_name_cols[0])
    cols_to_show += ["Outcode", "postalArea"]
    print(engineers_with_area[cols_to_show].head())

    # 5. Normalise schedule master into cycles
    print("\n--- Normalised schedule cycles (preview) ---")
    cycles_df = build_schedule_cycles(schedule)
    print(cycles_df.head(10))

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    cycles_path = OUTPUT_DIR / "schedule_cycles_preview.xlsx"
    cycles_df.to_excel(cycles_path, index=False)
    print(f"\nSaved schedule cycles preview to: {cycles_path}")

    # 6. Build geo maps (outcode coords & area centroids)
    print("\n--- Building geo maps (outcodes & area centroids) ---")
    outcode_coords, area_centroids = build_geo_maps(outcodes)
    print(f"Outcode coords count: {len(outcode_coords)}")
    print(f"Area centroids count: {len(area_centroids)}")

    # 7. Build distance-based engineer ordering per area
    print("\n--- Building distance-based engineer list per area ---")
    area_eng_map, all_engineers = build_area_engineer_distance_map(
        engineers_with_area, outcode_coords, area_centroids, cycles_df
    )
    example_area = next(iter(area_eng_map.keys()))
    print(f"Example area '{example_area}' nearest engineers:", area_eng_map[example_area][:5])

    # 8. Assign jobs to engineers based on cycles + job_data + distance
    print("\n--- Assigning jobs to engineers (distance-based, preview) ---")
    assignments_df, summary_df = assign_jobs_to_engineers(
        cycles_df, jobs_with_area, area_eng_map, all_engineers, monthly_cap
)


    assignments_path = OUTPUT_DIR / "engineer_assignments_preview.xlsx"
    summary_path = OUTPUT_DIR / "engineer_monthly_summary.xlsx"
    assignments_df.to_excel(assignments_path, index=False)
    summary_df.to_excel(summary_path, index=False)
    print(f"Saved engineer assignments preview to: {assignments_path}")
    print(f"Saved engineer monthly summary to:    {summary_path}")

    # 9. Rebuild Hayley-style schedule with updated job counts
    print("\n--- Rebuilding Hayley-style schedule with updated job counts ---")
    updated_schedule = rebuild_hayley_schedule(schedule, assignments_df)

    updated_schedule_path = OUTPUT_DIR / "SERVICE_SCHEDULE_UPDATED.xlsx"
    updated_schedule.to_excel(updated_schedule_path, index=False)
    print(f"Saved updated Hayley-style schedule to: {updated_schedule_path}")

    print("\n✅ Preview complete.")
    print("You can now open the files in data/output to inspect the results.")


if __name__ == "__main__":
    main()

