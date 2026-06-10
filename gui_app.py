from pathlib import Path
import traceback
import tempfile

import pandas as pd
import streamlit as st

import src.scheduler as scheduler

BASE_DIR = Path(__file__).resolve().parent


# Keep PM and calib in separate folders so their uploads don't clash - they end
# up with the same filenames once saved.
def get_session_input_dir(mode: str) -> Path:
    key = f"input_dir_{mode}"
    if key not in st.session_state:
        st.session_state[key] = tempfile.mkdtemp(prefix=f"scheduler_inputs_{mode}_")
    return Path(st.session_state[key])


def get_session_output_dir(mode: str) -> Path:
    key = f"output_dir_{mode}"
    if key not in st.session_state:
        st.session_state[key] = tempfile.mkdtemp(prefix=f"scheduler_outputs_{mode}_")
    return Path(st.session_state[key])


# Each tab asks for its own filenames, but the scheduler only knows the plain
# names (job_data.xlsx etc), so we save every upload under that plain name.
PM_REQUIRED_FILES = {
    "job_data_pm.xlsx": {
        "canonical": "job_data.xlsx",
        "desc": "Export from ServicePro - last 12 months of completed PM jobs.",
    },
    "engineers_pm.xlsx": {
        "canonical": "engineers.xlsx",
        "desc": "List of PM engineers with home postcodes.",
    },
    "SERVICE_SCHEDULE_MASTER_PM.xlsx": {
        "canonical": "SERVICE_SCHEDULE_MASTER.xlsx",
        "desc": "Master PM service schedule.",
    },
    "postcode-outcodes.csv": {
        "canonical": "postcode-outcodes.csv",
        "desc": "UK postcode data with latitude/longitude.",
    },
}

CALIB_REQUIRED_FILES = {
    "job_data_calib.xlsx": {
        "canonical": "job_data.xlsx",
        "desc": "Export from ServicePro - last 12 months of completed calibration jobs.",
    },
    "engineers_cal.xlsx": {
        "canonical": "engineers.xlsx",
        "desc": "List of calibration engineers with home postcodes.",
    },
    "SERVICE_SCHEDULE_MASTER_calib.xlsx": {
        "canonical": "SERVICE_SCHEDULE_MASTER.xlsx",
        "desc": "Master calibration service schedule.",
    },
    "postcode-outcodes.csv": {
        "canonical": "postcode-outcodes.csv",
        "desc": "UK postcode data with latitude/longitude.",
    },
}


def get_required_files(mode: str) -> dict:
    return PM_REQUIRED_FILES if mode == "pm" else CALIB_REQUIRED_FILES


def save_uploaded_file(uploaded_file, dest_dir: Path, canonical_name: str) -> Path:
    dest = dest_dir / canonical_name
    with open(dest, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return dest


def check_uploaded_files(input_dir: Path, required_files: dict):
    missing = [
        expected
        for expected, info in required_files.items()
        if not (input_dir / info["canonical"]).exists()
    ]
    return required_files, missing


def build_problem_alerts(assignments: pd.DataFrame, summary: pd.DataFrame):
    alerts = {}

    if not assignments.empty and "IsOverCap" in assignments.columns:
        overcap_area = assignments[assignments["IsOverCap"] == True]  # noqa: E712
        if not overcap_area.empty:
            area_counts = (
                overcap_area.groupby("postalArea")["IsOverCap"]
                .count()
                .sort_values(ascending=False)
            )
            alerts["overcap_areas"] = area_counts

    if not summary.empty and "Status" in summary.columns:
        over_eng = summary[summary["Status"] == "Over cap"]
        if not over_eng.empty:
            eng_counts = (
                over_eng.groupby("Engineer")["Status"]
                .count()
                .sort_values(ascending=False)
            )
            alerts["overcap_engineers"] = eng_counts

        avg_util = (
            summary.groupby("Engineer")["UtilisationPct"]
            .mean()
            .sort_values(ascending=False)
        )
        alerts["high_utilisation"] = avg_util

    return alerts


def create_scenario_engineers_file(first_name: str,
                                   last_name: str,
                                   postcode: str,
                                   input_dir: Path) -> Path:
    base_path = input_dir / "engineers.xlsx"
    if not base_path.exists():
        raise FileNotFoundError("engineers file not found in upload directory.")

    df = pd.read_excel(base_path)
    new_row = {}
    columns_lower = {c.lower(): c for c in df.columns}

    if "employee id" in columns_lower:
        new_row[columns_lower["employee id"]] = 999999
    if "first name" in columns_lower:
        new_row[columns_lower["first name"]] = first_name
    if "last name" in columns_lower:
        new_row[columns_lower["last name"]] = last_name
    if "post code" in columns_lower:
        new_row[columns_lower["post code"]] = postcode

    for col in df.columns:
        if col not in new_row:
            new_row[col] = None

    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

    scenario_path = input_dir / "engineers_scenario.xlsx"
    df.to_excel(scenario_path, index=False)
    return scenario_path


# Draws one whole tab. mode ("pm"/"calib") goes on the front of every widget key
# because Streamlit falls over if two widgets share a key.
def render_scheduler_tab(mode: str, label: str, intro_md: str):
    input_dir = get_session_input_dir(mode)
    output_dir = get_session_output_dir(mode)
    required_files = get_required_files(mode)

    st.markdown(intro_md)

    st.markdown(f"## {label} configuration")
    cols_cfg = st.columns(2)
    with cols_cfg[0]:
        monthly_cap = st.number_input(
            "Monthly jobs cap per engineer",
            min_value=20,
            max_value=100,
            value=45,
            step=1,
            help="Maximum jobs per engineer per month before being considered over capacity.",
            key=f"{mode}_monthly_cap",
        )
    with cols_cfg[1]:
        st.write(" ")
        st.write(" ")
        st.info("You can adjust the monthly cap to test different workload scenarios.")

    st.markdown("---")

    st.markdown("## 1. Upload required input files")
    st.write(
        f"Upload all four {label} files below, named exactly as shown. They are "
        "remembered for this session - you don't need to re-upload unless you refresh the page."
    )

    upload_cols = st.columns(2)
    file_slots = list(required_files.items())
    for i, (expected_name, info) in enumerate(file_slots):
        col = upload_cols[i % 2]
        canonical = info["canonical"]
        description = info["desc"]
        dest = input_dir / canonical
        already_uploaded = dest.exists()

        with col:
            status = "uploaded" if already_uploaded else "needed"
            uploaded = st.file_uploader(
                f"`{expected_name}` ({status})",
                key=f"{mode}_upload_{expected_name}",
                type=["xlsx", "csv"],
                help=description,
            )
            if uploaded is not None:
                if uploaded.name != expected_name:
                    st.error(
                        f"Expected `{expected_name}` but got `{uploaded.name}`. "
                        "Please rename the file or upload it to the correct slot."
                    )
                else:
                    save_uploaded_file(uploaded, input_dir, canonical)
                    st.success(f"Saved {expected_name}")
            elif already_uploaded:
                st.caption("Already uploaded this session")
            else:
                st.caption(f"_{description}_")

    _, missing = check_uploaded_files(input_dir, required_files)

    st.markdown("---")

    st.markdown("### File status")
    status_cols = st.columns(4)
    for i, (expected_name, info) in enumerate(required_files.items()):
        with status_cols[i]:
            if (input_dir / info["canonical"]).exists():
                st.success(expected_name)
            else:
                st.error(expected_name)

    if missing:
        st.warning(
            f"**{len(missing)} file(s) still needed:** {', '.join(missing)}. "
            "Please upload them above before running the scheduler."
        )

    with st.expander("What each input file is used for"):
        names = list(required_files.keys())
        st.markdown(
            f"""
- **{names[0]}** - Export from ServicePro. Last 12 months of completed jobs. Used to count real jobs per postcode area.
- **{names[1]}** - List of engineers with home postcodes. Used to assign jobs to the nearest engineer and manage capacity.
- **{names[2]}** - Master service schedule. Defines which months each postcode area should be visited.
- **{names[3]}** - UK postcode data with lat/lon. Used to calculate distances between areas and engineers.
"""
        )

    st.markdown("---")

    st.markdown("## Scenario mode - add a temporary engineer")

    scenario_enabled = st.checkbox(
        "Enable scenario: run with one extra engineer (temporary only)",
        value=False,
        help="Adds a temporary engineer to the model for 'what if' analysis. "
             "Does not change your uploaded engineers file.",
        key=f"{mode}_scenario_enabled",
    )

    scenario_desc = ""
    scen_first = scen_last = scen_postcode = ""

    if scenario_enabled:
        col_s1, col_s2, col_s3 = st.columns(3)
        with col_s1:
            scen_first = st.text_input("Scenario engineer first name", value="Scenario",
                                       key=f"{mode}_scen_first")
        with col_s2:
            scen_last = st.text_input("Scenario engineer last name", value="Engineer",
                                      key=f"{mode}_scen_last")
        with col_s3:
            scen_postcode = st.text_input(
                "Scenario engineer home postcode",
                value="B24 0AA",
                help="Full postcode e.g. 'B24 0AA'. Only the outcode is used for distance.",
                key=f"{mode}_scen_postcode",
            )

        scenario_desc = f"{scen_first} {scen_last} at {scen_postcode}"
        st.info(
            f"Scheduler will run with a temporary extra engineer: **{scenario_desc}**. "
            "This is only used for this run."
        )

    st.markdown("---")

    st.markdown("## 2. Run the scheduler")

    run_clicked = st.button("Run Scheduler", disabled=bool(missing),
                            key=f"{mode}_run_btn")

    if missing:
        st.caption("Upload all required files to enable the run button.")

    if run_clicked:
        _, missing_now = check_uploaded_files(input_dir, required_files)
        if missing_now:
            st.error("Cannot run - required input files are missing.")
            st.stop()

        try:
            engineers_override_path = None

            if scenario_enabled:
                if not scen_postcode.strip():
                    st.error("Please provide a postcode for the scenario engineer.")
                    st.stop()
                engineers_override_path = create_scenario_engineers_file(
                    scen_first.strip(), scen_last.strip(), scen_postcode.strip(),
                    input_dir,
                )

            st.info("Running scheduler, this may take a moment.")
            with st.spinner("Processing schedules, assigning jobs, and updating outputs..."):
                scheduler.main(
                    monthly_cap=int(monthly_cap),
                    engineers_override_path=engineers_override_path,
                    input_dir=input_dir,
                    output_dir=output_dir,
                )

            st.success("Scheduler finished successfully.")

            if scenario_enabled:
                st.warning(
                    f"You are viewing results for a scenario run including: {scenario_desc}"
                )

            _render_results(output_dir)

        except Exception as e:
            st.error("An error occurred while running the scheduler.")
            st.code("".join(traceback.format_exception(type(e), e, e.__traceback__)))

    elif not missing:
        st.info("All files uploaded. Adjust configuration if needed and click Run Scheduler.")


def _render_results(output_dir: Path):
    st.markdown("## 3. Review results")

    assignments_path = output_dir / "engineer_assignments_preview.xlsx"
    summary_path = output_dir / "engineer_monthly_summary.xlsx"
    schedule_path = output_dir / "SERVICE_SCHEDULE_UPDATED.xlsx"

    assignments = pd.read_excel(assignments_path) if assignments_path.exists() else pd.DataFrame()
    summary = pd.read_excel(summary_path) if summary_path.exists() else pd.DataFrame()
    schedule = pd.read_excel(schedule_path) if schedule_path.exists() else pd.DataFrame()

    if not assignments.empty and "Jobs_Assigned" in assignments.columns:
        total_jobs = int(assignments["Jobs_Assigned"].sum())
        total_spillovers = int(assignments["IsSpillover"].sum()) if "IsSpillover" in assignments.columns else 0
        total_overcap = int(assignments["IsOverCap"].sum()) if "IsOverCap" in assignments.columns else 0
    else:
        total_jobs = total_spillovers = total_overcap = 0

    num_engineers = summary["Engineer"].nunique() if not summary.empty and "Engineer" in summary.columns else 0

    st.markdown("### Summary at a glance")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total jobs scheduled", f"{total_jobs}")
    c2.metric("Engineers used", f"{num_engineers}")
    c3.metric("Spillover visits", f"{total_spillovers}")
    c4.metric("Over-cap events", f"{total_overcap}")

    if not summary.empty and "UtilisationPct" in summary.columns:
        st.markdown("### Average utilisation by engineer")
        util = (
            summary.groupby("Engineer")["UtilisationPct"]
            .mean()
            .reset_index()
            .sort_values("UtilisationPct", ascending=False)
        )
        if not util.empty:
            st.bar_chart(util.set_index("Engineer")["UtilisationPct"])

    st.markdown("### Problem alerts")
    alerts = build_problem_alerts(assignments, summary)

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**Areas with over-cap events**")
        overcap_areas = alerts.get("overcap_areas")
        if overcap_areas is not None and not overcap_areas.empty:
            st.dataframe(overcap_areas.rename("OverCap_Count").head(10).to_frame())
        else:
            st.write("No areas with over-capacity events.")

    with col_b:
        st.markdown("**Engineers with over-cap months**")
        overcap_eng = alerts.get("overcap_engineers")
        if overcap_eng is not None and not overcap_eng.empty:
            st.dataframe(overcap_eng.rename("OverCap_Months").head(10).to_frame())
        else:
            st.write("No engineers over capacity.")

    st.markdown("**Engineers by average utilisation**")
    high_util = alerts.get("high_utilisation")
    if high_util is not None and not high_util.empty:
        st.dataframe(high_util.rename("Avg_UtilisationPct").round(1).to_frame().head(20))

    if not schedule.empty:
        st.markdown("### Updated service schedule (sample)")
        cols = [
            c for c in
            ["Postcode", "Area Name", "Region",
             "Jan", "Feb", "Mar", "Apr", "May", "Jun",
             "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
            if c in schedule.columns
        ]
        st.dataframe(schedule[cols].head(20))

    st.markdown("### Download key outputs")
    col_d1, col_d2, col_d3 = st.columns(3)

    if schedule_path.exists():
        with open(schedule_path, "rb") as f:
            col_d1.download_button(
                label="Download updated service schedule",
                data=f,
                file_name="SERVICE_SCHEDULE_UPDATED.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_sched_{output_dir.name}",
            )
    else:
        col_d1.write("SERVICE_SCHEDULE_UPDATED.xlsx not found.")

    if assignments_path.exists():
        with open(assignments_path, "rb") as f:
            col_d2.download_button(
                label="Download engineer assignments",
                data=f,
                file_name="engineer_assignments_preview.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_assign_{output_dir.name}",
            )
    else:
        col_d2.write("engineer_assignments_preview.xlsx not found.")

    if summary_path.exists():
        with open(summary_path, "rb") as f:
            col_d3.download_button(
                label="Download engineer monthly summary",
                data=f,
                file_name="engineer_monthly_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_summary_{output_dir.name}",
            )
    else:
        col_d3.write("engineer_monthly_summary.xlsx not found.")


PM_INTRO = """
# Preventative Maintenance Scheduler

Internal tool for postcode-based PM job scheduling using historical job data
and engineer locations.

This app:
- Reads the last 12 months of PM jobs from ServicePro (job_data_pm.xlsx)
- Groups jobs by postcode area (e.g. B, LE, AB)
- Uses the service cycle rules (6-month spacing) from SERVICE_SCHEDULE_MASTER_PM.xlsx
- Assigns work to the nearest engineer by distance (postcode-outcodes.csv)
- Respects a monthly jobs cap per engineer where possible
- Highlights spillovers and any over-capacity months

---
"""

CALIB_INTRO = """
# Calibration Scheduler

Internal tool for postcode-based calibration job scheduling. Works the same way as
the PM scheduler but using calibration job data and calibration engineers.

This app:
- Reads the last 12 months of calibration jobs from ServicePro (job_data_calib.xlsx)
- Groups jobs by postcode area (e.g. B, LE, AB)
- Uses the service cycle rules (6-month spacing) from SERVICE_SCHEDULE_MASTER_calib.xlsx
- Assigns work to the nearest calibration engineer by distance (postcode-outcodes.csv)
- Respects a monthly jobs cap per engineer where possible
- Highlights spillovers and any over-capacity months

Upload the calibration-filtered ServicePro export and the calibration engineers
list here. PM files go in the other tab.

---
"""


def main():
    st.set_page_config(
        page_title="Service Scheduler",
        layout="wide",
    )

    logo_path = BASE_DIR / "totalkare.png"
    if logo_path.exists():
        from PIL import Image
        logo = Image.open(logo_path)
        logo_col1, logo_col2 = st.columns([5, 1])
        with logo_col2:
            st.image(logo, use_container_width=True)

    pm_tab, calib_tab = st.tabs(["Preventative Maintenance", "Calibration"])

    with pm_tab:
        render_scheduler_tab("pm", "PM", PM_INTRO)

    with calib_tab:
        render_scheduler_tab("calib", "Calibration", CALIB_INTRO)


if __name__ == "__main__":
    main()