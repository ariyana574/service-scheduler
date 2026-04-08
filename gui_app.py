from pathlib import Path
import traceback

import pandas as pd
import streamlit as st

# Import scheduler module (not just main)
import src.scheduler as scheduler

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "data" / "input"
OUTPUT_DIR = BASE_DIR / "data" / "output"


def check_input_files():
    """Check that all required input files exist."""
    required = {
        "job_data.xlsx": INPUT_DIR / "job_data.xlsx",
        "engineers.xlsx": INPUT_DIR / "engineers.xlsx",
        "SERVICE_SCHEDULE_MASTER.xlsx": INPUT_DIR / "SERVICE_SCHEDULE_MASTER.xlsx",
        "postcode-outcodes.csv": INPUT_DIR / "postcode-outcodes.csv",
    }
    missing = [name for name, path in required.items() if not path.exists()]
    return required, missing


def build_problem_alerts(assignments: pd.DataFrame, summary: pd.DataFrame):
    """Build key problem indicators for the alerts dashboard."""
    alerts = {}

    # Areas with over-cap events
    if not assignments.empty and "IsOverCap" in assignments.columns:
        overcap_area = assignments[assignments["IsOverCap"] == True]  # noqa: E712
        if not overcap_area.empty:
            area_counts = (
                overcap_area.groupby("postalArea")["IsOverCap"]
                .count()
                .sort_values(ascending=False)
            )
            alerts["overcap_areas"] = area_counts

    # Engineers with any over-cap months
    if not summary.empty and "Status" in summary.columns:
        over_eng = summary[summary["Status"] == "Over cap"]
        if not over_eng.empty:
            eng_counts = (
                over_eng.groupby("Engineer")["Status"]
                .count()
                .sort_values(ascending=False)
            )
            alerts["overcap_engineers"] = eng_counts

        # High average utilisation (e.g., > 90%)
        avg_util = (
            summary.groupby("Engineer")["UtilisationPct"]
            .mean()
            .sort_values(ascending=False)
        )
        alerts["high_utilisation"] = avg_util

    return alerts


def create_scenario_engineers_file(first_name: str,
                                   last_name: str,
                                   postcode: str) -> Path:
    """
    Create a temporary engineers file that includes all existing engineers plus
    one additional "scenario" engineer. Returns the path to the new file.
    """
    base_path = INPUT_DIR / "engineers.xlsx"
    if not base_path.exists():
        raise FileNotFoundError(f"Base engineers file not found: {base_path}")

    df = pd.read_excel(base_path)

    # Try to respect existing columns; fill minimal fields for scenario engineer
    new_row = {}
    columns_lower = {c.lower(): c for c in df.columns}

    # Employee Id – use some dummy high ID if column exists
    if "employee id" in columns_lower:
        new_row[columns_lower["employee id"]] = 999999

    # First Name / Last Name
    if "first name" in columns_lower:
        new_row[columns_lower["first name"]] = first_name
    if "last name" in columns_lower:
        new_row[columns_lower["last name"]] = last_name

    # Post Code
    if "post code" in columns_lower:
        new_row[columns_lower["post code"]] = postcode

    # Fill any other columns with blank / None
    for col in df.columns:
        if col not in new_row:
            new_row[col] = None

    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

    scenario_path = INPUT_DIR / "engineers_scenario.xlsx"
    df.to_excel(scenario_path, index=False)

    return scenario_path


def main():
    st.set_page_config(
        page_title="Service Scheduler",
        page_icon="🛠",
        layout="wide",
    )

    # === Company logo in the top-right corner (optional) ===
    logo_path = BASE_DIR / "totalkare.png"  # or change to your logo filename
    if logo_path.exists():
        from PIL import Image
        logo = Image.open(logo_path)

        logo_col1, logo_col2 = st.columns([5, 1])
        with logo_col2:
            st.image(logo, use_container_width=True)
    else:
        # Keep spacing even if no logo
        _, _ = st.columns([5, 1])

    # ===== Title & intro =====
    st.markdown(
        """
# 🛠 Service Scheduler

Internal tool for postcode-based job scheduling using historical job data and engineer locations.

This app:
- Reads the last 12 months of jobs from ServicePro (job_data.xlsx)
- Groups jobs by postcode area (e.g. B, LE, AB)
- Uses the service cycle rules (6-month spacing) from SERVICE_SCHEDULE_MASTER.xlsx
- Assigns work to the nearest engineer by distance (postcode-outcodes.csv)
- Respects a monthly jobs cap per engineer where possible
- Highlights spillovers and any over-capacity months

---
"""
    )

    # ===== Config panel =====
    st.markdown("## ⚙️ Configuration")

    cols_cfg = st.columns(2)

    with cols_cfg[0]:
        monthly_cap = st.number_input(
            "Monthly jobs cap per engineer",
            min_value=20,
            max_value=100,
            value=45,
            step=1,
            help="Maximum jobs per engineer per month before being considered over capacity.",
        )

    with cols_cfg[1]:
        st.write(" ")  # spacer
        st.write(" ")  # spacer
        st.info("You can adjust the monthly cap to test different workload scenarios.")

    st.markdown("---")

    # ===== 1️⃣ Input file status =====
    st.markdown("## 1️⃣ Check required input files")

    required, missing = check_input_files()

    st.write("The app expects these files in the `data/input` folder:")

    for name, path in required.items():
        exists = path.exists()
        status = "✅ found" if exists else "❌ missing"
        st.write(f"- `{name}` – {status}  \n  _{path}_")

    if missing:
        st.error(
            "Some required files are missing. Please place them in `data/input` "
            "before running the scheduler."
        )

    with st.expander("What each input file is used for"):
        st.markdown(
            """
- **job_data.xlsx**  
  Export from ServicePro – last 12 months of completed jobs.  
  Used to count real jobs per postcode area.

- **engineers.xlsx**  
  List of engineers with home postcodes.  
  Used to assign jobs to the nearest engineer and manage capacity.

- **SERVICE_SCHEDULE_MASTER.xlsx**  
  Master service schedule (Hayley’s file).  
  Defines which months each postcode area should be visited (6-month patterns).

- **postcode-outcodes.csv**  
  UK postcode data with latitude/longitude.  
  Used to calculate distance between areas and engineers.
"""
        )

    st.markdown("---")

    # ===== Scenario mode (optional extra engineer) =====
    st.markdown("## 🧪 Scenario mode – add a temporary engineer")

    scenario_enabled = st.checkbox(
        "Enable scenario: run with one extra engineer (temporary only)",
        value=False,
        help="Adds a temporary engineer to the model for 'what if' analysis. "
             "This does not change your original engineers.xlsx file.",
    )

    scenario_path = None
    scenario_desc = ""

    if scenario_enabled:
        col_s1, col_s2, col_s3 = st.columns(3)
        with col_s1:
            scen_first = st.text_input("Scenario engineer first name", value="Scenario")
        with col_s2:
            scen_last = st.text_input("Scenario engineer last name", value="Engineer")
        with col_s3:
            scen_postcode = st.text_input(
                "Scenario engineer home postcode",
                value="B24 0AA",
                help="Full postcode e.g. 'B24 0AA'. Only the outcode is used for distance.",
            )

        scenario_desc = f"{scen_first} {scen_last} at {scen_postcode}"

        st.info(
            "When you run the scheduler, a temporary copy of engineers.xlsx will be created "
            f"including this scenario engineer: **{scenario_desc}**. "
            "It is only used for this run."
        )

    st.markdown("---")

    # ===== 2️⃣ Run scheduler =====
    st.markdown("## 2️⃣ Run the scheduler")

    run_clicked = st.button("🚀 Run Scheduler")

    if run_clicked:
        # Re-check missing in case user added files after page load
        _, missing_now = check_input_files()

        if missing_now:
            st.error("Cannot run – required input files are missing.")
            st.stop()

        try:
            engineers_override_path = None

            if scenario_enabled:
                if not scen_postcode.strip():
                    st.error("Please provide a postcode for the scenario engineer.")
                    st.stop()
                # Create scenario engineers file
                engineers_override_path = create_scenario_engineers_file(
                    scen_first.strip(), scen_last.strip(), scen_postcode.strip()
                )

            st.info("Running scheduler… this may take a moment.")
            with st.spinner("Processing schedules, assigning jobs, and updating outputs..."):
                scheduler.main(
                    monthly_cap=int(monthly_cap),
                    engineers_override_path=engineers_override_path,
                )

            st.success("Scheduler finished successfully.")

            if scenario_enabled:
                st.warning(
                    f"You are viewing results for a **scenario run** including a temporary engineer: {scenario_desc}"
                )

            # ===== 3️⃣ Load key outputs =====
            st.markdown("## 3️⃣ Review results")

            assignments_path = OUTPUT_DIR / "engineer_assignments_preview.xlsx"
            summary_path = OUTPUT_DIR / "engineer_monthly_summary.xlsx"
            schedule_path = OUTPUT_DIR / "SERVICE_SCHEDULE_UPDATED.xlsx"

            assignments = pd.read_excel(assignments_path) if assignments_path.exists() else pd.DataFrame()
            summary = pd.read_excel(summary_path) if summary_path.exists() else pd.DataFrame()
            schedule = pd.read_excel(schedule_path) if schedule_path.exists() else pd.DataFrame()

            # --- KPIs ---
            if not assignments.empty and "Jobs_Assigned" in assignments.columns:
                total_jobs = int(assignments["Jobs_Assigned"].sum())
                total_spillovers = int(assignments["IsSpillover"].sum()) if "IsSpillover" in assignments.columns else 0
                total_overcap = int(assignments["IsOverCap"].sum()) if "IsOverCap" in assignments.columns else 0
            else:
                total_jobs = total_spillovers = total_overcap = 0

            if not summary.empty and "Engineer" in summary.columns:
                num_engineers = summary["Engineer"].nunique()
            else:
                num_engineers = 0

            st.markdown("### 🔍 Summary at a glance")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total jobs scheduled", f"{total_jobs}")
            c2.metric("Engineers used", f"{num_engineers}")
            c3.metric("Spillover visits", f"{total_spillovers}")
            c4.metric("Over-cap events", f"{total_overcap}")

            # --- Utilisation bar chart ---
            if not summary.empty and "UtilisationPct" in summary.columns:
                st.markdown("### 📊 Average utilisation by engineer")

                util = (
                    summary.groupby("Engineer")["UtilisationPct"]
                    .mean()
                    .reset_index()
                    .sort_values("UtilisationPct", ascending=False)
                )

                if not util.empty:
                    st.bar_chart(util.set_index("Engineer")["UtilisationPct"])
                else:
                    st.info("No utilisation data available.")
            else:
                st.info("Engineer monthly summary file not found or empty.")

            # --- Problem alerts dashboard ---
            st.markdown("### 🚨 Problem alerts")

            alerts = build_problem_alerts(assignments, summary)

            col_a, col_b = st.columns(2)

            # Over-cap areas
            with col_a:
                st.markdown("**Areas with over-cap events**")
                overcap_areas = alerts.get("overcap_areas")
                if overcap_areas is not None and not overcap_areas.empty:
                    st.dataframe(
                        overcap_areas.rename("OverCap_Count").head(10)
                        .to_frame()
                    )
                else:
                    st.write("No areas with over-capacity events.")

            # Engineers with over-cap months
            with col_b:
                st.markdown("**Engineers with over-cap months**")
                overcap_eng = alerts.get("overcap_engineers")
                if overcap_eng is not None and not overcap_eng.empty:
                    st.dataframe(
                        overcap_eng.rename("OverCap_Months").head(10)
                        .to_frame()
                    )
                else:
                    st.write("No engineers over capacity.")

            # High utilisation list
            st.markdown("**Engineers by average utilisation**")
            high_util = alerts.get("high_utilisation")
            if high_util is not None and not high_util.empty:
                st.dataframe(
                    high_util.rename("Avg_UtilisationPct")
                    .round(1)
                    .to_frame()
                    .head(20)
                )
            else:
                st.write("No utilisation data available.")

            # --- Preview updated service schedule ---
            if not schedule.empty:
                st.markdown("### 📑 Updated service schedule (sample)")

                cols = [
                    c
                    for c in [
                        "Postcode",
                        "Area Name",
                        "Region",
                        "Jan",
                        "Feb",
                        "Mar",
                        "Apr",
                        "May",
                        "Jun",
                        "Jul",
                        "Aug",
                        "Sep",
                        "Oct",
                        "Nov",
                        "Dec",
                    ]
                    if c in schedule.columns
                ]

                st.dataframe(schedule[cols].head(20))
            else:
                st.info("Updated service schedule file not found or empty.")

            # --- Download buttons ---
            st.markdown("### 💾 Download key outputs")

            col_d1, col_d2, col_d3 = st.columns(3)

            if schedule_path.exists():
                with open(schedule_path, "rb") as f:
                    col_d1.download_button(
                        label="📥 Download updated service schedule",
                        data=f,
                        file_name="SERVICE_SCHEDULE_UPDATED.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            else:
                col_d1.write("SERVICE_SCHEDULE_UPDATED.xlsx not found.")

            if assignments_path.exists():
                with open(assignments_path, "rb") as f:
                    col_d2.download_button(
                        label="📥 Download engineer assignments",
                        data=f,
                        file_name="engineer_assignments_preview.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            else:
                col_d2.write("engineer_assignments_preview.xlsx not found.")

            if summary_path.exists():
                with open(summary_path, "rb") as f:
                    col_d3.download_button(
                        label="📥 Download engineer monthly summary",
                        data=f,
                        file_name="engineer_monthly_summary.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            else:
                col_d3.write("engineer_monthly_summary.xlsx not found.")

        except Exception as e:
            st.error("An error occurred while running the scheduler.")
            st.code("".join(traceback.format_exception(type(e), e, e.__traceback__)))

    else:
        st.info("To run the scheduler, adjust the configuration if needed and click the **🚀 Run Scheduler** button above once all input files are ready.")


if __name__ == "__main__":
    main()
