import io
from pathlib import Path

import pandas as pd
from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    session,
    send_file,
    flash,
)
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = "change_this_to_a_strong_random_secret"  # IMPORTANT: change this

# ---- Single user login config ----
USERNAME = "admin"
PASSWORD_HASH = generate_password_hash("changeme123")  # change password before prod

# ---- Global store for latest report ----
LATEST_REPORT_DF = None  # pandas DataFrame


# ---- Agent mapping ----
AGENT_EMAILS = {
    "Jason Canales": "jason.canales@way.com",
    "Joseph Parker": "joseph.parker@way.com",
    "Ayrton Oneal": "ayrton.oneal@way.com",
    "Mike Zimmerman": "mike.zimmerman@way.com",
    "Felecia Boswell": "felecia.boswell@wayinsured.com",
    "Drew Backus": "drew.backus@way.com",
    "Joe Hodges": "joseph.hodges@way.com",
    "Matthew Laushman": "matthew.laushman@way.com",
    "Larry Johnson": "larry.johnson@way.com",
    "Matthew Mandra": "matthew.mandra@way.com",
    "Lee Oday": "lee.oday@way.com",
    "Bryant Hamman": "bryant.hamman@way.com",
    "Kenyatta Rutland": "kenyatta.rutland@way.com",
    "Carlos Villanueva": "carlos.villanueva@way.com",
    "Amr Mohamed": "amr.m@way.com",
    "Lucas Vargas": "lucas.vargas@way.com",
    "Mike Alvarez": "mike.alvarez@way.com",
    "Francesca Cristantiello": "francesca.cristantiello@way.com",
    "Alejandro Gonzalez": "alejandro.gonzalez@way.com",
    "Angie Covilla": "angie.covilla@way.com",
    "Laura Gutierrez": "laura.gutierrez@way.com",
    "Yulieth Jimenez": "yulieth.jimenez@way.com",
    "Julian Hernandez": "julian.hernandez@way.com",
    "Laura Vasquez": "laura.vasquez@way.com",
    "Linda Mayorquin": "linda.mayorquin@way.com",
    "Key Kim": "key.kim@way.com",
    # add more if needed
}


def login_required(f):
    """Simple decorator to protect routes."""
    from functools import wraps

    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)

    return wrapper


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")

        if username == USERNAME and check_password_hash(PASSWORD_HASH, password):
            session["logged_in"] = True
            return redirect(url_for("upload_files"))
        else:
            flash("Invalid username or password", "danger")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/", methods=["GET", "POST"])
@login_required
def upload_files():
    if request.method == "POST":
        new_tile_file = request.files.get("new_tile")
        call_log_file = request.files.get("call_log")

        if not new_tile_file or not call_log_file:
            flash("Please upload both new_tile.csv and call_log.csv", "warning")
            return redirect(url_for("upload_files"))

        try:
            new_tile_df = pd.read_csv(new_tile_file)
            call_log_df = pd.read_csv(call_log_file)
        except Exception as e:
            flash(f"Error reading CSV files: {e}", "danger")
            return redirect(url_for("upload_files"))

        global LATEST_REPORT_DF
        LATEST_REPORT_DF = generate_report(new_tile_df, call_log_df)

        flash("Report generated successfully!", "success")
        return redirect(url_for("dashboard"))

    return render_template("upload.html")


@app.route("/dashboard")
@login_required
def dashboard():
    global LATEST_REPORT_DF
    if LATEST_REPORT_DF is None:
        flash("No report generated yet. Please upload files first.", "info")
        return redirect(url_for("upload_files"))

    # Convert to list of dicts for table
    rows = LATEST_REPORT_DF.to_dict(orient="records")

    # Data for charts (example: Total Outbound Calls per agent)
    agents = LATEST_REPORT_DF["Agent Name"].tolist()
    total_outbound = LATEST_REPORT_DF["Total Outbound Calls"].tolist()
    manual_outbound = LATEST_REPORT_DF["No. of Manual Outbound Calls"].tolist()
    inbound_calls = LATEST_REPORT_DF["No. of Inbound Calls"].tolist()

    return render_template(
        "dashboard.html",
        rows=rows,
        agents=agents,
        total_outbound=total_outbound,
        manual_outbound=manual_outbound,
        inbound_calls=inbound_calls,
    )


@app.route("/agent/<name>")
@login_required
def agent_detail(name):
    global LATEST_REPORT_DF
    if LATEST_REPORT_DF is None:
        flash("No report generated yet.", "info")
        return redirect(url_for("upload_files"))

    # Name from URL may be encoded, decode spaces
    agent_name = name.replace("%20", " ")
    agent_row = LATEST_REPORT_DF[LATEST_REPORT_DF["Agent Name"] == agent_name]

    if agent_row.empty:
        flash(f"No data found for agent: {agent_name}", "warning")
        return redirect(url_for("dashboard"))

    row = agent_row.iloc[0].to_dict()
    return render_template("agent.html", row=row)


@app.route("/download")
@login_required
def download_report():
    global LATEST_REPORT_DF
    if LATEST_REPORT_DF is None:
        flash("No report generated yet.", "info")
        return redirect(url_for("upload_files"))

    # Create Excel in memory with conditional formatting
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        LATEST_REPORT_DF.to_excel(writer, index=False, sheet_name="Report")

    # Apply conditional formatting on Total Calls column (E)
    output.seek(0)
    wb = load_workbook(output)
    ws = wb.active

    total_calls_col = 5  # E
    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    for row_idx in range(2, ws.max_row + 1):
        cell = ws.cell(row=row_idx, column=total_calls_col)
        value = cell.value or 0
        if value < 75:
            cell.fill = red_fill
        else:
            cell.fill = green_fill

    # Save again into BytesIO
    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    return send_file(
        final_output,
        as_attachment=True,
        download_name="agent_summary_fixed.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# -------------------------------
# CORE REPORT LOGIC (YOUR RULES)
# -------------------------------
def generate_report(new_tile: pd.DataFrame, call_log: pd.DataFrame) -> pd.DataFrame:
    """
    Implements your final agreed logic:

    - Total Outbound Calls = Outbound Calls Completed (tile)
    - No. of Manual Outbound Calls = count from call_log email
    - No. of Inbound Calls = Inbound Calls Completed (tile)
    - Total Talk Time (min) = Total Talk Time (min) (tile)
    - If agent not in tile → tile metrics = 0, but manual outbound from log still counts
    - Total Calls = Total Outbound Calls + No. of Inbound Calls
    """

    # Normalize agent names in tile
    new_tile = new_tile.copy()
    new_tile["Agent Name"] = (
        new_tile["Agent Name"].astype(str).str.strip().str.title()
    )

    # Filter tile to only known agents
    tile_filtered = new_tile[new_tile["Agent Name"].isin(AGENT_EMAILS.keys())].copy()

    # Many teams -> duplicates, keep one row (your data for a day is identical per team)
    tile_filtered = tile_filtered.drop_duplicates(subset=["Agent Name"], keep="last")

    # Keep only needed columns from tile
    tile_filtered = tile_filtered[
        [
            "Agent Name",
            "Outbound Calls Completed",
            "Inbound Calls Completed",
            "Total Talk Time (min)",
        ]
    ]

    # Manual outbound from call_log
    call_log = call_log.copy()
    call_log["Handling Agent Email"] = (
        call_log["Handling Agent Email"].astype(str).str.lower().str.strip()
    )
    email_counts = call_log["Handling Agent Email"].value_counts().to_dict()

    # Base DF with all agents
    final_df = pd.DataFrame({"Agent Name": list(AGENT_EMAILS.keys())})
    final_df["Agent Email"] = final_df["Agent Name"].map(AGENT_EMAILS).str.lower()

    # Merge tile metrics
    final_df = final_df.merge(tile_filtered, on="Agent Name", how="left")

    # Fill NaNs for agents not in tile
    final_df["Outbound Calls Completed"] = final_df["Outbound Calls Completed"].fillna(0)
    final_df["Inbound Calls Completed"] = final_df["Inbound Calls Completed"].fillna(0)
    final_df["Total Talk Time (min)"] = final_df["Total Talk Time (min)"].fillna(0)

    # Total Outbound Calls (tile)
    final_df["Total Outbound Calls"] = final_df["Outbound Calls Completed"]

    # No. of Manual Outbound Calls (call_log count)
    final_df["No. of Manual Outbound Calls"] = final_df["Agent Email"].map(
        lambda e: email_counts.get(e, 0)
    )

    # No. of Inbound Calls (tile)
    final_df["No. of Inbound Calls"] = final_df["Inbound Calls Completed"]

    # Talk time
    final_df["Total Talk Time (min)"] = final_df["Total Talk Time (min)"].round(2)

    # Total Calls
    final_df["Total Calls"] = (
        final_df["Total Outbound Calls"] + final_df["No. of Inbound Calls"]
    )

    # Drop helper
    final_df = final_df.drop(columns=["Agent Email", "Outbound Calls Completed", "Inbound Calls Completed"])

    # Final column order
    final_df = final_df[
        [
            "Agent Name",
            "Total Outbound Calls",
            "No. of Manual Outbound Calls",
            "No. of Inbound Calls",
            "Total Calls",
            "Total Talk Time (min)",
        ]
    ]

    return final_df


if __name__ == "__main__":
    # For dev only; use gunicorn or similar in prod
    app.run(debug=True)
