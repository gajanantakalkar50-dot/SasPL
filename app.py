from flask import Flask, render_template, request, redirect, session, url_for, flash, send_file
import pandas as pd
import os
import io

app = Flask(__name__)
app.secret_key = "sas_secret_key"

# ---------- PATHS ----------
DATA_FOLDER = "data"
os.makedirs(DATA_FOLDER, exist_ok=True)
USERS_FILE = os.path.join(DATA_FOLDER, "users.xlsx")
EXCEL_FILE = os.path.join(DATA_FOLDER, "project_reports.xlsx")


# ---------- HELPER FUNCTIONS ----------
def read_excel_safe(path, sheet):
    """Safely read an Excel sheet; return empty DataFrame if missing."""
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except Exception as e:
        print(f"⚠️ Error reading {sheet}: {e}")
        return pd.DataFrame()


def save_excel(df, path, sheet):
    """Save a DataFrame to a specific sheet (replace)."""
    with pd.ExcelWriter(path, mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=sheet, index=False)


# ---------- AUTH ----------
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"].strip().lower()
        password = request.form["password"].strip()

        if not os.path.exists(USERS_FILE):
            flash("❌ 'users.xlsx' file not found in /data folder!", "danger")
            return redirect("/login")

        users = pd.read_excel(USERS_FILE)
        users.columns = [c.lower().strip() for c in users.columns]

        if not {"username", "password", "role"}.issubset(users.columns):
            flash("❌ users.xlsx must have columns: username, password, role", "danger")
            return redirect("/login")

        users["username"] = users["username"].astype(str).str.strip().str.lower()
        users["password"] = users["password"].astype(str).str.strip()
        users["role"] = users["role"].astype(str).str.strip().str.lower()

        match = users[
            (users["username"] == username) &
            (users["password"] == password)
        ]

        if not match.empty:
            role = match.iloc[0]["role"]
            session["username"] = username
            session["role"] = role
            flash(f"✅ Welcome {username.capitalize()}! Logged in as {role}.", "success")

            if role == "engineer":
                return redirect("/daily")
            elif role == "manager":
                return redirect("/approve")
            else:
                return redirect("/")
        else:
            flash("❌ Invalid username or password.", "danger")
            return redirect("/login")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("✅ Logged out successfully.", "info")
    return redirect("/login")


# ---------- LOGIN DECORATOR ----------
def login_required(role=None):
    def decorator(func):
        def wrapper(*args, **kwargs):
            if "username" not in session:
                flash("⚠️ Please login first!", "warning")
                return redirect("/login")
            if role and session.get("role") != role:
                flash("🚫 Unauthorized access!", "danger")
                return redirect("/")
            return func(*args, **kwargs)
        wrapper.__name__ = func.__name__
        return wrapper
    return decorator


# ---------- HOME ----------
@app.route("/")
def home():
    return render_template("menu.html")


# ---------- PROJECT REPORT FORM (MANAGER) ----------
@app.route("/project", methods=["GET", "POST"])
@login_required("manager")
def project_form():
    if request.method == "POST":
        project_name = request.form["project_name"]
        engineer_name = request.form["engineer_name"]
        assigned_date = request.form["assigned_date"]
        target_date = request.form["target_date"]

        new_entry = pd.DataFrame([{
            "Project Name": project_name,
            "Engineer Name": engineer_name,
            "Assigned Date": assigned_date,
            "Target Date": target_date
        }])

        if os.path.exists(EXCEL_FILE):
            df_old = read_excel_safe(EXCEL_FILE, "Projects")
            df_all = pd.concat([df_old, new_entry], ignore_index=True)
        else:
            df_all = new_entry

        with pd.ExcelWriter(EXCEL_FILE, mode="a", if_sheet_exists="overlay") as writer:
            df_all.to_excel(writer, sheet_name="Projects", index=False)

        flash("✅ Project report added successfully!", "success")
        return redirect("/project")

    return render_template("project_form.html")


# ---------- DAILY FORM (ENGINEER) ----------
@app.route("/daily")
@login_required("engineer")
def daily_form():
    df = read_excel_safe(EXCEL_FILE, "Projects")
    projects = sorted(df["Project Name"].dropna().unique().tolist()) if not df.empty else []
    return render_template("daily_form.html", project_names=projects)


@app.route("/submit_daily", methods=["POST"])
@login_required("engineer")
def submit_daily():
    username = session.get("username")
    project = request.form["project_name"]
    report_date = request.form["report_date"]

    descs = request.form.getlist("task_desc[]")
    status = request.form.getlist("task_status[]")
    percents = request.form.getlist("task_percent[]")
    eng_remarks = request.form.getlist("remark_engineer[]")

    rows = []
    for i in range(len(descs)):
        rows.append({
            "Engineer": username,
            "Project": project,
            "Date": report_date,
            "Task": descs[i],
            "Status": status[i],
            "%": percents[i],
            "Engineer Remark": eng_remarks[i],
            "Manager Remark": "",
            "Approval": "Pending"
        })

    df_new = pd.DataFrame(rows)

    if os.path.exists(EXCEL_FILE):
        df_old = read_excel_safe(EXCEL_FILE, "DailyChecks")
        df_all = pd.concat([df_old, df_new], ignore_index=True)
    else:
        df_all = df_new

    with pd.ExcelWriter(EXCEL_FILE, mode="a", if_sheet_exists="overlay") as writer:
        df_all.to_excel(writer, sheet_name="DailyChecks", index=False)

    flash("✅ Report submitted for manager approval.", "success")
    return redirect("/daily")


# ---------- MANAGER APPROVAL ----------
@app.route("/approve")
@login_required("manager")
def approve_page():
    df = read_excel_safe(EXCEL_FILE, "DailyChecks")

    if df.empty:
        flash("No reports available for approval.", "info")
        return render_template("approve_tasks.html", reports=[])

    if "Approval" not in df.columns:
        df["Approval"] = ""

    pending = df[df["Approval"].astype(str).str.lower().isin(["pending", ""])]
    if pending.empty:
        flash("✅ All tasks have been approved or rejected.", "info")

    pending = pending.copy()
    pending["real_index"] = pending.index

    reports = pending.to_dict(orient="records")
    return render_template("approve_tasks.html", reports=reports)


@app.route("/approve_task", methods=["POST"])
@login_required("manager")
def approve_task():
    action = request.form.get("action")
    remark = request.form.get("manager_remark", "")
    index_str = request.form.get("index", "")

    if not index_str.strip():
        flash("❌ Task index missing!", "danger")
        return redirect(url_for("approve_page"))

    try:
        idx = int(index_str)
    except ValueError:
        flash("❌ Invalid index format!", "danger")
        return redirect(url_for("approve_page"))

    df = read_excel_safe(EXCEL_FILE, "DailyChecks")

    if idx not in df.index:
        flash("❌ Invalid task reference!", "danger")
        return redirect(url_for("approve_page"))

    df.loc[idx, "Approval"] = action
    df.loc[idx, "Manager Remark"] = remark

    with pd.ExcelWriter(EXCEL_FILE, mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="DailyChecks", index=False)

    flash(f"✅ Task #{idx + 1} marked as {action}.", "success")
    return redirect(url_for("approve_page"))


# ---------- VIEW REPORTS ----------
@app.route("/view", methods=["GET", "POST"])
@login_required()
def view_reports():
    df = read_excel_safe(EXCEL_FILE, "DailyChecks")
    if df.empty:
        return render_template("view_reports.html", reports=[])

    if session.get("role") == "engineer":
        df = df[df["Engineer"] == session.get("username")]

    if request.method == "POST":
        project = request.form.get("project_name", "").strip()
        engineer = request.form.get("engineer_name", "").strip()
        start_date = request.form.get("start_date")
        end_date = request.form.get("end_date")

        if project:
            df = df[df["Project"].str.contains(project, case=False, na=False)]
        if engineer:
            df = df[df["Engineer"].str.contains(engineer, case=False, na=False)]
        if start_date:
            df = df[df["Date"] >= start_date]
        if end_date:
            df = df[df["Date"] <= end_date]

    session["filtered_data"] = df.to_dict(orient="records")
    return render_template("view_reports.html", reports=df.to_dict(orient="records"))


# ---------- EXPORT TO EXCEL ----------
@app.route("/export_excel")
@login_required()
def export_excel():
    if "filtered_data" not in session or not session["filtered_data"]:
        flash("No data to export!", "warning")
        return redirect("/view")

    df = pd.DataFrame(session["filtered_data"])
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Reports")

    output.seek(0)
    return send_file(
        output,
        download_name="Project_Reports.xlsx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ---------- RUN APP ----------
if __name__ == "__main__":
    app.run(debug=True)
