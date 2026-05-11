import pandas as pd
import random
from datetime import datetime, timedelta

# ==============================
# SOC INCIDENT TRACKER v1
# ==============================

# 1. Read CSV file
df = pd.read_csv("telecom_churn.csv")

# 2. Clean dataset
df = df.dropna(how="all")
df = df.drop_duplicates()
df = df.reset_index(drop=True)

# 3. Sample SOC data
site_names = ["Site-A", "Site-B", "Site-C", "Site-D", "Site-E"]

issue_types = [
    "Power Failure",
    "Network Down",
    "Low Battery",
    "Communication Loss",
    "High Temperature",
    "Router Offline"
]

engineers = ["Ali", "Ahmed", "Sara", "Usman", "Fatima"]

statuses = ["Open", "Closed", "In Progress"]

remarks_list = [
    "Issue reported by monitoring system",
    "Engineer assigned for checking",
    "Site restored after troubleshooting",
    "Further investigation required",
    "Remote support provided"
]

# 4. Create alarm and close times
base_time = datetime.now()

alarm_times = []
close_times = []

for i in range(len(df)):
    alarm_time = base_time - timedelta(hours=random.randint(1, 120))
    
    status = random.choice(statuses)

    if status == "Closed":
        close_time = alarm_time + timedelta(minutes=random.randint(15, 360))
    else:
        close_time = pd.NaT

    alarm_times.append(alarm_time)
    close_times.append(close_time)

# 5. Create SOC incident columns
df["Site Name"] = [random.choice(site_names) for _ in range(len(df))]
df["Issue Type"] = [random.choice(issue_types) for _ in range(len(df))]
df["Alarm Time"] = alarm_times
df["Close Time"] = close_times
df["Assigned Engineer"] = [random.choice(engineers) for _ in range(len(df))]
df["Status"] = [
    "Closed" if pd.notna(close_times[i]) else random.choice(["Open", "In Progress"])
    for i in range(len(df))
]

# 6. Calculate downtime only for closed incidents
df["Downtime (Minutes)"] = (
    df["Close Time"] - df["Alarm Time"]
).dt.total_seconds() / 60

df["Downtime (Minutes)"] = df["Downtime (Minutes)"].fillna(0).round(2)

# 7. Add remarks
df["Remarks"] = [random.choice(remarks_list) for _ in range(len(df))]

# 8. Create final SOC report with only useful columns
final_report = df[
    [
        "Site Name",
        "Issue Type",
        "Alarm Time",
        "Assigned Engineer",
        "Status",
        "Close Time",
        "Downtime (Minutes)",
        "Remarks"
    ]
]

# 9. Create summary report
summary = final_report.groupby("Status").agg(
    Total_Incidents=("Status", "count"),
    Average_Downtime_Minutes=("Downtime (Minutes)", "mean")
).reset_index()

# 10. Export to Excel with two sheets
with pd.ExcelWriter("soc_incident_report.xlsx", engine="openpyxl") as writer:
    final_report.to_excel(writer, sheet_name="Incident Report", index=False)
    summary.to_excel(writer, sheet_name="Summary", index=False)

print("SOC Incident Report created successfully.")
print("File name: soc_incident_report.xlsx")