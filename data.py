import pandas as pd
import datetime

grade = ["AM1", "AM2"]
parish = ["St. James", "Kingston", "Hanover"]
area = ["Montego Bay", "Falmouth", "Sandy Bay"]
visits = ["BEN.", "CON.", "REG.", "EDUC.", "R.I.", "COMP.", "C/L"]
keys = [
    "name", "grade", "period", "area", "parish_office", "week_ended", "files_carried_fwd", "files_received",
    "total", "files_cleared", "files_on_hand", "total_effective", "total_ineffective", "results_total", "objectives"
]


# Get the next specific weekday after today.
def next_weekday(d, weekday):
    days_ahead = weekday - d.weekday()
    if days_ahead < 0:  # Target day already happened this week
        days_ahead += 7
    return d + datetime.timedelta(days_ahead)


# Set period date to next Monday
period_start = next_weekday(datetime.date.today(), 0)
period_end = period_start + datetime.timedelta(days=4)
week_ended_date = datetime.date.today()

# Business data
df = pd.read_csv("files/Employers.csv")
df["NAME_OF_EMPLOYER"] = df["NAME_OF_EMPLOYER"].str.split().str.join(' ')
df["REF_NO"] = df["REF_NO"].str.strip()
df.sort_values(by=["NAME_OF_EMPLOYER"], ascending=True, inplace=True)
businesses_dict = df.to_dict("list")


# TODO #1 Add holidays and days off

