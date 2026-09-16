"""Quick test: generate sample Excel files and import them into the database."""

import sys
from pathlib import Path

import pandas as pd

OUT = Path("sample_data")
OUT.mkdir(exist_ok=True)

# --- Semesters ---
pd.DataFrame(
    {"academic_year": ["2026/2027", "2025/2026"], "semester": [1, 2]}
).to_excel(OUT / "semesters.xlsx", index=False)

# --- Programmes ---
pd.DataFrame(
    {"code": ["CE", "EE", "ME", "IE"], "name": [
        "Civil Engineering",
        "Electrical Engineering",
        "Mechanical Engineering",
        "Industrial Engineering",
    ]}
).to_excel(OUT / "programmes.xlsx", index=False)

# --- Student Groups ---
pd.DataFrame(
    {"programme_code": ["CE", "CE", "CE", "CE", "CE", "CE", "CE",
                         "EE", "EE", "EE",
                         "ME", "ME", "ME", "ME", "ME", "ME",
                         "IE", "IE"],
     "group_code": ["A1", "A2", "A3", "A4", "A5", "A6", "A7",
                     "C1", "C2", "C3",
                     "D1", "D2", "D3", "D4", "D5", "D6",
                     "E1", "E2"]}
).to_excel(OUT / "student_groups.xlsx", index=False)

# --- Programme Courses ---
pd.DataFrame(
    {"programme_code": ["CE", "CE", "CE", "EE", "EE", "ME", "ME", "IE"],
     "course_code": ["MT161", "TG201", "ST101", "MT161", "EE101", "MT161", "ME101", "IE101"]}
).to_excel(OUT / "programme_courses.xlsx", index=False)

# --- Venues ---
pd.DataFrame(
    {"name": ["NB102", "NB203", "EQ002", "EQ005", "TW101", "TW107"],
     "capacity": [200, 150, 80, 80, 60, 60]}
).to_excel(OUT / "venues.xlsx", index=False)

# --- Master Timetable ---
pd.DataFrame(
    {"course_code": ["MT161", "MT161", "MT161", "TG201", "TG201",
                     "EE101", "ME101", "IE101", "ST101"],
     "activity_type": ["Lecture", "Tutorial", "Practical", "Lecture", "Practical",
                       "Lecture", "Lecture", "Lecture", "Seminar"],
     "day": ["Monday", "Monday", "Wednesday", "Tuesday", "Thursday",
             "Friday", "Friday", "Wednesday", "Thursday"],
     "start_time": ["08:00", "14:00", "10:00", "09:00", "14:00",
                    "09:00", "09:00", "09:00", "15:00"],
     "end_time": ["10:00", "16:00", "13:00", "11:00", "17:00",
                  "11:00", "11:00", "11:00", "17:00"],
     "venue": ["NB102", "NB203", "EQ002", "NB102", "EQ005",
               "NB203", "NB203", "NB102", "NB203"],
     "group": ["ALL", "A1,A2,C1,D1", "A1,C1",
               "ALL", "C1,C2",
               "ALL", "ALL", "ALL", "C1,C2,D1,D2,E1,E2"]}
).to_excel(OUT / "master_timetable.xlsx", index=False)

# --- Workshop Allocation ---
pd.DataFrame(
    {"course_code": ["TG201", "TG201"],
     "group_code": ["C1", "C2"],
     "day": ["Tuesday", "Tuesday"],
     "start_time": ["14:00", "14:00"],
     "end_time": ["17:00", "17:00"],
     "venue": ["TW101", "TW107"]}
).to_excel(OUT / "workshop_allocation.xlsx", index=False)

# --- TD Allocation ---
pd.DataFrame(
    {"course_code": ["TG201"],
     "group_code": ["A1"],
     "day": ["Tuesday"],
     "start_time": ["14:00"],
     "end_time": ["17:00"],
     "venue": ["TW101"]}
).to_excel(OUT / "td_allocation.xlsx", index=False)

print("Sample Excel files created in sample_data/")
print("\nRun these commands in order:")
print("  python manage.py import_semesters       --file sample_data/semesters.xlsx")
print("  python manage.py import_programmes      --file sample_data/programmes.xlsx")
print("  python manage.py import_student_groups  --file sample_data/student_groups.xlsx")
print("  python manage.py import_programme_courses --file sample_data/programme_courses.xlsx")
print("  python manage.py import_venues          --file sample_data/venues.xlsx")
print("  python manage.py import_master_timetable --file sample_data/master_timetable.xlsx --semester 1")
print("  python manage.py import_workshop_allocation --file sample_data/workshop_allocation.xlsx --semester 1")
print("  python manage.py import_td_allocation   --file sample_data/td_allocation.xlsx --semester 1")