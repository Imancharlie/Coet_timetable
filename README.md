# COET Timetable

A Django-based timetable management and allocation system for the College of Engineering and Technology (COET).

## Features

- Programme, course, semester, and venue management
- Student group management
- Master timetable handling
- Teaching/Workshop allocation with import from Excel files
- Import sample data from `sample_data/` Excel files via management commands
- Report generation (PDF via reportlab/weasyprint)

## Requirements

- Python 3.11+
- Django 5.2
- pandas
- openpyxl
- psycopg2-binary
- reportlab
- weasyprint

## Installation

```bash
pip install -r requirements.txt
python manage.py migrate
python manage.py createsuperuser
python manage.py runserver
```

## Importing sample data

Sample Excel files are located in `sample_data/`. Use the provided management commands, for example:

```bash
python manage.py import_programmes
python manage.py import_semesters
python manage.py import_programme_courses
python manage.py import_venues
python manage.py import_student_groups
python manage.py import_td_allocation
python manage.py import_workshop_allocation
python manage.py import_master_timetable
```