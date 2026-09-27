import uuid

import django.db.models.deletion
from django.db import migrations, models


class Migration(migrations.Migration):
    initial = True
    dependencies = [("core", "0001_initial")]
    operations = [
        migrations.CreateModel(
            name="PortalSettings",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("current_semester", models.ForeignKey(help_text="Semester selected by default on the public student portal.", on_delete=django.db.models.deletion.PROTECT, related_name="current_for_portals", to="core.semester")),
            ],
            options={"verbose_name": "Student portal settings", "verbose_name_plural": "Student portal settings"},
        ),
        migrations.CreateModel(
            name="CollisionReport",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("reference", models.UUIDField(default=uuid.uuid4, editable=False, unique=True)),
                ("timetable_type", models.CharField(choices=[("TEACHING", "Teaching timetable"), ("EXAMINATION", "Examination timetable")], max_length=16)),
                ("course_or_exam", models.CharField(blank=True, max_length=100)),
                ("day", models.CharField(blank=True, choices=[("MONDAY", "Monday"), ("TUESDAY", "Tuesday"), ("WEDNESDAY", "Wednesday"), ("THURSDAY", "Thursday"), ("FRIDAY", "Friday"), ("SATURDAY", "Saturday"), ("SUNDAY", "Sunday")], max_length=10)),
                ("time_description", models.CharField(blank=True, max_length=80)),
                ("description", models.TextField()),
                ("contact_email", models.EmailField(blank=True, max_length=254)),
                ("status", models.CharField(choices=[("NEW", "New"), ("REVIEWING", "Under review"), ("RESOLVED", "Resolved")], default="NEW", max_length=12)),
                ("staff_notes", models.TextField(blank=True)),
                ("created_at", models.DateTimeField(auto_now_add=True)),
                ("updated_at", models.DateTimeField(auto_now=True)),
                ("group", models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, to="core.studentgroup")),
                ("programme", models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, to="core.programme")),
                ("semester", models.ForeignKey(on_delete=django.db.models.deletion.PROTECT, to="core.semester")),
            ],
            options={"ordering": ["-created_at"]},
        ),
    ]
