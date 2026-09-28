from django import forms

from core.models import Day, Programme, Semester, StudentGroup
from .models import CollisionReport


class CollisionReportForm(forms.ModelForm):
    class Meta:
        model = CollisionReport
        fields = (
            "timetable_type", "semester", "programme", "group", "course_or_exam",
            "day", "time_description", "description", "contact_email",
        )
        widgets = {
            "timetable_type": forms.Select(attrs={"class": "field-control"}),
            "semester": forms.Select(attrs={"class": "field-control"}),
            "programme": forms.Select(attrs={"class": "field-control"}),
            "group": forms.Select(attrs={"class": "field-control"}),
            "course_or_exam": forms.TextInput(attrs={"class": "field-control", "placeholder": "Course code or exam name"}),
            "day": forms.Select(attrs={"class": "field-control"}),
            "time_description": forms.TextInput(attrs={"class": "field-control", "placeholder": "For example, Monday 09:00–11:00"}),
            "description": forms.Textarea(attrs={"class": "field-control", "rows": 4, "placeholder": "Describe the two activities that overlap."}),
            "contact_email": forms.EmailInput(attrs={"class": "field-control", "autocomplete": "email", "placeholder": "Optional"}),
        }

    def __init__(self, *args, default_semester=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["semester"].queryset = Semester.objects.order_by(
            "-academic_year", "-semester"
        )
        self.fields["programme"].queryset = Programme.objects.order_by("name")
        self.fields["group"].queryset = StudentGroup.objects.select_related(
            "programme"
        ).order_by("programme__name", "code")
        self.fields["day"].choices = [("", "Not specified"), *Day.choices]
        self.fields["programme"].required = False
        self.fields["group"].required = False
        self.fields["course_or_exam"].required = False
        self.fields["day"].required = False
        self.fields["time_description"].required = False
        self.fields["contact_email"].required = False
        self.fields["description"].label = "What is conflicting?"
        self.fields["contact_email"].label = "Email for follow-up (optional)"
        self.fields["programme"].choices = [(p.pk, p.name) for p in self.fields["programme"].queryset]
        self.fields["group"].choices = [(g.pk, f"{g.programme.name} — {g.code}") for g in self.fields["group"].queryset]
        if default_semester and not self.is_bound:
            self.initial.setdefault("semester", default_semester.pk)

    def clean(self):
        cleaned = super().clean()
        programme, group = cleaned.get("programme"), cleaned.get("group")
        if group and programme and group.programme_id != programme.pk:
            self.add_error("group", "Choose a group belonging to the selected programme.")
        elif group and not programme:
            cleaned["programme"] = group.programme
        return cleaned
