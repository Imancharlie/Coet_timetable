from django import forms
from django.forms import inlineformset_factory

from .models import (
    Programme,
    ProgrammeCourse,
    Semester,
    Session,
    SessionGroup,
    StudentGroup,
    TechnicalDrawingAllocation,
    Venue,
    WorkshopAllocation,
)


class SemesterForm(forms.ModelForm):
    class Meta:
        model = Semester
        fields = ["academic_year", "semester"]
        widgets = {
            "academic_year": forms.TextInput(
                attrs={"placeholder": "e.g. 2026/2027"}
            ),
            "semester": forms.NumberInput(attrs={"min": 1, "max": 4}),
        }


class ProgrammeForm(forms.ModelForm):
    class Meta:
        model = Programme
        fields = ["code", "name"]
        widgets = {
            "code": forms.TextInput(attrs={"placeholder": "e.g. CE"}),
            "name": forms.TextInput(
                attrs={"placeholder": "e.g. Civil Engineering"}
            ),
        }


class StudentGroupForm(forms.ModelForm):
    class Meta:
        model = StudentGroup
        fields = ["programme", "code"]
        widgets = {
            "programme": forms.Select(attrs={"class": "w-full"}),
            "code": forms.TextInput(attrs={"placeholder": "e.g. A1"}),
        }


class ProgrammeCourseForm(forms.ModelForm):
    class Meta:
        model = ProgrammeCourse
        fields = ["programme", "course_code"]
        widgets = {
            "programme": forms.Select(attrs={"class": "w-full"}),
            "course_code": forms.TextInput(
                attrs={"placeholder": "e.g. MT161"}
            ),
        }


class VenueForm(forms.ModelForm):
    class Meta:
        model = Venue
        fields = ["name", "capacity"]
        widgets = {
            "name": forms.TextInput(attrs={"placeholder": "e.g. NB102"}),
            "capacity": forms.NumberInput(attrs={"min": 0}),
        }


class SessionForm(forms.ModelForm):
    class Meta:
        model = Session
        fields = [
            "semester",
            "course_code",
            "activity_type",
            "day",
            "start_time",
            "end_time",
            "venue",
        ]
        widgets = {
            "semester": forms.Select(attrs={"class": "w-full"}),
            "course_code": forms.TextInput(
                attrs={"placeholder": "e.g. MT161"}
            ),
            "activity_type": forms.Select(attrs={"class": "w-full"}),
            "day": forms.Select(attrs={"class": "w-full"}),
            "venue": forms.Select(attrs={"class": "w-full"}),
            "start_time": forms.TimeInput(
                attrs={"type": "time"}, format="%H:%M"
            ),
            "end_time": forms.TimeInput(
                attrs={"type": "time"}, format="%H:%M"
            ),
        }


class SessionGroupForm(forms.ModelForm):
    class Meta:
        model = SessionGroup
        fields = ["group"]
        widgets = {
            "group": forms.Select(attrs={"class": "w-full"}),
        }


SessionGroupFormSet = inlineformset_factory(
    Session,
    SessionGroup,
    form=SessionGroupForm,
    extra=1,
    can_delete=True,
)


class WorkshopAllocationForm(forms.ModelForm):
    class Meta:
        model = WorkshopAllocation
        fields = [
            "semester",
            "course_code",
            "group_code",
            "day",
            "start_time",
            "end_time",
            "venue",
        ]
        widgets = {
            "semester": forms.Select(attrs={"class": "w-full"}),
            "course_code": forms.TextInput(
                attrs={"placeholder": "e.g. TG201"}
            ),
            "group_code": forms.TextInput(
                attrs={"placeholder": "e.g. C1"}
            ),
            "day": forms.Select(attrs={"class": "w-full"}),
            "venue": forms.TextInput(
                attrs={"placeholder": "e.g. TW101"}
            ),
            "start_time": forms.TimeInput(
                attrs={"type": "time"}, format="%H:%M"
            ),
            "end_time": forms.TimeInput(
                attrs={"type": "time"}, format="%H:%M"
            ),
        }


class TechnicalDrawingAllocationForm(forms.ModelForm):
    class Meta:
        model = TechnicalDrawingAllocation
        fields = [
            "semester",
            "course_code",
            "group_code",
            "day",
            "start_time",
            "end_time",
            "venue",
        ]
        widgets = {
            "semester": forms.Select(attrs={"class": "w-full"}),
            "course_code": forms.TextInput(
                attrs={"placeholder": "e.g. TG201"}
            ),
            "group_code": forms.TextInput(
                attrs={"placeholder": "e.g. A1"}
            ),
            "day": forms.Select(attrs={"class": "w-full"}),
            "venue": forms.TextInput(
                attrs={"placeholder": "e.g. TW101"}
            ),
            "start_time": forms.TimeInput(
                attrs={"type": "time"}, format="%H:%M"
            ),
            "end_time": forms.TimeInput(
                attrs={"type": "time"}, format="%H:%M"
            ),
        }


class FileUploadForm(forms.Form):
    file = forms.FileField(
        widget=forms.ClearableFileInput(
            attrs={"accept": ".xlsx,.xls", "class": "block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"}
        )
    )
