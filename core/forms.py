from django import forms
from django.forms import inlineformset_factory

from .models import (
    ALLOCATED_ACTIVITY_TYPES,
    ActivityType,
    Course,
    CourseActivityRequirement,
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

INPUT_CLS = "w-full px-3 py-2 border border-gray-300 rounded-lg text-sm text-gray-800 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 bg-white"
SELECT_CLS = INPUT_CLS


class SemesterForm(forms.ModelForm):
    class Meta:
        model = Semester
        fields = ["academic_year", "semester"]
        widgets = {
            "academic_year": forms.TextInput(
                attrs={"class": INPUT_CLS, "placeholder": "e.g. 2026/2027"}
            ),
            "semester": forms.NumberInput(attrs={"class": INPUT_CLS, "min": 1, "max": 4}),
        }


class ProgrammeForm(forms.ModelForm):
    class Meta:
        model = Programme
        fields = ["code", "name"]
        widgets = {
            "code": forms.TextInput(attrs={"class": INPUT_CLS, "placeholder": "e.g. CE"}),
            "name": forms.TextInput(
                attrs={"class": INPUT_CLS, "placeholder": "e.g. Civil Engineering"}
            ),
        }


class StudentGroupForm(forms.ModelForm):
    class Meta:
        model = StudentGroup
        fields = ["programme", "code"]
        widgets = {
            "programme": forms.Select(attrs={"class": SELECT_CLS}),
            "code": forms.TextInput(attrs={"class": INPUT_CLS, "placeholder": "e.g. A1"}),
        }


class ProgrammeCourseForm(forms.ModelForm):
    """Programme -> shared course link.

    ``course_code``/``course_name`` are deliberately absent: they are mirrors of
    the shared ``Course`` and are filled in by ``ProgrammeCourse.save()``, so
    editing them here would only be overwritten.
    """

    class Meta:
        model = ProgrammeCourse
        fields = ["programme", "course", "semester"]
        widgets = {
            "programme": forms.Select(attrs={"class": SELECT_CLS}),
            "course": forms.Select(attrs={"class": SELECT_CLS}),
            "semester": forms.NumberInput(attrs={"class": INPUT_CLS, "min": 1}),
        }
        labels = {"course": "Course"}

    def clean(self):
        cleaned = super().clean()
        programme = cleaned.get("programme")
        course = cleaned.get("course")
        semester = cleaned.get("semester")
        if programme and course and semester:
            clash = ProgrammeCourse.objects.filter(
                programme=programme, semester=semester, course=course
            ).exclude(pk=self.instance.pk)
            if clash.exists():
                self.add_error(
                    "semester",
                    f"{programme.code} already studies {course.code} in "
                    f"semester {semester}.",
                )
        return cleaned


class CourseRequirementForm(forms.ModelForm):
    """One "how many sessions of this activity" row of a course."""

    class Meta:
        model = CourseActivityRequirement
        fields = ["activity_type", "count"]
        widgets = {
            "activity_type": forms.Select(attrs={"class": SELECT_CLS}),
            "count": forms.NumberInput(attrs={"class": INPUT_CLS, "min": 1}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Only the three allocatable activities can be required; a lecture or a
        # workshop is never a group-allocation requirement.
        self.fields["activity_type"].choices = [
            (value, label)
            for value, label in ActivityType.choices
            if value in ALLOCATED_ACTIVITY_TYPES
        ]


CourseRequirementFormSet = inlineformset_factory(
    Course,
    CourseActivityRequirement,
    form=CourseRequirementForm,
    extra=1,
    can_delete=True,
    min_num=0,
    validate_min=False,
)


class CourseForm(forms.ModelForm):
    """The shared course record: one code, one name, its requirements."""

    class Meta:
        model = Course
        fields = ["code", "name"]
        widgets = {
            "code": forms.TextInput(
                attrs={"class": INPUT_CLS, "placeholder": "e.g. MT161"}
            ),
            "name": forms.TextInput(
                attrs={"class": INPUT_CLS, "placeholder": "e.g. Mathematics 1"}
            ),
        }


class VenueForm(forms.ModelForm):
    class Meta:
        model = Venue
        fields = ["name", "capacity"]
        widgets = {
            "name": forms.TextInput(attrs={"class": INPUT_CLS, "placeholder": "e.g. NB102"}),
            "capacity": forms.NumberInput(attrs={"class": INPUT_CLS, "min": 0}),
        }


class VenueRecycleForm(forms.ModelForm):
    """Name-only validation scope for resolving a casing/spacing duplicate.

    Capacity is never required here: fixing a venue-name formatting problem is
    an independent action that must not block on (or validate) the capacity
    field. Only creating/finalising a venue record requires a capacity.
    """

    class Meta:
        model = Venue
        fields = ["name"]
        widgets = {
            "name": forms.TextInput(attrs={"class": INPUT_CLS, "placeholder": "e.g. NB102"}),
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
            "semester": forms.Select(attrs={"class": SELECT_CLS}),
            "course_code": forms.TextInput(
                attrs={"class": INPUT_CLS, "placeholder": "e.g. MT161"}
            ),
            "activity_type": forms.Select(attrs={"class": SELECT_CLS}),
            "day": forms.Select(attrs={"class": SELECT_CLS}),
            "venue": forms.Select(attrs={"class": SELECT_CLS}),
            "start_time": forms.TimeInput(
                attrs={"class": INPUT_CLS, "type": "time"}, format="%H:%M"
            ),
            "end_time": forms.TimeInput(
                attrs={"class": INPUT_CLS, "type": "time"}, format="%H:%M"
            ),
        }


class SessionGroupForm(forms.ModelForm):
    class Meta:
        model = SessionGroup
        fields = ["group"]
        widgets = {
            "group": forms.Select(attrs={"class": SELECT_CLS}),
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
            "time_period",
            "start_time",
            "end_time",
            "venue",
            "workshop",
            "position",
            "schedule_section",
            "week_start",
            "week_end",
            "year_of_study",
        ]
        widgets = {
            "semester": forms.Select(attrs={"class": SELECT_CLS}),
            "course_code": forms.TextInput(
                attrs={"class": INPUT_CLS, "placeholder": "e.g. Building or TG201"}
            ),
            "group_code": forms.TextInput(
                attrs={"class": INPUT_CLS, "placeholder": "e.g. C1"}
            ),
            "day": forms.Select(attrs={"class": SELECT_CLS}),
            "time_period": forms.Select(attrs={"class": SELECT_CLS}),
            "venue": forms.TextInput(attrs={"class": INPUT_CLS, "placeholder": "e.g. TW101"}),
            "workshop": forms.TextInput(
                attrs={"class": INPUT_CLS, "placeholder": "e.g. Building"}
            ),
            "position": forms.NumberInput(attrs={"class": INPUT_CLS, "min": 1, "max": 6}),
            "schedule_section": forms.TextInput(
                attrs={"class": INPUT_CLS, "placeholder": "e.g. SCHEDULE 1"}
            ),
            "week_start": forms.NumberInput(attrs={"class": INPUT_CLS, "min": 1}),
            "week_end": forms.NumberInput(attrs={"class": INPUT_CLS, "min": 1}),
            "year_of_study": forms.NumberInput(attrs={"class": INPUT_CLS, "min": 1}),
            "start_time": forms.TimeInput(
                attrs={"class": INPUT_CLS, "type": "time"}, format="%H:%M"
            ),
            "end_time": forms.TimeInput(
                attrs={"class": INPUT_CLS, "type": "time"}, format="%H:%M"
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
            "semester": forms.Select(attrs={"class": SELECT_CLS}),
            "course_code": forms.TextInput(
                attrs={"class": INPUT_CLS, "placeholder": "e.g. TG201"}
            ),
            "group_code": forms.TextInput(
                attrs={"class": INPUT_CLS, "placeholder": "e.g. A1"}
            ),
            "day": forms.Select(attrs={"class": SELECT_CLS}),
            "venue": forms.TextInput(attrs={"class": INPUT_CLS, "placeholder": "e.g. TW101"}),
            "start_time": forms.TimeInput(
                attrs={"class": INPUT_CLS, "type": "time"}, format="%H:%M"
            ),
            "end_time": forms.TimeInput(
                attrs={"class": INPUT_CLS, "type": "time"}, format="%H:%M"
            ),
        }


class FileUploadForm(forms.Form):
    file = forms.FileField(
        widget=forms.ClearableFileInput(
            attrs={"accept": ".xlsx,.xls", "class": "block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"}
        )
    )