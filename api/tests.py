import datetime

from django.conf import settings
from django.test import TestCase

from core.models import (
    ActivityType,
    Course,
    Day,
    Programme,
    ProgrammeCourse,
    Semester,
    Session,
    SessionGroup,
    StudentGroup,
    Venue,
)


class TimetableApiTests(TestCase):
    def setUp(self):
        self.semester = Semester.objects.create(
            academic_year="2026/2027", semester=1, is_current=True
        )
        self.other = Semester.objects.create(
            academic_year="2025/2026", semester=2
        )
        self.programme = Programme.objects.create(
            code="EE", name="BSc. in Electrical Engineering"
        )
        self.group = StudentGroup.objects.create(programme=self.programme, code="C1")
        self.venue = Venue.objects.create(name="R217", capacity=40)
        Venue.objects.create(name="Unused Hall", capacity=200)
        self.course = Course.objects.create(code="EE150", name="Circuits")
        self.course.set_requirements({ActivityType.TUTORIAL: 1})
        ProgrammeCourse.objects.create(
            programme=self.programme, course=self.course, semester=1
        )
        self.session = Session.objects.create(
            semester=self.semester,
            course_code="EE150",
            activity_type=ActivityType.TUTORIAL,
            day=Day.MONDAY,
            start_time=datetime.time(10, 0),
            end_time=datetime.time(11, 0),
            venue=self.venue,
        )
        SessionGroup.objects.create(session=self.session, group=self.group)
        old_venue = Venue.objects.create(name="Old Room", capacity=10)
        Session.objects.create(
            semester=self.other,
            course_code="EE150",
            activity_type=ActivityType.TUTORIAL,
            day=Day.TUESDAY,
            start_time=datetime.time(8, 0),
            end_time=datetime.time(9, 0),
            venue=old_venue,
        )
        self.key = {"HTTP_X_API_KEY": settings.API_KEY}

    def test_a_missing_key_is_rejected_and_does_not_open_the_staff_login(self):
        response = self.client.get("/api/venues/")
        self.assertEqual(response.status_code, 401)
        self.assertNotIn("Staff sign in", response.content.decode())

    def test_the_group_timetable_is_the_current_semester(self):
        response = self.client.get(
            "/api/timetable/?programme=ee&group=c1", **self.key
        )
        self.assertEqual(response.status_code, 200)
        body = response.json()
        self.assertEqual(body["semester"]["id"], self.semester.pk)
        self.assertEqual(body["group"]["code"], "C1")
        self.assertEqual(body["count"], 1)
        entry = body["entries"][0]
        self.assertEqual(entry["course_code"], "EE150")
        self.assertEqual(entry["type"], "Tutorial")
        self.assertEqual(entry["venue"], "R217")
        self.assertEqual(entry["day_label"], "Monday")
        self.assertEqual(entry["start"], "10:00")

    def test_venues_are_only_those_used_this_semester(self):
        response = self.client.get("/api/venues/", **self.key)
        names = {row["name"] for row in response.json()["venues"]}
        self.assertEqual(names, {"R217"})

    def test_programme_groups_follow_the_current_semester(self):
        response = self.client.get(
            "/api/programme-groups/?programme=EE", **self.key
        )
        body = response.json()
        self.assertEqual(body["count"], 1)
        self.assertEqual(body["programmes"][0]["groups"][0]["code"], "C1")

    def test_fully_set_is_true_when_every_required_group_is_complete(self):
        response = self.client.get("/api/timetable/complete/", **self.key)
        body = response.json()
        self.assertTrue(body["fully_set"])
        self.assertEqual(body["complete_count"], 1)
        self.assertEqual(body["incomplete_count"], 0)
        self.assertEqual(body["groups"][0]["percent"], 100)

    def test_swagger_lists_the_endpoints_without_a_key(self):
        schema = self.client.get("/api/schema/")
        self.assertEqual(schema.status_code, 200)
        text = schema.content.decode()
        self.assertIn("/api/timetable/", text)
        self.assertIn("/api/venues/", text)
        self.assertIn("/api/programme-groups/", text)
        self.assertIn("/api/timetable/complete/", text)
        docs = self.client.get("/api/docs/")
        self.assertEqual(docs.status_code, 200)
        self.assertIn("swagger", docs.content.decode().lower())
