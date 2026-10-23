from django.urls import path

from . import views

urlpatterns = [
    path("timetable/", views.group_timetable, name="api-timetable"),
    path("venues/", views.venue_list, name="api-venues"),
    path("programme-groups/", views.programme_groups, name="api-programme-groups"),
    path("timetable/complete/", views.timetable_complete, name="api-timetable-complete"),
    path("schema/", views.schema, name="api-schema"),
    path("docs/", views.docs, name="api-docs"),
]
