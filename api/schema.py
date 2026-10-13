"""OpenAPI description of the timetable API, served to Swagger UI."""



def openapi_schema(request):
    return {
        "openapi": "3.0.3",
        "info": {
            "title": "CoET Timetable API",
            "version": "1.0.0",
            "description": (
                "Read the current semester: one group timetable, the venues "
                "that timetable uses, the programme groups, and whether every "
                "group allocation is complete. Send the API key in the "
                "X-API-Key header. The key is settings.API_KEY."
            ),
        },
        "servers": [{"url": request.build_absolute_uri("/").rstrip("/")}],
        "tags": [
            {"name": "Timetable"},
            {"name": "Venues"},
            {"name": "Groups"},
        ],
        "components": {
            "securitySchemes": {
                "ApiKeyAuth": {
                    "type": "apiKey",
                    "in": "header",
                    "name": "X-API-Key",
                }
            },
            "schemas": {
                "Semester": {
                    "type": "object",
                    "properties": {
                        "id": {"type": "integer"},
                        "academic_year": {"type": "string"},
                        "semester": {"type": "integer"},
                        "label": {"type": "string"},
                    },
                },
                "Error": {
                    "type": "object",
                    "properties": {"detail": {"type": "string"}},
                },
            },
        },
        "security": [{"ApiKeyAuth": []}],
        "paths": {
            "/api/timetable/": {
                "get": {
                    "tags": ["Timetable"],
                    "summary": "Group timetable",
                    "description": (
                        "Sessions, workshops and technical drawing for one "
                        "group, in the same shape the timetable view loads. "
                        "Always the current semester."
                    ),
                    "operationId": "groupTimetable",
                    "parameters": [
                        _query("programme", "Programme code, for example EE.", required=True),
                        _query("group", "Group code, for example C1.", required=True),
                        _query("year", "Year of study. Defaults to 1.", schema={"type": "integer", "default": 1}),
                    ],
                    "responses": _responses("The group's timetable."),
                }
            },
            "/api/venues/": {
                "get": {
                    "tags": ["Venues"],
                    "summary": "Venues for the current semester",
                    "description": (
                        "Rooms named on a session, workshop or technical "
                        "drawing in the current semester."
                    ),
                    "operationId": "venueList",
                    "responses": _responses("Venues used this semester."),
                }
            },
            "/api/programme-groups/": {
                "get": {
                    "tags": ["Groups"],
                    "summary": "Programme groups",
                    "description": (
                        "Groups on programmes that have courses in the current "
                        "semester. Pass programme to limit the list to one programme."
                    ),
                    "operationId": "programmeGroups",
                    "parameters": [
                        _query("programme", "Limit the list to one programme code."),
                    ],
                    "responses": _responses("Programmes and their groups."),
                }
            },
            "/api/timetable/complete/": {
                "get": {
                    "tags": ["Timetable"],
                    "summary": "Whether the timetable allocation is complete",
                    "description": (
                        "fully_set is true when every group that has a required "
                        "activity in the current semester is 100% complete, and "
                        "at least one group has a requirement. A group with "
                        "nothing required does not count as unfinished. This is "
                        "the same complete rule as Group Progress."
                    ),
                    "operationId": "timetableComplete",
                    "responses": _responses("Allocation progress for the current semester."),
                }
            },
        },
    }


def _query(name, description, required=False, schema=None):
    return {
        "name": name,
        "in": "query",
        "required": required,
        "description": description,
        "schema": schema or {"type": "string"},
    }


def _responses(description):
    return {
        "200": {
            "description": description,
            "content": {"application/json": {"schema": {"type": "object"}}},
        },
        "400": {
            "description": "The query is missing something the endpoint needs.",
            "content": {"application/json": {"schema": {"$ref": "#/components/schemas/Error"}}},
        },
        "401": {
            "description": "The API key is missing or wrong.",
            "content": {"application/json": {"schema": {"$ref": "#/components/schemas/Error"}}},
        },
        "404": {
            "description": "No current semester, or the programme or group does not exist.",
            "content": {"application/json": {"schema": {"$ref": "#/components/schemas/Error"}}},
        },
    }
