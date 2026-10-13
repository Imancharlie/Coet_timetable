"""Post-allocation audit and reporting.

A **read-only** layer over the final applied timetable: it queries allocation
and timetable data, never writes to it. The only rows this app owns are its own
:class:`audit.models.AuditReport` versions and the additive
:class:`audit.models.AllocationRunMetrics` instrumentation, both of which live
outside every table the allocator touches.
"""

default_app_config = "audit.apps.AuditConfig"
