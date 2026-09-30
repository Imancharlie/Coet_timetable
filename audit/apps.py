from django.apps import AppConfig


class AuditConfig(AppConfig):
    default_auto_field = "django.db.models.BigAutoField"
    name = "audit"
    verbose_name = "Post-allocation audit"

    def ready(self):
        # Registers the post_save listener that records run performance. The
        # listener only ever *reads* the allocator's plan snapshot.
        from audit import signals  # noqa: F401
