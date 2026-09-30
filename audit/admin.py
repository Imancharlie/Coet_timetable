from django.contrib import admin

from audit.models import AllocationRunMetrics, AuditReport


@admin.register(AuditReport)
class AuditReportAdmin(admin.ModelAdmin):
    list_display = ("id", "allocation_run", "version", "status", "verdict",
                    "completion_pct", "created_at")
    list_filter = ("status",)
    search_fields = ("data_hash",)
    readonly_fields = ("data_hash", "thresholds_json", "summary_json", "pdf_path")
    ordering = ("-created_at",)

    @admin.display(description="verdict")
    def verdict(self, obj):
        return obj.summary().get("verdict", "—")

    @admin.display(description="completion %")
    def completion_pct(self, obj):
        value = obj.summary().get("completion_pct")
        return "—" if value is None else f"{value:.1f}"


@admin.register(AllocationRunMetrics)
class AllocationRunMetricsAdmin(admin.ModelAdmin):
    list_display = ("run", "duration_s", "candidates_examined", "backtracks",
                    "requirements_processed", "search_limit_hit")
    list_filter = ("search_limit_hit",)
