from rctreportviewer.html.compliance_calculations import write_compliance_calculations
from rctreportviewer.html.component_summary import write_component_summary
from rctreportviewer.html.envelope_summary import write_envelope_summary
from rctreportviewer.html.evaluations import write_evaluations_section
from rctreportviewer.html.hvac_summary import write_hvac_summary
from rctreportviewer.html.interior_loads_summary import write_interior_loads_summary
from rctreportviewer.html.results_summary import write_results_summary
from rctreportviewer.html.swh_summary import write_swh_summary
from rctreportviewer.html.js import write_javascript


def write_html_file(rct_detailed_report):
    """
    Writes the extracted data to an HTML file for easy viewing with Bootstrap styling.
    """

    with open(rct_detailed_report.output_file_path, "w", encoding="utf-8") as file:
        # ---------- HEAD ----------
        file.write(
            """<!DOCTYPE html>
<html lang="en" style="scrollbar-gutter: stable;">
<head>
    <meta charset="UTF-8">
    <title>SIMcheck Detailed Evaluation Report</title>

    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.css" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

    <style>
        .vertical-header {
            writing-mode: vertical-rl;
            transform: rotate(180deg);
            white-space: nowrap;
        }
        td.rule-id { white-space: nowrap; }
        td.outcome-summary { white-space: pre-wrap; }

        td {
            font-size: 0.95rem;
            font-family: system-ui, -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            line-height: 1.4;
        }
    </style>
</head>
"""
        )

        # ---------- BODY / CONTAINER ----------
        file.write(
            f"""
<body>
    <div class="container-xl py-3">

        <header class="mb-4">
            <h2 class="text-center">
                {rct_detailed_report.evaluation_data["ruleset"]} Model Report
            </h2>
            <p class="text-center text-muted mb-0">
                Generated on: {rct_detailed_report.evaluation_data["date_run"]}
            </p>
        </header>
"""
        )

        # ---------- SECTIONS ----------
        write_component_summary(file, rct_detailed_report)
        write_compliance_calculations(file, rct_detailed_report)
        write_results_summary(file, rct_detailed_report)
        write_envelope_summary(file, rct_detailed_report)
        write_interior_loads_summary(file, rct_detailed_report)
        write_hvac_summary(file, rct_detailed_report)
        write_swh_summary(file, rct_detailed_report)
        write_evaluations_section(file, rct_detailed_report)

        # ---------- FOOTER / UI ----------
        file.write(
            """
    </div>

    <div class="position-fixed bottom-0 end-0 mb-3 me-3" style="z-index: 1050;">
        <button id="back-to-top"
                class="btn btn-primary"
                onclick="scrollToTop()"
                style="opacity: 0; visibility: hidden;">
            ↑
        </button>
    </div>
</body>
"""
        )

        write_javascript(file, rct_detailed_report)

        file.write("</html>")
