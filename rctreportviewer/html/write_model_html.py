from rctreportviewer.html.compliance_calculations import write_compliance_calculations
from rctreportviewer.html.component_summary import write_component_summary
from rctreportviewer.html.envelope_summary import write_envelope_summary
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
        file.write("""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>SIMcheck Detailed Evaluation Report</title>

    <meta name="viewport" content="width=device-width, initial-scale=1">

    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

    <style>
        body {
            overflow-x: hidden;
        }

        td.rule-id {
            white-space: nowrap;
        }

        td.outcome-summary {
            white-space: pre-wrap;
        }

        .sticky-top-2 {
            top: 37px;
            z-index: 1029;
        }
    </style>
</head>
<body>
""")

        file.write("""
<div class="container-fluid px-3 py-2">
  <div class="mx-auto" style="max-width: 1600px;">

    <header class="mb-4">
      <h2 class="text-center mb-1">Model Report</h2>
    </header>
""")

        write_component_summary(file, rct_detailed_report)
        write_compliance_calculations(file, rct_detailed_report)
        write_results_summary(file, rct_detailed_report)
        write_envelope_summary(file, rct_detailed_report)
        write_interior_loads_summary(file, rct_detailed_report)
        write_hvac_summary(file, rct_detailed_report)
        write_swh_summary(file, rct_detailed_report)

        file.write("""
  </div>
</div>

<button
    id="back-to-top"
    class="btn btn-primary position-fixed bottom-0 end-0 m-3"
    onclick="scrollToTop()"
    style="opacity:0; visibility:hidden; z-index:1050;">
    ↑
</button>
""")

        write_javascript(file, rct_detailed_report)

        file.write("""
</body>
</html>
""")
