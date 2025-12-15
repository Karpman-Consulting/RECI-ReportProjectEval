from rctreportviewer.html.evaluations import write_evaluations_section


def write_html_file(rct_detailed_report):
    """
    Writes the extracted data to an HTML file for easy viewing with Bootstrap styling.
    """

    with open(rct_detailed_report.output_file_path, "w", encoding="utf-8") as file:
        file.write(
            """
        <!DOCTYPE HTML>
        <html style="scrollbar-gutter: stable;">
        <head>
            <meta charset="UTF-8">
            <title>SIMcheck Detailed Evaluation Report</title>
            <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
            <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
            <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
            <style>
                td.rule-id { white-space: nowrap; }
                td.outcome-summary { white-space: pre-wrap; }
                .sticky-top-2 {
                    top: 37px;
                    z-index: 1029;
                }
            </style>
        </head>
        """
        )

        file.write(
            f"""
        <body class="mt-2 ms-2">
            <div class="d-flex flex-nowrap">

                <div class="flex-grow-1">
                    <h2 class="text-center mb-4">{rct_detailed_report.evaluation_data["ruleset"]} Model Report</h2>
                    <div class="mb-3">
                        <p><strong>Generated on:</strong> {rct_detailed_report.evaluation_data["date_run"]}</p>
                    </div>
        """
        )
        write_evaluations_section(file, rct_detailed_report)

        file.write(
            """
                </div>
            </div>
            <div class="position-fixed bottom-0 end-0 mb-2 me-2" style="z-index: 1050;">
                <button id="back-to-top" class="btn btn-primary" onclick="scrollToTop()" style="opacity: 0; visibility: hidden;"> ↑ </button>
            </div>
        </body>
        """
        )

        file.write("</html>")
