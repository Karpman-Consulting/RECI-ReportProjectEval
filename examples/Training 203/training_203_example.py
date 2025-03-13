import rctreportviewer.main as rctreportviewer


simcheck_detailed_report = rctreportviewer.RCTDetailedReport(
    r"ASHRAE9012019DetailReport.json",
    [
        r"Training 203 Demo.json"
    ],
    r"ASHRAE9012019DetailReport.html",
)
simcheck_detailed_report.run()
