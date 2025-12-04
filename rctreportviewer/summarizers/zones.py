from rctreportviewer.power import determine_fan_power
from rctreportviewer.summarizers.spaces import summarize_rmd_space_data
from rctreportviewer.summarizers.surfaces import summarize_rmd_surface_data
from rctreportviewer.summarizers.terminals import summarize_rmd_terminal_data


def summarize_rmd_zone_data(rct_report_viewer, building_segment, rmd_building_summary):
    for zone in building_segment.get("zones", []):
        rmd_building_summary["space_count"] += len(zone.get("spaces", []))

        infiltration = zone.get("infiltration")
        if infiltration and "flow_rate" in infiltration:
            rmd_building_summary["total_infiltration"] += infiltration["flow_rate"]

        zonal_exhaust_fan = zone.get("zonal_exhaust_fan")
        if zonal_exhaust_fan:
            fan_power = determine_fan_power(zonal_exhaust_fan)
            if fan_power:
                if (
                    "Undefined"
                    not in rmd_building_summary[
                        "total_fan_power_by_fan_control_by_fan_type"
                    ]
                ):
                    rmd_building_summary["total_fan_power_by_fan_control_by_fan_type"][
                        "Undefined"
                    ] = {
                        "Supply": 0,
                        "Return/Relief": 0,
                        "Exhaust": 0,
                        "Zonal Exhaust": 0,
                        "Terminal Unit": 0,
                    }

                rmd_building_summary["total_fan_power_by_fan_control_by_fan_type"][
                    "Undefined"
                ]["Zonal Exhaust"] += fan_power
                rmd_building_summary["total_fan_power"] += fan_power

        summarize_rmd_space_data(
            rct_report_viewer, building_segment, zone, rmd_building_summary
        )

        summarize_rmd_surface_data(building_segment, zone, rmd_building_summary)

        summarize_rmd_terminal_data(zone, rmd_building_summary)
