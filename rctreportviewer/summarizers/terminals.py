from rctreportviewer.power import determine_fan_power


def summarize_rmd_terminal_data(zone, rmd_building_summary):
    for terminal in zone["terminals"]:
        if "minimum_outdoor_airflow" in terminal:
            rmd_building_summary["total_zone_minimum_oa_flow"] += terminal[
                "minimum_outdoor_airflow"
            ]

        if "fan" in terminal:
            fan_power = determine_fan_power(terminal["fan"])
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
                ]["Terminal Unit"] += fan_power
                rmd_building_summary["total_fan_power"] += fan_power
