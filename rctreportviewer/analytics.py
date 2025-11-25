import re

from rctreportviewer.constants import bpf_area_type_map
from rctreportviewer.units import convert_unit


def perform_analytic_calculations(rct_report_viewer):
    """
    Perform calculations on the model data to extract additional information.
    """

    # Calculate area-weighted BPF from baseline model summary
    def compute_area_weighted_bpf_for_metric(metric_bpf_data):
        total_weighted = 0
        total_area = 0
        for building_segment_id in rct_report_viewer.baseline_model_summary[
            "lighting_area_type_by_building_segment"
        ]:
            lighting_building_area_type = rct_report_viewer.baseline_model_summary[
                "lighting_area_type_by_building_segment"
            ][building_segment_id]
            bpf_area_type = bpf_area_type_map.get(
                lighting_building_area_type, "All Others"
            )
            if not bpf_area_type:
                continue
            try:
                bpf = metric_bpf_data[bpf_area_type]
                area = rct_report_viewer.baseline_model_summary[
                    "total_floor_area_by_building_segment"
                ].get(building_segment_id, 0)
                total_weighted += bpf * area
                total_area += area
            except KeyError:
                continue
        return total_weighted / total_area if total_area else None

    def compute_u_factors(model_summary):
        for surface in ["wall", "roof", "window", "skylight"]:
            ua_key = f"overall_{surface}_ua_by_building_segment"
            area_key = f"total_{surface}_area_by_building_segment"
            u_factor_key = f"overall_{surface}_u_factor_by_building_segment"

            ua_data = model_summary.get(ua_key, {})
            area_data = model_summary.get(area_key, {})
            u_factor_data = model_summary.setdefault(u_factor_key, {})

            for segment_id, ua_value in ua_data.items():
                area_value = area_data.get(segment_id)
                if area_value:  # Avoid division by zero or None
                    u_factor_data[segment_id] = ua_value / area_value

    climate_zone_vals = set(
        model.get("weather", {}).get("climate_zone")
        for model in rct_report_viewer.rpd_data["ruleset_model_descriptions"]
    )
    if len(climate_zone_vals) != 1:
        print("Multiple climate zones found in the RPD data, using the first one.")

    climate_zone = next(iter(climate_zone_vals), None)

    if climate_zone:
        climate_zone = climate_zone.split("CZ")[
            1
        ]  # Extract the climate zone without 'CZ'
        for metric in ["Cost", "Site Energy", "Source Energy", "GHG Emissions"]:
            # TODO: Replace hardcoded 90.1-2019 with dynamic ruleset version
            metric_key = f"90.1-2019 {metric}"
            bpf_data = rct_report_viewer.bpf_data[metric_key][climate_zone]
            rct_report_viewer.bpfs_by_metric[
                metric
            ] = compute_area_weighted_bpf_for_metric(bpf_data)

    # Calculate the LPD allowance based on evaluation data + RPD data combined
    for space_id in rct_report_viewer.space_areas:
        rct_report_viewer.baseline_total_lighting_power_allowance += (
            rct_report_viewer.space_lpd_allowances.get(space_id, 0)
            * convert_unit(rct_report_viewer.space_areas[space_id], "m2", "ft2")
        )
        space_type = rct_report_viewer.baseline_space_types.get(space_id)
        if space_type:
            rct_report_viewer.baseline_lighting_power_allowance_by_space_type[
                space_type
            ] = rct_report_viewer.baseline_lighting_power_allowance_by_space_type.get(
                space_type, 0
            ) + rct_report_viewer.space_lpd_allowances.get(
                space_id, 0
            ) * convert_unit(
                rct_report_viewer.space_areas[space_id], "m2", "ft2"
            )

    compute_u_factors(rct_report_viewer.baseline_model_summary)
    compute_u_factors(rct_report_viewer.proposed_model_summary)

    for lighting_space_type in rct_report_viewer.baseline_model_summary[
        "total_lighting_power_by_space_type"
    ]:
        rct_report_viewer.baseline_model_summary[
            "average_lighting_power_by_space_type"
        ][lighting_space_type] = (
            rct_report_viewer.baseline_model_summary[
                "total_lighting_power_by_space_type"
            ][lighting_space_type]
            / rct_report_viewer.baseline_model_summary[
                "total_floor_area_by_space_type"
            ][lighting_space_type]
        )
    for lighting_space_type in rct_report_viewer.proposed_model_summary[
        "total_lighting_power_by_space_type"
    ]:
        rct_report_viewer.proposed_model_summary[
            "average_lighting_power_by_space_type"
        ][lighting_space_type] = (
            rct_report_viewer.proposed_model_summary[
                "total_lighting_power_by_space_type"
            ][lighting_space_type]
            / rct_report_viewer.proposed_model_summary[
                "total_floor_area_by_space_type"
            ][lighting_space_type]
        )

    rct_report_viewer.baseline_model_summary["total_fan_power_by_fan_type"] = {
        ft: 0
        for ft in [
            "Supply",
            "Return/Relief",
            "Exhaust",
            "Zonal Exhaust",
            "Terminal Unit",
        ]
    }
    rct_report_viewer.baseline_model_summary["total_air_flow_by_fan_type"] = {
        ft: 0
        for ft in [
            "Supply",
            "Return/Relief",
            "Exhaust",
            "Zonal Exhaust",
            "Terminal Unit",
        ]
    }
    rct_report_viewer.proposed_model_summary["total_fan_power_by_fan_type"] = {
        ft: 0
        for ft in [
            "Supply",
            "Return/Relief",
            "Exhaust",
            "Zonal Exhaust",
            "Terminal Unit",
        ]
    }
    rct_report_viewer.proposed_model_summary["total_air_flow_by_fan_type"] = {
        ft: 0
        for ft in [
            "Supply",
            "Return/Relief",
            "Exhaust",
            "Zonal Exhaust",
            "Terminal Unit",
        ]
    }

    for fan_control in rct_report_viewer.baseline_model_summary[
        "total_fan_power_by_fan_control_by_fan_type"
    ]:
        for fan_type in rct_report_viewer.baseline_model_summary[
            "total_fan_power_by_fan_control_by_fan_type"
        ][fan_control]:
            rct_report_viewer.baseline_model_summary["total_fan_power_by_fan_type"][
                fan_type
            ] += rct_report_viewer.baseline_model_summary[
                "total_fan_power_by_fan_control_by_fan_type"
            ][
                fan_control
            ][
                fan_type
            ]
            rct_report_viewer.baseline_model_summary[
                "total_fan_power"
            ] += rct_report_viewer.baseline_model_summary[
                "total_fan_power_by_fan_control_by_fan_type"
            ][
                fan_control
            ][
                fan_type
            ]

    for fan_control in rct_report_viewer.baseline_model_summary[
        "total_air_flow_by_fan_control_by_fan_type"
    ]:
        for fan_type in rct_report_viewer.baseline_model_summary[
            "total_air_flow_by_fan_control_by_fan_type"
        ][fan_control]:
            rct_report_viewer.baseline_model_summary["total_air_flow_by_fan_type"][
                fan_type
            ] += rct_report_viewer.baseline_model_summary[
                "total_air_flow_by_fan_control_by_fan_type"
            ][
                fan_control
            ][
                fan_type
            ]

    for fan_control in rct_report_viewer.proposed_model_summary[
        "total_fan_power_by_fan_control_by_fan_type"
    ]:
        for fan_type in rct_report_viewer.proposed_model_summary[
            "total_fan_power_by_fan_control_by_fan_type"
        ][fan_control]:
            rct_report_viewer.proposed_model_summary["total_fan_power_by_fan_type"][
                fan_type
            ] += rct_report_viewer.proposed_model_summary[
                "total_fan_power_by_fan_control_by_fan_type"
            ][
                fan_control
            ][
                fan_type
            ]
            rct_report_viewer.proposed_model_summary[
                "total_fan_power"
            ] += rct_report_viewer.proposed_model_summary[
                "total_fan_power_by_fan_control_by_fan_type"
            ][
                fan_control
            ][
                fan_type
            ]

    for fan_control in rct_report_viewer.proposed_model_summary[
        "total_air_flow_by_fan_control_by_fan_type"
    ]:
        for fan_type in rct_report_viewer.proposed_model_summary[
            "total_air_flow_by_fan_control_by_fan_type"
        ][fan_control]:
            rct_report_viewer.proposed_model_summary["total_air_flow_by_fan_type"][
                fan_type
            ] += rct_report_viewer.proposed_model_summary[
                "total_air_flow_by_fan_control_by_fan_type"
            ][
                fan_control
            ][
                fan_type
            ]

    rct_report_viewer.baseline_model_summary["other_fan_power_by_fan_type"] = {
        ft: 0
        for ft in [
            "Supply",
            "Return/Relief",
            "Exhaust",
            "Zonal Exhaust",
            "Terminal Unit",
        ]
    }
    rct_report_viewer.baseline_model_summary["other_air_flow_by_fan_type"] = {
        ft: 0
        for ft in [
            "Supply",
            "Return/Relief",
            "Exhaust",
            "Zonal Exhaust",
            "Terminal Unit",
        ]
    }
    rct_report_viewer.proposed_model_summary["other_fan_power_by_fan_type"] = {
        ft: 0
        for ft in [
            "Supply",
            "Return/Relief",
            "Exhaust",
            "Zonal Exhaust",
            "Terminal Unit",
        ]
    }
    rct_report_viewer.proposed_model_summary["other_air_flow_by_fan_type"] = {
        ft: 0
        for ft in [
            "Supply",
            "Return/Relief",
            "Exhaust",
            "Zonal Exhaust",
            "Terminal Unit",
        ]
    }

    for fan_control in rct_report_viewer.baseline_model_summary[
        "total_fan_power_by_fan_control_by_fan_type"
    ]:
        if fan_control in ["Undefined", "INLET_VANE", "DISCHARGE_DAMPER", "OTHER"]:
            for fan_type in rct_report_viewer.baseline_model_summary[
                "total_fan_power_by_fan_control_by_fan_type"
            ][fan_control]:
                rct_report_viewer.baseline_model_summary["other_fan_power_by_fan_type"][
                    fan_type
                ] += rct_report_viewer.baseline_model_summary[
                    "total_fan_power_by_fan_control_by_fan_type"
                ][
                    fan_control
                ][
                    fan_type
                ]

    for fan_control in rct_report_viewer.baseline_model_summary[
        "total_air_flow_by_fan_control_by_fan_type"
    ]:
        if fan_control in ["Undefined", "INLET_VANE", "DISCHARGE_DAMPER", "OTHER"]:
            for fan_type in rct_report_viewer.baseline_model_summary[
                "total_air_flow_by_fan_control_by_fan_type"
            ][fan_control]:
                rct_report_viewer.baseline_model_summary["other_air_flow_by_fan_type"][
                    fan_type
                ] += rct_report_viewer.baseline_model_summary[
                    "total_air_flow_by_fan_control_by_fan_type"
                ][
                    fan_control
                ][
                    fan_type
                ]

    for fan_control in rct_report_viewer.proposed_model_summary[
        "total_fan_power_by_fan_control_by_fan_type"
    ]:
        if fan_control in ["Undefined", "INLET_VANE", "DISCHARGE_DAMPER", "OTHER"]:
            for fan_type in rct_report_viewer.proposed_model_summary[
                "total_fan_power_by_fan_control_by_fan_type"
            ][fan_control]:
                rct_report_viewer.proposed_model_summary["other_fan_power_by_fan_type"][
                    fan_type
                ] += rct_report_viewer.proposed_model_summary[
                    "total_fan_power_by_fan_control_by_fan_type"
                ][
                    fan_control
                ][
                    fan_type
                ]

    for fan_control in rct_report_viewer.proposed_model_summary[
        "total_air_flow_by_fan_control_by_fan_type"
    ]:
        if fan_control in ["Undefined", "INLET_VANE", "DISCHARGE_DAMPER", "OTHER"]:
            for fan_type in rct_report_viewer.proposed_model_summary[
                "total_air_flow_by_fan_control_by_fan_type"
            ][fan_control]:
                rct_report_viewer.proposed_model_summary["other_air_flow_by_fan_type"][
                    fan_type
                ] += rct_report_viewer.proposed_model_summary[
                    "total_air_flow_by_fan_control_by_fan_type"
                ][
                    fan_control
                ][
                    fan_type
                ]
