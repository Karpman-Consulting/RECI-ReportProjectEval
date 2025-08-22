from rctreportviewer.units import convert_unit


def summarize_schedule_data(schedule, rmd_building_summary):
    schedule_id = schedule.get("id")
    schedule_area = rmd_building_summary.get("floor_area_by_schedule", {}).get(
        schedule_id
    )

    # If the schedule area is not defined, skip summarizing this schedule
    if not schedule_area:
        return
    schedule_area = convert_unit(schedule_area, "m2", "ft2")
    rmd_building_summary["schedule_summaries"][schedule_id] = {
        "EFLH": sum(schedule.get("hourly_values", [])),
        "associated_floor_area": schedule_area,
        "percent_total_lighting_power": (
            rmd_building_summary.get("int_ltg_power_by_schedule", {}).get(
                schedule_id, 0.0
            )
            / rmd_building_summary.get("total_lighting_power", 1.0)
        )
        * 100,
        "percent_total_equipment_power": (
            rmd_building_summary.get("equip_power_by_schedule", {}).get(
                schedule_id, 0.0
            )
            / rmd_building_summary.get("total_equipment_power", 1.0)
        )
        * 100,
        "associated_peak_internal_gain": (
            rmd_building_summary.get("int_ltg_power_by_schedule", {}).get(
                schedule_id, 0.0
            )
            + rmd_building_summary.get("equip_power_by_schedule", {}).get(
                schedule_id, 0.0
            )
            + rmd_building_summary.get("occ_peak_internal_gain_by_schedule", {}).get(
                schedule_id, 0.0
            )
        ),
    }
