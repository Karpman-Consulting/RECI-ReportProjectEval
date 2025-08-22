

def summarize_water_heater_data(water_heater, rmd_building_summary):
    fuel_type = water_heater.get("heater_fuel_type", "Unknown")
    efficiencies = list(
        zip(
            water_heater.get("efficiency_metric_types", []),
            water_heater.get("efficiency_metric_values", []),
        )
    )

    # Get area types via use → distribution → heater
    dist_sys_id = water_heater.get("distribution_system")
    area_types = set()
    if dist_sys_id:
        use_ids = [
            use.get("id")
            for use in rmd_building_summary.get("service_water_heating_uses", [])
            if use.get("served_by_distribution_system") == dist_sys_id
        ]
        for use_id in use_ids:
            area_types.update(
                rmd_building_summary["swh_use_id_to_area_types"].get(use_id, [])
            )

    rmd_building_summary["water_heater_summary"][water_heater["id"]] = {
        "id": water_heater["id"],
        "fuel_type": fuel_type,
        "efficiencies": efficiencies,
        "area_types": sorted(area_types),
    }
