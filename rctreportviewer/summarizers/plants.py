from rctreportviewer.constants import fuel_type_map
from rctreportviewer.power import determine_fan_power


def summarize_heating_plant_data(boiler, rmd_building_summary):
    fuel = fuel_type_map.get(boiler.get("energy_source_type"))

    if fuel == "Electricity":
        rmd_building_summary["electric_boiler_count"] += 1
        rmd_building_summary["electric_boiler_plant_capacity"] += boiler.get(
            "design_capacity", 0.0
        )

    elif fuel == "Fossil Fuel":
        rmd_building_summary["fossil_fuel_boiler_count"] += 1
        rmd_building_summary["fossil_fuel_boiler_plant_capacity"] += boiler.get(
            "design_capacity", 0.0
        )

    loop = boiler.get("loop")
    if loop:
        rmd_building_summary["boiler_loops"].append(loop)


def summarize_cooling_plant_data(chiller, cooling_towers, rmd_building_summary):
    fuel = fuel_type_map.get(chiller.get("energy_source_type"))

    if fuel == "Electricity":
        rmd_building_summary["electric_chiller_count"] += 1
        rmd_building_summary["electric_chiller_plant_capacity"] += chiller.get(
            "design_capacity", 0.0
        )

    elif fuel == "Fossil Fuel":
        rmd_building_summary["fossil_fuel_chiller_count"] += 1
        rmd_building_summary["fossil_fuel_chiller_plant_capacity"] += chiller.get(
            "design_capacity", 0.0
        )

    for cooling_tower in cooling_towers:
        rmd_building_summary["cooling_tower_gpm"] += cooling_tower.get(
            "rated_water_flowrate", 0.0
        )
        fan = cooling_tower.get("fan")
        if fan:
            rmd_building_summary["cooling_tower_hp"] += determine_fan_power(fan)

    loop = chiller.get("loop")
    if loop:
        rmd_building_summary["chw_loops"].append(loop)
