from rctreportviewer.summarizers.output import summarize_output_data
from rctreportviewer.summarizers.building_segments import (
    summarize_building_segment_data,
)
from rctreportviewer.summarizers.schedules import summarize_schedule_data
from rctreportviewer.summarizers.swh import summarize_water_heater_data
from rctreportviewer.summarizers.plants import (
    summarize_heating_plant_data,
    summarize_cooling_plant_data,
)
from rctreportviewer.power import determine_pump_power


def summarize_rmd_data(rct_report_viewer, rmd_data, model_type):
    rmd_building_summary = {
        "rmd_type": model_type,
        "building_count": len(rmd_data.get("buildings", [])),
        "building_segment_count": 0,
        "zone_count": 0,
        "space_count": 0,
        "system_count": 0,
        "boiler_count": len(rmd_data.get("boilers", [])),
        "electric_boiler_count": 0,
        "fossil_fuel_boiler_count": 0,
        "chiller_count": len(rmd_data.get("chillers", [])),
        "electric_chiller_count": 0,
        "fossil_fuel_chiller_count": 0,
        "water_heater_count": len(rmd_data.get("service_water_heating_equipment")),
        "heat_rejection_count": len(rmd_data.get("heat_rejections", [])),
        "constructions": rmd_data.get("constructions", []),
        "hvac_system_summaries": [],
        "electric_boiler_plant_capacity": 0.0,
        "fossil_fuel_boiler_plant_capacity": 0.0,
        "electric_chiller_plant_capacity": 0.0,
        "fossil_fuel_chiller_plant_capacity": 0.0,
        "cooling_tower_gpm": 0.0,
        "cooling_tower_hp": 0.0,
        "design_flow_by_loop_id": {},
        "pump_count": len(rmd_data.get("pumps", [])),
        "fluid_loop_types": {
            proposed_fluid_loop.get("type")
            for proposed_fluid_loop in rmd_data.get("fluid_loops", [])
        },
        "heating_capacity_by_fuel_type": {},
        "cooling_capacity_by_fuel_type": {},
        "external_fluid_sources": rmd_data.get("external_fluid_sources", []),
        "service_water_heating_uses": rmd_data.get("service_water_heating_uses", []),
        "overall_wall_ua_by_building_segment": {},
        "overall_wall_u_factor_by_building_segment": {},
        "overall_roof_ua_by_building_segment": {},
        "overall_roof_u_factor_by_building_segment": {},
        "overall_window_ua_by_building_segment": {},
        "overall_window_u_factor_by_building_segment": {},
        "overall_skylight_ua_by_building_segment": {},
        "overall_skylight_u_factor_by_building_segment": {},
        "lighting_area_type_by_building_segment": {},
        "total_floor_area_by_building_segment": {},
        "total_wall_area_by_building_segment": {},
        "total_roof_area_by_building_segment": {},
        "total_window_area_by_building_segment": {},
        "total_skylight_area_by_building_segment": {},
        "total_floor_area_by_space_type": {},
        "total_occupants_by_space_type": {},
        "total_lighting_power_by_space_type": {},
        "total_miscellaneous_equipment_power_by_space_type": {},
        "average_occupancy_by_space_type": {},
        "average_lighting_power_by_space_type": {},
        "average_miscellaneous_equipment_power_by_space_type": {},
        "total_fan_power_by_fan_control_by_fan_type": {},
        "total_air_flow_by_fan_control_by_fan_type": {},
        "other_fan_power_by_fan_type": {},
        "other_air_flow_by_fan_type": {},
        "total_fan_power_by_fan_type": {},
        "total_air_flow_by_fan_type": {},
        "energy_by_fuel_type": {},
        "cost_by_fuel_type": {},
        "energy_by_end_use": {},
        "elec_by_end_use": {},
        "gas_by_end_use": {},
        "other_by_end_use": {},
        "energy_by_end_use_eui": {},
        "elec_by_end_use_eui": {},
        "gas_by_end_use_eui": {},
        "cost_by_end_use": {},
        "total_floor_area": 0,
        "total_exterior_wall_area": 0,
        "total_roof_area": 0,
        "total_window_area": 0,
        "total_skylight_area": 0,
        "total_occupants": 0,
        "total_lighting_power": 0,
        "total_equipment_power": 0,
        "total_pump_power": 0,
        "total_fan_power": 0,
        "total_zone_minimum_oa_flow": 0,
        "total_infiltration": 0,
        "unmet_heating_hours": 0,
        "unmet_cooling_hours": 0,
        "total_energy": 0,
        "compliance_calcs_by_parameter": {},
        "total_cost": 0,
        "int_ltg_power_by_schedule": {},
        "equip_power_by_schedule": {},
        "floor_area_by_schedule": {},
        "occ_peak_internal_gain_by_schedule": {},
        "schedule_summaries": {},
        "swh_use_id_to_area_types": {},
        "water_heater_summary": {},
        "boiler_loops": [],
        "chw_loops": [],
        "int_ltg_summaries": {},
    }

    output = rmd_data.get("model_output")
    if output is not None:
        summarize_output_data(output, rmd_building_summary)

    for chiller in rmd_data.get("chillers", []):
        condensing_loop = chiller.get("condensing_loop")
        cooling_towers = []
        if condensing_loop:
            for heat_rejection in rmd_data.get("heat_rejections", []):
                if heat_rejection.get("loop") == condensing_loop:
                    cooling_towers.append(heat_rejection)

        summarize_cooling_plant_data(chiller, cooling_towers, rmd_building_summary)

    for boiler in rmd_data.get("boilers", []):
        summarize_heating_plant_data(boiler, rmd_building_summary)

    for building in rmd_data.get("buildings", []):
        rmd_building_summary["building_segment_count"] += len(
            building.get("building_segments", [])
        )

        summarize_building_segment_data(
            rct_report_viewer, building, rmd_building_summary
        )

    for pump in rmd_data.get("pumps", []):
        pump_power = determine_pump_power(pump)
        if pump_power:
            rmd_building_summary["total_pump_power"] += pump_power
            rmd_building_summary[pump.get("loop_or_piping", "Undefined")] = pump_power

    for schedule in rmd_data.get("schedules", []):
        # Skip temperature schedules
        if schedule.get("type") in [None, "TEMPERATURE"]:
            continue
        # Skip flag schedules
        if any(hourly_val < 0 for hourly_val in schedule.get("hourly_values", [])):
            continue

        summarize_schedule_data(schedule, rmd_building_summary)

    for water_heater in rmd_data.get("service_water_heating_equipment", []):
        summarize_water_heater_data(water_heater, rmd_building_summary)

    return rmd_building_summary
