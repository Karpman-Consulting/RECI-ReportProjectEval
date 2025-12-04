from rctreportviewer.power import determine_fan_power
from rctreportviewer.constants import fuel_type_map


def summarize_rmd_system_data(
    rct_report_viewer, building_segment, rmd_building_summary
):
    def get_system_type(system_id):
        for system_type, system_names in rct_report_viewer.hvac_system_types_b.items():
            if system_id in system_names:
                return system_type
        return None

    for hvac_system in building_segment.get(
        "heating_ventilating_air_conditioning_systems", []
    ):
        # Add hvac system to the summary list if not already present
        system_in_summaries = False
        system_summary = {}
        system_name = hvac_system.get("id")
        for system in rmd_building_summary["hvac_system_summaries"]:
            if system.get("name") == system_name:
                system_in_summaries = True
                break
        if not system_in_summaries:
            system_summary["name"] = system_name
            system_summary["type"] = get_system_type(system_name)

        hvac_fan_system = hvac_system.get("fan_system")
        if hvac_fan_system:
            supply_fan_controls = hvac_fan_system.get("fan_control", "Undefined")
            if supply_fan_controls == "CONSTANT":
                occupied_operation = hvac_fan_system.get(
                    "operation_during_occupied", "Undefined"
                )
                if occupied_operation == "CYCLING":
                    supply_fan_controls = "Constant Cycling"

            if (
                supply_fan_controls
                not in rmd_building_summary[
                    "total_fan_power_by_fan_control_by_fan_type"
                ]
            ):
                rmd_building_summary["total_fan_power_by_fan_control_by_fan_type"][
                    supply_fan_controls
                ] = {
                    "Supply": 0,
                    "Return/Relief": 0,
                    "Exhaust": 0,
                    "Zonal Exhaust": 0,
                    "Terminal Unit": 0,
                }
            if (
                supply_fan_controls
                not in rmd_building_summary["total_air_flow_by_fan_control_by_fan_type"]
            ):
                rmd_building_summary["total_air_flow_by_fan_control_by_fan_type"][
                    supply_fan_controls
                ] = {
                    "Supply": 0,
                    "Return/Relief": 0,
                    "Exhaust": 0,
                    "Zonal Exhaust": 0,
                    "Terminal Unit": 0,
                }

            for fan in hvac_fan_system.get("supply_fans", []):
                fan_power = determine_fan_power(fan)
                if fan_power:
                    rmd_building_summary["total_fan_power_by_fan_control_by_fan_type"][
                        supply_fan_controls
                    ]["Supply"] += fan_power
                    rmd_building_summary["total_fan_power"] += fan_power
                if "design_airflow" in fan:
                    rmd_building_summary["total_air_flow_by_fan_control_by_fan_type"][
                        supply_fan_controls
                    ]["Supply"] += fan["design_airflow"]

            for fan in hvac_fan_system.get("return_fans", []) + hvac_fan_system.get(
                "relief_fans", []
            ):
                fan_power = determine_fan_power(fan)
                if fan_power:
                    rmd_building_summary["total_fan_power_by_fan_control_by_fan_type"][
                        supply_fan_controls
                    ]["Return/Relief"] += fan_power
                    rmd_building_summary["total_fan_power"] += fan_power
                if "design_airflow" in fan:
                    rmd_building_summary["total_air_flow_by_fan_control_by_fan_type"][
                        supply_fan_controls
                    ]["Return/Relief"] += fan["design_airflow"]

            for fan in hvac_fan_system.get("exhaust_fans", []):
                fan_power = determine_fan_power(fan)
                if fan_power:
                    rmd_building_summary["total_fan_power_by_fan_control_by_fan_type"][
                        supply_fan_controls
                    ]["Exhaust"] += fan_power
                    rmd_building_summary["total_fan_power"] += fan_power
                if "design_airflow" in fan:
                    rmd_building_summary["total_air_flow_by_fan_control_by_fan_type"][
                        supply_fan_controls
                    ]["Exhaust"] += fan["design_airflow"]

        hvac_heating_system = hvac_system.get("heating_system")
        if hvac_heating_system:
            # Add heating system info to the summary list if it exists
            system_summary["heating_equipment_type"] = hvac_heating_system.get("type")
            system_summary["heating_energy_source"] = hvac_heating_system.get(
                "energy_source_type"
            )
            system_summary["heating_capacity"] = hvac_heating_system.get(
                "design_capacity", 0.0
            )
            system_summary["heating_capacity_units"] = "Btu/h"
            system_summary["heating_efficiency_metric_types"] = hvac_heating_system.get(
                "efficiency_metric_types", []
            )
            system_summary[
                "heating_efficiency_metric_values"
            ] = hvac_heating_system.get("efficiency_metric_values", [])

        hvac_cooling_system = hvac_system.get("cooling_system")
        if hvac_cooling_system:
            # Add cooling system to the summary list if it exists
            system_summary["cooling_equipment_type"] = hvac_cooling_system.get("type")
            system_summary["cooling_capacity"] = hvac_cooling_system.get(
                "design_total_cool_capacity", 0.0
            )
            system_summary["cooling_capacity_units"] = "Btu/h"
            system_summary["cooling_efficiency_metric_types"] = hvac_cooling_system.get(
                "efficiency_metric_types", []
            )
            system_summary[
                "cooling_efficiency_metric_values"
            ] = hvac_cooling_system.get("efficiency_metric_values", [])

        # Count the number of zones served by system
        for zone in building_segment.get("zones", []):
            for terminal in zone.get("terminals", []):
                if (
                    terminal.get(
                        "served_by_heating_ventilating_air_conditioning_system"
                    )
                    == system_name
                ):
                    system_summary["zone_qty"] = system_summary.get("zone_qty", 0) + 1

        if system_summary:
            rmd_building_summary["hvac_system_summaries"].append(system_summary)


def summarize_heating_cooling_capacity_data(building_segment, rmd_building_summary):
    # Ensure capacity dicts exist and have a Total bucket
    heating = rmd_building_summary.setdefault("heating_capacity_by_fuel_type", {})
    cooling = rmd_building_summary.setdefault("cooling_capacity_by_fuel_type", {})
    heating.setdefault("Total", 0.0)
    cooling.setdefault("Total", 0.0)

    boiler_loops = set(rmd_building_summary.get("boiler_loops", []))
    chw_loops = set(rmd_building_summary.get("chw_loops", []))
    external_fluid_sources = rmd_building_summary.get("external_fluid_sources", [])

    def is_external_loop(loop_id):
        if not loop_id:
            return False
        for fs in external_fluid_sources:
            if loop_id == fs.get("loop"):
                return True
        return False

    def add_capacity(bucket: dict, key: str, amount: float):
        if not amount:
            return
        bucket[key] = bucket.get(key, 0.0) + amount
        bucket["Total"] += amount

    # HVAC systems
    for hvac_system in building_segment.get(
        "heating_ventilating_air_conditioning_systems", []
    ):
        # Heating
        hs = hvac_system.get("heating_system")
        if hs:
            cap = hs.get("design_capacity", 0.0)
            loop = hs.get("hot_water_loop")
            if loop:
                if is_external_loop(loop):
                    add_capacity(heating, "Purchased Heat", cap)
                elif loop in boiler_loops:
                    add_capacity(heating, "On-site Boiler Plant", cap)
            else:
                fuel = fuel_type_map.get(hs.get("energy_source_type"))
                if fuel:
                    add_capacity(heating, fuel, cap)

        # Cooling
        cs = hvac_system.get("cooling_system")
        if cs:
            cap = cs.get("design_total_cool_capacity", 0.0)
            loop = cs.get("chilled_water_loop")
            if loop:
                if is_external_loop(loop):
                    add_capacity(cooling, "Purchased CHW", cap)
                elif loop in chw_loops:
                    add_capacity(cooling, "On-site Chiller Plant", cap)
            else:
                # DX cooling assumed electric if no CHW loop present
                add_capacity(cooling, "Electricity", cap)

    # Terminal Heating and Cooling
    for zone in building_segment.get("zones", []):
        for terminal in zone.get("terminals", []):
            hcap = terminal.get("heating_capacity", 0.0)
            hloop = terminal.get("heating_from_loop")
            if hcap and hloop:
                if is_external_loop(hloop):
                    add_capacity(heating, "Purchased Heat", hcap)
                else:
                    add_capacity(heating, "On-site Boiler Plant", hcap)

            ccap = terminal.get("cooling_capacity", 0.0)
            cloop = terminal.get("cooling_from_loop")
            if ccap and cloop:
                if is_external_loop(cloop):
                    add_capacity(cooling, "Purchased CHW", ccap)
                else:
                    add_capacity(cooling, "On-site Chiller Plant", ccap)
