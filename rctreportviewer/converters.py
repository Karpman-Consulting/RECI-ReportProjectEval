from rctreportviewer.units import convert_unit


def convert_schedule_summaries_internal_gain(summary: dict):
    for schedule_data in summary.get("schedule_summaries", {}).values():
        if "associated_peak_internal_gain" in schedule_data:
            schedule_data["associated_peak_internal_gain"] = convert_unit(
                schedule_data["associated_peak_internal_gain"], "W", "kBtu / h"
            )


def convert_compliance_summary_energies(summary: dict):
    for parameter_data in summary.get("compliance_calcs_by_parameter", {}).values():
        for data_name, data in parameter_data.items():
            if data_name in ["source_energy", "site_energy"]:
                parameter_data[data_name] = convert_unit(data, "Btu", "MMBtu")


def convert_int_ltg_summary_units(summary: dict):
    for int_ltg_summary in summary.get("interior_lighting_summaries", {}).values():
        if "floor_area" in int_ltg_summary:
            int_ltg_summary["floor_area"] = convert_unit(
                int_ltg_summary["floor_area"], "m2", "ft2"
            )
        if "power_per_area" in int_ltg_summary:
            int_ltg_summary["power_per_area"] = convert_unit(
                int_ltg_summary["power_per_area"], "W / m2", "W / ft2"
            )


def convert_summary_units(summary: dict, units_dict: dict):
    for key, value in summary.items():
        if key not in units_dict:
            continue
        from_unit, to_unit = units_dict[key]

        if isinstance(value, dict):
            for sub_key, sub_value in value.items():
                if isinstance(sub_value, dict):
                    for sub_sub_key, sub_sub_value in sub_value.items():
                        value[sub_key][sub_sub_key] = convert_unit(
                            sub_sub_value, from_unit, to_unit
                        )
                else:
                    value[sub_key] = convert_unit(sub_value, from_unit, to_unit)
        else:
            summary[key] = convert_unit(value, from_unit, to_unit)


def calculate_eui(summary: dict):
    floor_area = summary.get("total_floor_area", 1)  # avoid division by zero

    elec_by_end_use = summary.get("elec_by_end_use", {})  # kWh by end use
    gas_by_end_use = summary.get("gas_by_end_use", {})  # therms by end use
    total_energy_by_end_use = summary.get("energy_by_end_use", {})  # kBtu by end use
    cost_by_fuel_type = summary.get("cost_by_fuel_type", {})  # $ by fuel

    # Ensure output dicts exist
    summary.setdefault("elec_by_end_use_eui", {})
    summary.setdefault("gas_by_end_use_eui", {})
    summary.setdefault("energy_by_end_use_eui", {})
    summary.setdefault("cost_by_end_use", {})

    # --- EUI calculations (kBtu/ft²) ---
    for end_use, kwh in elec_by_end_use.items():
        summary["elec_by_end_use_eui"][end_use] = (kwh * 3.412) / floor_area

    for end_use, therms in gas_by_end_use.items():
        summary["gas_by_end_use_eui"][end_use] = (therms * 100.0) / floor_area

    for end_use, kbtu in total_energy_by_end_use.items():
        summary["energy_by_end_use_eui"][end_use] = kbtu / floor_area

    total_kwh = sum(elec_by_end_use.values()) or 0.0
    total_therms = sum(gas_by_end_use.values()) or 0.0

    elec_total_or_rate = cost_by_fuel_type.get("ELECTRICITY", 0.0)
    gas_total_or_rate = cost_by_fuel_type.get("NATURAL_GAS", 0.0)

    elec_rate_per_kwh = (
        (elec_total_or_rate / total_kwh) if total_kwh > 0 else elec_total_or_rate
    )
    gas_rate_per_therm = (
        (gas_total_or_rate / total_therms) if total_therms > 0 else gas_total_or_rate
    )

    # Compute cost per end use: ($) = kWh*$/(kWh) + therms*$/(therm)
    for end_use in set(list(elec_by_end_use.keys()) + list(gas_by_end_use.keys())):
        kwh = elec_by_end_use.get(end_use, 0.0)
        therms = gas_by_end_use.get(end_use, 0.0)
        end_use_cost = kwh * elec_rate_per_kwh + therms * gas_rate_per_therm
        summary["cost_by_end_use"][end_use] = end_use_cost


def convert_model_data_units(
    baseline_model_summary: dict, proposed_model_summary: dict
):
    """
    Converts the baseline and proposed model summary values to the desired units
    and calculates EUI (Energy Use Intensity) values for electricity, gas, and total energy.
    """
    units_dict = {
        "overall_wall_ua_by_building_segment": ("W / K", "Btu / h / degR"),
        "overall_wall_u_factor_by_building_segment": (
            "W / m2 / K",
            "Btu / h / ft2 / degR",
        ),
        "overall_roof_ua_by_building_segment": ("W / K", "Btu / h / degR"),
        "overall_roof_u_factor_by_building_segment": (
            "W / m2 / K",
            "Btu / h / ft2 / degR",
        ),
        "overall_window_ua_by_building_segment": ("W / K", "Btu / h / degR"),
        "overall_window_u_factor_by_building_segment": (
            "W / m2 / K",
            "Btu / h / ft2 / degR",
        ),
        "overall_skylight_ua_by_building_segment": ("W / K", "Btu / h / degR"),
        "overall_skylight_u_factor_by_building_segment": (
            "W / m2 / K",
            "Btu / h / ft2 / degR",
        ),
        "average_lighting_power_by_space_type": ("W / m2", "W / ft2"),
        "total_floor_area_by_building_segment": ("m2", "ft2"),
        "floor_area_by_schedule": ("m2", "ft2"),
        "total_wall_area_by_building_segment": ("m2", "ft2"),
        "total_roof_area_by_building_segment": ("m2", "ft2"),
        "total_window_area_by_building_segment": ("m2", "ft2"),
        "total_floor_area_by_space_type": ("m2", "ft2"),
        "total_floor_area": ("m2", "ft2"),
        "total_exterior_wall_area": ("m2", "ft2"),
        "total_roof_area": ("m2", "ft2"),
        "total_window_area": ("m2", "ft2"),
        "total_zone_minimum_oa_flow": ("L / s", "cfm"),
        "total_infiltration": ("L / s", "cfm"),
        "total_air_flow_by_fan_control_by_fan_type": ("L / s", "cfm"),
        "total_air_flow_by_fan_type": ("L / s", "cfm"),
        "total_energy": ("J", "kBtu"),
        "total_site_energy": ("J", "MMBtu"),
        "total_source_energy": ("J", "MMBtu"),
        "total_site_energy_regulated": ("Btu", "MMBtu"),
        "total_site_energy_unregulated": ("Btu", "MMBtu"),
        "energy_by_fuel_type": ("J", "kBtu"),
        "energy_by_end_use": ("J", "kBtu"),
        "elec_by_end_use": ("J", "kWh"),
        "gas_by_end_use": ("J", "therm"),
        "other_by_end_use": ("J", "kBtu"),
        "heating_capacity_by_fuel_type": ("W", "kBtu / h"),
        "cooling_capacity_by_fuel_type": ("W", "kBtu / h"),
        "electric_chiller_plant_capacity": ("W", "ton"),
        "fossil_fuel_chiller_plant_capacity": ("W", "ton"),
        "cooling_tower_gpm": ("L/s", "gpm"),
        "cooling_tower_hp": ("W", "hp"),
        "electric_boiler_plant_capacity": ("W", "kBtu / h"),
        "fossil_fuel_boiler_plant_capacity": ("W", "kBtu / h"),
    }

    convert_summary_units(baseline_model_summary, units_dict)
    convert_summary_units(proposed_model_summary, units_dict)

    calculate_eui(baseline_model_summary)
    calculate_eui(proposed_model_summary)

    convert_schedule_summaries_internal_gain(baseline_model_summary)
    convert_schedule_summaries_internal_gain(proposed_model_summary)

    convert_compliance_summary_energies(baseline_model_summary)
    convert_compliance_summary_energies(proposed_model_summary)

    convert_int_ltg_summary_units(baseline_model_summary)
    convert_int_ltg_summary_units(proposed_model_summary)
