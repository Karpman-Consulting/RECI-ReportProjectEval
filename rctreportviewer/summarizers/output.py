def summarize_output_data(output, rmd_building_summary):
    if output is None:
        return

    # unmet hours
    rmd_building_summary["unmet_heating_hours"] = output.get("unmet_heating_hours", 0)
    rmd_building_summary["unmet_cooling_hours"] = output.get("unmet_cooling_hours", 0)

    bbp_summary = rmd_building_summary["compliance_calcs_by_parameter"].get("bbp", {})
    bbuec_summary = rmd_building_summary["compliance_calcs_by_parameter"].get(
        "bbuec", {}
    )
    bbrec_summary = rmd_building_summary["compliance_calcs_by_parameter"].get(
        "bbrec", {}
    )
    pbp_summary = rmd_building_summary["compliance_calcs_by_parameter"].get("pbp", {})
    pbp_nre_summary = rmd_building_summary["compliance_calcs_by_parameter"].get(
        "pbp_nre", {}
    )

    # --- Annual totals by fuel ---
    for source_result in output.get("annual_source_results", []):
        source = source_result.get("energy_source")
        annual_consumption = source_result.get("annual_consumption", 0)
        annual_cost = source_result.get("annual_cost", 0)

        rmd_building_summary["total_energy"] += annual_consumption
        rmd_building_summary["total_cost"] += annual_cost

        rmd_building_summary["energy_by_fuel_type"][source] = (
            rmd_building_summary["energy_by_fuel_type"].get(source, 0)
            + annual_consumption
        )
        rmd_building_summary["cost_by_fuel_type"][source] = (
            rmd_building_summary["cost_by_fuel_type"].get(source, 0) + annual_cost
        )

    # --- Annual totals by end use (keep elec/gas, add OTHER bucket) ---
    for end_use in output.get("annual_end_use_results", []):
        end_use_name = end_use.get("type")
        energy_use = end_use.get("annual_site_energy_use", 0)
        source = end_use.get("energy_source")

        # regulated vs unregulated site energy
        if end_use.get("is_regulated"):
            bbrec_summary["site_energy"] = (
                bbrec_summary.get("site_energy", 0) + energy_use
            )
            bbrec_summary[source] = bbrec_summary.get(source, 0) + energy_use
        else:
            bbuec_summary["site_energy"] = (
                bbuec_summary.get("site_energy", 0) + energy_use
            )
            bbuec_summary[source] = bbuec_summary.get(source, 0) + energy_use

        # always update BBP/PBP_NRE
        bbp_summary["site_energy"] = bbp_summary.get("site_energy", 0) + energy_use
        pbp_nre_summary["site_energy"] = (
            pbp_nre_summary.get("site_energy", 0) + energy_use
        )

        # building totals
        rmd_building_summary["total_energy"] += energy_use
        rmd_building_summary["energy_by_end_use"][end_use_name] = (
            rmd_building_summary["energy_by_end_use"].get(end_use_name, 0) + energy_use
        )

        # per-fuel end-use splits
        if source == "ELECTRICITY":
            rmd_building_summary["elec_by_end_use"][end_use_name] = (
                rmd_building_summary["elec_by_end_use"].get(end_use_name, 0)
                + energy_use
            )
        elif source == "NATURAL_GAS":
            rmd_building_summary["gas_by_end_use"][end_use_name] = (
                rmd_building_summary["gas_by_end_use"].get(end_use_name, 0) + energy_use
            )
        else:
            rmd_building_summary["other_by_end_use"][end_use_name] = (
                rmd_building_summary["other_by_end_use"].get(end_use_name, 0)
                + energy_use
            )

    # Update compliance calcs
    if rmd_building_summary["rmd_type"] == "Baseline":
        rmd_building_summary["compliance_calcs_by_parameter"]["bbp"] = bbp_summary
        rmd_building_summary["compliance_calcs_by_parameter"]["bbuec"] = bbuec_summary
        rmd_building_summary["compliance_calcs_by_parameter"]["bbrec"] = bbrec_summary
    elif rmd_building_summary["rmd_type"] == "Proposed":
        rmd_building_summary["compliance_calcs_by_parameter"]["pbp"] = pbp_summary
        rmd_building_summary["compliance_calcs_by_parameter"][
            "pbp_nre"
        ] = pbp_nre_summary
