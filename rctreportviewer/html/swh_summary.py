from rctreportviewer.constants import efficiency_display_map


def write_swh_summary(file, rct_detailed_report):
    file.write("""      
        <div class="mb-3 me-4">
            <button class="btn btn-info collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapse-swh-summary" aria-expanded="false">
                Service Water Heating Summary
            </button>

            <div id="collapse-swh-summary" class="accordion-collapse collapse">
                <div class="accordion-body">
                    <table class="table table-sm table-borderless" style="width: 1250px;">
                        <thead>
                            <tr class="text-center">
                                <th colspan="1" class="col-4"></th>
                                <th colspan="3" class="col-4" style="border: 2px solid black;">Proposed Water Heater</th>
                                <th colspan="3" class="col-4" style="border: 2px solid black;">Baseline Water Heater</th>
                            </tr>
                            <tr class="text-center">
                                <th style="border: 2px solid black;">Water Heater</th>
                                <th style="border: 2px solid black;">Area Type</th>
                                <th style="border: 2px solid black;">Fuel</th>
                                <th style="border: 2px solid black;">Efficiency</th>
                                <th style="border: 2px solid black;">Area Type</th>
                                <th style="border: 2px solid black;">Fuel</th>
                                <th style="border: 2px solid black;">Efficiency</th>
                            </tr>
                        </thead>
                        <tbody style="border: 2px solid black;">
            """
               )

    proposed_water_heater_summary = rct_detailed_report.proposed_model_summary.get(
        "water_heater_summary", {}
    )
    baseline_water_heater_summary = rct_detailed_report.baseline_model_summary.get(
        "water_heater_summary", {}
    )
    combined_water_heater_ids = set(
        wh_id
        for wh_id in (
                list(proposed_water_heater_summary.keys())
                + list(baseline_water_heater_summary.keys())
        )
    )
    for water_heater_id in combined_water_heater_ids:
        proposed_wh_id_match = next(
            (
                pwh_id
                for pwh_id in proposed_water_heater_summary
                if pwh_id == water_heater_id
            ),
            None,
        )
        baseline_wh_id_match = next(
            (
                bwh_id
                for bwh_id in baseline_water_heater_summary
                if bwh_id == water_heater_id
            ),
            None,
        )

        def format_efficiencies(eff_list):
            if not isinstance(eff_list, list) or not eff_list:
                return "-"
            return "; ".join(
                f"{value:.2f} {efficiency_display_map.get(metric, metric.replace('_', ' ').title())}"
                for metric, value in eff_list
            )

        # Set safe defaults
        proposed_area_type = "-"
        proposed_fuel = "-"
        proposed_efficiency = "-"

        baseline_area_type = "-"
        baseline_fuel = "-"
        baseline_efficiency = "-"

        # Populate if matching proposed WH found
        if proposed_wh_id_match:
            proposed_wh_data = proposed_water_heater_summary[proposed_wh_id_match]
            proposed_area_type = ", ".join(
                proposed_wh_data.get("area_types", ["-"])
            )
            proposed_fuel = proposed_wh_data.get("fuel_type", "-")
            proposed_efficiency = format_efficiencies(
                proposed_wh_data.get("efficiencies", [])
            )

        # Populate if matching baseline WH found
        if baseline_wh_id_match:
            baseline_wh_data = baseline_water_heater_summary[baseline_wh_id_match]
            baseline_area_type = ", ".join(
                baseline_wh_data.get("area_types", ["-"])
            )
            baseline_fuel = baseline_wh_data.get("fuel_type", "-")
            baseline_efficiency = format_efficiencies(
                baseline_wh_data.get("efficiencies", [])
            )

        file.write(f"""
                            <tr style="font-size: 12px;" class="lh-1 text-center">
                                <td style="border-right: 2px solid black;">{water_heater_id}</td>
                                <td>{proposed_area_type.replace("_", " ").title()}</td>
                                <td>{proposed_fuel.replace("_", " ").title()}</td>
                                <td style="border-right: 2px solid black;">{proposed_efficiency}</td>
                                <td>{baseline_area_type.replace("_", " ").title()}</td>
                                <td>{baseline_fuel.replace("_", " ").title()}</td>
                                <td style="border-right: 2px solid black;">{baseline_efficiency}</td>
                            </tr>
        """)

    file.write(
        f"""
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    """)
