import math


def write_html_file(rct_detailed_report):
    """
    Writes the extracted data to an HTML file for easy viewing with Bootstrap styling.
    """
    section_titles_with_colors = {
        1: ("Design Model and Compliance Calculations", "#D8BFD8"),
        2: ("Additions and Alterations", "#66b3ff"),
        3: ("Space Use Classification", "#99ff99"),
        4: ("Schedules", "#ffcc99"),
        5: ("Envelope", "#f4a460"),
        6: ("Lighting", "#ffd700"),
        7: ("Thermal Blocks - HVAC Zones Designed", "#c2f0c2"),
        8: ("Thermal Blocks - HVAC Zones Not Designed", "#f0c2c2"),
        9: ("Thermal Blocks - Multifamily Residential Buildings", "#f0e68c"),
        10: ("HVAC Systems", "#4682b4"),
        11: ("Service Water Heating Systems", "#E97451"),
        12: ("Receptacles and Other Loads", "#d3d3d3"),
        13: ("Modeling Limitations to the Simulation Program", "#f4cccc"),
        14: ("Exterior Conditions", "#87ceeb"),
        15: ("Distribution Transformers", "#d9ead3"),
        16: ("Elevators", "#c0c0c0"),
        17: ("Refrigeration", "#5f9ea0"),
        18: ("Baseline HVAC Selection", "#ead1dc"),
        19: ("General Baseline HVAC System Requirements", "#778899"),
        20: ("System-Specific Baseline HVAC System Requirements", "#ffdab9"),
        21: ("Baseline HVAC - Water Side Requirements: Hot Water", "#ff6347"),
        22: ("Baseline HVAC - Water Side Requirements: Chilled Water", "#6495ED"),
        23: ("Baseline HVAC - Air Side Requirements", "#F0FFFF"),
    }

    with open(rct_detailed_report.output_file_path, "w", encoding="utf-8") as file:
        file.write(
            """
        <html style="scrollbar-gutter: stable;">
        <head>
            <meta charset="UTF-8">
            <title>SIMcheck Detailed Evaluation Report</title>
            <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
            <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
            <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
            <style>
                td.rule-id { white-space: nowrap; }
                td.outcome-summary { white-space: pre-wrap; }
                .sticky-top-2 {
                    top: 37px;
                    z-index: 1029;
                }
            </style>
        </head>
        """
        )
        file.write(
            f"""
        <body class="mt-2 ms-2">
            <div class="d-flex flex-nowrap">

                <div class="flex-grow-1">
                    <h1 class="text-center mb-4">RECI - Project Evaluation Report</h1>
                    <div class="mb-3">
                        <p><strong>Ruleset:</strong> {rct_detailed_report.evaluation_data["ruleset"]}</p>
                        <p><strong>Generated on:</strong> {rct_detailed_report.evaluation_data["date_run"]}</p>
                        <p><strong>Models Analyzed:</strong> {", ".join(rct_detailed_report.model_types)}</p>
                    </div>

        """
        )

        rule_categories = {
            "Failing": rct_detailed_report.rules_failed,
            "Passing": rct_detailed_report.rules_passed,
            "Undetermined": rct_detailed_report.full_eval_rules_undetermined + rct_detailed_report.appl_eval_rules_undetermined,
            "N/A": rct_detailed_report.rules_not_applicable,
        }

        file.write(
            f"""
                <div class="mb-3 me-4">
                    <button class="btn btn-info collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapse-model-component-summary" aria-expanded="false">
                        Model Component Summary
                    </button>

                    <div id="collapse-model-component-summary" class="accordion-collapse collapse">
                        <div class="accordion-body">
                            <table class="table table-sm table-borderless" style="width: 400px;">
                                <thead>
                                    <tr style="border-bottom: 2px solid black;"><th class="col-4 text-end"></th><th class="col-4 text-center">Baseline</th><th class="col-4 text-center">Proposed</th></tr>
                                </thead>
                                <tbody>
                                    <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Building Qty</td><td class="col-4 text-center">{rct_detailed_report.baseline_model_summary["building_count"]}</td><td class="col-4 text-center">{rct_detailed_report.proposed_model_summary["building_count"]}</td></tr>
                                    <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Total Floor Area</td><td class="col-4 text-center">{round(rct_detailed_report.baseline_model_summary['total_floor_area']):,}</td><td class="col-4 text-center">{round(rct_detailed_report.proposed_model_summary["total_floor_area"]):,}</td></tr>
                                    <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Building Area Qty</td><td class="col-4 text-center">{rct_detailed_report.baseline_model_summary["building_segment_count"]}</td><td class="col-4 text-center">{rct_detailed_report.proposed_model_summary["building_segment_count"]}</td></tr>
                                    <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">System Qty</td><td class="col-4 text-center">{rct_detailed_report.baseline_model_summary["system_count"]}</td><td class="col-4 text-center">{rct_detailed_report.proposed_model_summary["system_count"]}</td></tr>
                                    <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Zone Qty</td><td class="col-4 text-center">{rct_detailed_report.baseline_model_summary["zone_count"]}</td><td class="col-4 text-center">{rct_detailed_report.proposed_model_summary["zone_count"]}</td></tr>
                                    <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Space Qty</td><td class="col-4 text-center">{rct_detailed_report.baseline_model_summary["space_count"]}</td><td class="col-4 text-center">{rct_detailed_report.proposed_model_summary["space_count"]}</td></tr>
                                    <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Fluid Loops</td><td class="col-4 text-center">{", ".join(s.title() for s in rct_detailed_report.baseline_model_summary["fluid_loop_types"])}</td><td class="col-4 text-center">{", ".join(s.title() for s in rct_detailed_report.proposed_model_summary["fluid_loop_types"])}</td></tr>
                                    <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Pump Qty</td><td class="col-4 text-center">{rct_detailed_report.baseline_model_summary["pump_count"]}</td><td class="col-4 text-center">{rct_detailed_report.proposed_model_summary["pump_count"]}</td></tr>
                                    <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Boiler Qty</td><td class="col-4 text-center">{rct_detailed_report.baseline_model_summary["boiler_count"]}</td><td class="col-4 text-center">{rct_detailed_report.proposed_model_summary["boiler_count"]}</td></tr>
                                    <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Chiller Qty</td><td class="col-4 text-center">{rct_detailed_report.baseline_model_summary["chiller_count"]}</td><td class="col-4 text-center">{rct_detailed_report.proposed_model_summary["chiller_count"]}</td></tr>
                                    <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Heat Rejection Qty</td><td class="col-4 text-center">{rct_detailed_report.baseline_model_summary["heat_rejection_count"]}</td><td class="col-4 text-center">{rct_detailed_report.proposed_model_summary["heat_rejection_count"]}</td></tr>
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>

                <div class="mb-3 me-4">
                    <button class="btn btn-info collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapse-model-results-summary" aria-expanded="false">
                        Results Summary
                    </button>

                    <div id="collapse-model-results-summary" class="accordion-collapse collapse">
                        <div class="accordion-body">
                            <div style="position: relative; left: 360px;" class="mb-3">
                                <div class="btn-group" role="group" aria-label="Chart toggle">
                                    <input type="radio" class="btn-check" name="chartOptions" id="btn-elec" autocomplete="off" checked>
                                    <label style="width: 95px;" class="btn btn-outline-primary" for="btn-elec" onclick="showChart('elec')">Electricity</label>

                                    <input type="radio" class="btn-check" name="chartOptions" id="btn-gas" autocomplete="off">
                                    <label style="width: 95px;" class="btn btn-outline-danger" for="btn-gas" onclick="showChart('gas')">Gas</label>

                                    <input type="radio" class="btn-check" name="chartOptions" id="btn-energy" autocomplete="off">
                                    <label style="width: 95px;" class="btn btn-outline-success" for="btn-energy" onclick="showChart('energy')">Total</label>
                                </div>
                            </div>

                            <div class="form-check form-switch mb-3" style="margin-left: 725px;">
                              <input class="form-check-input" type="checkbox" id="unitToggle" onchange="toggleUnits()">
                              <label class="form-check-label" for="unitToggle">Show EUI (kBtu/ft²)</label>
                            </div>

                            <div class="mb-3" style="position: relative; left: 260px;">
                              <span id="baselineTotal" class="me-4 fw-bold">Baseline Total: </span>
                              <span id="proposedTotal" class="fw-bold">Proposed Total: </span>
                            </div>

                            <div id="elecChartContainer" style="width: 900px; height: 500px;">
                              <canvas id="elecByEndUse"></canvas>
                            </div>
                            <div id="gasChartContainer" style="width: 900px; height: 500px; display: none;">
                              <canvas id="gasByEndUse"></canvas>
                            </div>
                            <div id="energyChartContainer" style="width: 900px; height: 500px; display: none;">
                              <canvas id="energyByEndUse"></canvas>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="mb-3 me-4">
                    <button class="btn btn-info collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapse-envelope-summary" aria-expanded="false">
                        Envelope Summary
                    </button>

                    <div id="collapse-envelope-summary" class="accordion-collapse collapse">
                        <div class="accordion-body">
                            <table class="table table-sm table-borderless mb-0" style="width: 1300px;">
                                <thead>
                                    <tr class="text-center">
                                        <th colspan="2"></th>
                                        <th colspan="6" style="border: 2px solid black;">Baseline</th>
                                        <th colspan="6" style="border: 2px solid black;">Proposed</th>
                                    </tr>
                                    <tr class="text-center">
                                        <th rowspan="2" style="border: 2px solid black;">Building Area</th>
                                        <th rowspan="2" style="border: 2px solid black;">Surface Type</th>
                                        <th colspan="3" style="border: 2px solid black;">Opaque Surface</th>
                                        <th colspan="3" style="border: 2px solid black;">Fenestration</th>
                                        <th colspan="3" style="border: 2px solid black;">Opaque Surface</th>
                                        <th colspan="3" style="border: 2px solid black;">Fenestration</th>
                                    </tr>
                                    <tr class="text-center">
                                        <th style="border: 2px solid black;">Area (ft<sup>2</sup>)</th>
                                        <th style="border: 2px solid black;"> % </th>
                                        <th style="border: 2px solid black;"> U-Factor </th>
                                        <th style="border: 2px solid black;">Area (ft<sup>2</sup>)</th>
                                        <th style="border: 2px solid black;"> % </th>
                                        <th style="border: 2px solid black;"> U-Factor </th>
                                        <th style="border: 2px solid black;">Area (ft<sup>2</sup>)</th>
                                        <th style="border: 2px solid black;"> % </th>
                                        <th style="border: 2px solid black;"> U-Factor </th>
                                        <th style="border: 2px solid black;">Area (ft<sup>2</sup>)</th>
                                        <th style="border: 2px solid black;"> % </th>
                                        <th style="border: 2px solid black;"> U-Factor </th>
                                    </tr>
                                </thead>
                                <tbody style="border: 2px solid black;">
            """)

        for building_segment_id in rct_detailed_report.baseline_model_summary["total_floor_area_by_building_segment"]:
            if building_segment_id in rct_detailed_report.baseline_model_summary["total_roof_area_by_building_segment"]:
                file.write(
                    f"""
                                    <tr style="font-size: 12px;" class="lh-1 text-center">
                                        <td>{building_segment_id}</td>
                                        <td style="border-right: 2px solid black;">Roof</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_roof_area_by_building_segment'].get(building_segment_id, 0) - rct_detailed_report.baseline_model_summary['total_skylight_area_by_building_segment'].get(building_segment_id, 0)):,}</td>
                                        <td>{round((rct_detailed_report.baseline_model_summary['total_roof_area_by_building_segment'].get(building_segment_id, 0) - rct_detailed_report.baseline_model_summary['total_skylight_area_by_building_segment'].get(building_segment_id, 0)) / rct_detailed_report.baseline_model_summary['total_roof_area_by_building_segment'][building_segment_id] * 100, 1)}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary["overall_roof_u_factor_by_building_segment"].get(building_segment_id, 0), 3)}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary["total_skylight_area_by_building_segment"].get(building_segment_id, 0)):,}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary["total_skylight_area_by_building_segment"].get(building_segment_id, 0) / rct_detailed_report.baseline_model_summary['total_roof_area_by_building_segment'][building_segment_id] * 100, 1)}</td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary["overall_skylight_u_factor_by_building_segment"].get(building_segment_id, 0), 3)}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary["total_roof_area_by_building_segment"].get(building_segment_id, 0) - rct_detailed_report.proposed_model_summary['total_skylight_area_by_building_segment'].get(building_segment_id, 0)):,}</td>
                                        <td>{round((rct_detailed_report.proposed_model_summary["total_roof_area_by_building_segment"].get(building_segment_id, 0) - rct_detailed_report.proposed_model_summary['total_skylight_area_by_building_segment'].get(building_segment_id, 0)) / rct_detailed_report.proposed_model_summary['total_roof_area_by_building_segment'][building_segment_id] * 100, 1)}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary["overall_roof_u_factor_by_building_segment"].get(building_segment_id, 0), 3)}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary["total_skylight_area_by_building_segment"].get(building_segment_id, 0)):,}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary["total_skylight_area_by_building_segment"].get(building_segment_id, 0) / rct_detailed_report.proposed_model_summary['total_roof_area_by_building_segment'][building_segment_id] * 100, 1)}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary["overall_skylight_u_factor_by_building_segment"].get(building_segment_id, 0), 3)}</td>
                                    </tr>
                    """
                )
            if building_segment_id in rct_detailed_report.baseline_model_summary["total_wall_area_by_building_segment"]:
                file.write(
                    f"""
                                    <tr style="font-size: 12px;" class="lh-1 text-center">
                                        <td>{building_segment_id}</td>
                                        <td style="border-right: 2px solid black;">Ext. Wall</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_wall_area_by_building_segment'].get(building_segment_id, 0) - rct_detailed_report.baseline_model_summary["total_window_area_by_building_segment"].get(building_segment_id, 0)):,}</td>
                                        <td>{round((rct_detailed_report.baseline_model_summary['total_wall_area_by_building_segment'].get(building_segment_id, 0) - rct_detailed_report.baseline_model_summary["total_window_area_by_building_segment"].get(building_segment_id, 0)) / rct_detailed_report.baseline_model_summary['total_wall_area_by_building_segment'][building_segment_id] * 100, 1)}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary["overall_wall_u_factor_by_building_segment"].get(building_segment_id, 0), 3)}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary["total_window_area_by_building_segment"].get(building_segment_id, 0)):,}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary["total_window_area_by_building_segment"].get(building_segment_id, 0) / rct_detailed_report.baseline_model_summary['total_wall_area_by_building_segment'][building_segment_id] * 100, 1)}</td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary["overall_window_u_factor_by_building_segment"].get(building_segment_id, 0), 3)}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary["total_wall_area_by_building_segment"].get(building_segment_id, 0) - rct_detailed_report.proposed_model_summary["total_window_area_by_building_segment"].get(building_segment_id, 0)):,}</td>
                                        <td>{round((rct_detailed_report.proposed_model_summary["total_wall_area_by_building_segment"].get(building_segment_id, 0) - rct_detailed_report.proposed_model_summary["total_window_area_by_building_segment"].get(building_segment_id, 0)) / rct_detailed_report.proposed_model_summary['total_wall_area_by_building_segment'][building_segment_id] * 100, 1)}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary["overall_wall_u_factor_by_building_segment"].get(building_segment_id, 0), 3)}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary["total_window_area_by_building_segment"].get(building_segment_id, 0)):,}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary["total_window_area_by_building_segment"].get(building_segment_id, 0) / rct_detailed_report.proposed_model_summary['total_wall_area_by_building_segment'][building_segment_id] * 100, 1)}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary["overall_window_u_factor_by_building_segment"].get(building_segment_id, 0), 3)}</td>
                                    </tr>
                    """
                )

        file.write("""          </tbody>
                            </table>
                            <p style="font-size: 0.75rem;" class="ms-2">*U-Factors represent area-weighted averages for the corresponding Building Area & Surface Type</p>
                        </div>
                    </div>
                </div>

                <div class="mb-3 me-4">
                    <button class="btn btn-info collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapse-internal-loads-summary" aria-expanded="false">
                        Internal Loads Summary
                    </button>

                    <div id="collapse-internal-loads-summary" class="accordion-collapse collapse">
                        <div class="accordion-body">
                            <table class="table table-sm table-borderless" style="width: 900px;">
                                <thead>
                                    <tr class="text-center">
                                        <th colspan="2" class="col-4"></th>
                                        <th colspan="4" class="col-4" style="border: 2px solid black;">Baseline</th>
                                        <th colspan="3" class="col-4" style="border: 2px solid black;">Proposed</th>
                                    </tr>
                                    <tr class="text-center">
                                        <th style="border: 2px solid black;">Space Type</th>
                                        <th style="border: 2px solid black;">Area (ft<sup>2</sup>)</th>
                                        <th style="border: 2px solid black;">Occupancy Density (ft<sup>2</sup>/person)</th>
                                        <th style="border: 2px solid black;">Equipment Power Density (W/ft<sup>2</sup>)</th>
                                        <th style="border: 2px solid black;">Allowed Lighting Power Density (W/ft<sup>2</sup>)</th>
                                        <th style="border: 2px solid black;">Lighting Power Density (W/ft<sup>2</sup>)</th>
                                        <th style="border: 2px solid black;">Lighting Power Density (W/ft<sup>2</sup>)</th>
                                        <th style="border: 2px solid black;">Equipment Power Density (W/ft<sup>2</sup>)</th>
                                        <th style="border: 2px solid black;">Occupancy Density (ft<sup>2</sup>/person)</th>
                                    </tr>
                                </thead>
                                <tbody style="border: 2px solid black;">
        """)

        for space_type in rct_detailed_report.baseline_model_summary["total_floor_area_by_space_type"]:
            file.write(
                f"""
                                    <tr style="font-size: 12px;" class="lh-1 text-center">
                                        <td>{space_type.replace("_", " ").title()}</td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['total_floor_area_by_space_type'].get(space_type, 0)):,}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_floor_area_by_space_type'][space_type] / rct_detailed_report.baseline_model_summary['total_occupants_by_space_type'].get(space_type, math.inf))}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_miscellaneous_equipment_power_by_space_type'].get(space_type, 0) / rct_detailed_report.baseline_model_summary['total_floor_area_by_space_type'][space_type], 2)}</td>
                                        <td>{round(rct_detailed_report.baseline_lighting_power_allowance_by_space_type.get(space_type, 0) / rct_detailed_report.baseline_model_summary['total_floor_area_by_space_type'][space_type], 2)}</td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['total_lighting_power_by_space_type'].get(space_type, 0) / rct_detailed_report.baseline_model_summary['total_floor_area_by_space_type'][space_type], 2)}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_lighting_power_by_space_type'].get(space_type, 0) / rct_detailed_report.proposed_model_summary['total_floor_area_by_space_type'][space_type], 2)}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_miscellaneous_equipment_power_by_space_type'].get(space_type, 0) / rct_detailed_report.proposed_model_summary['total_floor_area_by_space_type'][space_type], 2)}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_floor_area_by_space_type'][space_type] / rct_detailed_report.proposed_model_summary['total_occupants_by_space_type'].get(space_type, math.inf))}</td>
                                    </tr>
                """
            )
        file.write(f"""
                                    <tr  style="font-size: 12px; border-top: 1px solid black;" class="lh-1 fw-bold text-center">
                                        <td>Total</td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['total_floor_area']):,}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_floor_area'] / rct_detailed_report.baseline_model_summary['total_occupants'], 2)}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_equipment_power'] / rct_detailed_report.baseline_model_summary['total_floor_area'], 2)}</td>
                                        <td>{round(rct_detailed_report.baseline_total_lighting_power_allowance / rct_detailed_report.baseline_model_summary['total_floor_area'], 2)}</td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['total_lighting_power'] / rct_detailed_report.baseline_model_summary['total_floor_area'], 2)}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_lighting_power'] / rct_detailed_report.proposed_model_summary['total_floor_area'], 2)}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_equipment_power'] / rct_detailed_report.proposed_model_summary['total_floor_area'], 2)}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_floor_area'] / rct_detailed_report.proposed_model_summary['total_occupants'], 2)}</td>
                                    </tr>
                                </tbody>
                            </table>
        """)

        # ----------------------- Schedule Summary Table -----------------------
        file.write(f"""
                            <h3>Schedule Summary</h3>
                            <table class="table table-sm table-borderless" style="width: 1250px;">
                                <thead>
                                    <tr class="text-center">
                                        <th colspan="1" class="col-4"></th>
                                        <th colspan="5" class="col-4" style="border: 2px solid black;">Baseline</th>
                                        <th colspan="5" class="col-4" style="border: 2px solid black;">Proposed</th>
                                    </tr>
                                    <tr class="text-center">
                                        <th style="border: 2px solid black;">Schedule</th>
                                        <th style="border: 2px solid black;">EFLH</th>
                                        <th style="border: 2px solid black;">Associated Floor Area (ft<sup>2</sup>)</th>
                                        <th style="border: 2px solid black;">% of Total Lighting Watts Controlled</th>
                                        <th style="border: 2px solid black;">% of Total Equipment Watts Controlled</th>
                                        <th style="border: 2px solid black;">Associated Peak Internal Gain (kBtu/hr)</th>
                                        <th style="border: 2px solid black;">EFLH</th>
                                        <th style="border: 2px solid black;">Associated Floor Area (ft<sup>2</sup>)</th>
                                        <th style="border: 2px solid black;">% of Total Lighting Watts Controlled</th>
                                        <th style="border: 2px solid black;">% of Total Equipment Watts Controlled</th>
                                        <th style="border: 2px solid black;">Associated Peak Internal Gain (kBtu/hr)</th>
                                    </tr>
                                </thead>
                                <tbody style="border: 2px solid black;">
        """)

        baseline_schedule_summaries = rct_detailed_report.baseline_model_summary["schedule_summaries"]
        for schedule_id in baseline_schedule_summaries.keys():
            baseline_schedule_summary = (
                rct_detailed_report.baseline_model_summary["schedule_summaries"].get(schedule_id, {})
            )
            proposed_schedule_summary = (
                rct_detailed_report.proposed_model_summary["schedule_summaries"].get(schedule_id, {})
            )
            file.write(
                f"""
                                    <tr style="font-size: 12px;" class="lh-1 text-center">
                                        <td style="border-right: 2px solid black;">{schedule_id}</td>
                                        <td>{round(baseline_schedule_summary.get("EFLH", 0)):,}</td>
                                        <td>{round(baseline_schedule_summary.get("associated_floor_area", 0.0), 4):,}</td>
                                        <td>{round(baseline_schedule_summary.get("percent_total_lighting_power", 0.0), 4):,}</td>
                                        <td>{round(baseline_schedule_summary.get("percent_total_equipment_power", 0.0), 4):,}</td>
                                        <td style="border-right: 2px solid black;">{round(baseline_schedule_summary.get("associated_peak_internal_gain", 0.0), 4):,}</td>
                                        <td>{round(proposed_schedule_summary.get("EFLH", 0)):,}</td>
                                        <td>{round(proposed_schedule_summary.get("associated_floor_area", 0.0), 4):,}</td>
                                        <td>{round(proposed_schedule_summary.get("percent_total_lighting_power", 0.0), 4):,}</td>
                                        <td>{round(proposed_schedule_summary.get("percent_total_equipment_power", 0.0), 4):,}</td>
                                        <td>{round(proposed_schedule_summary.get("associated_peak_internal_gain", 0.0), 4):,}</td>
                                    </tr>
                """
            )
        file.write(f"""
                                    
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>

                <div class="mb-3 me-4">
                    <button class="btn btn-info collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapse-hvac-summary" aria-expanded="false">
                        HVAC Summary
                    </button>

                    <div id="collapse-hvac-summary" class="accordion-collapse collapse">
                        <div class="accordion-body">
        """)

        if (rct_detailed_report.proposed_model_summary["chiller_count"] + rct_detailed_report.baseline_model_summary["chiller_count"]) > 0:
            # -------------------------- Cooling Plant Summary Table-------------------------
            file.write(f"""   
                                <h3> Cooling Plant Summary</h3>
                                <table class="table table-sm table-borderless fan-summary" style="width: 1150px;">
                                    <thead>
                                        <tr class="text-center">
                                            <th style="border: 2px solid black; width: 12%;" rowspan="2">Fuel Type</th>
                                            <th style="border: 2px solid black; width: 14%;" colspan="4">Baseline Design</th>
                                            <th style="border: 2px solid black; width: 14%;" colspan="4">Proposed Design</th>
                                        </tr>
                                        <tr class="text-center">
                                            <th style="border: 2px solid black;">Total Quantity of Chillers</th>
                                            <th style="border: 2px solid black;">Total Chiller Plant Capacity [ton]</th>
                                            <th style="border: 2px solid black;">Total Cooling Tower GPM</th>
                                            <th style="border: 2px solid black;">Total Cooling Tower HP</th>
                                            <th style="border: 2px solid black;">Total Quantity of Chillers</th>
                                            <th style="border: 2px solid black;">Total Chiller Plant Capacity [ton]</th>
                                            <th style="border: 2px solid black;">Total Cooling Tower GPM</th>
                                            <th style="border: 2px solid black;">Total Cooling Tower HP</th>
                                        </tr>
                                    </thead>
                                    <tbody style="border: 2px solid black;">
                                        <tr style="font-size: 12px;" class="text-center">
                                            <td style="border-right: 2px solid black;">Electricity</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary.get("electric_chiller_count", 0),):,}</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary.get("electric_chiller_plant_capacity", 0), 1):,}</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary.get("cooling_tower_gpm", 0), 1):,}</td>
                                            <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary.get("cooling_tower_hp", 0), 1):,}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary.get("electric_chiller_count", 0),):,}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary.get("electric_chiller_plant_capacity", 0), 1):,}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary.get("cooling_tower_gpm", 0), 1):,}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary.get("cooling_tower_hp", 0), 1):,}</td>
                                        </tr>
                                        <tr style="font-size: 12px;" class="text-center">
                                            <td style="border-right: 2px solid black;">Fossil Fuel</td>
                                            <td style="background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(rct_detailed_report.proposed_model_summary.get("fossil_fuel_chiller_count", 0),):,}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary.get("fossil_fuel_chiller_plant_capacity", 0.0), 1):,}</td>
                                            <td style="background: black;"></td>
                                            <td style="background: black;"></td>
                                        </tr>
                                        <tr style="font-size: 12px; border-top: 1px solid black;" class="fw-bold text-center subtotal">
                                            <td style="border-right: 2px solid black;">Total</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary.get("electric_chiller_count", 0),):,}</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary.get("electric_chiller_plant_capacity", 0), 1):,}</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary.get("cooling_tower_gpm", 0), 1):,}</td>
                                            <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary.get("cooling_tower_hp", 0), 1):,}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary.get("chiller_count", 0),):,}</td>
                                            <td>{round((rct_detailed_report.proposed_model_summary.get("electric_chiller_plant_capacity", 0) +
                                                    rct_detailed_report.proposed_model_summary.get("fossil_fuel_chiller_plant_capacity", 0)), 1):,}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary.get("cooling_tower_gpm", 0), 1):,}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary.get("cooling_tower_hp", 0), 1):,}</td>
                                        </tr>
                                    </tbody>
                                </table>
            """)

        # -------------------------- Heating Plant Summary Table-------------------------
        if (rct_detailed_report.proposed_model_summary["boiler_count"] + rct_detailed_report.baseline_model_summary["boiler_count"]) > 0:
            file.write(f"""   
                                        <h3> Heating Plant Summary</h3>
                                        <table class="table table-sm table-borderless fan-summary" style="width: 800px;">
                                            <thead>
                                                <tr class="text-center">
                                                    <th style="border: 2px solid black; width: 12%;" rowspan="2">Fuel Type</th>
                                                    <th style="border: 2px solid black; width: 14%;" colspan="2">Baseline Design</th>
                                                    <th style="border: 2px solid black; width: 14%;" colspan="2">Proposed Design</th>
                                                </tr>
                                                <tr class="text-center">
                                                    <th style="border: 2px solid black;">Total Quantity of Boilers</th>
                                                    <th style="border: 2px solid black;">Total Boiler Plant Capacity [Btu/hr]</th>
                                                    <th style="border: 2px solid black;">Total Quantity of Boilers</th>
                                                    <th style="border: 2px solid black;">Total Boiler Plant Capacity [Btu/hr]</th>
                                                </tr>
                                            </thead>
                                            <tbody style="border: 2px solid black;">
                                                <tr style="font-size: 12px;" class="text-center">
                                                    <td style="border-right: 2px solid black;">Electricity</td>
                                                    <td style="background: black;"></td>
                                                    <td style="border-right: 2px solid black; background: black;"></td>
                                                    <td>{round(rct_detailed_report.proposed_model_summary.get("electric_boiler_count", 0),):,}</td>
                                                    <td>{round(rct_detailed_report.proposed_model_summary.get("electric_boiler_plant_capacity", 0)):,}</td>
                                                </tr>
                                                <tr style="font-size: 12px;" class="text-center">
                                                    <td style="border-right: 2px solid black;">Fossil Fuel</td>
                                                    <td>{round(rct_detailed_report.baseline_model_summary.get("fossil_fuel_boiler_count", 0),):,}</td>
                                                    <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary.get("fossil_fuel_boiler_plant_capacity", 0)):,}</td>
                                                    <td>{round(rct_detailed_report.proposed_model_summary.get("fossil_fuel_boiler_count", 0),):,}</td>
                                                    <td>{round(rct_detailed_report.proposed_model_summary.get("fossil_fuel_boiler_plant_capacity", 0)):,}</td>
                                                </tr>
                                                <tr style="font-size: 12px; border-top: 1px solid black;" class="fw-bold text-center subtotal">
                                                    <td style="border-right: 2px solid black;">Total</td>
                                                    <td>{round(rct_detailed_report.baseline_model_summary.get("fossil_fuel_boiler_count", 0)):,}</td>
                                                    <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary.get("fossil_fuel_boiler_plant_capacity", 0)):,}</td>
                                                    <td>{round(rct_detailed_report.proposed_model_summary.get("boiler_count", 0),):,}</td>
                                                    <td>{round((rct_detailed_report.proposed_model_summary.get("electric_boiler_plant_capacity", 0) +
                                                                rct_detailed_report.proposed_model_summary.get("fossil_fuel_boiler_plant_capacity", 0))):,}</td>
                                                </tr>
                                            </tbody>
                                        </table>
            """)

        # -------------------------- Air-Side HVAC Capacity Summary Table-------------------------
        file.write(f"""   
                            <h3> Air-side HVAC Capacity Summary</h3>
                            <table class="table table-sm table-borderless fan-summary" style="width: 750px;">
                                <thead>
                                    <tr class="text-center">
                                        <th style="border: 2px solid black; width: 12%;" rowspan="2">Fuel Type</th>
                                        <th style="border: 2px solid black; width: 14%;" colspan="2">Baseline Design</th>
                                        <th style="border: 2px solid black; width: 14%;" colspan="2">Proposed Design</th>
                                    </tr>
                                    <tr class="text-center">
                                        <th style="border: 2px solid black;">Heating Capacity [kBtu/hr]</th>
                                        <th style="border: 2px solid black;">Cooling Capacity [kBtu/hr]</th>
                                        <th style="border: 2px solid black;">Heating Capacity [kBtu/hr]</th>
                                        <th style="border: 2px solid black;">Cooling Capacity [kBtu/hr]</th>
                                    </tr>
                                </thead>
                                <tbody style="border: 2px solid black;">
                                    <tr style="font-size: 12px;" class="text-center">
                                        <td style="border-right: 2px solid black;">Electricity</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['heating_capacity_by_fuel_type'].get("Electricity", 0.0)):,}</td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['cooling_capacity_by_fuel_type'].get("Electricity", 0.0)):,}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['heating_capacity_by_fuel_type'].get("Electricity", 0.0)):,}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['cooling_capacity_by_fuel_type'].get("Electricity", 0.0)):,}</td>
                                    </tr>
                                    <tr style="font-size: 12px;" class="text-center">
                                        <td style="border-right: 2px solid black;">Fossil Fuel</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['heating_capacity_by_fuel_type'].get("Fossil Fuel", 0.0)):,}</td>
                                        <td style="background: black; border-right: 2px solid black;"></td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['heating_capacity_by_fuel_type'].get("Fossil Fuel", 0.0)):,}</td>
                                        <td style="background: black;"></td>
                                    </tr>
                                    <tr style="font-size: 12px;" class="text-center">
                                        <td style="border-right: 2px solid black;">On-site Boiler Plant</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['heating_capacity_by_fuel_type'].get("On-site Boiler Plant", 0.0)):,}</td>
                                        <td style="background: black; border-right: 2px solid black;"></td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['heating_capacity_by_fuel_type'].get("On-site Boiler Plant", 0.0)):,}</td>
                                        <td style="background: black;"></td>
                                    </tr>
                                    <tr style="font-size: 12px;" class="text-center">
                                        <td style="border-right: 2px solid black;">Purchased Heat</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['heating_capacity_by_fuel_type'].get("Purchased Heat", 0.0)):,}</td>
                                        <td style="background: black; border-right: 2px solid black;"></td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['heating_capacity_by_fuel_type'].get("Purchased Heat", 0.0)):,}</td>
                                        <td style="background: black;"></td>
                                    </tr>
                                    <tr style="font-size: 12px;" class="text-center">
                                        <td style="border-right: 2px solid black;">On-site Chiller Plant</td>
                                        <td style="background: black;"></td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['cooling_capacity_by_fuel_type'].get("On-site Chiller Plant", 0.0)):,}</td>
                                        <td style="background: black;"></td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['cooling_capacity_by_fuel_type'].get("On-site Chiller Plant", 0.0)):,}</td>
                                    </tr>
                                    <tr style="font-size: 12px;" class="text-center">
                                        <td style="border-right: 2px solid black;">Purchased CHW</td>
                                        <td style="background: black;"></td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['cooling_capacity_by_fuel_type'].get("Purchased CHW", 0.0)):,}</td>
                                        <td style="background: black;"></td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['cooling_capacity_by_fuel_type'].get("Purchased CHW", 0.0)):,}</td>
                                    </tr>
                                    <tr style="font-size: 12px; border-top: 1px solid black;" class="fw-bold text-center subtotal">
                                        <td style="border-right: 2px solid black;">Total</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['heating_capacity_by_fuel_type'].get("Total", 0.0)):,}</td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['cooling_capacity_by_fuel_type'].get("Total", 0.0)):,}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['heating_capacity_by_fuel_type'].get("Total", 0.0)):,}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['cooling_capacity_by_fuel_type'].get("Total", 0.0)):,}</td>
                                    </tr>
                                </tbody>
                            </table>
        """)

        # ----------------------- HVAC Fan Summary Table -----------------------
        file.write(f"""
                            <h3>Baseline HVAC Fan Summary</h3>
                            <p><strong>Outdoor Airflow:</strong> {round(rct_detailed_report.baseline_model_summary['total_zone_minimum_oa_flow']):,} CFM</p>
                            <table class="table table-sm table-borderless fan-summary" style="width: 1250px;">
                                <thead>
                                    <tr class="text-center">
                                        <th style="border: 2px solid black; width: 12%;" rowspan="2">Fan Type</th>
                                        <th style="border: 2px solid black; width: 14%;" colspan="3">Constant Volume</th>
                                        <th style="border: 2px solid black; width: 14%;" colspan="3">Variable Volume</th>
                                        <th style="border: 2px solid black; width: 14%;" colspan="3">Multispeed</th>
                                        <th style="border: 2px solid black; width: 14%;" colspan="3">Constant Volume, Cycling</th>
                                        <th style="border: 2px solid black; width: 14%;" colspan="3">Other</th>
                                        <th style="border: 2px solid black; width: 18%;" colspan="4">Total</th>
                                    </tr>
                                    <tr class="text-center">
                                        <th style="border: 2px solid black;">CFM</th>
                                        <th style="border: 2px solid black;">kW</th>
                                        <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                        <th style="border: 2px solid black;">CFM</th>
                                        <th style="border: 2px solid black;">kW</th>
                                        <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                        <th style="border: 2px solid black;">CFM</th>
                                        <th style="border: 2px solid black;">kW</th>
                                        <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                        <th style="border: 2px solid black;">CFM</th>
                                        <th style="border: 2px solid black;">kW</th>
                                        <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                        <th style="border: 2px solid black;">CFM</th>
                                        <th style="border: 2px solid black;">kW</th>
                                        <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                        <th style="border: 2px solid black;">CFM</th>
                                        <th style="border: 2px solid black;">kW</th>
                                        <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                        <th style="border: 2px solid black;">% of Subtotal kW</th>
                                    </tr>
                                </thead>
                                <tbody style="border: 2px solid black;">
        """)

        for fan_type in ["Supply", "Return/Relief", "Exhaust", "Zonal Exhaust"]:
            file.write(
                f"""
                                    <tr style="font-size: 12px;" class="text-center">
                                        <td style="border-right: 2px solid black;">{fan_type}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("CONSTANT", {}).get(fan_type, 0)):,}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("CONSTANT", {}).get(fan_type, 0) / 1000, 2):,}</td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("CONSTANT", {}).get(fan_type, 0) / (rct_detailed_report.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("CONSTANT", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get(fan_type, 0)):,}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get(fan_type, 0) / 1000, 2):,}</td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get(fan_type, 0) / (rct_detailed_report.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get(fan_type, 0)):,}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get(fan_type, 0) / 1000, 2):,}</td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get(fan_type, 0) / (rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get(fan_type, 0)):,}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get(fan_type, 0) / 1000, 2):,}</td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get(fan_type, 0) / (rct_detailed_report.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['other_air_flow_by_fan_type'].get(fan_type, 0)):,}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['other_fan_power_by_fan_type'].get(fan_type, 0) / 1000, 2):,}</td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['other_fan_power_by_fan_type'].get(fan_type, 0) / (rct_detailed_report.baseline_model_summary['other_air_flow_by_fan_type'].get("Supply", 99999999) or 99999999), 4)}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_air_flow_by_fan_type'].get(fan_type, 0)):,}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_type'].get(fan_type, 0) / 1000, 2):,}</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_type'].get(fan_type, 0) / (rct_detailed_report.baseline_model_summary['total_air_flow_by_fan_type'].get("Supply", 99999999) or 99999999), 4)}</td>
                                        <td>{round(100 * rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_type'].get(fan_type, 0) / sum(rct_detailed_report.baseline_model_summary["total_fan_power_by_fan_type"].values()))}</td>
                                    </tr>
                """
            )
        # --------- Subtotal Row -------------
        file.write(f"""
                                    <tr style="font-size: 12px; border-top: 1px solid black;" class="fw-bold text-center subtotal">
                                        <td style="border-right: 2px solid black;">Subtotal</td>
                                        <td></td>
                                        <td></td>
                                        <td style="border-right: 2px solid black;"></td>
                                        <td></td>
                                        <td></td>
                                        <td style="border-right: 2px solid black;"></td>
                                        <td></td>
                                        <td></td>
                                        <td style="border-right: 2px solid black;"></td>
                                        <td></td>
                                        <td></td>
                                        <td style="border-right: 2px solid black;"></td>
                                        <td></td>
                                        <td></td>
                                        <td style="border-right: 2px solid black;"></td>
                                        <td></td>
                                        <td></td>
                                        <td></td>
                                        <td>0</td>
                                    </tr>
                    """)
        # --------- Terminal Units Row ------------
        file.write(f"""
                                    <tr style="font-size: 12px; border-top: 1px solid black;" class="text-center">
                                        <td style="border-right: 2px solid black;">Terminal Units</td>
                                        <td style="background: black;"></td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("CONSTANT", {}).get("Terminal Unit", 0)):,}</td>
                                        <td style="border-right: 2px solid black; background: black;"></td>
                                        <td style="background: black;"></td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get("Terminal Unit", 0)):,}</td>
                                        <td style="border-right: 2px solid black; background: black;"></td>
                                        <td style="background: black;"></td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get("Terminal Unit", 0)):,}</td>
                                        <td style="border-right: 2px solid black; background: black;"></td>
                                        <td style="background: black;"></td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get("Terminal Unit", 0)):,}</td>
                                        <td style="border-right: 2px solid black; background: black;"></td>
                                        <td style="background: black;"></td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['other_fan_power_by_fan_type'].get("Terminal Unit", 0)):,}</td>
                                        <td style="border-right: 2px solid black; background: black;"></td>
                                        <td style="background: black;"></td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_type'].get("Terminal Unit", 0)):,}</td>
                                        <td style="background: black;"></td>
                                        <td style="background: black;"></td>
                                    </tr>
        """)
        file.write(f"""
                                </tbody>
                            </table>

                            <h3>Proposed HVAC Fan Summary</h3>
                            <p><strong>Outdoor Airflow:</strong> {round(rct_detailed_report.baseline_model_summary['total_zone_minimum_oa_flow']):,} CFM</p>
                            <table class="table table-sm table-borderless fan-summary" style="width: 1250px;">
                                <thead>
                                    <tr class="text-center">
                                        <th style="border: 2px solid black; width: 12%;" rowspan="2">Fan Type</th>
                                        <th style="border: 2px solid black; width: 14%;" colspan="3">Constant Volume</th>
                                        <th style="border: 2px solid black; width: 14%;" colspan="3">Variable Volume</th>
                                        <th style="border: 2px solid black; width: 14%;" colspan="3">Multispeed</th>
                                        <th style="border: 2px solid black; width: 14%;" colspan="3">Constant Volume, Cycling</th>
                                        <th style="border: 2px solid black; width: 14%;" colspan="3">Other</th>
                                        <th style="border: 2px solid black; width: 18%;" colspan="4">Total</th>
                                    </tr>
                                    <tr class="text-center">
                                        <th style="border: 2px solid black;">CFM</th>
                                        <th style="border: 2px solid black;">kW</th>
                                        <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                        <th style="border: 2px solid black;">CFM</th>
                                        <th style="border: 2px solid black;">kW</th>
                                        <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                        <th style="border: 2px solid black;">CFM</th>
                                        <th style="border: 2px solid black;">kW</th>
                                        <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                        <th style="border: 2px solid black;">CFM</th>
                                        <th style="border: 2px solid black;">kW</th>
                                        <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                        <th style="border: 2px solid black;">CFM</th>
                                        <th style="border: 2px solid black;">kW</th>
                                        <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                        <th style="border: 2px solid black;">CFM</th>
                                        <th style="border: 2px solid black;">kW</th>
                                        <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                        <th style="border: 2px solid black;">% of Subtotal kW</th>
                                    </tr>
                                </thead>
                                <tbody style="border: 2px solid black;">
        """)

        for fan_type in ["Supply", "Return/Relief", "Exhaust", "Zonal Exhaust"]:
            file.write(
                f"""
                                    <tr style="font-size: 12px;" class="text-center">
                                        <td style="border-right: 2px solid black;">{fan_type}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("CONSTANT", {}).get(fan_type, 0)):,}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("CONSTANT", {}).get(fan_type, 0) / 1000, 2):,}</td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("CONSTANT", {}).get(fan_type, 0) / (rct_detailed_report.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("CONSTANT", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get(fan_type, 0)):,}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get(fan_type, 0) / 1000, 2):,}</td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get(fan_type, 0) / (rct_detailed_report.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get(fan_type, 0)):,}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get(fan_type, 0) / 1000, 2):,}</td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get(fan_type, 0) / (rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get(fan_type, 0)):,}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get(fan_type, 0) / 1000, 2):,}</td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get(fan_type, 0) / (rct_detailed_report.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['other_air_flow_by_fan_type'].get(fan_type, 0)):,}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['other_fan_power_by_fan_type'].get(fan_type, 0) / 1000, 2):,}</td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.proposed_model_summary['other_fan_power_by_fan_type'].get(fan_type, 0) / (rct_detailed_report.proposed_model_summary['other_air_flow_by_fan_type'].get("Supply", 99999999) or 99999999), 4)}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_air_flow_by_fan_type'].get(fan_type, 0)):,}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_type'].get(fan_type, 0) / 1000, 2):,}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_type'].get(fan_type, 0) / (rct_detailed_report.proposed_model_summary['total_air_flow_by_fan_type'].get("Supply", 99999999) or 99999999), 4)}</td>
                                        <td>{round(100 * rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_type'].get(fan_type, 0) / sum(rct_detailed_report.proposed_model_summary["total_fan_power_by_fan_type"].values()))}</td>
                                    </tr>
                """
            )
        # ---------- Subtotal Row -------------
        file.write(f"""
                                    <tr style="font-size: 12px; border-top: 1px solid black;" class="fw-bold text-center subtotal">
                                        <td style="border-right: 2px solid black;">Subtotal</td>
                                        <td></td>
                                        <td></td>
                                        <td style="border-right: 2px solid black;"></td>
                                        <td></td>
                                        <td></td>
                                        <td style="border-right: 2px solid black;"></td>
                                        <td></td>
                                        <td></td>
                                        <td style="border-right: 2px solid black;"></td>
                                        <td></td>
                                        <td></td>
                                        <td style="border-right: 2px solid black;"></td>
                                        <td></td>
                                        <td></td>
                                        <td style="border-right: 2px solid black;"></td>
                                        <td></td>
                                        <td></td>
                                        <td></td>
                                        <td>0</td>
                                    </tr>
        """)
        # --------- Terminal Units Row -----------
        file.write(f"""
                                    <tr style="font-size: 12px; border-top: 1px solid black;" class="text-center">
                                        <td>Terminal Units</td>
                                        <td style="background: black;"></td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("CONSTANT", {}).get("Terminal Unit", 0)):,}</td>
                                        <td style="border-right: 2px solid black; background: black;"></td>
                                        <td style="background: black;"></td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get("Terminal Unit", 0)):,}</td>
                                        <td style="border-right: 2px solid black; background: black;"></td>
                                        <td style="background: black;"></td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get("Terminal Unit", 0)):,}</td>
                                        <td style="border-right: 2px solid black; background: black;"></td>
                                        <td style="background: black;"></td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get("Terminal Unit", 0)):,}</td>
                                        <td style="border-right: 2px solid black; background: black;"></td>
                                        <td style="background: black;"></td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['other_fan_power_by_fan_type'].get("Terminal Unit", 0)):,}</td>
                                        <td style="border-right: 2px solid black; background: black;"></td>
                                        <td style="background: black;"></td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_type'].get("Terminal Unit", 0)):,}</td>
                                        <td style="background: black;"></td>
                                        <td style="background: black;"></td>
                                    </tr>
                                </tbody>
                            </table>
        """)

        file.write(f""" 
                    </div>
                </div>
            </div>
        """
       )

        for category, rules in rule_categories.items():
            btn_class = (
                "btn-danger"
                if category == "Failing"
                else "btn-warning"
                if category == "Undetermined"
                else "btn-success"
                if category == "Passing"
                else "btn-secondary"
            )
            file.write(
                f"""
                    <div class="mb-3 me-4">
                        <button class="btn {btn_class} w-100 text-start sticky-top" 
                            type="button" data-bs-toggle="collapse" data-bs-target="#collapse_fully_{category.replace(' ', '_')}">
                            <strong>{category} Rules ({len(rules)})</strong>
                        </button>
                        <div class="collapse mx-4" id="collapse_fully_{category.replace(' ', '_')}">
                """
            )
            if category == "Undetermined":
                file.write(
                    f"""
                    <h3 class="mt-4">Rules Fully Evaluated</h3>
                    """
                )
            file.write(
                f"""
                            <table class="table table-bordered table-striped mt-2">
                                <thead class="table-dark">
                                    <tr>
                                        <th rowspan='2'>Rule ID</th>
                                        <th>Description</th>
                                        <th>Standard Section</th>
                                        <th>Outcome Counts</th>
                                    </tr>
                                    <tr><th colspan='3'>Evaluations</th></tr>
                                </thead>
                                <tbody>
                    """
            )

            if category == "Undetermined":
                sections_seen = set()
                for rule_id in rct_detailed_report.full_eval_rules_undetermined:
                    rule_data = next(
                        rule
                        for rule in rct_detailed_report.evaluation_data["rules"]
                        if rule["rule_id"] == rule_id
                    )
                    section = rule_id.split("-")[0]
                    if section not in sections_seen:
                        sections_seen.add(section)
                        section_title = section_titles_with_colors.get(
                            int(section)
                        )[0]
                        section_color = section_titles_with_colors.get(
                            int(section)
                        )[1]
                        file.write(
                            f"""
                            </tbody>
                                <thead class="table-group-divider">
                                    <tr>
                                        <td colspan="4" class="section-title sticky-top sticky-top-2" style="background-color: {section_color} !important;">{section_title}</td>
                                    </tr>
                                </thead>
                            <tbody>
                            """
                        )

                    description = rule_data.get("description", "N/A")
                    standard_section = rule_data.get("standard_section", "N/A")
                    outcome_summary = " | ".join([f"{k}: {v}" for k, v in rct_detailed_report.rule_evaluation_outcome_counts[rule_id].items()])

                    file.write(
                        f"""
                            <tr>
                                <td class="rule-id" rowspan='2'>{rule_id}</td>
                                <td>{description}</td>
                                <td>{standard_section}</td>
                                <td class="outcome-summary">{outcome_summary}</td>
                            </tr>
                            <tr>
                                <td colspan='3'>
                                    <button class="btn btn-primary" type="button" data-bs-toggle="collapse" data-bs-target="#eval_{rule_id}">
                                        View Evaluations
                                    </button>
                                    <div class="collapse" id="eval_{rule_id}">
                                        <ul>
                        """
                    )
                    outcome_order = {
                        "FAILED": 0,
                        "UNDETERMINED": 1,
                        "PASS": 2,
                        "NOT_APPLICABLE": 3,
                    }

                    # Sort evaluations based on outcome priority
                    sorted_evaluations = sorted(
                        rule_data["evaluations"],
                        key=lambda e: outcome_order.get(e["outcome"], 3),
                    )

                    for evaluation in sorted_evaluations:
                        has_any_units = False
                        styles = {
                            "FAILED": "background-color: #ffcccc; color: black; font-weight: bold; padding-left: 10px; border-radius: 8px; border: 2px solid #ff0000;",
                            "PASS": "background-color: #ccffcc; color: black; font-weight: bold; padding-left: 10px; border-radius: 8px; border: 2px solid #008000;",
                            "UNDETERMINED": "background-color: #ffffcc; color: black; font-weight: bold; padding-left: 10px; border-radius: 8px; border: 2px solid #ffcc00;",
                            "DEFAULT": "padding-left: 10px; border: 2px solid #ccc; border-radius: 8px;",
                        }

                        # Select the appropriate style based on outcome
                        li_style = styles.get(
                            evaluation["outcome"], styles["DEFAULT"]
                        )
                        file.write(
                            f"""
                                <li style=\"{li_style}\"  class=\"p-2 m-1\">{evaluation['data_group_id']}
                                    <ul>
                                        <li><strong>Outcome:</strong> {evaluation['outcome']}</li>
                                """
                        )
                        if evaluation["messages"]:
                            messages = set()
                            if isinstance(evaluation["messages"], str):
                                messages.add(evaluation["messages"])
                            if isinstance(evaluation["messages"], dict):
                                for key, message in evaluation["messages"].items():
                                    messages.add(f"{key}: {message}")
                            if isinstance(evaluation["messages"], list):
                                for message in evaluation["messages"]:
                                    messages.add(message)
                            file.write(
                                f"<li><strong>Messages:</strong> {', '.join(messages)}</li>"
                            )
                        if evaluation["calculated_values"]:
                            file.write(
                                """
                                    <li><strong>Calculated Values:</strong>
                                        <table class="mb-2 me-2 table table-sm table-bordered">
                                            <thead>
                                                <tr><th>Variable</th><th>Value</th>
                                """
                            )
                            if any(
                                    cv.get("unit")
                                    for cv in evaluation["calculated_values"]
                            ):
                                has_any_units = True
                                file.write("<th>Unit</th>")
                            file.write("</tr></thead><tbody>")

                            for calculated_value in evaluation["calculated_values"]:
                                file.write(
                                    f"""
                                    <tr>
                                    <td>{calculated_value['variable']}</td>
                                    <td>{calculated_value['value'][0] if len(calculated_value['value']) == 1
                                    else calculated_value['value']}
                                    </td>
                                    """
                                )
                                if calculated_value.get("unit"):
                                    file.write(
                                        f"<td>{calculated_value['unit']}</td>"
                                    )
                                elif has_any_units:
                                    file.write("<td></td>")
                                file.write("</tr>")
                            file.write("</tbody></table></li>")
                        file.write("</ul></li>")
                    file.write("</ul></div></td></tr>")
                file.write(
                    f"""
                        </tbody>
                        </table>
                        <h3 class="mt-4">Rules Evaluated for Applicability Only</h3>
                        <table class="table table-bordered table-striped mt-2">
                            <thead class="table-dark">
                                <tr>
                                    <th rowspan='2'>Rule ID</th>
                                    <th>Description</th>
                                    <th>Standard Section</th>
                                    <th>Outcome Counts</th>
                                </tr>
                                <tr><th colspan='3'>Evaluations</th></tr>
                            </thead>
                            <tbody>
                    """
                )
                sections_seen = set()
                for rule_id in rct_detailed_report.appl_eval_rules_undetermined:
                    rule_data = next(
                        rule
                        for rule in rct_detailed_report.evaluation_data["rules"]
                        if rule["rule_id"] == rule_id
                    )
                    section = rule_id.split("-")[0]
                    if section not in sections_seen:
                        sections_seen.add(section)
                        section_title = section_titles_with_colors.get(
                            int(section)
                        )[0]
                        section_color = section_titles_with_colors.get(
                            int(section)
                        )[1]
                        file.write(
                            f"""
                            </tbody>
                                <thead class="table-group-divider">
                                    <tr>
                                        <td colspan="4" class="section-title sticky-top sticky-top-2" style="background-color: {section_color} !important;">{section_title}</td>
                                    </tr>
                                </thead>
                            <tbody>
                            """
                        )

                    description = rule_data.get("description", "N/A")
                    standard_section = rule_data.get("standard_section", "N/A")
                    outcome_summary = " | ".join([f"{k}: {v}" for k, v in rct_detailed_report.rule_evaluation_outcome_counts[rule_id].items()])

                    file.write(
                        f"""
                            <tr>
                                <td class="rule-id" rowspan='2'>{rule_id}</td>
                                <td>{description}</td>
                                <td>{standard_section}</td>
                                <td class="outcome-summary">{outcome_summary}</td>
                            </tr>
                            <tr>
                                <td colspan='3'>
                                    <button class="btn btn-primary" type="button" data-bs-toggle="collapse" data-bs-target="#eval_{rule_id}">
                                        View Evaluations
                                    </button>
                                    <div class="collapse" id="eval_{rule_id}">
                                        <ul>
                            """
                    )
                    outcome_order = {
                        "FAILED": 0,
                        "UNDETERMINED": 1,
                        "PASS": 2,
                        "NOT_APPLICABLE": 3,
                    }

                    # Sort evaluations based on outcome priority
                    sorted_evaluations = sorted(
                        rule_data["evaluations"],
                        key=lambda e: outcome_order.get(e["outcome"], 3),
                    )

                    for evaluation in sorted_evaluations:
                        has_any_units = False
                        styles = {
                            "FAILED": "background-color: #ffcccc; color: black; font-weight: bold; padding-left: 10px; border-radius: 8px; border: 2px solid #ff0000;",
                            "PASS": "background-color: #ccffcc; color: black; font-weight: bold; padding-left: 10px; border-radius: 8px; border: 2px solid #008000;",
                            "UNDETERMINED": "background-color: #ffffcc; color: black; font-weight: bold; padding-left: 10px; border-radius: 8px; border: 2px solid #ffcc00;",
                            "DEFAULT": "padding-left: 10px; border: 2px solid #ccc; border-radius: 8px;",
                        }

                        # Select the appropriate style based on outcome
                        li_style = styles.get(
                            evaluation["outcome"], styles["DEFAULT"]
                        )
                        file.write(
                            f"""
                                <li style=\"{li_style}\"  class=\"p-2 m-1\">{evaluation['data_group_id']}
                                    <ul>
                                        <li><strong>Outcome:</strong> {evaluation['outcome']}</li>
                                """
                        )
                        if evaluation["messages"]:
                            messages = set()
                            if isinstance(evaluation["messages"], str):
                                messages.add(evaluation["messages"])
                            if isinstance(evaluation["messages"], dict):
                                for key, message in evaluation["messages"].items():
                                    messages.add(f"{key}: {message}")
                            if isinstance(evaluation["messages"], list):
                                for message in evaluation["messages"]:
                                    messages.add(message)
                            file.write(
                                f"<li><strong>Messages:</strong> {', '.join(messages)}</li>"
                            )
                        if evaluation["calculated_values"]:
                            file.write(
                                """
                                    <li><strong>Calculated Values:</strong>
                                        <table class="mb-2 me-2 table table-sm table-bordered">
                                            <thead>
                                                <tr><th>Variable</th><th>Value</th>
                                """
                            )
                            if any(
                                    cv.get("unit")
                                    for cv in evaluation["calculated_values"]
                            ):
                                has_any_units = True
                                file.write("<th>Unit</th>")
                            file.write("</tr></thead><tbody>")

                            for calculated_value in evaluation["calculated_values"]:
                                file.write(
                                    f"""
                                    <tr>
                                    <td>{calculated_value['variable']}</td>
                                    <td>{calculated_value['value'][0] if len(calculated_value['value']) == 1
                                    else calculated_value['value']}
                                    </td>
                                    """
                                )
                                if calculated_value.get("unit"):
                                    file.write(
                                        f"<td>{calculated_value['unit']}</td>"
                                    )
                                elif has_any_units:
                                    file.write("<td></td>")
                                file.write("</tr>")
                            file.write("</tbody></table></li>")
                        file.write("</ul></li>")
                    file.write("</ul></div></td></tr>")
            else:
                sections_seen = set()
                for rule_id in rules:
                    rule_data = next(
                        rule
                        for rule in rct_detailed_report.evaluation_data["rules"]
                        if rule["rule_id"] == rule_id
                    )
                    section = rule_id.split("-")[0]
                    if section not in sections_seen:
                        sections_seen.add(section)
                        section_title = section_titles_with_colors.get(
                            int(section)
                        )[0]
                        section_color = section_titles_with_colors.get(
                            int(section)
                        )[1]
                        file.write(
                            f"""
                            </tbody>
                                <thead class="table-group-divider">
                                    <tr>
                                        <th colspan="4" class="section-title sticky-top sticky-top-2" style="background-color: {section_color} !important;">{section_title}</th>
                                    </tr>
                                </thead>
                            <tbody>
                            """
                        )

                    description = rule_data.get("description", "N/A")
                    standard_section = rule_data.get("standard_section", "N/A")
                    outcome_summary = " | ".join([f"{k}: {v}" for k, v in rct_detailed_report.rule_evaluation_outcome_counts[rule_id].items()])

                    file.write(
                        f"""
                            <tr>
                                <td class="rule-id" rowspan='2'>{rule_id}</td>
                                <td>{description}</td>
                                <td>{standard_section}</td>
                                <td class="outcome-summary">{outcome_summary}</td>
                            </tr>
                            <tr>
                                <td colspan='3'>
                                    <button class="btn btn-primary" type="button" data-bs-toggle="collapse" data-bs-target="#eval_{rule_id}">
                                        View Evaluations
                                    </button>
                                    <div class="collapse" id="eval_{rule_id}">
                                        <ul>
                        """
                    )
                    outcome_order = {
                        "FAILED": 0,
                        "UNDETERMINED": 1,
                        "PASS": 2,
                        "NOT_APPLICABLE": 3,
                    }

                    # Sort evaluations based on outcome priority
                    sorted_evaluations = sorted(
                        rule_data["evaluations"],
                        key=lambda e: outcome_order.get(e["outcome"], 3),
                    )

                    for evaluation in sorted_evaluations:
                        has_any_units = False
                        styles = {
                            "FAILED": "background-color: #ffcccc; color: black; font-weight: bold; padding-left: 10px; border-radius: 8px; border: 2px solid #ff0000;",
                            "PASS": "background-color: #ccffcc; color: black; font-weight: bold; padding-left: 10px; border-radius: 8px; border: 2px solid #008000;",
                            "UNDETERMINED": "background-color: #ffffcc; color: black; font-weight: bold; padding-left: 10px; border-radius: 8px; border: 2px solid #ffcc00;",
                            "DEFAULT": "padding-left: 10px; border: 2px solid #ccc; border-radius: 8px;",
                        }

                        # Select the appropriate style based on outcome
                        li_style = styles.get(
                            evaluation["outcome"], styles["DEFAULT"]
                        )
                        file.write(
                            f"""
                                <li style=\"{li_style}\"  class=\"p-2 m-1\">{evaluation['data_group_id']}
                                    <ul>
                                        <li><strong>Outcome:</strong> {evaluation['outcome']}</li>
                                """
                        )
                        if evaluation["messages"]:
                            messages = set()
                            if isinstance(evaluation["messages"], str):
                                messages.add(evaluation["messages"])
                            if isinstance(evaluation["messages"], dict):
                                for key, message in evaluation["messages"].items():
                                    messages.add(f"{key}: {message}")
                            if isinstance(evaluation["messages"], list):
                                for message in evaluation["messages"]:
                                    messages.add(message)
                            file.write(
                                f"<li><strong>Messages:</strong> {', '.join(messages)}</li>"
                            )
                        if evaluation["calculated_values"]:
                            file.write(
                                """
                                    <li><strong>Calculated Values:</strong>
                                        <table class="mb-2 me-2 table table-sm table-bordered">
                                            <thead>
                                                <tr><th>Variable</th><th>Value</th>
                                """
                            )
                            if any(
                                    cv.get("unit")
                                    for cv in evaluation["calculated_values"]
                            ):
                                has_any_units = True
                                file.write("<th>Unit</th>")
                            file.write("</tr></thead><tbody>")

                            for calculated_value in evaluation["calculated_values"]:
                                file.write(
                                    f"""
                                    <tr>
                                    <td>{calculated_value['variable']}</td>
                                    <td>{calculated_value['value'][0] if len(calculated_value['value']) == 1
                                    else calculated_value['value']}
                                    </td>
                                    """
                                )
                                if calculated_value.get("unit"):
                                    file.write(
                                        f"<td>{calculated_value['unit']}</td>"
                                    )
                                elif has_any_units:
                                    file.write("<td></td>")
                                file.write("</tr>")
                            file.write("</tbody></table></li>")
                        file.write("</ul></li>")
                    file.write("</ul></div></td></tr>")

            file.write("</tbody></table></div></div>")

        file.write("</div></div>")
        file.write(
            """
        <div class="position-fixed bottom-0 end-0 mb-2 me-2" style="z-index: 1050;">
            <button id="back-to-top" class="btn btn-primary" onclick="scrollToTop()" style="opacity: 0; visibility: hidden;"> ↑ </button>
        </div>
        """
        )
        file.write("</body>")
        file.write(
            f"""
            <script>
            window.onscroll = function() {{
                toggleBackToTopButton();
            }};

            function toggleBackToTopButton() {{
                const backToTopButton = document.getElementById("back-to-top");
                if (document.body.scrollTop > 100 || document.documentElement.scrollTop > 100) {{
                    backToTopButton.style.opacity = "1";
                    backToTopButton.style.visibility = "visible";
                }}  
                else {{
                    backToTopButton.style.opacity = "0";
                    backToTopButton.style.visibility = "hidden";
                }}
            }}

            function scrollToTop() {{
                window.scrollTo({{
                    top: 0,
                    behavior: 'smooth'
                }});
            }}

            function calculateSubtotals() {{
                document.querySelectorAll(".fan-summary").forEach(table => {{
                    let columnSums = [];
                    let columnPrecisions = [];

                    table.querySelectorAll("tr").forEach(row => {{
                        if (row.classList.contains("subtotal")) {{
                            row.querySelectorAll("td").forEach((td, colIndex) => {{
                                if (colIndex === 0) return;
                                let sum = columnSums[colIndex] || 0;
                                let precision = columnPrecisions[colIndex] || 0;
                                td.textContent = sum.toLocaleString(undefined, {{ minimumFractionDigits: precision, maximumFractionDigits: precision }});
                            }});
                            columnSums = [];
                            columnPrecisions = [];
                        }} else {{
                            row.querySelectorAll("td").forEach((td, colIndex) => {{
                                let cleanedText = td.textContent.replace(/,/g, "").trim();
                                let value = parseFloat(cleanedText) || 0;
                                let decimalPlaces = (cleanedText.split(".")[1] || "").length;
                                columnPrecisions[colIndex] = Math.max(columnPrecisions[colIndex] || 0, decimalPlaces);
                                columnSums[colIndex] = (columnSums[colIndex] || 0) + value;
                            }});
                        }}
                    }});
                }});
            }}

            document.addEventListener("DOMContentLoaded", () => {{
                calculateSubtotals();

                // Chart labels
                const labels = {[label.replace('_', ' ').title() for label in rct_detailed_report.baseline_model_summary["elec_by_end_use"].keys()]};

                const elecDataRaw = {{
                  consumption: {{
                    baseline: {list(rct_detailed_report.baseline_model_summary["elec_by_end_use"].values())},
                    proposed: {list(rct_detailed_report.proposed_model_summary["elec_by_end_use"].values())}
                  }},
                  eui: {{
                    baseline: {list(rct_detailed_report.baseline_model_summary["elec_by_end_use_eui"].values())},
                    proposed: {list(rct_detailed_report.proposed_model_summary["elec_by_end_use_eui"].values())}
                  }}
                }};

                const gasDataRaw = {{
                  consumption: {{
                    baseline: {list(rct_detailed_report.baseline_model_summary["gas_by_end_use"].values())},
                    proposed: {list(rct_detailed_report.proposed_model_summary["gas_by_end_use"].values())}
                  }},
                  eui: {{
                    baseline: {list(rct_detailed_report.baseline_model_summary["gas_by_end_use_eui"].values())},
                    proposed: {list(rct_detailed_report.proposed_model_summary["gas_by_end_use_eui"].values())}
                  }}
                }};

                const energyDataRaw = {{
                  consumption: {{
                    baseline: {list(rct_detailed_report.baseline_model_summary["energy_by_end_use"].values())},
                    proposed: {list(rct_detailed_report.proposed_model_summary["energy_by_end_use"].values())}
                  }},
                  eui: {{
                    baseline: {list(rct_detailed_report.baseline_model_summary["energy_by_end_use_eui"].values())},
                    proposed: {list(rct_detailed_report.proposed_model_summary["energy_by_end_use_eui"].values())}
                  }}
                }};

                // Electricity Datasets
                const elecData = {{
                    labels: labels,
                    datasets: [
                        {{
                            label: 'Baseline',
                            data: {list(rct_detailed_report.baseline_model_summary["elec_by_end_use"].values())},
                            backgroundColor: 'rgba(54, 162, 235, 0.7)'
                        }},
                        {{
                            label: 'Proposed',
                            data: {list(rct_detailed_report.proposed_model_summary["elec_by_end_use"].values())},
                            backgroundColor: 'rgba(75, 192, 75, 0.7)'
                        }}
                    ]
                }};

                const elecConfig = {{
                    type: 'bar',
                    data: elecData,
                    options: {{
                        responsive: true,
                        plugins: {{
                            title: {{
                                display: true,
                                text: 'Electricity By End Use'
                            }},
                            tooltip: {{
                                mode: 'index',
                                intersect: false
                            }}
                        }},
                        interaction: {{
                            mode: 'index',
                            intersect: false
                        }},
                        scales: {{
                            x: {{
                                stacked: false,
                                ticks: {{
                                    minRotation: 60,
                                    maxRotation: 60
                                }}
                            }},
                            y: {{
                                beginAtZero: true,
                                title: {{
                                    display: true,
                                    text: 'kWh',
                                    font: {{
                                        size: 14
                                    }}
                                }}
                            }}
                        }}
                    }}
                }};

                const gasData = {{
                    labels: labels,
                    datasets: [
                        {{
                            label: 'Baseline',
                            data: {list(rct_detailed_report.baseline_model_summary["gas_by_end_use"].values())},
                            backgroundColor: 'rgba(255, 180, 80, 0.5)'
                        }},
                        {{
                            label: 'Proposed',
                            data: {list(rct_detailed_report.proposed_model_summary["gas_by_end_use"].values())},
                            backgroundColor: 'rgba(255, 100, 100, 0.5)'
                        }}
                    ]
                }};

                const gasConfig = {{
                    type: 'bar',
                    data: gasData,
                    options: {{
                        responsive: true,
                        plugins: {{
                            title: {{
                                display: true,
                                text: 'Natural Gas By End Use'
                            }},
                            tooltip: {{
                                mode: 'index',
                                intersect: false
                            }}
                        }},
                        interaction: {{
                            mode: 'index',
                            intersect: false
                        }},
                        scales: {{
                            x: {{
                                stacked: false,
                                ticks: {{
                                    minRotation: 60,
                                    maxRotation: 60
                                }}
                            }},
                            y: {{
                                beginAtZero: true,
                                title: {{
                                    display: true,
                                    text: 'Therms',
                                    font: {{
                                        size: 14
                                    }}
                                }}
                            }}
                        }}
                    }}
                }};

                const energyData = {{
                    labels: labels,
                    datasets: [
                        {{
                            label: 'Baseline',
                            data: {list(rct_detailed_report.baseline_model_summary["energy_by_end_use"].values())},
                            backgroundColor: 'rgba(128, 0, 64, 0.6)'
                        }},
                        {{
                            label: 'Proposed',
                            data: {list(rct_detailed_report.proposed_model_summary["energy_by_end_use"].values())},
                            backgroundColor: 'rgba(0, 128, 128, 0.6)'
                        }}
                    ]
                }};

                const energyConfig = {{
                    type: 'bar',
                    data: energyData,
                    options: {{
                        responsive: true,
                        plugins: {{
                            title: {{
                                display: true,
                                text: 'Total Site Energy By End Use'
                            }},
                            tooltip: {{
                                mode: 'index',
                                intersect: false
                            }}
                        }},
                        interaction: {{
                            mode: 'index',
                            intersect: false
                        }},
                        scales: {{
                            x: {{
                                stacked: false,
                                ticks: {{
                                    minRotation: 60,
                                    maxRotation: 60
                                }}
                            }},
                            y: {{
                                beginAtZero: true,
                                title: {{
                                    display: true,
                                    text: 'kBtu',
                                    font: {{
                                        size: 14
                                    }}
                                }}
                            }}
                        }}
                    }}
                }};

                const elecChart = new Chart(document.getElementById('elecByEndUse'), elecConfig);
                const gasChart = new Chart(document.getElementById('gasByEndUse'), gasConfig);
                const energyChart = new Chart(document.getElementById('energyByEndUse'), energyConfig);

                function updateCharts(unitType) {{
                  // Update Electricity
                  elecChart.data.datasets[0].data = elecDataRaw[unitType].baseline;
                  elecChart.data.datasets[1].data = elecDataRaw[unitType].proposed;
                  elecChart.options.scales.y.title.text = unitType === 'consumption' ? 'kWh' : 'kBtu/ft²';
                  elecChart.update();

                  // Update Gas
                  gasChart.data.datasets[0].data = gasDataRaw[unitType].baseline;
                  gasChart.data.datasets[1].data = gasDataRaw[unitType].proposed;
                  gasChart.options.scales.y.title.text = unitType === 'consumption' ? 'Therms' : 'kBtu/ft²';
                  gasChart.update();

                  // Update Total Energy
                  energyChart.data.datasets[0].data = energyDataRaw[unitType].baseline;
                  energyChart.data.datasets[1].data = energyDataRaw[unitType].proposed;
                  energyChart.options.scales.y.title.text = unitType === 'consumption' ? 'kBtu' : 'kBtu/ft²';
                  energyChart.update();
                }}

                function sumArray(arr) {{
                  return arr.reduce((acc, val) => acc + val, 0);
                }}

                function getUnitLabel(source, unitType) {{
                  if (unitType === 'eui') {{
                    return 'kBtu/ft²';
                  }} else {{
                    return source === 'elec' ? 'kWh' : source === 'gas' ? 'Therms' : 'kBtu';
                  }}
                }}

                function updateTotalColors(source) {{
                  console.log(source);
                  const baselineEl = document.getElementById('baselineTotal');
                  const proposedEl = document.getElementById('proposedTotal');

                  if (source === 'elec') {{
                    baselineEl.style.color = 'rgb(54, 162, 235)'; // Blue
                    proposedEl.style.color = 'rgb(75, 192, 75)';  // Green
                  }} else if (source === 'gas') {{
                    baselineEl.style.color = 'rgb(255, 180, 80)'; // Orange
                    proposedEl.style.color = 'rgb(255, 100, 100)'; // Red
                  }} else if (source === 'energy') {{
                      baselineEl.style.color = 'rgb(128, 0, 64)';   // Maroon
                      proposedEl.style.color = 'rgb(0, 128, 128)';  // Teal
                    }}
                }}

                function updateTotals(source, unitType) {{
                  console.log(source);
                  let baseline, proposed;

                  if (source === 'elec') {{
                    baseline = unitType === 'eui'
                      ? {list(rct_detailed_report.baseline_model_summary["elec_by_end_use_eui"].values())}
                      : {list(rct_detailed_report.baseline_model_summary["elec_by_end_use"].values())};

                    proposed = unitType === 'eui'
                      ? {list(rct_detailed_report.proposed_model_summary["elec_by_end_use_eui"].values())}
                      : {list(rct_detailed_report.proposed_model_summary["elec_by_end_use"].values())};

                  }} else if (source === 'gas') {{
                    baseline = unitType === 'eui'
                      ? {list(rct_detailed_report.baseline_model_summary["gas_by_end_use_eui"].values())}
                      : {list(rct_detailed_report.baseline_model_summary["gas_by_end_use"].values())};

                    proposed = unitType === 'eui'
                      ? {list(rct_detailed_report.proposed_model_summary["gas_by_end_use_eui"].values())}
                      : {list(rct_detailed_report.proposed_model_summary["gas_by_end_use"].values())};

                  }} else if (source === 'energy') {{
                    baseline = unitType === 'eui'
                      ? {list(rct_detailed_report.baseline_model_summary["energy_by_end_use_eui"].values())}
                      : {list(rct_detailed_report.baseline_model_summary["energy_by_end_use"].values())};

                    proposed = unitType === 'eui'
                      ? {list(rct_detailed_report.proposed_model_summary["energy_by_end_use_eui"].values())}
                      : {list(rct_detailed_report.proposed_model_summary["energy_by_end_use"].values())};
                  }}

                  const unit = getUnitLabel(source, unitType);
                  const baselineSum = sumArray(baseline).toLocaleString(undefined, {{ maximumFractionDigits: 0 }});
                  const proposedSum = sumArray(proposed).toLocaleString(undefined, {{ maximumFractionDigits: 0 }});

                  document.getElementById('baselineTotal').textContent = `Baseline Total: ${{baselineSum}} ${{unit}}`;
                  document.getElementById('proposedTotal').textContent = `Proposed Total: ${{proposedSum}} ${{unit}}`;
                }}

                let currentChart = 'elec';

                window.toggleUnits = function() {{
                  const useEUI = document.getElementById('unitToggle').checked;
                  const unitType = useEUI ? 'eui' : 'consumption';
                  updateCharts(unitType);
                  updateTotals(currentChart, unitType);
                }};

                window.showChart = function(type) {{
                  const elecContainer = document.getElementById('elecChartContainer');
                  const gasContainer = document.getElementById('gasChartContainer');
                  const energyContainer = document.getElementById('energyChartContainer');
                  elecContainer.style.display = type === 'elec' ? 'block' : 'none';
                  gasContainer.style.display = type === 'gas' ? 'block' : 'none';
                  energyContainer.style.display = type === 'energy' ? 'block' : 'none';
                  currentChart = type;
                  const useEUI = document.getElementById('unitToggle').checked;
                  const unitType = useEUI ? 'eui' : 'consumption';
                  updateTotals(type, unitType);
                  updateTotalColors(type);
                }};

                // Initial total update
                updateTotals(currentChart, 'consumption');
                updateTotalColors(currentChart);

            }});
            </script>
            </html>
            """
        )

