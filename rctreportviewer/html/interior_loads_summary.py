import math


def write_interior_loads_summary(file, rct_detailed_report):
    file.write(
        f"""
        <div class="mb-3 me-4">
            <button class="btn btn-info collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapse-internal-loads-summary" aria-expanded="false">
                Internal Loads Summary
            </button>

            <div id="collapse-internal-loads-summary" class="accordion-collapse collapse">
                <div class="accordion-body">
                    <h3>Space Type Summary</h3>
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
                                <th style="border: 2px solid black;">Occupancy Density (ft<sup>2</sup>/person)</th>
                                <th style="border: 2px solid black;">Equipment Power Density (W/ft<sup>2</sup>)</th>
                                <th style="border: 2px solid black;">Lighting Power Density (W/ft<sup>2</sup>)</th>
                            </tr>
                        </thead>
                        <tbody style="border: 2px solid black;">
    """
    )

    for space_type in rct_detailed_report.baseline_model_summary[
        "total_floor_area_by_space_type"
    ]:
        area = rct_detailed_report.baseline_model_summary['total_floor_area_by_space_type'].get(space_type, 0)
        occupants_b = rct_detailed_report.baseline_model_summary['total_occupants_by_space_type'].get(space_type, 0)
        occ_density_b = area / (occupants_b or math.inf)
        eqp_density_b = rct_detailed_report.baseline_model_summary['total_miscellaneous_equipment_power_by_space_type'].get(space_type, 0) / (area or math.inf)
        ltg_density_allowed_b = rct_detailed_report.baseline_lighting_power_allowance_by_space_type.get(space_type, 0) / (area or math.inf)
        ltg_density_b = rct_detailed_report.baseline_model_summary['total_lighting_power_by_space_type'].get(space_type, 0) / (area or math.inf)
        occupants_p = rct_detailed_report.proposed_model_summary['total_occupants_by_space_type'].get(space_type, 0)
        occ_density_p = area / (occupants_p or math.inf)
        eqp_density_p = rct_detailed_report.proposed_model_summary['total_miscellaneous_equipment_power_by_space_type'].get(space_type, 0) /(area or math.inf)
        ltg_density_p = rct_detailed_report.proposed_model_summary['total_lighting_power_by_space_type'].get(space_type, 0) / (area or math.inf)
        file.write(
            f"""
                            <tr style="font-size: 12px;" class="lh-1 text-center">
                                <td>{space_type.replace("_", " ").title()}</td>
                                <td style="border-right: 2px solid black;">{round(area):,}</td>
                                <td>{round(occ_density_b)}</td>
                                <td>{round(eqp_density_b, 2)}</td>
                                <td>{round(ltg_density_allowed_b, 2)}</td>
                                <td style="border-right: 2px solid black;">{round(ltg_density_b, 2)}</td>
                                <td>{round(occ_density_p)}</td>
                                <td>{round(eqp_density_p, 2)}</td>
                                <td>{round(ltg_density_p, 2)}</td>
                            </tr>
            """
        )

    file.write(
        f"""
                            <tr  style="font-size: 12px; border-top: 1px solid black;" class="lh-1 fw-bold text-center">
                                <td>Total</td>
                                <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['total_floor_area']):,}</td>
                                <td>{round(rct_detailed_report.baseline_model_summary['total_floor_area'] / rct_detailed_report.baseline_model_summary['total_occupants'], 2)}</td>
                                <td>{round(rct_detailed_report.baseline_model_summary['total_equipment_power'] / rct_detailed_report.baseline_model_summary['total_floor_area'], 2)}</td>
                                <td>{round(rct_detailed_report.baseline_total_lighting_power_allowance / rct_detailed_report.baseline_model_summary['total_floor_area'], 2)}</td>
                                <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['total_lighting_power'] / rct_detailed_report.baseline_model_summary['total_floor_area'], 2)}</td>
                                <td>{round(rct_detailed_report.proposed_model_summary['total_floor_area'] / rct_detailed_report.proposed_model_summary['total_occupants'], 2)}</td>
                                <td>{round(rct_detailed_report.proposed_model_summary['total_equipment_power'] / rct_detailed_report.proposed_model_summary['total_floor_area'], 2)}</td>
                                <td>{round(rct_detailed_report.proposed_model_summary['total_lighting_power'] / rct_detailed_report.proposed_model_summary['total_floor_area'], 2)}</td>
                            </tr>
                        </tbody>
                    </table>
    """
    )

    # ----------------------- Schedule Summary Table -----------------------
    file.write(
        f"""
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
    """
    )

    baseline_schedule_summaries = rct_detailed_report.baseline_model_summary[
        "schedule_summaries"
    ]
    for schedule_id in baseline_schedule_summaries.keys():
        baseline_schedule_summary = rct_detailed_report.baseline_model_summary[
            "schedule_summaries"
        ].get(schedule_id, {})
        proposed_schedule_summary = rct_detailed_report.proposed_model_summary[
            "schedule_summaries"
        ].get(schedule_id, {})
        file.write(
            f"""
                                <tr style="font-size: 12px;" class="lh-1 text-center">
                                    <td style="border-right: 2px solid black;">{schedule_id}</td>
                                    <td>{round(baseline_schedule_summary.get("EFLH", 0)):,}</td>
                                    <td>{round(baseline_schedule_summary.get("associated_floor_area", 0.0)):,}</td>
                                    <td>{round(baseline_schedule_summary.get("percent_total_lighting_power", 0.0), 1):,}</td>
                                    <td>{round(baseline_schedule_summary.get("percent_total_equipment_power", 0.0), 1):,}</td>
                                    <td style="border-right: 2px solid black;">{round(baseline_schedule_summary.get("associated_peak_internal_gain", 0.0), 1):,}</td>
                                    <td>{round(proposed_schedule_summary.get("EFLH", 0)):,}</td>
                                    <td>{round(proposed_schedule_summary.get("associated_floor_area", 0.0)):,}</td>
                                    <td>{round(proposed_schedule_summary.get("percent_total_lighting_power", 0.0), 1):,}</td>
                                    <td>{round(proposed_schedule_summary.get("percent_total_equipment_power", 0.0), 1):,}</td>
                                    <td>{round(proposed_schedule_summary.get("associated_peak_internal_gain", 0.0), 1):,}</td>
                                </tr>
            """
        )
    file.write(
        f"""

                            </tbody>
                        </table>
                        <p style="font-size: 0.75rem;" class="ms-2">*Peak Internal Gain = Internal Gain when hourly fractional value is 1 or 100%</p>
                    </div>
                </div>
            </div>
        """
    )
