

def write_envelope_summary(
    file, rct_detailed_report
):
    """
    Write the envelope summary section of the RCT detailed report.

    :param file: The file object to write the HTML content to.
    :param rct_detailed_report: The RCT detailed report object containing model summaries.
    """

    # Write the envelope summary section
    file.write("""
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

    for building_segment_id in rct_detailed_report.baseline_model_summary[
        "total_floor_area_by_building_segment"
    ]:
        if (
            building_segment_id
            in rct_detailed_report.baseline_model_summary[
                "total_roof_area_by_building_segment"
            ]
        ):
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

        if (
            building_segment_id
            in rct_detailed_report.baseline_model_summary[
                "total_wall_area_by_building_segment"
            ]
        ):
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

    file.write(
        """          </tbody>
                        </table>
                        <p style="font-size: 0.75rem;" class="ms-2">*U-Factors represent area-weighted averages for the corresponding Building Area & Surface Type</p>
                    </div>
                </div>
            </div>
        """)
