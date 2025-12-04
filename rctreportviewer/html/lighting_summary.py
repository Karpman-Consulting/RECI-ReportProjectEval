def write_interior_lighting_summary(file, rct_detailed_report):
    """
    Write the interior lighting summary section to the HTML file.
    """

    file.write(
        """
        <div class="mb-3 me-4">
            <button class="btn btn-info collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapse-interior-lighting-summary" aria-expanded="false">
                Interior Lighting Summary
            </button>

            <div id="collapse-interior-lighting-summary" class="accordion-collapse collapse">
                <div class="accordion-body">
                    <table class="table table-sm table-borderless mb-0">
                        <thead>
                            <tr class="text-center">
                                <th colspan="2"></th>
                                <th colspan="10" style="border: 2px solid black;">Lighting Power [Watt]</th>
                                <th colspan="10" style="border: 2px solid black;">Lighting Power Density [Watt/ft<sup>2</sup>]</th>
                                <th></th>
                            </tr>
                            <tr class="text-center">
                                <th colspan="2"></th>
                                <th colspan="5" style="border: 2px solid black;">Baseline Design</th>
                                <th colspan="5" style="border: 2px solid black;">Proposed Design</th>
                                <th colspan="5" style="border: 2px solid black;">Baseline Design</th>
                                <th colspan="5" style="border: 2px solid black;">Proposed Design</th>
                                <th></th>
                            </tr>
                            <tr class="text-end align-middle">
                                <th style="border: 2px solid black;">Interior Lighting ID</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Floor Area [ft<sup>2</sup>]</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Non-exempt General Lighting</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Additional Decorative Lighting</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Additional Retail Lighting</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Exempt Lighting</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Total</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Non-exempt General Lighting</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Additional Decorative Lighting</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Additional Retail Lighting</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Exempt Lighting</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Total</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Non-exempt General Lighting</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Additional Decorative Lighting</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Additional Retail Lighting</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Exempt Lighting</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Total</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Non-exempt General Lighting</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Additional Decorative Lighting</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Additional Retail Lighting</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Exempt Lighting</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">Total</th>
                                <th style="border: 2px solid black; writing-mode: vertical-lr;">% Savings of Proposed vs Baseline</th>
                            </tr>
                        </thead>
                        <tbody style="border: 2px solid black;">
    """
    )

    int_ltg_totals = {
        "baseline": {
            "floor_area": 0,
            "int_ltg_power_general": 0,
            "int_ltg_power_retail": 0,
            "int_ltg_power_decorative": 0,
            "int_ltg_power_exempt": 0,
            "int_ltg_power_total": 0,
        },
        "proposed": {
            "floor_area": 0,
            "int_ltg_power_general": 0,
            "int_ltg_power_retail": 0,
            "int_ltg_power_decorative": 0,
            "int_ltg_power_exempt": 0,
            "int_ltg_power_total": 0,
        },
    }

    for interior_lighting in set(
        list(rct_detailed_report.baseline_model_summary["int_ltg_summaries"].keys())
        + list(rct_detailed_report.proposed_model_summary["int_ltg_summaries"].keys())
    ):
        if (
            interior_lighting
            in rct_detailed_report.baseline_model_summary["int_ltg_summaries"]
            and interior_lighting
            in rct_detailed_report.proposed_model_summary["int_ltg_summaries"]
        ):
            baseline_lighting = rct_detailed_report.baseline_model_summary[
                "int_ltg_summaries"
            ][interior_lighting]
            proposed_lighting = rct_detailed_report.proposed_model_summary[
                "int_ltg_summaries"
            ][interior_lighting]
            int_ltg_totals["baseline"]["floor_area"] += baseline_lighting["floor_area"]
            int_ltg_totals["baseline"]["int_ltg_power_general"] += baseline_lighting[
                "int_ltg_power_general"
            ]
            int_ltg_totals["baseline"]["int_ltg_power_retail"] += baseline_lighting[
                "int_ltg_power_retail"
            ]
            int_ltg_totals["baseline"]["int_ltg_power_decorative"] += baseline_lighting[
                "int_ltg_power_decorative"
            ]
            int_ltg_totals["baseline"]["int_ltg_power_exempt"] += baseline_lighting[
                "int_ltg_power_exempt"
            ]
            int_ltg_totals["baseline"]["int_ltg_power_total"] += baseline_lighting[
                "int_ltg_power_total"
            ]
            int_ltg_totals["proposed"]["floor_area"] += proposed_lighting["floor_area"]
            int_ltg_totals["proposed"]["int_ltg_power_general"] += proposed_lighting[
                "int_ltg_power_general"
            ]
            int_ltg_totals["proposed"]["int_ltg_power_retail"] += proposed_lighting[
                "int_ltg_power_retail"
            ]
            int_ltg_totals["proposed"]["int_ltg_power_decorative"] += proposed_lighting[
                "int_ltg_power_decorative"
            ]
            int_ltg_totals["proposed"]["int_ltg_power_exempt"] += proposed_lighting[
                "int_ltg_power_exempt"
            ]
            int_ltg_totals["proposed"]["int_ltg_power_total"] += proposed_lighting[
                "int_ltg_power_total"
            ]
            file.write(
                f"""
                            <tr style="font-size: 12px;" class="lh-1 text-center">
                                <td>{interior_lighting}</td>
                                <td style="border-right: 2px solid black;">{round(baseline_lighting["floor_area"]):,}</td>
                                <td>{round(baseline_lighting["int_ltg_power_general"]):,}</td>
                                <td>{round(baseline_lighting["int_ltg_power_retail"]):,}</td>
                                <td>{round(baseline_lighting["int_ltg_power_decorative"]):,}</td>
                                <td>{round(baseline_lighting["int_ltg_power_exempt"]):,}</td>
                                <td style="border-right: 2px solid black;">{round(baseline_lighting["int_ltg_power_total"]):,}</td>
                                <td>{round(proposed_lighting["int_ltg_power_general"]):,}</td>
                                <td>{round(proposed_lighting["int_ltg_power_retail"]):,}</td>
                                <td>{round(proposed_lighting["int_ltg_power_decorative"]):,}</td>
                                <td>{round(proposed_lighting["int_ltg_power_exempt"]):,}</td>
                                <td style="border-right: 2px solid black;">{round(proposed_lighting["int_ltg_power_total"]):,}</td>
                                <td>{round(baseline_lighting["int_ltg_power_general"] / baseline_lighting["floor_area"], 2):,}</td>
                                <td>{round(baseline_lighting["int_ltg_power_retail"] / baseline_lighting["floor_area"], 2):,}</td>
                                <td>{round(baseline_lighting["int_ltg_power_decorative"] / baseline_lighting["floor_area"], 2):,}</td>
                                <td>{round(baseline_lighting["int_ltg_power_exempt"] / baseline_lighting["floor_area"], 2):,}</td>
                                <td style="border-right: 2px solid black;">{round(baseline_lighting["int_ltg_power_total"] / baseline_lighting["floor_area"], 2):,}</td>
                                <td>{round(proposed_lighting["int_ltg_power_general"] / proposed_lighting["floor_area"], 2):,}</td>
                                <td>{round(proposed_lighting["int_ltg_power_retail"] / proposed_lighting["floor_area"], 2):,}</td>
                                <td>{round(proposed_lighting["int_ltg_power_decorative"] / proposed_lighting["floor_area"], 2):,}</td>
                                <td>{round(proposed_lighting["int_ltg_power_exempt"] / proposed_lighting["floor_area"], 2):,}</td>
                                <td style="border-right: 2px solid black;">{round(proposed_lighting["int_ltg_power_total"] / proposed_lighting["floor_area"], 2):,}</td>
                                <td>{round((baseline_lighting["int_ltg_power_total"] - proposed_lighting["int_ltg_power_total"]) / baseline_lighting["int_ltg_power_total"] * 100, 1) if baseline_lighting["int_ltg_power_total"] > 0 else 0}%</td>
                            </tr>
                """
            )
        elif (
            interior_lighting
            in rct_detailed_report.baseline_model_summary["int_ltg_summaries"]
        ):
            baseline_lighting = rct_detailed_report.baseline_model_summary[
                "int_ltg_summaries"
            ][interior_lighting]
            int_ltg_totals["baseline"]["floor_area"] += baseline_lighting["floor_area"]
            int_ltg_totals["baseline"]["int_ltg_power_general"] += baseline_lighting[
                "int_ltg_power_general"
            ]
            int_ltg_totals["baseline"]["int_ltg_power_retail"] += baseline_lighting[
                "int_ltg_power_retail"
            ]
            int_ltg_totals["baseline"]["int_ltg_power_decorative"] += baseline_lighting[
                "int_ltg_power_decorative"
            ]
            int_ltg_totals["baseline"]["int_ltg_power_exempt"] += baseline_lighting[
                "int_ltg_power_exempt"
            ]
            int_ltg_totals["baseline"]["int_ltg_power_total"] += baseline_lighting[
                "int_ltg_power_total"
            ]
            file.write(
                f"""
                            <tr style="font-size: 12px;" class="lh-1 text-center">
                                <td>{interior_lighting}</td>
                                <td style="border-right: 2px solid black;">{round(baseline_lighting["floor_area"]):,}</td>
                                <td>{round(baseline_lighting["int_ltg_power_general"]):,}</td>
                                <td>{round(baseline_lighting["int_ltg_power_retail"]):,}</td>
                                <td>{round(baseline_lighting["int_ltg_power_decorative"]):,}</td>
                                <td>{round(baseline_lighting["int_ltg_power_exempt"]):,}</td>
                                <td style="border-right: 2px solid black;">{round(baseline_lighting["int_ltg_power_total"]):,}</td>
                                <td>0</td>
                                <td>0</td>
                                <td>0</td>
                                <td>0</td>
                                <td style="border-right: 2px solid black;">0</td>
                                <td>{round(baseline_lighting["int_ltg_power_general"] / baseline_lighting["floor_area"], 2):,}</td>
                                <td>{round(baseline_lighting["int_ltg_power_retail"] / baseline_lighting["floor_area"], 2):,}</td>
                                <td>{round(baseline_lighting["int_ltg_power_decorative"] / baseline_lighting["floor_area"], 2):,}</td>
                                <td>{round(baseline_lighting["int_ltg_power_exempt"] / baseline_lighting["floor_area"], 2):,}</td>
                                <td style="border-right: 2px solid black;">{round(baseline_lighting["int_ltg_power_total"] / baseline_lighting["floor_area"], 2):,}</td>
                                <td>0</td>
                                <td>0</td>
                                <td>0</td>
                                <td>0</td>
                                <td style="border-right: 2px solid black;">0</td>
                                <td>-</td>
                            </tr>
                """
            )

        elif (
            interior_lighting
            in rct_detailed_report.proposed_model_summary["int_ltg_summaries"]
        ):
            proposed_lighting = rct_detailed_report.proposed_model_summary[
                "int_ltg_summaries"
            ][interior_lighting]
            int_ltg_totals["proposed"]["floor_area"] += proposed_lighting["floor_area"]
            int_ltg_totals["proposed"]["int_ltg_power_general"] += proposed_lighting[
                "int_ltg_power_general"
            ]
            int_ltg_totals["proposed"]["int_ltg_power_retail"] += proposed_lighting[
                "int_ltg_power_retail"
            ]
            int_ltg_totals["proposed"]["int_ltg_power_decorative"] += proposed_lighting[
                "int_ltg_power_decorative"
            ]
            int_ltg_totals["proposed"]["int_ltg_power_exempt"] += proposed_lighting[
                "int_ltg_power_exempt"
            ]
            int_ltg_totals["proposed"]["int_ltg_power_total"] += proposed_lighting[
                "int_ltg_power_total"
            ]
            file.write(
                f"""
                            <tr style="font-size: 12px;" class="lh-1 text-center">
                                <td>{interior_lighting}</td>
                                <td style="border-right: 2px solid black;">{round(proposed_lighting["floor_area"]):,}</td>
                                <td>0</td>
                                <td>0</td>
                                <td>0</td>
                                <td>0</td>
                                <td style="border-right: 2px solid black;">0</td>
                                <td>{round(proposed_lighting["int_ltg_power_general"]):,}</td>
                                <td>{round(proposed_lighting["int_ltg_power_retail"]):,}</td>
                                <td>{round(proposed_lighting["int_ltg_power_decorative"]):,}</td>
                                <td>{round(proposed_lighting["int_ltg_power_exempt"]):,}</td>
                                <td style="border-right: 2px solid black;">{round(proposed_lighting["int_ltg_power_total"]):,}</td>
                                <td>0</td>
                                <td>0</td>
                                <td>0</td>
                                <td>0</td>
                                <td style="border-right: 2px solid black;">0</td>
                                <td>{round(proposed_lighting["int_ltg_power_general"] / proposed_lighting["floor_area"], 2):,}</td>
                                <td>{round(proposed_lighting["int_ltg_power_retail"] / proposed_lighting["floor_area"], 2):,}</td>
                                <td>{round(proposed_lighting["int_ltg_power_decorative"] / proposed_lighting["floor_area"], 2):,}</td>
                                <td>{round(proposed_lighting["int_ltg_power_exempt"] / proposed_lighting["floor_area"], 2):,}</td>
                                <td style="border-right: 2px solid black;">{round(proposed_lighting["int_ltg_power_total"] / proposed_lighting["floor_area"], 2):,}</td>
                                <td>-</td>
                            </tr>
                """
            )

    file.write(
        f"""          
                        <tr style="font-size: 12px; border-top: 1px solid black;" class="lh-1 fw-bold text-center">
                            <td>Total</td>
                            <td style="border-right: 2px solid black;">{round(int_ltg_totals['baseline']['floor_area']):,}</td>
                            <td>{round(int_ltg_totals['baseline']['int_ltg_power_general']):,}</td>
                            <td>{round(int_ltg_totals['baseline']['int_ltg_power_retail']):,}</td>
                            <td>{round(int_ltg_totals['baseline']['int_ltg_power_decorative']):,}</td>
                            <td>{round(int_ltg_totals['baseline']['int_ltg_power_exempt']):,}</td>
                            <td style="border-right: 2px solid black;">{round(int_ltg_totals['baseline']['int_ltg_power_total']):,}</td>
                            <td>{round(int_ltg_totals['proposed']['int_ltg_power_general']):,}</td>
                            <td>{round(int_ltg_totals['proposed']['int_ltg_power_retail']):,}</td>
                            <td>{round(int_ltg_totals['proposed']['int_ltg_power_decorative']):,}</td>
                            <td>{round(int_ltg_totals['proposed']['int_ltg_power_exempt']):,}</td>
                            <td style="border-right: 2px solid black;">{round(int_ltg_totals['proposed']['int_ltg_power_total']):,}</td>
                            <td>{round(int_ltg_totals['baseline']['int_ltg_power_general'] / int_ltg_totals['baseline']['floor_area'], 2):,}</td>
                            <td>{round(int_ltg_totals['baseline']['int_ltg_power_retail'] / int_ltg_totals['baseline']['floor_area'], 2):,}</td>
                            <td>{round(int_ltg_totals['baseline']['int_ltg_power_decorative'] / int_ltg_totals['baseline']['floor_area'], 2):,}</td>
                            <td>{round(int_ltg_totals['baseline']['int_ltg_power_exempt'] / int_ltg_totals['baseline']['floor_area'], 2):,}</td>
                            <td style="border-right: 2px solid black;">{round(int_ltg_totals['baseline']['int_ltg_power_total'] / int_ltg_totals['baseline']['floor_area'], 2):,}</td>
                            <td>{round(int_ltg_totals['proposed']['int_ltg_power_general'] / int_ltg_totals['proposed']['floor_area'], 2):,}</td>
                            <td>{round(int_ltg_totals['proposed']['int_ltg_power_retail'] / int_ltg_totals['proposed']['floor_area'], 2):,}</td>
                            <td>{round(int_ltg_totals['proposed']['int_ltg_power_decorative'] / int_ltg_totals['proposed']['floor_area'], 2):,}</td>
                            <td>{round(int_ltg_totals['proposed']['int_ltg_power_exempt'] / int_ltg_totals['proposed']['floor_area'], 2):,}</td>
                            <td style="border-right: 2px solid black;">{round(int_ltg_totals['proposed']['int_ltg_power_total'] / int_ltg_totals['proposed']['floor_area'], 2):,}</td>
                            <td>{round((int_ltg_totals['baseline']['int_ltg_power_total'] - int_ltg_totals['proposed']['int_ltg_power_total']) / int_ltg_totals['baseline']['int_ltg_power_total'] * 100, 1) if int_ltg_totals['baseline']['int_ltg_power_total'] > 0 else 0}%</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
"""
    )
