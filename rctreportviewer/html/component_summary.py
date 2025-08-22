

def write_component_summary(file, rct_detailed_report):
    # ----------------------- HVAC System Type Summary Tooltip -----------------------
    tooltip_lines = []
    total_qty = 0

    for system_type, systems in rct_detailed_report.hvac_system_types_b.items():
        qty = len(systems)
        total_qty += qty
        tooltip_lines.append(
            f"<div class='text-start'><b>{system_type}</b>: {qty}</div>"
        )

    tooltip_html = "".join(tooltip_lines)

    file.write(f"""
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
                            <tr style="font-size: 12px;" class="lh-1">
                                <td class="col-3 text-end">System Qty</td>
                                <td class="col-4 text-center">
                                    <span class="d-inline-block" data-bs-toggle="tooltip" data-bs-placement="top" data-bs-html="true" data-bs-title="{tooltip_html}" style="text-decoration: underline dotted; text-underline-offset: 3px; cursor: help;">
                                        {rct_detailed_report.baseline_model_summary["system_count"]}
                                    </span>
                                </td>
                                <td class="col-4 text-center">{rct_detailed_report.proposed_model_summary["system_count"]}</td>
                            </tr>
                            <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Zone Qty</td><td class="col-4 text-center">{rct_detailed_report.baseline_model_summary["zone_count"]}</td><td class="col-4 text-center">{rct_detailed_report.proposed_model_summary["zone_count"]}</td></tr>
                            <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Space Qty</td><td class="col-4 text-center">{rct_detailed_report.baseline_model_summary["space_count"]}</td><td class="col-4 text-center">{rct_detailed_report.proposed_model_summary["space_count"]}</td></tr>
                            <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Fluid Loops</td><td class="col-4 text-center">{", ".join(s.title() for s in rct_detailed_report.baseline_model_summary["fluid_loop_types"])}</td><td class="col-4 text-center">{", ".join(s.title() for s in rct_detailed_report.proposed_model_summary["fluid_loop_types"])}</td></tr>
                            <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Pump Qty</td><td class="col-4 text-center">{rct_detailed_report.baseline_model_summary["pump_count"]}</td><td class="col-4 text-center">{rct_detailed_report.proposed_model_summary["pump_count"]}</td></tr>
                            <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Boiler Qty</td><td class="col-4 text-center">{rct_detailed_report.baseline_model_summary["boiler_count"]}</td><td class="col-4 text-center">{rct_detailed_report.proposed_model_summary["boiler_count"]}</td></tr>
                            <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Chiller Qty</td><td class="col-4 text-center">{rct_detailed_report.baseline_model_summary["chiller_count"]}</td><td class="col-4 text-center">{rct_detailed_report.proposed_model_summary["chiller_count"]}</td></tr>
                            <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Heat Rejection Qty</td><td class="col-4 text-center">{rct_detailed_report.baseline_model_summary["heat_rejection_count"]}</td><td class="col-4 text-center">{rct_detailed_report.proposed_model_summary["heat_rejection_count"]}</td></tr>
                            <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">SWH Heater Qty</td><td class="col-4 text-center">{rct_detailed_report.baseline_model_summary["water_heater_count"]}</td><td class="col-4 text-center">{rct_detailed_report.proposed_model_summary["water_heater_count"]}</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    """)
