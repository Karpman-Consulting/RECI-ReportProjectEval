def write_component_summary(file, rct_detailed_report):
    # ----------------------- HVAC System Type Summary Tooltip -----------------------
    tooltip_lines = []
    # total_qty = 0

    # for system_type, systems in rct_detailed_report.hvac_system_types_b.items():
    #     qty = len(systems)
    #     total_qty += qty
    #     tooltip_lines.append(
    #         f"<div class='text-start'><b>{system_type}</b>: {qty}</div>"
    #     )

    tooltip_html = "".join(tooltip_lines)

    b = rct_detailed_report.baseline_model_summary
    p = rct_detailed_report.proposed_model_summary

    file.write(
        f"""
<section class="mb-4">
    <div class="card shadow-sm">
        <div class="card-header bg-light">
            <button class="btn btn-info"
                    type="button"
                    data-bs-toggle="collapse"
                    data-bs-target="#collapse-model-component-summary"
                    aria-expanded="false">
                Model Component Summary
            </button>
        </div>

        <div id="collapse-model-component-summary" class="collapse">
            <div class="card-body">

                <div class="table-responsive">
                    <table class="table table-sm align-middle">
                        <thead class="border-bottom">
                            <tr>
                                <th></th>
                                <th class="text-center">Baseline</th>
                                <th class="text-center">Proposed</th>
                            </tr>
                        </thead>
                        <tbody class="small">

                            <tr>
                                <td class="text-end">Building Qty</td>
                                <td class="text-center">{b["building_count"]}</td>
                                <td class="text-center">{p["building_count"]}</td>
                            </tr>

                            <tr>
                                <td class="text-end">Total Floor Area</td>
                                <td class="text-center">{round(b["total_floor_area"]):,}</td>
                                <td class="text-center">{round(p["total_floor_area"]):,}</td>
                            </tr>

                            <tr>
                                <td class="text-end">Building Area Qty</td>
                                <td class="text-center">{b["building_segment_count"]}</td>
                                <td class="text-center">{p["building_segment_count"]}</td>
                            </tr>

                            <tr>
                                <td class="text-end">System Qty</td>
                                <td class="text-center">
                                    <span class="d-inline-block"
                                          data-bs-toggle="tooltip"
                                          data-bs-html="true"
                                          title="{tooltip_html}"
                                          style="text-decoration: underline dotted; cursor: help;">
                                        {b["system_count"]}
                                    </span>
                                </td>
                                <td class="text-center">{p["system_count"]}</td>
                            </tr>

                            <tr>
                                <td class="text-end">Zone Qty</td>
                                <td class="text-center">{b["zone_count"]}</td>
                                <td class="text-center">{p["zone_count"]}</td>
                            </tr>

                            <tr>
                                <td class="text-end">Space Qty</td>
                                <td class="text-center">{b["space_count"]}</td>
                                <td class="text-center">{p["space_count"]}</td>
                            </tr>

                            <tr>
                                <td class="text-end">Fluid Loops</td>
                                <td class="text-center">{", ".join(s.title() for s in b["fluid_loop_types"])}</td>
                                <td class="text-center">{", ".join(s.title() for s in p["fluid_loop_types"])}</td>
                            </tr>

                            <tr>
                                <td class="text-end">Pump Qty</td>
                                <td class="text-center">{b["pump_count"]}</td>
                                <td class="text-center">{p["pump_count"]}</td>
                            </tr>

                            <tr>
                                <td class="text-end">Boiler Qty</td>
                                <td class="text-center">{b["boiler_count"]}</td>
                                <td class="text-center">{p["boiler_count"]}</td>
                            </tr>

                            <tr>
                                <td class="text-end">Chiller Qty</td>
                                <td class="text-center">{b["chiller_count"]}</td>
                                <td class="text-center">{p["chiller_count"]}</td>
                            </tr>

                            <tr>
                                <td class="text-end">Heat Rejection Qty</td>
                                <td class="text-center">{b["heat_rejection_count"]}</td>
                                <td class="text-center">{p["heat_rejection_count"]}</td>
                            </tr>

                            <tr>
                                <td class="text-end">SWH Heater Qty</td>
                                <td class="text-center">{b["water_heater_count"]}</td>
                                <td class="text-center">{p["water_heater_count"]}</td>
                            </tr>

                        </tbody>
                    </table>
                </div>

            </div>
        </div>
    </div>
</section>
"""
    )
