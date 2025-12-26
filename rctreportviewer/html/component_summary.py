def write_component_summary(file, rct_detailed_report):
    tooltip_html = ""

    b = rct_detailed_report.baseline_model_summary
    p = rct_detailed_report.proposed_model_summary

    file.write(
        f"""
<section class="mb-4">
  <div class="card shadow-sm">

    <!-- CLICKABLE HEADER -->
    <div class="card-header bg-light d-flex align-items-center"
         role="button"
         data-bs-toggle="collapse"
         data-bs-target="#collapse-model-component-summary"
         aria-expanded="false"
         style="cursor: pointer;">
      <span class="fw-semibold">Model Component Summary</span>
    </div>

    <div id="collapse-model-component-summary" class="collapse">
      <div class="card-body">

        <div class="table-responsive">
          <table class="table table-sm align-middle">

            <!-- COLUMN SIZING -->
            <colgroup>
              <!-- Label -->
              <col style="width:1%; white-space:nowrap;">
              <!-- Baseline -->
              <col style="width:1%; white-space:nowrap;">
              <!-- Proposed -->
              <col style="width:1%; white-space:nowrap;">
              <!-- FILLER -->
              <col style="width:auto;">
            </colgroup>

            <thead class="border-bottom">
              <tr>
                <th class="ps-3"></th>
                <th class="text-center px-3">Baseline</th>
                <th class="text-center px-3">Proposed</th>
                <th></th>
              </tr>
            </thead>

            <tbody class="small">

              <tr>
                <td class="text-start ps-3 pe-4 text-nowrap">Building Qty</td>
                <td class="text-center px-3 text-nowrap">{b["building_count"]}</td>
                <td class="text-center px-3 text-nowrap">{p["building_count"]}</td>
                <td></td>
              </tr>

              <tr>
                <td class="text-start ps-3 pe-4 text-nowrap">Total Floor Area</td>
                <td class="text-center px-3 text-nowrap">{round(b["total_floor_area"]):,}</td>
                <td class="text-center px-3 text-nowrap">{round(p["total_floor_area"]):,}</td>
                <td></td>
              </tr>

              <tr>
                <td class="text-start ps-3 pe-4 text-nowrap">Building Area Qty</td>
                <td class="text-center px-3 text-nowrap">{b["building_segment_count"]}</td>
                <td class="text-center px-3 text-nowrap">{p["building_segment_count"]}</td>
                <td></td>
              </tr>

              <tr>
                <td class="text-start ps-3 pe-4 text-nowrap">System Qty</td>
                <td class="text-center px-3 text-nowrap">
                  <span class="d-inline-block"
                        data-bs-toggle="tooltip"
                        data-bs-html="true"
                        title="{tooltip_html}"
                        style="text-decoration: underline dotted; cursor: help;">
                    {b["system_count"]}
                  </span>
                </td>
                <td class="text-center px-3 text-nowrap">{p["system_count"]}</td>
                <td></td>
              </tr>

              <tr>
                <td class="text-start ps-3 pe-4 text-nowrap">Zone Qty</td>
                <td class="text-center px-3 text-nowrap">{b["zone_count"]}</td>
                <td class="text-center px-3 text-nowrap">{p["zone_count"]}</td>
                <td></td>
              </tr>

              <tr>
                <td class="text-start ps-3 pe-4 text-nowrap">Space Qty</td>
                <td class="text-center px-3 text-nowrap">{b["space_count"]}</td>
                <td class="text-center px-3 text-nowrap">{p["space_count"]}</td>
                <td></td>
              </tr>

              <!-- Fluid Loops: allow wrapping -->
              <tr>
                <td class="text-start ps-3 pe-4 text-nowrap">Fluid Loops</td>
                <td class="text-center px-3">{", ".join(s.title() for s in b["fluid_loop_types"])}</td>
                <td class="text-center px-3">{", ".join(s.title() for s in p["fluid_loop_types"])}</td>
                <td></td>
              </tr>

              <tr>
                <td class="text-start ps-3 pe-4 text-nowrap">Pump Qty</td>
                <td class="text-center px-3 text-nowrap">{b["pump_count"]}</td>
                <td class="text-center px-3 text-nowrap">{p["pump_count"]}</td>
                <td></td>
              </tr>

              <tr>
                <td class="text-start ps-3 pe-4 text-nowrap">Boiler Qty</td>
                <td class="text-center px-3 text-nowrap">{b["boiler_count"]}</td>
                <td class="text-center px-3 text-nowrap">{p["boiler_count"]}</td>
                <td></td>
              </tr>

              <tr>
                <td class="text-start ps-3 pe-4 text-nowrap">Chiller Qty</td>
                <td class="text-center px-3 text-nowrap">{b["chiller_count"]}</td>
                <td class="text-center px-3 text-nowrap">{p["chiller_count"]}</td>
                <td></td>
              </tr>

              <tr>
                <td class="text-start ps-3 pe-4 text-nowrap">Heat Rejection Qty</td>
                <td class="text-center px-3 text-nowrap">{b["heat_rejection_count"]}</td>
                <td class="text-center px-3 text-nowrap">{p["heat_rejection_count"]}</td>
                <td></td>
              </tr>

              <tr>
                <td class="text-start ps-3 pe-4 text-nowrap">SWH Heater Qty</td>
                <td class="text-center px-3 text-nowrap">{b["water_heater_count"]}</td>
                <td class="text-center px-3 text-nowrap">{p["water_heater_count"]}</td>
                <td></td>
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
