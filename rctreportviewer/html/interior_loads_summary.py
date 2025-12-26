import math
from rctreportviewer.html.interior_lighting_details import write_interior_lighting_details


def write_interior_loads_summary(file, rct_detailed_report):
    baseline = rct_detailed_report.baseline_model_summary
    proposed = rct_detailed_report.proposed_model_summary

    file.write("""
<section class="mb-4">
  <div class="card shadow-sm">

    <!-- CLICKABLE HEADER -->
    <div class="card-header bg-light d-flex align-items-center"
         role="button"
         data-bs-toggle="collapse"
         data-bs-target="#collapse-internal-loads-summary"
         aria-expanded="false"
         style="cursor: pointer;">
      <span class="fw-semibold">Internal Loads Summary</span>
    </div>

    <div id="collapse-internal-loads-summary" class="collapse">
      <div class="card-body">

        <h5 class="mb-3">Space Type Summary</h5>
        <div class="table-responsive">
          <table class="table table-sm table-bordered align-middle text-center">
            <thead class="table-light">
              <tr>
                <th rowspan="2">Space Type</th>
                <th rowspan="2">Area (ft²)</th>
                <th colspan="4">Baseline</th>
                <th colspan="3">Proposed</th>
              </tr>
              <tr>
                <th>Occ. Density<br>(ft²/person)</th>
                <th>Equip. Power<br>(W/ft²)</th>
                <th>Allowed LPD<br>(W/ft²)</th>
                <th>LPD<br>(W/ft²)</th>
                <th>Occ. Density<br>(ft²/person)</th>
                <th>Equip. Power<br>(W/ft²)</th>
                <th>LPD<br>(W/ft²)</th>
              </tr>
            </thead>
            <tbody class="small">
""")

    for space_type, area in baseline["total_floor_area_by_space_type"].items():
        occupants_b = baseline["total_occupants_by_space_type"].get(space_type, 0)
        occupants_p = proposed["total_occupants_by_space_type"].get(space_type, 0)

        occ_density_b = area / (occupants_b or math.inf)
        occ_density_p = area / (occupants_p or math.inf)

        eqp_density_b = baseline["total_miscellaneous_equipment_power_by_space_type"].get(space_type, 0) / (area or math.inf)
        eqp_density_p = proposed["total_miscellaneous_equipment_power_by_space_type"].get(space_type, 0) / (area or math.inf)

        lpd_allowed_b = (
            rct_detailed_report.baseline_lighting_power_allowance_by_space_type
            .get(space_type, 0)
        )
        lpd_b = baseline["total_lighting_power_by_space_type"].get(space_type, 0) / (area or math.inf)
        lpd_p = proposed["total_lighting_power_by_space_type"].get(space_type, 0) / (area or math.inf)

        file.write(f"""
<tr>
  <td>{space_type.replace("_", " ").title()}</td>
  <td>{round(area):,}</td>
  <td>{round(occ_density_b)}</td>
  <td>{round(eqp_density_b, 2)}</td>
  <td>{round(lpd_allowed_b, 2)}</td>
  <td>{round(lpd_b, 2)}</td>
  <td>{round(occ_density_p)}</td>
  <td>{round(eqp_density_p, 2)}</td>
  <td>{round(lpd_p, 2)}</td>
</tr>
""")

    file.write(f"""
<tr class="fw-bold border-top">
  <td>Total</td>
  <td>{round(baseline["total_floor_area"]):,}</td>
  <td>{round(baseline["total_floor_area"] / baseline["total_occupants"], 2)}</td>
  <td>{round(baseline["total_equipment_power"] / baseline["total_floor_area"], 2)}</td>
  <td>{round(rct_detailed_report.baseline_total_lighting_power_allowance / baseline["total_floor_area"], 2)}</td>
  <td>{round(baseline["total_lighting_power"] / baseline["total_floor_area"], 2)}</td>
  <td>{round(proposed["total_floor_area"] / proposed["total_occupants"], 2)}</td>
  <td>{round(proposed["total_equipment_power"] / proposed["total_floor_area"], 2)}</td>
  <td>{round(proposed["total_lighting_power"] / proposed["total_floor_area"], 2)}</td>
</tr>
            </tbody>
          </table>
        </div>

        <h5 class="mt-4 mb-3">Schedule Summary</h5>
        <div class="table-responsive">
          <table class="table table-sm table-bordered align-middle text-center">
            <thead class="table-light">
              <tr>
                <th rowspan="2">Schedule</th>
                <th colspan="5">Baseline</th>
                <th colspan="5">Proposed</th>
              </tr>
              <tr>
                <th>EFLH</th>
                <th>Floor Area (ft²)</th>
                <th>% Lighting</th>
                <th>% Equipment</th>
                <th>Peak Gain (kBtu/hr)</th>
                <th>EFLH</th>
                <th>Floor Area (ft²)</th>
                <th>% Lighting</th>
                <th>% Equipment</th>
                <th>Peak Gain (kBtu/hr)</th>
              </tr>
            </thead>
            <tbody class="small">
""")

    for sched_id, base_sched in baseline["schedule_summaries"].items():
        prop_sched = proposed["schedule_summaries"].get(sched_id, {})
        file.write(f"""
<tr>
  <td>{sched_id}</td>
  <td>{round(base_sched.get("EFLH", 0)):,}</td>
  <td>{round(base_sched.get("associated_floor_area", 0)):,}</td>
  <td>{round(base_sched.get("percent_total_lighting_power", 0), 1)}</td>
  <td>{round(base_sched.get("percent_total_equipment_power", 0), 1)}</td>
  <td>{round(base_sched.get("associated_peak_internal_gain", 0), 1)}</td>
  <td>{round(prop_sched.get("EFLH", 0)):,}</td>
  <td>{round(prop_sched.get("associated_floor_area", 0)):,}</td>
  <td>{round(prop_sched.get("percent_total_lighting_power", 0), 1)}</td>
  <td>{round(prop_sched.get("percent_total_equipment_power", 0), 1)}</td>
  <td>{round(prop_sched.get("associated_peak_internal_gain", 0), 1)}</td>
</tr>
""")

    file.write("""
            </tbody>
          </table>
        </div>

        <p class="small text-muted mt-2">
          * Peak Internal Gain occurs when schedule fraction equals 1.0
        </p>
""")

    # === INTERIOR LIGHTING DETAILS (EMBEDDED) ===
    write_interior_lighting_details(file, rct_detailed_report)

    file.write("""
      </div>
    </div>
  </div>
</section>
""")
