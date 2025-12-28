from rctreportviewer.constants import efficiency_display_map


def write_swh_summary(file, rct_detailed_report):
    proposed = rct_detailed_report.proposed_model_summary.get(
        "water_heater_summary", {}
    )
    baseline = rct_detailed_report.baseline_model_summary.get(
        "water_heater_summary", {}
    )

    def format_efficiencies(eff_list):
        if not eff_list:
            return "-"
        return "; ".join(
            f"{value:.2f} {efficiency_display_map.get(metric, metric.replace('_', ' ').title())}"
            for metric, value in eff_list
        )

    all_ids = sorted(set(proposed) | set(baseline))

    file.write(
        """
<section class="mb-4">
  <div class="card shadow-sm">

    <!-- CLICKABLE HEADER -->
    <div class="card-header bg-light d-flex align-items-center sticky-top"
         role="button"
         data-bs-toggle="collapse"
         data-bs-target="#collapse-swh-summary"
         aria-expanded="false"
         style="cursor: pointer;">
      <span class="fw-semibold">Service Water Heating Summary</span>
    </div>

    <div id="collapse-swh-summary" class="collapse">
      <div class="card-body">
        <div class="table-responsive">
          <table class="table table-sm table-borderless align-middle">
            <thead class="table-light text-center">
              <tr class="border-top">
                <th class="border-start"></th>
                <th colspan="3" class="border-start">Proposed Water Heaters</th>
                <th colspan="3" class="border-start border-end">Baseline Water Heaters</th>
              </tr>
              <tr class="border-top border-bottom">
                <th class="border-start">Water Heater</th>
                <th class="border-start">Area Type</th>
                <th class="border-start">Fuel</th>
                <th class="border-start">Efficiency</th>
                <th class="border-start">Area Type</th>
                <th class="border-start">Fuel</th>
                <th class="border-start border-end">Efficiency</th>
              </tr>
            </thead>
            <tbody class="text-center small">
"""
    )

    for wh_id in all_ids:
        p = proposed.get(wh_id, {})
        b = baseline.get(wh_id, {})

        file.write(
            f"""
<tr class="border-bottom">
  <td class="border-start fw-semibold">{wh_id}</td>
  <td class="border-start">{", ".join(p.get("area_types", ["-"])).replace("_", " ").title()}</td>
  <td class="border-start">{p.get("fuel_type", "-").replace("_", " ").title()}</td>
  <td class="border-start">{format_efficiencies(p.get("efficiencies"))}</td>
  <td class="border-start">{", ".join(b.get("area_types", ["-"])).replace("_", " ").title()}</td>
  <td class="border-start">{b.get("fuel_type", "-").replace("_", " ").title()}</td>
  <td class="border-start border-end">{format_efficiencies(b.get("efficiencies"))}</td>
</tr>
"""
        )

    file.write(
        """
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </div>
</section>
"""
    )
