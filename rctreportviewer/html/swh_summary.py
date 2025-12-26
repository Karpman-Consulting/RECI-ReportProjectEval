from rctreportviewer.constants import efficiency_display_map


def write_swh_summary(file, rct_detailed_report):
    proposed = rct_detailed_report.proposed_model_summary.get("water_heater_summary", {})
    baseline = rct_detailed_report.baseline_model_summary.get("water_heater_summary", {})

    def format_efficiencies(eff_list):
        if not eff_list:
            return "-"
        return "; ".join(
            f"{value:.2f} {efficiency_display_map.get(metric, metric.replace('_', ' ').title())}"
            for metric, value in eff_list
        )

    all_ids = sorted(set(proposed) | set(baseline))

    file.write("""
<section class="mb-4">
  <div class="card shadow-sm">
    <div class="card-header bg-light">
      <button class="btn btn-info"
              type="button"
              data-bs-toggle="collapse"
              data-bs-target="#collapse-swh-summary">
        Service Water Heating Summary
      </button>
    </div>

    <div id="collapse-swh-summary" class="collapse">
      <div class="card-body">
        <div class="table-responsive">
          <table class="table table-sm table-borderless align-middle">
            <thead class="table-light text-center">
              <tr>
                <th></th>
                <th colspan="3">Proposed Water Heater</th>
                <th colspan="3">Baseline Water Heater</th>
              </tr>
              <tr>
                <th>Water Heater</th>
                <th>Area Type</th>
                <th>Fuel</th>
                <th>Efficiency</th>
                <th>Area Type</th>
                <th>Fuel</th>
                <th>Efficiency</th>
              </tr>
            </thead>
            <tbody class="text-center small">
""")

    for wh_id in all_ids:
        p = proposed.get(wh_id, {})
        b = baseline.get(wh_id, {})

        file.write(f"""
<tr>
  <td class="fw-semibold">{wh_id}</td>
  <td>{", ".join(p.get("area_types", ["-"])).replace("_", " ").title()}</td>
  <td>{p.get("fuel_type", "-").replace("_", " ").title()}</td>
  <td>{format_efficiencies(p.get("efficiencies"))}</td>
  <td>{", ".join(b.get("area_types", ["-"])).replace("_", " ").title()}</td>
  <td>{b.get("fuel_type", "-").replace("_", " ").title()}</td>
  <td>{format_efficiencies(b.get("efficiencies"))}</td>
</tr>
""")

    file.write("""
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </div>
</section>
""")
