def write_compliance_calculations(file, rct_detailed_report):
    b = rct_detailed_report.baseline_model_summary
    p = rct_detailed_report.proposed_model_summary
    output = rct_detailed_report.rpd_data.get("output", {})

    baseline_compliance = b.get("compliance_calcs_by_parameter", {})
    proposed_compliance = p.get("compliance_calcs_by_parameter", {})

    file.write(
        """
<section class="mb-4">
  <div class="card shadow-sm">

    <!-- CLICKABLE HEADER -->
    <div class="card-header bg-light d-flex align-items-center"
         role="button"
         data-bs-toggle="collapse"
         data-bs-target="#collapse-compliance-calcs"
         aria-expanded="false"
         style="cursor: pointer;">
      <span class="fw-semibold">Compliance Calculations</span>
    </div>

    <div id="collapse-compliance-calcs" class="collapse">
      <div class="card-body">

        <h5 class="mb-3">Energy by Source</h5>
        <div class="table-responsive">
          <table class="table table-sm align-middle" id="energySourceTable">
            <thead class="table-light text-center">
              <tr>
                <th>Energy Source</th>
                <th>Baseline Unregulated (MMBtu)</th>
                <th>Baseline Regulated (MMBtu)</th>
                <th>Total Baseline (MMBtu)</th>
                <th>Total Proposed (MMBtu)</th>
                <th>Source–Site Ratio</th>
                <th>GHG Factor (tCO₂/MMBtu)</th>
              </tr>
            </thead>
            <tbody class="small text-center">
"""
    )

    for fuel, proposed_energy in p.get("energy_by_fuel_type", {}).items():
        baseline_energy = b.get("energy_by_fuel_type", {}).get(fuel, 0)
        bbuec = baseline_compliance.get("bbuec", {}).get(fuel, 0)
        bbrec = baseline_compliance.get("bbrec", {}).get(fuel, 0)

        if fuel == "ELECTRICITY":
            label, ratio, ghg = "Electricity", 2.80, 0.037
        elif fuel == "NATURAL_GAS":
            label, ratio, ghg = "Natural Gas", 1.05, 0.053
        else:
            label, ratio, ghg = fuel, 0.0, 0.0

        file.write(
            f"""
              <tr>
                <td>{label}</td>
                <td class="baselineUnregulatedEnergy">{round(bbuec,1):,}</td>
                <td class="baselineRegulatedEnergy">{round(bbrec,1):,}</td>
                <td class="baselineEnergyUse">{round(baseline_energy,1):,}</td>
                <td class="proposedEnergyUse">{round(proposed_energy,1):,}</td>
                <td>
                  <input type="number" class="siteSourceRatio form-control form-control-sm text-center"
                         value="{ratio}" step="0.01">
                </td>
                <td>
                  <input type="number" class="ghgEmissionFactor form-control form-control-sm text-center"
                         value="{ghg}" step="0.001">
                </td>
              </tr>
"""
        )

    file.write(
        """
            </tbody>
          </table>
        </div>

        <h5 class="mt-4 mb-3">Performance Metrics</h5>
        <div class="table-responsive">
          <table class="table table-sm align-middle" id="complianceCalcsTable">
            <thead class="table-light text-center">
              <tr>
                <th>Parameter</th>
                <th>Symbol</th>
                <th>Cost ($)</th>
                <th>Site Energy (MMBtu)</th>
                <th>Source Energy (MMBtu)</th>
                <th>GHG (tCO₂e)</th>
              </tr>
            </thead>
            <tbody class="small text-center">
"""
    )

    def row(label, symbol, cost_id, site_id, source_id, ghg_id, cost, site):
        return f"""
<tr>
  <td class="text-start">{label}</td>
  <td><strong>{symbol}</strong></td>
  <td id="{cost_id}">{cost}</td>
  <td id="{site_id}">{site}</td>
  <td id="{source_id}">-</td>
  <td id="{ghg_id}">-</td>
</tr>
"""

    file.write(row(
        "Proposed building performance before site-generated renewable energy",
        "PBP<sub>nre</sub>",
        "pbp_nre_cost", "pbp_nre_site_energy",
        "pbp_nre_source_energy", "pbp_nre_ghg",
        f"${round(output.get('total_proposed_building_energy_cost_excluding_renewable_energy',0)):,}",
        round(proposed_compliance.get("pbp_nre",{}).get("site_energy",0))
    ))

    file.write("""
<tr>
  <td class="text-start">Proposed design on-site renewable savings</td>
  <td>-</td>
  <td id="proposed_cost_savings">$0</td>
  <td id="proposed_site_energy_savings">0</td>
  <td id="proposed_source_energy_savings">0</td>
  <td id="proposed_ghg_savings">0</td>
</tr>
""")

    file.write(row(
        "Proposed building performance including on-site renewable energy",
        "PBP",
        "pbp_cost", "pbp_site_energy",
        "pbp_source_energy", "pbp_ghg",
        f"${round(output.get('total_proposed_building_energy_cost_including_renewable_energy',0)):,}",
        "-"
    ))

    file.write(row(
        "Baseline building unregulated energy",
        "BBUEC",
        "bbuec_cost", "bbuec_site_energy",
        "bbuec_source_energy", "bbuec_ghg",
        f"${round(output.get('baseline_building_unregulated_energy_cost',0)):,}",
        round(baseline_compliance.get("bbuec",{}).get("site_energy",0))
    ))

    file.write(row(
        "Baseline building regulated energy",
        "BBREC",
        "bbrec_cost", "bbrec_site_energy",
        "bbrec_source_energy", "bbrec_ghg",
        f"${round(output.get('baseline_building_regulated_energy_cost',0)):,}",
        round(baseline_compliance.get("bbrec",{}).get("site_energy",0))
    ))

    file.write(row(
        "Baseline building performance",
        "BBP",
        "bbp_cost", "bbp_site_energy",
        "bbp_source_energy", "bbp_ghg",
        f"${round(output.get('baseline_building_performance_energy_cost',0)):,}",
        round(baseline_compliance.get("bbp",{}).get("site_energy",0))
    ))

    file.write(f"""
<tr>
  <td class="text-start">Building Performance Factor</td>
  <td><strong>BPF</strong></td>
  <td id="bpf_energy_cost">{round(output.get("total_area_weighted_building_performance_factor",0),2)}</td>
  <td id="bpf_site_energy">{round(rct_detailed_report.bpfs_by_metric.get("Site Energy",0),2)}</td>
  <td id="bpf_source_energy">{round(rct_detailed_report.bpfs_by_metric.get("Source Energy",0),2)}</td>
  <td id="bpf_ghg_emissions">{round(rct_detailed_report.bpfs_by_metric.get("GHG Emissions",0),2)}</td>
</tr>
<tr>
  <td class="text-start">Performance Index Target</td>
  <td><strong>PCI<sub>t</sub></strong></td>
  <td id="pcit_cost">{round(output.get("performance_cost_index_target",0),2)}</td>
  <td id="pcit_site_energy">-</td>
  <td id="pcit_source_energy">-</td>
  <td id="pcit_ghg_emissions">-</td>
</tr>
<tr>
  <td class="text-start">Performance Index without on-site renewable energy</td>
  <td><strong>PCI<sub>nre</sub></strong></td>
  <td id="pci_nre_cost">{round(output.get("performance_cost_index_target",0),2)}</td>
  <td id="pci_nre_site_energy">-</td>
  <td id="pci_nre_source_energy">-</td>
  <td id="pci_nre_ghg">-</td>
</tr>
<tr>
  <td class="text-start">Performance Index adjusted</td>
  <td><strong>PCI</strong></td>
  <td id="pci_adjusted_cost">{round(output.get("performance_cost_index",0),2)}</td>
  <td id="pci_adjusted_site_energy">-</td>
  <td id="pci_adjusted_source_energy">-</td>
  <td id="pci_adjusted_ghg">-</td>
</tr>
<tr class="fw-bold">
  <td class="text-start">% improvement excluding renewables</td>
  <td>-</td>
  <td id="cost_savings_nre">-</td>
  <td id="site_savings_nre">-</td>
  <td id="source_savings_nre">-</td>
  <td id="ghg_savings_nre">-</td>
</tr>
<tr class="fw-bold">
  <td class="text-start">% improvement including renewables</td>
  <td>-</td>
  <td id="cost_savings">-</td>
  <td id="site_savings">-</td>
  <td id="source_savings">-</td>
  <td id="ghg_savings">-</td>
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
