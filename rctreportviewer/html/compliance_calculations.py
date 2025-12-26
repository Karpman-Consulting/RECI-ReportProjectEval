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
        <div class="card-header bg-light">
            <button class="btn btn-info"
                    type="button"
                    data-bs-toggle="collapse"
                    data-bs-target="#collapse-compliance-calcs"
                    aria-expanded="false">
                Compliance Calculations
            </button>
        </div>

        <div id="collapse-compliance-calcs" class="collapse">
            <div class="card-body">

                <h5 class="mb-3">Energy by Source</h5>
                <div class="table-responsive">
                    <table class="table table-sm align-middle" id="energySourceTable">
                        <thead class="table-light border-bottom">
                            <tr class="text-center">
                                <th>Energy Source</th>
                                <th>Baseline Unregulated<br>(MMBtu)</th>
                                <th>Baseline Regulated<br>(MMBtu)</th>
                                <th>Total Baseline<br>(MMBtu)</th>
                                <th>Total Proposed<br>(MMBtu)</th>
                                <th>Source–Site Ratio</th>
                                <th>GHG Factor<br>(tCO₂/MMBtu)</th>
                            </tr>
                        </thead>
                        <tbody class="small">
"""
    )

    for energy_source, proposed_energy in p.get("energy_by_fuel_type", {}).items():
        baseline_energy = b.get("energy_by_fuel_type", {}).get(energy_source, 0)

        bbuec = baseline_compliance.get("bbuec", {}).get(energy_source, 0)
        bbrec = baseline_compliance.get("bbrec", {}).get(energy_source, 0)

        if energy_source == "ELECTRICITY":
            ratio, ghg = 2.80, 0.037
            label = "Electricity"
        elif energy_source == "NATURAL_GAS":
            ratio, ghg = 1.05, 0.053
            label = "Natural Gas"
        else:
            ratio, ghg = 0.0, 0.0
            label = energy_source

        file.write(
            f"""
                            <tr class="text-center">
                                <td>{label}</td>
                                <td class="baselineUnregulatedEnergy">{round(bbuec, 1):,}</td>
                                <td class="baselineRegulatedEnergy">{round(bbrec, 1):,}</td>
                                <td class="baselineEnergyUse">{round(baseline_energy, 1):,}</td>
                                <td class="proposedEnergyUse">{round(proposed_energy, 1):,}</td>
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
                        <thead class="table-light border-bottom">
                            <tr class="text-center">
                                <th>Parameter</th>
                                <th>Symbol</th>
                                <th>Cost ($)</th>
                                <th>Site Energy (MMBtu)</th>
                                <th>Source Energy (MMBtu)</th>
                                <th>GHG (tCO₂e)</th>
                            </tr>
                        </thead>
                        <tbody class="small">
"""
    )

    rows = [
        (
            "Proposed building performance before site-generated renewable energy",
            "PBP<sub>nre</sub>",
            f"${round(output.get('total_proposed_building_energy_cost_excluding_renewable_energy', 0)):,}",
            round(proposed_compliance.get("pbp_nre", {}).get("site_energy", 0)),
            "-", "-"
        ),
        (
            "Proposed building performance including on-site renewable energy",
            "PBP",
            f"${round(output.get('total_proposed_building_energy_cost_including_renewable_energy', 0)):,}",
            "-", "-", "-"
        ),
        (
            "Baseline building unregulated energy",
            "BBUEC",
            f"${round(output.get('baseline_building_unregulated_energy_cost', 0)):,}",
            round(baseline_compliance.get("bbuec", {}).get("site_energy", 0)),
            "-", "-"
        ),
        (
            "Baseline building regulated energy",
            "BBREC",
            f"${round(output.get('baseline_building_regulated_energy_cost', 0)):,}",
            round(baseline_compliance.get("bbrec", {}).get("site_energy", 0)),
            "-", "-"
        ),
        (
            "Baseline building performance",
            "BBP",
            f"${round(output.get('baseline_building_performance_energy_cost', 0)):,}",
            round(baseline_compliance.get("bbp", {}).get("site_energy", 0)),
            "-", "-"
        ),
    ]

    for label, symbol, cost, site, source, ghg in rows:
        file.write(
            f"""
                            <tr class="text-center">
                                <td class="text-start">{label}</td>
                                <td><strong>{symbol}</strong></td>
                                <td>{cost}</td>
                                <td>{site}</td>
                                <td>{source}</td>
                                <td>{ghg}</td>
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
