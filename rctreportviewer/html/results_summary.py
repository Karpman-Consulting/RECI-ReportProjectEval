

def write_results_summary(file, rct_detailed_report):
    # ----------------------- Model Results Summary -----------------------
    file.write(f"""
            <div class="mb-3 me-4">
                <button class="btn btn-info collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapse-model-results-summary" aria-expanded="false">
                    Results Summary
                </button>

                <div id="collapse-model-results-summary" class="accordion-collapse collapse">
                    <div class="accordion-body">
    """)

    file.write(f"""
                        <div style="position: relative; left: 360px;" class="mb-3">
                            <div class="btn-group" role="group" aria-label="Chart toggle">
                                <input type="radio" class="btn-check" name="chartOptions" id="btn-elec" autocomplete="off" checked>
                                <label style="width: 95px;" class="btn btn-outline-primary" for="btn-elec" onclick="showChart('elec')">Electricity</label>

                                <input type="radio" class="btn-check" name="chartOptions" id="btn-gas" autocomplete="off">
                                <label style="width: 95px;" class="btn btn-outline-danger" for="btn-gas" onclick="showChart('gas')">Gas</label>

                                <input type="radio" class="btn-check" name="chartOptions" id="btn-energy" autocomplete="off">
                                <label style="width: 95px;" class="btn btn-outline-success" for="btn-energy" onclick="showChart('energy')">Total</label>
                            </div>
                        </div>

                        <div class="form-check form-switch mb-3" style="margin-left: 725px;">
                          <input class="form-check-input" type="checkbox" id="unitToggle" onchange="toggleUnits()">
                          <label class="form-check-label" for="unitToggle">Show EUI (kBtu/ft²)</label>
                        </div>

                        <div class="mb-3" style="position: relative; left: 260px;">
                          <span id="baselineTotal" class="me-4 fw-bold">Baseline Total: </span>
                          <span id="proposedTotal" class="fw-bold">Proposed Total: </span>
                        </div>

                        <div id="elecChartContainer" style="width: 900px; height: 500px;">
                          <canvas id="elecByEndUse"></canvas>
                        </div>
                        <div id="gasChartContainer" style="width: 900px; height: 500px; display: none;">
                          <canvas id="gasByEndUse"></canvas>
                        </div>
                        <div id="energyChartContainer" style="width: 900px; height: 500px; display: none;">
                          <canvas id="energyByEndUse"></canvas>
                        </div>

                        <h3>Energy Performance Summary</h3>
                        <table class="table table-sm table-borderless" style="width: 600px;" id="energySourcePerformanceTable">
                            <thead>
                                <tr class="text-center">
                                    <th style="border: 2px solid black;">Energy Source</th>
                                    <th style="border: 2px solid black;">Source-Site<br>Ratio</th>
                                    <th style="border: 2px solid black;">GHG Emission Factor<br>(Metric Ton CO<sub>2</sub>/kBtu)</th>
                                </tr>
                            </thead>
                            <tbody style="border: 2px solid black;">
                                <tr style="font-size: 12px;" class="lh-1 text-center">
                                        <td class="align-middle">Electricity</td>
                                        <td><input type="number" class="electricitySiteSourceRatio" value="2.80" step="0.01" style="width: 60px;"></td>
                                        <td><input type="number" class="electricityGhgEmissionFactor" value="0.37" step="0.01" style="width: 60px;"></td>
                                </tr>
                                <tr style="font-size: 12px;" class="lh-1 text-center">
                                        <td class="align-middle">Natural Gas</td>
                                        <td><input type="number" class="naturalGasSiteSourceRatio" value="1.05" step="0.01" style="width: 60px;"></td>
                                        <td><input type="number" class="naturalGasGhgEmissionFactor" value="0.53" step="0.01" style="width: 60px;"></td>
                                </tr>
                            </tbody>
                        </table>

                        <table class="table table-sm table-borderless" style="width: 1250px;" id="energyPerformanceTable">
                            <thead>
                                <tr class="text-center">
                                    <th colspan="1" class="col-4"></th>
                                    <th colspan="2" class="col-4" style="border: 2px solid black; display:none;">Energy Use from Electricity</th>
                                    <th colspan="2" class="col-4" style="border: 2px solid black; display:none;">Energy Use from Natural Gas</th>
                                    <th colspan="3" class="col-4" style="border: 2px solid black;">Site Energy Use Intensity (kBtu/sf/yr)</th>
                                    <th colspan="3" class="col-4" style="border: 2px solid black;">Source Energy Use Intensity (kBtu/sf/yr)</th>
                                    <th colspan="3" class="col-4" style="border: 2px solid black;">Energy Cost ($/yr)</th>
                                    <th colspan="3" class="col-4" style="border: 2px solid black;">GHG Emission Intensity (kg CO<sub>2</sub>/sf/yr)</th>
                                </tr>
                                <tr class="text-center">
                                    <th style="border: 2px solid black;">End-use</th>
                                    <th style="border: 2px solid black; display:none;">Proposed</th>
                                    <th style="border: 2px solid black; display:none;">Baseline</th>
                                    <th style="border: 2px solid black; display:none;">Proposed</th>
                                    <th style="border: 2px solid black; display:none;">Baseline</th>
                                    <th style="border: 2px solid black;">Proposed</th>
                                    <th style="border: 2px solid black;">Baseline</th>
                                    <th style="border: 2px solid black;">% Savings</th>
                                    <th style="border: 2px solid black;">Proposed</th>
                                    <th style="border: 2px solid black;">Baseline</th>
                                    <th style="border: 2px solid black;">% Savings</th>
                                    <th style="border: 2px solid black;">Proposed</th>
                                    <th style="border: 2px solid black;">Baseline</th>
                                    <th style="border: 2px solid black;">% Savings</th>
                                    <th style="border: 2px solid black;">Proposed</th>
                                    <th style="border: 2px solid black;">Baseline</th>
                                    <th style="border: 2px solid black;">% Savings</th>
                                </tr>
                            </thead>
                            <tbody style="border: 2px solid black;">
    """)
    proposed_energy_by_end_use = rct_detailed_report.proposed_model_summary[
        "energy_by_end_use_eui"
    ]
    baseline_energy_by_end_use = rct_detailed_report.baseline_model_summary[
        "energy_by_end_use_eui"
    ]
    end_uses = set((baseline_energy_by_end_use | proposed_energy_by_end_use).keys())
    for end_use in end_uses:
        baseline_site_energy = baseline_energy_by_end_use.get(end_use, 0)
        proposed_site_energy = proposed_energy_by_end_use.get(end_use, 0)
        proposed_electricity = rct_detailed_report.proposed_model_summary[
            "elec_by_end_use_eui"
        ].get(end_use, 0)
        baseline_electricity = rct_detailed_report.baseline_model_summary[
            "elec_by_end_use_eui"
        ].get(end_use, 0)
        proposed_natural_gas = rct_detailed_report.proposed_model_summary[
            "gas_by_end_use_eui"
        ].get(end_use, 0)
        baseline_natural_gas = rct_detailed_report.baseline_model_summary[
            "gas_by_end_use_eui"
        ].get(end_use, 0)
        proposed_cost = rct_detailed_report.proposed_model_summary[
            "cost_by_end_use"
        ].get(end_use, 0)
        baseline_cost = rct_detailed_report.baseline_model_summary[
            "cost_by_end_use"
        ].get(end_use, 0)
        if not baseline_site_energy and not proposed_site_energy:
            continue
        site_improvement = (
            (baseline_site_energy - proposed_site_energy)
            / baseline_site_energy
            * 100
            if baseline_site_energy
            else (0 - proposed_site_energy) * 100  # Avoid division by zero
        )
        cost_improvement = (
            (baseline_cost - proposed_cost) / baseline_cost * 100
            if baseline_cost
            else (0 - proposed_cost) * 100  # Avoid division by zero
        )
        file.write(f"""
                                <tr style="font-size: 12px;" class="lh-1 text-center">
                                    <td style="border-right: 2px solid black;">{end_use.replace('_', ' ').title()}</td>
                                    <td class="electricityProposed" style="display:none;">{round(proposed_electricity, 1):,}</td>
                                    <td class="electricityBaseline" style="display:none;">{round(baseline_electricity, 1):,}</td>
                                    <td class="naturalGasProposed" style="display:none;">{round(proposed_natural_gas, 1):,}</td>
                                    <td class="naturalGasBaseline" style="display:none;">{round(baseline_natural_gas, 1):,}</td>
                                    <td class="siteEnergyProposed">{round(proposed_site_energy, 1):,}</td>
                                    <td class="siteEnergyBaseline">{round(baseline_site_energy, 1):,}</td>
                                    <td class="siteEnergySavings" style="border-right: 2px solid black;">{round(site_improvement, 1):,}%</td>
                                    <td class="sourceEnergyProposed">-</td>
                                    <td class="sourceEnergyBaseline">-</td>
                                    <td class="sourceEnergySavings" style="border-right: 2px solid black;">-</td>
                                    <td class="energyCostProposed">${round(proposed_cost):,}</td>
                                    <td class="energyCostBaseline">${round(baseline_cost):,}</td>
                                    <td class="energyCostSavings" style="border-right: 2px solid black;">{round(cost_improvement, 1):,}%</td>
                                    <td class="ghgEmissionsProposed">-</td>
                                    <td class="ghgEmissionsBaseline">-</td>
                                    <td class="ghgEmissionsSavings" style="border-right: 2px solid black;">-</td>
                                </tr>
    """)

    file.write(f"""
                                <tr style="font-size: 12px;" class="energyPerformanceTotals lh-1 text-center">
                                    <td style="border-right: 2px solid black;">Total</td>
                                    <td class="totSiteEnergyProposed">-</td>
                                    <td class="totSiteEnergyBaseline">-</td>
                                    <td class="totSiteEnergySavings"style="border-right: 2px solid black;">-</td>
                                    <td class="totSourceEnergyProposed">-</td>
                                    <td class="totSourceEnergyBaseline">-</td>
                                    <td class="totSourceEnergySavings" style="border-right: 2px solid black;">-</td>
                                    <td class="totCostProposed">-</td>
                                    <td class="totCostBaseline">-</td>
                                    <td class="totCostSavings" style="border-right: 2px solid black;">-</td>
                                    <td class="totGhgEmissionsProposed">-</td>
                                    <td class="totGhgEmissionsBaseline">-</td>
                                    <td class="totGhgEmissionsSavings" style="border-right: 2px solid black;">-</td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
    """)
