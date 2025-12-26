def write_results_summary(file, rct_detailed_report):
    file.write(
        """
<section class="mb-4">
    <div class="card shadow-sm">
        <div class="card-header bg-light">
            <button class="btn btn-info"
                    type="button"
                    data-bs-toggle="collapse"
                    data-bs-target="#collapse-model-results-summary"
                    aria-expanded="false">
                Results Summary
            </button>
        </div>

        <div id="collapse-model-results-summary" class="collapse">
            <div class="card-body">
"""
    )

    # ---------- Chart Controls ----------
    file.write(
        """
                <div class="row align-items-center mb-3">
                    <div class="col-md-8 text-center">
                        <div class="btn-group" role="group" aria-label="Chart toggle">
                            <input type="radio" class="btn-check" name="chartOptions" id="btn-elec" autocomplete="off" checked>
                            <label class="btn btn-outline-primary" for="btn-elec" onclick="showChart('elec')">Electricity</label>

                            <input type="radio" class="btn-check" name="chartOptions" id="btn-gas" autocomplete="off">
                            <label class="btn btn-outline-danger" for="btn-gas" onclick="showChart('gas')">Gas</label>

                            <input type="radio" class="btn-check" name="chartOptions" id="btn-energy" autocomplete="off">
                            <label class="btn btn-outline-success" for="btn-energy" onclick="showChart('energy')">Total</label>
                        </div>
                    </div>

                    <div class="col-md-4 text-md-end text-center mt-2 mt-md-0">
                        <div class="form-check form-switch d-inline-block">
                            <input class="form-check-input" type="checkbox" id="unitToggle" onchange="toggleUnits()">
                            <label class="form-check-label" for="unitToggle">
                                Show EUI (kBtu/ft²)
                            </label>
                        </div>
                    </div>
                </div>

                <div class="text-center mb-3">
                    <span id="baselineTotal" class="me-4 fw-bold">Baseline Total:</span>
                    <span id="proposedTotal" class="fw-bold">Proposed Total:</span>
                </div>
"""
    )

    # ---------- Charts ----------
    file.write(
        """
                <div class="row justify-content-center mb-4">
                    <div class="col-lg-10">
                        <div id="elecChartContainer">
                            <canvas id="elecByEndUse"></canvas>
                        </div>
                        <div id="gasChartContainer" class="d-none">
                            <canvas id="gasByEndUse"></canvas>
                        </div>
                        <div id="energyChartContainer" class="d-none">
                            <canvas id="energyByEndUse"></canvas>
                        </div>
                    </div>
                </div>
"""
    )

    # ---------- Energy Source Performance ----------
    file.write(
        """
                <h5 class="mb-3">Energy Performance Summary</h5>

                <div class="table-responsive mb-4">
                    <table class="table table-sm align-middle" id="energySourcePerformanceTable">
                        <thead class="table-light border-bottom text-center">
                            <tr>
                                <th>Energy Source</th>
                                <th>Source–Site Ratio</th>
                                <th>GHG Emission Factor<br>(tCO₂/kBtu)</th>
                            </tr>
                        </thead>
                        <tbody class="small text-center">
                            <tr>
                                <td>Electricity</td>
                                <td>
                                    <input type="number"
                                           class="electricitySiteSourceRatio form-control form-control-sm text-center"
                                           value="2.80" step="0.01">
                                </td>
                                <td>
                                    <input type="number"
                                           class="electricityGhgEmissionFactor form-control form-control-sm text-center"
                                           value="0.37" step="0.01">
                                </td>
                            </tr>
                            <tr>
                                <td>Natural Gas</td>
                                <td>
                                    <input type="number"
                                           class="naturalGasSiteSourceRatio form-control form-control-sm text-center"
                                           value="1.05" step="0.01">
                                </td>
                                <td>
                                    <input type="number"
                                           class="naturalGasGhgEmissionFactor form-control form-control-sm text-center"
                                           value="0.53" step="0.01">
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </div>
"""
    )

    # ---------- Energy Performance Table ----------
    file.write(
        """
                <div class="table-responsive">
                    <table class="table table-sm align-middle" id="energyPerformanceTable">
                        <thead class="table-light border-bottom text-center">
                            <tr>
                                <th rowspan="2">End-use</th>
                                <th colspan="3">Site EUI (kBtu/sf/yr)</th>
                                <th colspan="3">Source EUI</th>
                                <th colspan="3">Energy Cost ($/yr)</th>
                                <th colspan="3">GHG Intensity</th>
                            </tr>
                            <tr>
                                <th>Proposed</th><th>Baseline</th><th>% Savings</th>
                                <th>Proposed</th><th>Baseline</th><th>% Savings</th>
                                <th>Proposed</th><th>Baseline</th><th>% Savings</th>
                                <th>Proposed</th><th>Baseline</th><th>% Savings</th>
                            </tr>
                        </thead>
                        <tbody class="small text-center">
"""
    )

    proposed_energy = rct_detailed_report.proposed_model_summary["energy_by_end_use_eui"]
    baseline_energy = rct_detailed_report.baseline_model_summary["energy_by_end_use_eui"]

    end_uses = set(proposed_energy) | set(baseline_energy)

    for end_use in end_uses:
        b_site = baseline_energy.get(end_use, 0)
        p_site = proposed_energy.get(end_use, 0)

        if not b_site and not p_site:
            continue

        site_savings = ((b_site - p_site) / b_site * 100) if b_site else 0

        file.write(
            f"""
                            <tr>
                                <td class="text-start">{end_use.replace('_', ' ').title()}</td>
                                <td>{round(p_site, 1)}</td>
                                <td>{round(b_site, 1)}</td>
                                <td>{round(site_savings, 1)}%</td>
                                <td>-</td><td>-</td><td>-</td>
                                <td>-</td><td>-</td><td>-</td>
                                <td>-</td><td>-</td><td>-</td>
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
