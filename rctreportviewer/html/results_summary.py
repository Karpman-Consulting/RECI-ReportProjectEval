def write_results_summary(file, rct_detailed_report):
    file.write(
        """
<section class="mb-4">
  <div class="card shadow-sm">

    <!-- CLICKABLE HEADER -->
    <div class="card-header bg-light d-flex align-items-center"
         role="button"
         data-bs-toggle="collapse"
         data-bs-target="#collapse-model-results-summary"
         aria-expanded="false"
         style="cursor: pointer;">
      <span class="fw-semibold">Results Summary</span>
    </div>

    <div id="collapse-model-results-summary" class="collapse">
      <div class="card-body">
"""
    )

    # ---------- Chart Controls ----------
    file.write(
        """
        <div class="row align-items-center mb-3">
          <div class="col-md-4"></div>

          <div class="col-md-4 d-flex justify-content-center">
            <div class="btn-group" role="group">
              <input type="radio" class="btn-check" name="chartOptions" id="btn-elec" checked>
              <label class="btn btn-outline-primary" for="btn-elec">Electricity</label>

              <input type="radio" class="btn-check" name="chartOptions" id="btn-gas">
              <label class="btn btn-outline-danger" for="btn-gas">Gas</label>

              <input type="radio" class="btn-check" name="chartOptions" id="btn-energy">
              <label class="btn btn-outline-success" for="btn-energy">Total</label>
            </div>
          </div>

          <div class="col-md-4 text-md-end text-center mt-2 mt-md-0">
            <div class="form-check form-switch d-inline-block">
              <input class="form-check-input" type="checkbox" id="unitToggle">
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
            <div id="elecChartContainer"><canvas id="elecByEndUse"></canvas></div>
            <div id="gasChartContainer" class="d-none"><canvas id="gasByEndUse"></canvas></div>
            <div id="energyChartContainer" class="d-none"><canvas id="energyByEndUse"></canvas></div>
          </div>
        </div>
"""
    )

    # ---------- Pie Charts ----------
    file.write(
        """
        <div class="text-center mb-3">
          <div id="pieLegend" class="d-flex flex-wrap justify-content-center gap-3"></div>
        </div>
        
        <div class="row justify-content-center mb-4">
          <div class="col-lg-10">
            <div id="elecPieContainer" class="mt-4 d-none">
              <div class="row">
                <div class="col-md-6"><canvas id="elecPieBaseline"></canvas></div>
                <div class="col-md-6"><canvas id="elecPieProposed"></canvas></div>
              </div>
            </div>

            <div id="gasPieContainer" class="mt-4 d-none">
              <div class="row">
                <div class="col-md-6"><canvas id="gasPieBaseline"></canvas></div>
                <div class="col-md-6"><canvas id="gasPieProposed"></canvas></div>
              </div>
            </div>

            <div id="energyPieContainer" class="mt-4">
              <div class="row">
                <div class="col-md-6"><canvas id="energyPieBaseline"></canvas></div>
                <div class="col-md-6"><canvas id="energyPieProposed"></canvas></div>
              </div>
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
            <thead class="table-light text-center">
              <tr>
                <th>Energy Source</th>
                <th>Source–Site Ratio</th>
                <th>GHG Emission Factor (tCO₂/kBtu)</th>
              </tr>
            </thead>
            <tbody class="text-center small">
              <tr>
                <td>Electricity</td>
                <td><input type="number" class="electricitySiteSourceRatio form-control form-control-sm text-center" value="2.80" step="0.01"></td>
                <td><input type="number" class="electricityGhgEmissionFactor form-control form-control-sm text-center" value="0.37" step="0.01"></td>
              </tr>
              <tr>
                <td>Natural Gas</td>
                <td><input type="number" class="naturalGasSiteSourceRatio form-control form-control-sm text-center" value="1.05" step="0.01"></td>
                <td><input type="number" class="naturalGasGhgEmissionFactor form-control form-control-sm text-center" value="0.53" step="0.01"></td>
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
            <thead class="table-light text-center">
              <tr class="border-top">
                <th rowspan="2" class="border-start">End-use</th>
                <th colspan="3" class="border-start">Site EUI</th>
                <th colspan="3" class="border-start">Source EUI</th>
                <th colspan="3" class="border-start">Energy Cost</th>
                <th colspan="3" class="border-start border-end">GHG Intensity</th>
              </tr>
              <tr>
                <th class="border-start">Proposed</th><th class="border-start">Baseline</th><th class="border-start">%</th>
                <th class="border-start">Proposed</th><th class="border-start">Baseline</th><th class="border-start">%</th>
                <th class="border-start">Proposed</th><th class="border-start">Baseline</th><th class="border-start">%</th>
                <th class="border-start">Proposed</th><th class="border-start">Baseline</th><th class="border-start border-end">%</th>
              </tr>
            </thead>
            <tbody class="small text-center">
"""
    )

    # ---- End-use rows ----
    for end_use in set(
        rct_detailed_report.proposed_model_summary["energy_by_end_use_eui"]
    ):
        p_site = rct_detailed_report.proposed_model_summary["energy_by_end_use_eui"][
            end_use
        ]
        b_site = rct_detailed_report.baseline_model_summary[
            "energy_by_end_use_eui"
        ].get(end_use, 0)
        s = ((b_site - p_site) / b_site * 100) if b_site else 0

        elec_p = rct_detailed_report.proposed_model_summary["elec_by_end_use_eui"].get(
            end_use, 0
        )
        gas_p = rct_detailed_report.proposed_model_summary["gas_by_end_use_eui"].get(
            end_use, 0
        )
        elec_b = rct_detailed_report.baseline_model_summary["elec_by_end_use_eui"].get(
            end_use, 0
        )
        gas_b = rct_detailed_report.baseline_model_summary["gas_by_end_use_eui"].get(
            end_use, 0
        )

        cost_p = rct_detailed_report.proposed_model_summary["cost_by_end_use"].get(
            end_use, 0
        )
        cost_b = rct_detailed_report.baseline_model_summary["cost_by_end_use"].get(
            end_use, 0
        )
        cost_s = ((cost_b - cost_p) / cost_b * 100) if cost_b else 0

        file.write(
            f"""
<tr>
  <td class="border-start text-start">{end_use.replace('_',' ').title()}</td>

  <td class="border-start siteEnergyProposed">{p_site:.1f}</td>
  <td class="border-start siteEnergyBaseline">{b_site:.1f}</td>
  <td class="border-start siteEnergySavings">{s:.1f}%</td>

  <td class="border-start sourceEnergyProposed">-</td>
  <td class="border-start sourceEnergyBaseline">-</td>
  <td class="border-start sourceEnergySavings">-</td>

  <td class="border-start energyCostProposed">${cost_p:,.0f}</td>
  <td class="border-start energyCostBaseline">${cost_b:,.0f}</td>
  <td class="border-start energyCostSavings">{cost_s:.1f}%</td>

  <td class="border-start ghgEmissionsProposed">-</td>
  <td class="border-start ghgEmissionsBaseline">-</td>
  <td class="border-start border-end ghgEmissionsSavings">-</td>

  <td class="electricityProposed d-none">{elec_p:.3f}</td>
  <td class="naturalGasProposed d-none">{gas_p:.3f}</td>
  <td class="electricityBaseline d-none">{elec_b:.3f}</td>
  <td class="naturalGasBaseline d-none">{gas_b:.3f}</td>
</tr>
"""
        )

    # ---- REQUIRED Total row ----
    file.write(
        """
<tr class="energyPerformanceTotals fw-bold">
  <td class="border-start text-start">Total</td>

  <td class="border-start totSiteEnergyProposed">-</td>
  <td class="border-start totSiteEnergyBaseline">-</td>
  <td class="border-start totSiteEnergySavings">-</td>

  <td class="border-start totSourceEnergyProposed">-</td>
  <td class="border-start totSourceEnergyBaseline">-</td>
  <td class="border-start totSourceEnergySavings">-</td>

  <td class="border-start totCostProposed">-</td>
  <td class="border-start totCostBaseline">-</td>
  <td class="border-start totCostSavings">-</td>

  <td class="border-start totGhgEmissionsProposed">-</td>
  <td class="border-start totGhgEmissionsBaseline">-</td>
  <td class="border-start border-end totGhgEmissionsSavings">-</td>
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
