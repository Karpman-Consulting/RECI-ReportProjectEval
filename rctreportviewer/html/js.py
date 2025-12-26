def write_javascript(file, rct_detailed_report):
    file.write(
        f"""
<script>
(() => {{
  "use strict";

  document.addEventListener("DOMContentLoaded", () => {{

    /* ==================== Utilities ==================== */

    const $ = (id) => document.getElementById(id);

    const parseNumber = (str) => {{
      if (str == null) return 0;
      return parseFloat(String(str).replace(/[$,]/g, "")) || 0;
    }};

    const getTextNumber = (id) => {{
      const el = $(id);
      return el ? parseNumber(el.textContent) : 0;
    }};

    const setText = (id, value) => {{
      const el = $(id);
      if (el) el.textContent = value;
    }};

    const setRatio = (id, numerator, denominator) => {{
      const el = $(id);
      if (!el) return;
      el.textContent = denominator !== 0 ? (numerator / denominator).toFixed(2) : "0.00";
    }};

    const sum = (arr) => (arr || []).reduce((a, b) => a + (Number(b) || 0), 0);
    
    const formatNumber = (v) => unitType() === "eui" ? Number(v).toFixed(1) : Math.round(v).toLocaleString();
    
    const formatTooltip = (ctx) => {{
      const v = ctx.parsed.y;
      return unitType() === "eui" ? v.toFixed(1) : Math.round(v).toLocaleString();
    }};
    
    const formatTotal = (v) => unitType() === "eui"
        ? v.toLocaleString(undefined, {{ minimumFractionDigits: 1, maximumFractionDigits: 1 }})
        : v.toLocaleString(undefined, {{ maximumFractionDigits: 0 }});

    /* ==================== Back to top ==================== */

    const backToTopBtn = $("back-to-top");
    const toggleBackToTopButton = () => {{
      if (!backToTopBtn) return;
      const show = (document.body.scrollTop > 100) || (document.documentElement.scrollTop > 100) || (window.scrollY > 100);
      backToTopBtn.style.opacity = show ? "1" : "0";
      backToTopBtn.style.visibility = show ? "visible" : "hidden";
    }};
    window.addEventListener("scroll", toggleBackToTopButton);

    window.scrollToTop = () => {{
      window.scrollTo({{ top: 0, behavior: "smooth" }});
    }};

    /* ==================== Bootstrap tooltips ==================== */

    document.querySelectorAll('[data-bs-toggle="tooltip"]').forEach(el => {{
      // ensure title exists so bootstrap doesn't throw
      if (el.getAttribute("title") == null) el.setAttribute("title", "");
      new bootstrap.Tooltip(el, {{ container: "body" }});
    }});

    /* ==================== Energy metrics recalculation (safe) ==================== */

    const energySourceTable = $("energySourceTable");
    const inputs = energySourceTable ? energySourceTable.querySelectorAll("input") : [];
    const rows = energySourceTable ? energySourceTable.querySelectorAll("tbody tr") : [];

    function recalculateEnergyMetrics() {{
      if (!rows || rows.length === 0) return;

      let proposedSourceEnergy = 0;
      let proposedGHGEmissions = 0;
      let baselineUnregulatedSourceEnergy = 0;
      let baselineUnregulatedGHGEmissions = 0;
      let baselineRegulatedSourceEnergy = 0;
      let baselineRegulatedGHGEmissions = 0;

      rows.forEach(row => {{
        const getVal = (cls) => parseNumber(row.querySelector(`.${{cls}}`)?.textContent || "0");
        const getInputVal = (cls) => parseNumber(row.querySelector(`.${{cls}}`)?.value || "0");

        const proposed = getVal("proposedEnergyUse");
        const unreg = getVal("baselineUnregulatedEnergy");
        const reg = getVal("baselineRegulatedEnergy");
        const ssr = getInputVal("siteSourceRatio");
        const ghg = getInputVal("ghgEmissionFactor");

        proposedSourceEnergy += proposed * ssr;
        proposedGHGEmissions += proposed * ghg;
        baselineUnregulatedSourceEnergy += unreg * ssr;
        baselineUnregulatedGHGEmissions += unreg * ghg;
        baselineRegulatedSourceEnergy += reg * ssr;
        baselineRegulatedGHGEmissions += reg * ghg;
      }});

      const baselineSourceEnergy = baselineUnregulatedSourceEnergy + baselineRegulatedSourceEnergy;
      const baselineGHGEmissions = baselineUnregulatedGHGEmissions + baselineRegulatedGHGEmissions;

      // All getTextNumber() are safe (0 if missing) to avoid console null errors
      const proposedSiteEnergy = getTextNumber("pbp_nre_site_energy") - getTextNumber("proposed_site_energy_savings");
      const proposedSrcEnergy = proposedSourceEnergy - getTextNumber("proposed_source_energy_savings");
      const proposedGHG = proposedGHGEmissions - getTextNumber("proposed_ghg_savings");

      const baselineSiteEnergy = getTextNumber("bbp_site_energy");
      const baselineUnregulatedSiteEnergy = getTextNumber("bbuec_site_energy");
      const baselineRegulatedSiteEnergy = getTextNumber("bbrec_site_energy");

      const bpfSite = getTextNumber("bpf_site_energy");
      const bpfSource = getTextNumber("bpf_source_energy");
      const bpfGHG = getTextNumber("bpf_ghg_emissions");

      // write totals (as strings, because some targets might be ratios/percent strings elsewhere)
      setText("pbp_nre_source_energy", proposedSourceEnergy.toLocaleString(undefined, {{ maximumFractionDigits: 0 }}));
      setText("pbp_nre_ghg", proposedGHGEmissions.toLocaleString(undefined, {{ maximumFractionDigits: 0 }}));
      setText("pbp_site_energy", proposedSiteEnergy.toLocaleString(undefined, {{ maximumFractionDigits: 0 }}));
      setText("pbp_source_energy", proposedSrcEnergy.toLocaleString(undefined, {{ maximumFractionDigits: 0 }}));
      setText("pbp_ghg", proposedGHG.toLocaleString(undefined, {{ maximumFractionDigits: 0 }}));

      setText("bbuec_source_energy", baselineUnregulatedSourceEnergy.toLocaleString(undefined, {{ maximumFractionDigits: 0 }}));
      setText("bbuec_ghg", baselineUnregulatedGHGEmissions.toLocaleString(undefined, {{ maximumFractionDigits: 0 }}));
      setText("bbrec_source_energy", baselineRegulatedSourceEnergy.toLocaleString(undefined, {{ maximumFractionDigits: 0 }}));
      setText("bbrec_ghg", baselineRegulatedGHGEmissions.toLocaleString(undefined, {{ maximumFractionDigits: 0 }}));
      setText("bbp_source_energy", baselineSourceEnergy.toLocaleString(undefined, {{ maximumFractionDigits: 0 }}));
      setText("bbp_ghg", baselineGHGEmissions.toLocaleString(undefined, {{ maximumFractionDigits: 0 }}));

      // PCIt ratios
      setRatio("pcit_site_energy",
        baselineUnregulatedSiteEnergy + bpfSite * baselineRegulatedSiteEnergy,
        baselineSiteEnergy
      );
      setRatio("pcit_source_energy",
        baselineUnregulatedSourceEnergy + bpfSource * baselineRegulatedSourceEnergy,
        baselineSourceEnergy
      );
      setRatio("pcit_ghg_emissions",
        baselineUnregulatedGHGEmissions + bpfGHG * baselineRegulatedGHGEmissions,
        baselineGHGEmissions
      );

      setRatio("pci_nre_site_energy", getTextNumber("pbp_nre_site_energy"), baselineSiteEnergy);
      setRatio("pci_nre_source_energy", proposedSourceEnergy, baselineSourceEnergy);
      setRatio("pci_nre_ghg", proposedGHGEmissions, baselineGHGEmissions);

      // PCI adjusted
      const capFraction = 0.05;
      const adjustedSiteSavings = Math.min(getTextNumber("proposed_site_energy_savings"), capFraction * baselineSiteEnergy);
      const adjustedSourceSavings = Math.min(getTextNumber("proposed_source_energy_savings"), capFraction * baselineSourceEnergy);
      const adjustedGHGSavings = Math.min(getTextNumber("proposed_ghg_savings"), capFraction * baselineGHGEmissions);

      const adjustedPBPSite = getTextNumber("pbp_nre_site_energy") - adjustedSiteSavings;
      const adjustedPBPSource = proposedSourceEnergy - adjustedSourceSavings;
      const adjustedPBPGHG = proposedGHGEmissions - adjustedGHGSavings;

      const pciAdjustedSite = baselineSiteEnergy ? adjustedPBPSite / baselineSiteEnergy : 0;
      const pciAdjustedSource = baselineSourceEnergy ? adjustedPBPSource / baselineSourceEnergy : 0;
      const pciAdjustedGHG = baselineGHGEmissions ? adjustedPBPGHG / baselineGHGEmissions : 0;

      setText("pci_adjusted_site_energy", pciAdjustedSite.toFixed(2));
      setText("pci_adjusted_source_energy", pciAdjustedSource.toFixed(2));
      setText("pci_adjusted_ghg", pciAdjustedGHG.toFixed(2));

      const getCost = (id) => {{
        const el = $(id);
        return el ? parseNumber(el.textContent) : 0;
      }};

      const baselineCost = getCost("bbp_cost");
      const proposedCost = getCost("pbp_cost");
      const proposedNRECost = getCost("pbp_nre_cost");

      // % Improvement excluding renewables
      const cost_savings_nre = baselineCost ? ((baselineCost - proposedNRECost) / baselineCost) * 100 : 0;
      const site_savings_nre = baselineSiteEnergy ? ((baselineSiteEnergy - getTextNumber("pbp_nre_site_energy")) / baselineSiteEnergy) * 100 : 0;
      const source_savings_nre = baselineSourceEnergy ? ((baselineSourceEnergy - proposedSourceEnergy) / baselineSourceEnergy) * 100 : 0;
      const ghg_savings_nre = baselineGHGEmissions ? ((baselineGHGEmissions - proposedGHGEmissions) / baselineGHGEmissions) * 100 : 0;

      setText("cost_savings_nre", cost_savings_nre.toFixed(1) + "%");
      setText("site_savings_nre", site_savings_nre.toFixed(1) + "%");
      setText("source_savings_nre", source_savings_nre.toFixed(1) + "%");
      setText("ghg_savings_nre", ghg_savings_nre.toFixed(1) + "%");

      // % Improvement including renewables
      const cost_savings = baselineCost ? ((baselineCost - proposedCost) / baselineCost) * 100 : 0;
      const site_savings = baselineSiteEnergy ? ((baselineSiteEnergy - getTextNumber("pbp_site_energy")) / baselineSiteEnergy) * 100 : 0;
      const source_savings = baselineSourceEnergy ? ((baselineSourceEnergy - proposedSrcEnergy) / baselineSourceEnergy) * 100 : 0;
      const ghg_savings = baselineGHGEmissions ? ((baselineGHGEmissions - proposedGHG) / baselineGHGEmissions) * 100 : 0;

      setText("cost_savings", cost_savings.toFixed(1) + "%");
      setText("site_savings", site_savings.toFixed(1) + "%");
      setText("source_savings", source_savings.toFixed(1) + "%");
      setText("ghg_savings", ghg_savings.toFixed(1) + "%");
    }}

    // wire energySourceTable inputs
    if (inputs && inputs.length) {{
      inputs.forEach(inp => inp.addEventListener("input", recalculateEnergyMetrics));
      recalculateEnergyMetrics();
    }}

    /* ==================== Energy performance metrics ==================== */

    const espTable = $("energySourcePerformanceTable");
    const energyPerformanceTable = $("energyPerformanceTable");

    const energyPerformanceInputs = espTable ? espTable.querySelectorAll("input") : [];
    const energyPerformanceRows = energyPerformanceTable ? energyPerformanceTable.querySelectorAll("tbody tr") : [];

    function recalculateEnergyPerformanceMetrics() {{
      if (!energyPerformanceRows || energyPerformanceRows.length === 0) return;

      let totProposedSiteEnergy = 0;
      let totBaselineSiteEnergy = 0;
      let totProposedSourceEnergy = 0;
      let totBaselineSourceEnergy = 0;
      let totProposedCost = 0;
      let totBaselineCost = 0;
      let totProposedGHGEmissions = 0;
      let totBaselineGHGEmissions = 0;

      const electricitySiteSourceRatio = parseFloat(document.querySelector(".electricitySiteSourceRatio")?.value || 2.80);
      const naturalGasSiteSourceRatio = parseFloat(document.querySelector(".naturalGasSiteSourceRatio")?.value || 1.05);
      const electricityGHGEmissionFactor = parseFloat(document.querySelector(".electricityGhgEmissionFactor")?.value || 0.37);
      const naturalGasGHGEmissionFactor = parseFloat(document.querySelector(".naturalGasGhgEmissionFactor")?.value || 0.53);

      const rowsArr = Array.from(energyPerformanceRows);

      rowsArr.forEach((row, idx) => {{
        const getCell = (cls) => row.getElementsByClassName(cls)[0];
        const getRowText = (cls) => parseNumber(getCell(cls)?.textContent || "0");

        const setRowText = (cls, value, asPercent=false, asCurrency=false) => {{
          const cell = getCell(cls);
          if (!cell) return;
          if (asPercent) {{
            cell.textContent = value.toFixed(1) + "%";
          }} else if (asCurrency) {{
            cell.textContent = value.toLocaleString("en-US", {{
              style: "currency",
              currency: "USD",
              minimumFractionDigits: 0,
              maximumFractionDigits: 0
            }});
          }} else {{
            cell.textContent = value.toFixed(1).toLocaleString();
          }}
        }};

        if (idx === rowsArr.length - 1) {{
          const totSiteSavings = totBaselineSiteEnergy ? (totBaselineSiteEnergy - totProposedSiteEnergy) / totBaselineSiteEnergy * 100 : 0;
          const totSourceSavings = totBaselineSourceEnergy ? (totBaselineSourceEnergy - totProposedSourceEnergy) / totBaselineSourceEnergy * 100 : 0;
          const totCostSavings = totBaselineCost ? (totBaselineCost - totProposedCost) / totBaselineCost * 100 : 0;
          const totGHGSavings = totBaselineGHGEmissions ? (totBaselineGHGEmissions - totProposedGHGEmissions) / totBaselineGHGEmissions * 100 : 0;

          setRowText("totSiteEnergyProposed", totProposedSiteEnergy);
          setRowText("totSiteEnergyBaseline", totBaselineSiteEnergy);
          setRowText("totSiteEnergySavings", totSiteSavings, true);

          setRowText("totSourceEnergyProposed", totProposedSourceEnergy);
          setRowText("totSourceEnergyBaseline", totBaselineSourceEnergy);
          setRowText("totSourceEnergySavings", totSourceSavings, true);

          setRowText("totCostProposed", totProposedCost, false, true);
          setRowText("totCostBaseline", totBaselineCost, false, true);
          setRowText("totCostSavings", totCostSavings, true);

          setRowText("totGhgEmissionsProposed", totProposedGHGEmissions);
          setRowText("totGhgEmissionsBaseline", totBaselineGHGEmissions);
          setRowText("totGhgEmissionsSavings", totGHGSavings, true);
          return;
        }}

        const proposedSiteEnergy = getRowText("siteEnergyProposed");
        const baselineSiteEnergy = getRowText("siteEnergyBaseline");
        const proposedCost = getRowText("energyCostProposed");
        const baselineCost = getRowText("energyCostBaseline");

        const elecP = getRowText("electricityProposed");
        const gasP = getRowText("naturalGasProposed");
        const elecB = getRowText("electricityBaseline");
        const gasB = getRowText("naturalGasBaseline");

        const proposedSourceEnergy = (elecP * electricitySiteSourceRatio) + (gasP * naturalGasSiteSourceRatio);
        const baselineSourceEnergy = (elecB * electricitySiteSourceRatio) + (gasB * naturalGasSiteSourceRatio);
        const sourceEnergySavings = baselineSourceEnergy ? ((baselineSourceEnergy - proposedSourceEnergy) / baselineSourceEnergy * 100) : 0;

        setRowText("sourceEnergyProposed", proposedSourceEnergy);
        setRowText("sourceEnergyBaseline", baselineSourceEnergy);
        setRowText("sourceEnergySavings", sourceEnergySavings, true);

        const proposedGHG = (elecP * electricityGHGEmissionFactor) + (gasP * naturalGasGHGEmissionFactor);
        const baselineGHG = (elecB * electricityGHGEmissionFactor) + (gasB * naturalGasGHGEmissionFactor);
        const ghgSavings = baselineGHG ? ((baselineGHG - proposedGHG) / baselineGHG * 100) : 0;

        setRowText("ghgEmissionsProposed", proposedGHG);
        setRowText("ghgEmissionsBaseline", baselineGHG);
        setRowText("ghgEmissionsSavings", ghgSavings, true);

        totProposedSiteEnergy += proposedSiteEnergy;
        totBaselineSiteEnergy += baselineSiteEnergy;
        totProposedSourceEnergy += proposedSourceEnergy;
        totBaselineSourceEnergy += baselineSourceEnergy;
        totProposedCost += proposedCost;
        totBaselineCost += baselineCost;
        totProposedGHGEmissions += proposedGHG;
        totBaselineGHGEmissions += baselineGHG;
      }});
    }}

    if (energyPerformanceInputs && energyPerformanceInputs.length) {{
      energyPerformanceInputs.forEach(inp => inp.addEventListener("input", recalculateEnergyPerformanceMetrics));
      recalculateEnergyPerformanceMetrics();
    }}

    /* ==================== Subtotals (fan-summary tables) ==================== */

    function calculateSubtotals() {{
      document.querySelectorAll(".fan-summary").forEach(table => {{
        let columnSums = [];
        let columnPrecisions = [];

        table.querySelectorAll("tr").forEach(row => {{
          if (row.classList.contains("subtotal")) {{
            row.querySelectorAll("td").forEach((td, colIndex) => {{
              if (colIndex === 0) return;
              const sumVal = columnSums[colIndex] || 0;
              const precision = columnPrecisions[colIndex] || 0;
              td.textContent = sumVal.toLocaleString(undefined, {{
                minimumFractionDigits: precision,
                maximumFractionDigits: precision
              }});
            }});
            columnSums = [];
            columnPrecisions = [];
          }} else {{
            row.querySelectorAll("td").forEach((td, colIndex) => {{
              let cleaned = (td.textContent || "").replace(/,/g, "").trim();
              let val = parseFloat(cleaned);
              if (!Number.isFinite(val)) val = 0;
              let decimals = (cleaned.split(".")[1] || "").length;
              columnPrecisions[colIndex] = Math.max(columnPrecisions[colIndex] || 0, decimals);
              columnSums[colIndex] = (columnSums[colIndex] || 0) + val;
            }});
          }}
        }});
      }});
    }}
    calculateSubtotals();

    /* ==================== Charts ==================== */

    const labels = {[
      k.replace("_", " ").title()
      for k in rct_detailed_report.baseline_model_summary["elec_by_end_use"].keys()
    ]};

    const elecDataRaw = {{
      consumption: {{
        baseline: {list(rct_detailed_report.baseline_model_summary["elec_by_end_use"].values())},
        proposed: {list(rct_detailed_report.proposed_model_summary["elec_by_end_use"].values())}
      }},
      eui: {{
        baseline: {list(rct_detailed_report.baseline_model_summary["elec_by_end_use_eui"].values())},
        proposed: {list(rct_detailed_report.proposed_model_summary["elec_by_end_use_eui"].values())}
      }}
    }};

    const gasDataRaw = {{
      consumption: {{
        baseline: {list(rct_detailed_report.baseline_model_summary["gas_by_end_use"].values())},
        proposed: {list(rct_detailed_report.proposed_model_summary["gas_by_end_use"].values())}
      }},
      eui: {{
        baseline: {list(rct_detailed_report.baseline_model_summary["gas_by_end_use_eui"].values())},
        proposed: {list(rct_detailed_report.proposed_model_summary["gas_by_end_use_eui"].values())}
      }}
    }};

    const energyDataRaw = {{
      consumption: {{
        baseline: {list(rct_detailed_report.baseline_model_summary["energy_by_end_use"].values())},
        proposed: {list(rct_detailed_report.proposed_model_summary["energy_by_end_use"].values())}
      }},
      eui: {{
        baseline: {list(rct_detailed_report.baseline_model_summary["energy_by_end_use_eui"].values())},
        proposed: {list(rct_detailed_report.proposed_model_summary["energy_by_end_use_eui"].values())}
      }}
    }};

    // Your original palette:
    // elec: baseline blue, proposed green
    // gas: baseline orange, proposed red
    // energy: baseline maroon, proposed teal
    const COLORS = {{
      elec: {{
        baselineBg: "rgba(54, 162, 235, 0.7)",
        proposedBg: "rgba(75, 192, 75, 0.7)",
        baselineText: "rgb(54, 162, 235)",
        proposedText: "rgb(75, 192, 75)"
      }},
      gas: {{
        baselineBg: "rgba(255, 180, 80, 0.5)",
        proposedBg: "rgba(255, 100, 100, 0.5)",
        baselineText: "rgb(255, 180, 80)",
        proposedText: "rgb(255, 100, 100)"
      }},
      energy: {{
        baselineBg: "rgba(128, 0, 64, 0.6)",
        proposedBg: "rgba(0, 128, 128, 0.6)",
        baselineText: "rgb(128, 0, 64)",
        proposedText: "rgb(0, 128, 128)"
      }}
    }};

    function makeChart(canvasId, title, unit, palette, dataPair) {{
      const canvas = $(canvasId);
      if (!canvas) return null;

      return new Chart(canvas, {{
        type: "bar",
        data: {{
          labels,
          datasets: [
            {{
              label: "Baseline",
              data: dataPair.baseline,
              backgroundColor: palette.baselineBg
            }},
            {{
              label: "Proposed",
              data: dataPair.proposed,
              backgroundColor: palette.proposedBg
            }}
          ]
        }},
        options: {{
          responsive: true,
          interaction: {{ mode: "index", intersect: false }},
          plugins: {{
            title: {{ display: true, text: title }},
            tooltip: {{
              callbacks: {{
                label: (ctx) => {{
                  const v = ctx.parsed.y;
                  const val = unitType() === "eui"
                    ? v.toFixed(1)
                    : Math.round(v).toLocaleString();
                  return `${{ctx.dataset.label}}: ${{val}}`;
                }}
              }}
            }}
          }},
          scales: {{
            x: {{
              stacked: false,
              ticks: {{
                minRotation: 60,
                maxRotation: 60
              }}
            }},
            y: {{
              beginAtZero: true,
              title: {{ display: true, text: unit, font: {{ size: 14 }} }},
              ticks: {{
                callback: (value) =>
                  unitType() === "eui" ? value.toFixed(1) : Math.round(value)
              }}
            }}
          }}
        }}
      }});
    }}

    const elecChart = makeChart("elecByEndUse", "Electricity By End Use", "kWh", COLORS.elec, elecDataRaw.consumption);
    const gasChart = makeChart("gasByEndUse", "Natural Gas By End Use", "Therms", COLORS.gas, gasDataRaw.consumption);
    const energyChart = makeChart("energyByEndUse", "Total Site Energy By End Use", "kBtu", COLORS.energy, energyDataRaw.consumption);

    let currentChart = "elec";

    const containers = {{
      elec: $("elecChartContainer"),
      gas: $("gasChartContainer"),
      energy: $("energyChartContainer")
    }};

    function unitType() {{
      return $("unitToggle")?.checked ? "eui" : "consumption";
    }}

    function updateTotals() {{
      const u = unitType();
      let baseline = [];
      let proposed = [];
      let unitLabel = "";

      if (currentChart === "elec") {{
        baseline = elecDataRaw[u].baseline;
        proposed = elecDataRaw[u].proposed;
        unitLabel = (u === "eui") ? "kBtu/ft²" : "kWh";
      }} else if (currentChart === "gas") {{
        baseline = gasDataRaw[u].baseline;
        proposed = gasDataRaw[u].proposed;
        unitLabel = (u === "eui") ? "kBtu/ft²" : "Therms";
      }} else {{
        baseline = energyDataRaw[u].baseline;
        proposed = energyDataRaw[u].proposed;
        unitLabel = (u === "eui") ? "kBtu/ft²" : "kBtu";
      }}

      setText("baselineTotal", `Baseline Total: ${{formatTotal(sum(baseline))}} ${{unitLabel}}`);
      setText("proposedTotal", `Proposed Total: ${{formatTotal(sum(proposed))}} ${{unitLabel}}`);

      const baselineEl = $("baselineTotal");
      const proposedEl = $("proposedTotal");
      if (baselineEl) baselineEl.style.color = COLORS[currentChart].baselineText;
      if (proposedEl) proposedEl.style.color = COLORS[currentChart].proposedText;
    }}

    function updateChartsForUnits() {{
      const u = unitType();

      if (elecChart) {{
        elecChart.data.datasets[0].data = elecDataRaw[u].baseline;
        elecChart.data.datasets[1].data = elecDataRaw[u].proposed;
        elecChart.options.scales.y.title.text = (u === "eui") ? "kBtu/ft²" : "kWh";
        elecChart.update();
      }}

      if (gasChart) {{
        gasChart.data.datasets[0].data = gasDataRaw[u].baseline;
        gasChart.data.datasets[1].data = gasDataRaw[u].proposed;
        gasChart.options.scales.y.title.text = (u === "eui") ? "kBtu/ft²" : "Therms";
        gasChart.update();
      }}

      if (energyChart) {{
        energyChart.data.datasets[0].data = energyDataRaw[u].baseline;
        energyChart.data.datasets[1].data = energyDataRaw[u].proposed;
        energyChart.options.scales.y.title.text = (u === "eui") ? "kBtu/ft²" : "kBtu";
        energyChart.update();
      }}

      updateTotals();
    }}

    function showChart(type) {{
      currentChart = type;

      Object.entries(containers).forEach(([k, el]) => {{
        if (!el) return;
        el.classList.toggle("d-none", k !== type);
      }});

      // Some browsers/containers need a tick before resize renders correctly
      setTimeout(() => {{
        if (type === "elec" && elecChart) elecChart.resize();
        if (type === "gas" && gasChart) gasChart.resize();
        if (type === "energy" && energyChart) energyChart.resize();
      }}, 0);

      updateTotals();
    }}

    /* ==================== Wire chart controls (NO inline handlers) ==================== */

    const btnElec = $("btn-elec");
    const btnGas = $("btn-gas");
    const btnEnergy = $("btn-energy");
    const unitToggle = $("unitToggle");

    const lblElec = document.querySelector('label[for="btn-elec"]');
    const lblGas = document.querySelector('label[for="btn-gas"]');
    const lblEnergy = document.querySelector('label[for="btn-energy"]');

    const wireShow = (radioEl, labelEl, type) => {{
      if (radioEl) {{
        radioEl.addEventListener("change", () => {{
          if (radioEl.checked) showChart(type);
        }});
      }}
      if (labelEl) {{
        // Some Bootstrap btn-check setups don't reliably fire "change" in all cases;
        // clicking label always works.
        labelEl.addEventListener("click", () => showChart(type));
      }}
    }};

    wireShow(btnElec, lblElec, "elec");
    wireShow(btnGas, lblGas, "gas");
    wireShow(btnEnergy, lblEnergy, "energy");

    if (unitToggle) {{
      unitToggle.addEventListener("change", updateChartsForUnits);
    }}

    // Init (default checked is elec)
    showChart("elec");
    updateChartsForUnits();

  }});
}})();
</script>
"""
    )
