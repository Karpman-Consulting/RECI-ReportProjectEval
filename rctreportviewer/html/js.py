def write_javascript(file, rct_detailed_report):
    file.write(
        f"""
<script>
(function () {{

  // -----------------------
  // Safe helpers
  // -----------------------
  const parseNumber = (val) => {{
    if (val == null) return 0;
    return parseFloat(String(val).replace(/[$,]/g, "")) || 0;
  }};

  const getEl = (id) => document.getElementById(id);

  const getText = (id) => {{
    const el = getEl(id);
    if (!el) return 0;
    return parseNumber(el.textContent);
  }};

  const setText = (id, value, opts = {{ minimumFractionDigits: 0, maximumFractionDigits: 0 }}) => {{
    const el = getEl(id);
    if (!el) return;
    if (typeof value === "number" && Number.isFinite(value)) {{
      el.textContent = value.toLocaleString(undefined, opts);
    }} else {{
      el.textContent = String(value ?? "");
    }}
  }};

  const setRatio = (id, num, den) => {{
    const el = getEl(id);
    if (!el) return;
    el.textContent = den ? (num / den).toFixed(2) : "0.00";
  }};

  // Shared state for charts/toggles
  let elecChart = null;
  let gasChart = null;
  let energyChart = null;
  let currentChart = 'elec';  // 'elec' | 'gas' | 'energy'

  // -----------------------
  // Data injected from Python
  // -----------------------
  const labels = {[
      label.replace('_', ' ').title()
      for label in rct_detailed_report.baseline_model_summary["elec_by_end_use"].keys()
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

  function sumArray(arr) {{
    return (arr || []).reduce((acc, v) => acc + (Number(v) || 0), 0);
  }}

  function getUnitLabel(source, unitType) {{
    if (unitType === 'eui') return 'kBtu/ft²';
    if (source === 'elec') return 'kWh';
    if (source === 'gas') return 'Therms';
    return 'kBtu';
  }}

  function updateTotalColors(source) {{
    const baselineEl = getEl('baselineTotal');
    const proposedEl = getEl('proposedTotal');
    if (!baselineEl || !proposedEl) return;

    if (source === 'elec') {{
      baselineEl.style.color = 'rgb(54, 162, 235)';
      proposedEl.style.color = 'rgb(75, 192, 75)';
    }} else if (source === 'gas') {{
      baselineEl.style.color = 'rgb(255, 180, 80)';
      proposedEl.style.color = 'rgb(255, 100, 100)';
    }} else {{
      baselineEl.style.color = 'rgb(128, 0, 64)';
      proposedEl.style.color = 'rgb(0, 128, 128)';
    }}
  }}

  function updateTotals(source, unitType) {{
    const baselineEl = getEl('baselineTotal');
    const proposedEl = getEl('proposedTotal');
    if (!baselineEl || !proposedEl) return;

    let raw;
    if (source === 'elec') raw = elecDataRaw;
    else if (source === 'gas') raw = gasDataRaw;
    else raw = energyDataRaw;

    const baseline = raw[unitType]?.baseline || [];
    const proposed = raw[unitType]?.proposed || [];

    const unit = getUnitLabel(source, unitType);
    const baselineSum = sumArray(baseline).toLocaleString(undefined, {{ maximumFractionDigits: 0 }});
    const proposedSum = sumArray(proposed).toLocaleString(undefined, {{ maximumFractionDigits: 0 }});

    baselineEl.textContent = `Baseline Total: ${{baselineSum}} ${{unit}}`;
    proposedEl.textContent = `Proposed Total: ${{proposedSum}} ${{unit}}`;
  }}

  function setChartYAxisTitle(chart, unitType, consumptionUnit) {{
    if (!chart) return;
    chart.options.scales.y.title.text = (unitType === 'consumption') ? consumptionUnit : 'kBtu/ft²';
  }}

  function updateCharts(unitType) {{
    if (elecChart) {{
      elecChart.data.datasets[0].data = elecDataRaw[unitType].baseline;
      elecChart.data.datasets[1].data = elecDataRaw[unitType].proposed;
      setChartYAxisTitle(elecChart, unitType, 'kWh');
      elecChart.update();
    }}

    if (gasChart) {{
      gasChart.data.datasets[0].data = gasDataRaw[unitType].baseline;
      gasChart.data.datasets[1].data = gasDataRaw[unitType].proposed;
      setChartYAxisTitle(gasChart, unitType, 'Therms');
      gasChart.update();
    }}

    if (energyChart) {{
      energyChart.data.datasets[0].data = energyDataRaw[unitType].baseline;
      energyChart.data.datasets[1].data = energyDataRaw[unitType].proposed;
      setChartYAxisTitle(energyChart, unitType, 'kBtu');
      energyChart.update();
    }}
  }}

  // -----------------------
  // Functions required by inline HTML handlers
  // -----------------------
  window.toggleUnits = function () {{
    const toggle = getEl('unitToggle');
    const useEUI = toggle ? toggle.checked : false;
    const unitType = useEUI ? 'eui' : 'consumption';

    updateCharts(unitType);
    updateTotals(currentChart, unitType);
  }};

  window.showChart = function (type) {{
    const elecContainer = getEl('elecChartContainer');
    const gasContainer = getEl('gasChartContainer');
    const energyContainer = getEl('energyChartContainer');

    if (elecContainer) elecContainer.style.display = (type === 'elec') ? 'block' : 'none';
    if (gasContainer) gasContainer.style.display = (type === 'gas') ? 'block' : 'none';
    if (energyContainer) energyContainer.style.display = (type === 'energy') ? 'block' : 'none';

    currentChart = type;

    const toggle = getEl('unitToggle');
    const unitType = (toggle && toggle.checked) ? 'eui' : 'consumption';
    updateTotals(type, unitType);
    updateTotalColors(type);
  }};

  // -----------------------
  // Energy metrics (null-safe)
  // -----------------------
  function recalculateEnergyMetrics() {{
    const rows = document.querySelectorAll('#energySourceTable tbody tr');
    if (!rows.length) return;

    let pSrc = 0, pGHG = 0;
    let buSrc = 0, buGHG = 0;
    let brSrc = 0, brGHG = 0;

    rows.forEach(row => {{
      const txt = cls => parseNumber(row.querySelector(`.${{cls}}`)?.textContent);
      const inp = cls => parseNumber(row.querySelector(`.${{cls}}`)?.value);

      const proposed = txt('proposedEnergyUse');
      const unreg = txt('baselineUnregulatedEnergy');
      const reg = txt('baselineRegulatedEnergy');
      const ssr = inp('siteSourceRatio');
      const ghg = inp('ghgEmissionFactor');

      pSrc += proposed * ssr;
      pGHG += proposed * ghg;
      buSrc += unreg * ssr;
      buGHG += unreg * ghg;
      brSrc += reg * ssr;
      brGHG += reg * ghg;
    }});

    const bSrc = buSrc + brSrc;
    const bGHG = buGHG + brGHG;

    // These IDs are not always present depending on report layout,
    // so all reads/writes are guarded via getText/setText.
    const pSite = getText('pbp_nre_site_energy') - getText('proposed_site_energy_savings');
    const pSrcAdj = pSrc - getText('proposed_source_energy_savings');
    const pGHGAdj = pGHG - getText('proposed_ghg_savings');

    setText('pbp_nre_source_energy', pSrc);
    setText('pbp_nre_ghg', pGHG);
    setText('pbp_site_energy', pSite);
    setText('pbp_source_energy', pSrcAdj);
    setText('pbp_ghg', pGHGAdj);

    setText('bbuec_source_energy', buSrc);
    setText('bbuec_ghg', buGHG);
    setText('bbrec_source_energy', brSrc);
    setText('bbrec_ghg', brGHG);
    setText('bbp_source_energy', bSrc);
    setText('bbp_ghg', bGHG);

    const bSite = getText('bbp_site_energy');
    setRatio('pci_nre_site_energy', getText('pbp_nre_site_energy'), bSite);
    setRatio('pci_nre_source_energy', pSrc, bSrc);
    setRatio('pci_nre_ghg', pGHG, bGHG);
  }}

  // -----------------------
  // Fan subtotals (unchanged behavior, safe)
  // -----------------------
  function calculateSubtotals() {{
    document.querySelectorAll(".fan-summary").forEach(table => {{
      let columnSums = [];
      let columnPrecisions = [];

      table.querySelectorAll("tr").forEach(row => {{
        if (row.classList.contains("subtotal")) {{
          row.querySelectorAll("td").forEach((td, colIndex) => {{
            if (colIndex === 0) return;
            let sum = columnSums[colIndex] || 0;
            let precision = columnPrecisions[colIndex] || 0;
            td.textContent = sum.toLocaleString(undefined, {{ minimumFractionDigits: precision, maximumFractionDigits: precision }});
          }});
          columnSums = [];
          columnPrecisions = [];
        }} else {{
          row.querySelectorAll("td").forEach((td, colIndex) => {{
            let cleanedText = (td.textContent || "").replace(/,/g, "").trim();
            let value = parseFloat(cleanedText);
            if (!Number.isFinite(value)) return;
            let decimalPlaces = (cleanedText.split(".")[1] || "").length;
            columnPrecisions[colIndex] = Math.max(columnPrecisions[colIndex] || 0, decimalPlaces);
            columnSums[colIndex] = (columnSums[colIndex] || 0) + value;
          }});
        }}
      }});
    }});
  }}

  // -----------------------
  // DOM Ready init
  // -----------------------
  document.addEventListener("DOMContentLoaded", () => {{

    // Tooltips
    if (window.bootstrap?.Tooltip) {{
      document.querySelectorAll("[data-bs-toggle='tooltip']").forEach(el => {{
        if (el.getAttribute("title") == null) el.setAttribute("title", "");
      }});
      document.querySelectorAll('[data-bs-toggle="tooltip"]').forEach(el => {{
        new bootstrap.Tooltip(el, {{ container: 'body' }});
      }});
    }}

    // Hook energy table inputs
    const inputs = document.querySelectorAll('#energySourceTable input');
    inputs.forEach(input => input.addEventListener('input', recalculateEnergyMetrics));
    if (inputs.length) recalculateEnergyMetrics();

    // Fan subtotals
    calculateSubtotals();

    // Charts (guard canvases)
    if (window.Chart) {{
      const elecCanvas = getEl('elecByEndUse');
      const gasCanvas = getEl('gasByEndUse');
      const energyCanvas = getEl('energyByEndUse');

      if (elecCanvas) {{
        elecChart = new Chart(elecCanvas, {{
          type: 'bar',
          data: {{
            labels: labels,
            datasets: [
              {{ label: 'Baseline', data: elecDataRaw.consumption.baseline, backgroundColor: 'rgba(54, 162, 235, 0.7)' }},
              {{ label: 'Proposed', data: elecDataRaw.consumption.proposed, backgroundColor: 'rgba(75, 192, 75, 0.7)' }}
            ]
          }},
          options: {{
            responsive: true,
            plugins: {{
              title: {{ display: true, text: 'Electricity By End Use' }},
              tooltip: {{ mode: 'index', intersect: false }}
            }},
            interaction: {{ mode: 'index', intersect: false }},
            scales: {{
              x: {{ ticks: {{ minRotation: 60, maxRotation: 60 }} }},
              y: {{ beginAtZero: true, title: {{ display: true, text: 'kWh' }} }}
            }}
          }}
        }});
      }}

      if (gasCanvas) {{
        gasChart = new Chart(gasCanvas, {{
          type: 'bar',
          data: {{
            labels: labels,
            datasets: [
              {{ label: 'Baseline', data: gasDataRaw.consumption.baseline, backgroundColor: 'rgba(255, 180, 80, 0.5)' }},
              {{ label: 'Proposed', data: gasDataRaw.consumption.proposed, backgroundColor: 'rgba(255, 100, 100, 0.5)' }}
            ]
          }},
          options: {{
            responsive: true,
            plugins: {{
              title: {{ display: true, text: 'Natural Gas By End Use' }},
              tooltip: {{ mode: 'index', intersect: false }}
            }},
            interaction: {{ mode: 'index', intersect: false }},
            scales: {{
              x: {{ ticks: {{ minRotation: 60, maxRotation: 60 }} }},
              y: {{ beginAtZero: true, title: {{ display: true, text: 'Therms' }} }}
            }}
          }}
        }});
      }}

      if (energyCanvas) {{
        energyChart = new Chart(energyCanvas, {{
          type: 'bar',
          data: {{
            labels: labels,
            datasets: [
              {{ label: 'Baseline', data: energyDataRaw.consumption.baseline, backgroundColor: 'rgba(128, 0, 64, 0.6)' }},
              {{ label: 'Proposed', data: energyDataRaw.consumption.proposed, backgroundColor: 'rgba(0, 128, 128, 0.6)' }}
            ]
          }},
          options: {{
            responsive: true,
            plugins: {{
              title: {{ display: true, text: 'Total Site Energy By End Use' }},
              tooltip: {{ mode: 'index', intersect: false }}
            }},
            interaction: {{ mode: 'index', intersect: false }},
            scales: {{
              x: {{ ticks: {{ minRotation: 60, maxRotation: 60 }} }},
              y: {{ beginAtZero: true, title: {{ display: true, text: 'kBtu' }} }}
            }}
          }}
        }});
      }}
    }}

    // Initialize chart view + totals/colors to match your original behavior:
    // - showChart('elec') updates containers
    // - totals reflect consumption by default
    window.showChart(currentChart);
    updateTotals(currentChart, 'consumption');
    updateTotalColors(currentChart);

    // Back-to-top handler (safe)
    window.onscroll = function () {{
      const btn = getEl("back-to-top");
      if (!btn) return;
      const show = (document.body.scrollTop > 100 || document.documentElement.scrollTop > 100);
      btn.style.opacity = show ? "1" : "0";
      btn.style.visibility = show ? "visible" : "hidden";
    }};
  }});

  // Public API used by HTML
  // (already attached above: window.toggleUnits, window.showChart)
  window.scrollToTop = function () {{
    window.scrollTo({{ top: 0, behavior: 'smooth' }});
  }};

}})();
</script>
"""
    )
