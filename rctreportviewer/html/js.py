def write_javascript(file, rct_detailed_report):
    file.write(
        f"""
<script>
(function () {{

  // -----------------------
  // Safe helpers (NO crashes)
  // -----------------------
  const parseNumber = (val) => {{
    if (val == null) return 0;
    return parseFloat(String(val).replace(/[$,]/g, "")) || 0;
  }};

  const getText = (id) => {{
    const el = document.getElementById(id);
    if (!el) return 0;
    return parseNumber(el.textContent);
  }};

  const setText = (id, value, opts = {{ minimumFractionDigits: 0, maximumFractionDigits: 0 }}) => {{
    const el = document.getElementById(id);
    if (!el) return;
    if (typeof value === "number" && Number.isFinite(value)) {{
      el.textContent = value.toLocaleString(undefined, opts);
    }} else {{
      el.textContent = String(value ?? "");
    }}
  }};

  const setRatio = (id, num, den) => {{
    const el = document.getElementById(id);
    if (!el) return;
    el.textContent = den ? (num / den).toFixed(2) : "0.00";
  }};

  // -----------------------
  // DOM ready
  // -----------------------
  document.addEventListener("DOMContentLoaded", () => {{

    const inputs = document.querySelectorAll('#energySourceTable input');
    const rows = document.querySelectorAll('#energySourceTable tbody tr');
    const perfInputs = document.querySelectorAll('#energySourcePerformanceTable input');
    const perfRows = document.querySelectorAll('#energySourcePerformanceTable tbody tr');

    // -----------------------
    // Tooltips (safe)
    // -----------------------
    if (window.bootstrap?.Tooltip) {{
      document.querySelectorAll('[data-bs-toggle="tooltip"]').forEach(el => {{
        new bootstrap.Tooltip(el, {{ container: 'body' }});
      }});
    }}

    // -----------------------
    // Energy Metrics
    // -----------------------
    function recalculateEnergyMetrics() {{
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
      const bSite = getText('bbp_site_energy');

      const pSite = getText('pbp_nre_site_energy') - getText('proposed_site_energy_savings');
      const pSrcAdj = pSrc - getText('proposed_source_energy_savings');
      const pGHGAdj = pGHG - getText('proposed_ghg_savings');

      setText('pbp_nre_source_energy', pSrc);
      setText('pbp_nre_ghg', pGHG);
      setText('pbp_site_energy', pSite);
      setText('pbp_source_energy', pSrcAdj);
      setText('pbp_ghg', pGHGAdj);

      setText('bbp_source_energy', bSrc);
      setText('bbp_ghg', bGHG);

      setRatio('pci_nre_site_energy', getText('pbp_nre_site_energy'), bSite);
      setRatio('pci_nre_source_energy', pSrc, bSrc);
      setRatio('pci_nre_ghg', pGHG, bGHG);
    }}

    inputs.forEach(i => i.addEventListener('input', recalculateEnergyMetrics));
    recalculateEnergyMetrics();

    // -----------------------
    // Back to top
    // -----------------------
    window.onscroll = () => {{
      const btn = document.getElementById("back-to-top");
      if (!btn) return;
      const show = document.documentElement.scrollTop > 100;
      btn.style.opacity = show ? "1" : "0";
      btn.style.visibility = show ? "visible" : "hidden";
    }};

    window.scrollToTop = () => {{
      window.scrollTo({{ top: 0, behavior: 'smooth' }});
    }};

    // -----------------------
    // Charts (SAFE)
    // -----------------------
    const labels = {[
      label.replace('_', ' ').title()
      for label in rct_detailed_report.baseline_model_summary["elec_by_end_use"].keys()
    ]};

    function makeBarChart(id, title, unit, baseline, proposed, colors) {{
      const canvas = document.getElementById(id);
      if (!canvas) return null;
      return new Chart(canvas, {{
        type: 'bar',
        data: {{
          labels,
          datasets: [
            {{ label: 'Baseline', data: baseline, backgroundColor: colors[0] }},
            {{ label: 'Proposed', data: proposed, backgroundColor: colors[1] }}
          ]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          plugins: {{
            title: {{ display: true, text: title }},
            tooltip: {{ mode: 'index', intersect: false }}
          }},
          scales: {{
            y: {{
              beginAtZero: true,
              title: {{ display: true, text: unit }}
            }},
            x: {{
              ticks: {{ minRotation: 60, maxRotation: 60 }}
            }}
          }}
        }}
      }});
    }}

    makeBarChart(
      'elecByEndUse',
      'Electricity By End Use',
      'kWh',
      {list(rct_detailed_report.baseline_model_summary["elec_by_end_use"].values())},
      {list(rct_detailed_report.proposed_model_summary["elec_by_end_use"].values())},
      ['rgba(54,162,235,0.7)', 'rgba(75,192,75,0.7)']
    );

    makeBarChart(
      'gasByEndUse',
      'Natural Gas By End Use',
      'Therms',
      {list(rct_detailed_report.baseline_model_summary["gas_by_end_use"].values())},
      {list(rct_detailed_report.proposed_model_summary["gas_by_end_use"].values())},
      ['rgba(255,180,80,0.6)', 'rgba(255,100,100,0.6)']
    );

    makeBarChart(
      'energyByEndUse',
      'Total Site Energy By End Use',
      'kBtu',
      {list(rct_detailed_report.baseline_model_summary["energy_by_end_use"].values())},
      {list(rct_detailed_report.proposed_model_summary["energy_by_end_use"].values())},
      ['rgba(128,0,64,0.6)', 'rgba(0,128,128,0.6)']
    );

  }});
}})();
</script>
"""
    )
