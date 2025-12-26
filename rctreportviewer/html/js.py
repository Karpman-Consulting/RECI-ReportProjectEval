def write_javascript(file, rct_detailed_report):
    file.write(
        f"""
<script>
document.addEventListener("DOMContentLoaded", () => {{

    /* -------------------- Utilities -------------------- */

    const parseNumber = (str) =>
        parseFloat((str || "0").replace(/[$,]/g, "")) || 0;

    const getEl = (id) => document.getElementById(id);

    const setText = (id, value) => {{
        const el = getEl(id);
        if (el) el.textContent = value;
    }};

    const setRatio = (id, numerator, denominator) => {{
        const el = getEl(id);
        if (!el) return;
        el.textContent = denominator
            ? (numerator / denominator).toFixed(2)
            : "0.00";
    }};

    /* -------------------- Back to top -------------------- */

    const backToTopBtn = getEl("back-to-top");

    window.addEventListener("scroll", () => {{
        if (!backToTopBtn) return;
        const show = window.scrollY > 100;
        backToTopBtn.style.opacity = show ? "1" : "0";
        backToTopBtn.style.visibility = show ? "visible" : "hidden";
    }});

    window.scrollToTop = () =>
        window.scrollTo({{ top: 0, behavior: "smooth" }});

    /* -------------------- Tooltips -------------------- */

    document
        .querySelectorAll('[data-bs-toggle="tooltip"]')
        .forEach(el => new bootstrap.Tooltip(el, {{ container: "body" }}));

    /* -------------------- Charts -------------------- */

    const labels = {[
        label.replace("_", " ").title()
        for label in rct_detailed_report.baseline_model_summary["elec_by_end_use"].keys()
    ]};

    const elecDataRaw = {{
        consumption: {{
            baseline: {list(rct_detailed_report.baseline_model_summary["elec_by_end_use"].values())},
            proposed: {list(rct_detailed_report.proposed_model_summary["elec_by_end_use"].values())},
        }},
        eui: {{
            baseline: {list(rct_detailed_report.baseline_model_summary["elec_by_end_use_eui"].values())},
            proposed: {list(rct_detailed_report.proposed_model_summary["elec_by_end_use_eui"].values())},
        }},
    }};

    const gasDataRaw = {{
        consumption: {{
            baseline: {list(rct_detailed_report.baseline_model_summary["gas_by_end_use"].values())},
            proposed: {list(rct_detailed_report.proposed_model_summary["gas_by_end_use"].values())},
        }},
        eui: {{
            baseline: {list(rct_detailed_report.baseline_model_summary["gas_by_end_use_eui"].values())},
            proposed: {list(rct_detailed_report.proposed_model_summary["gas_by_end_use_eui"].values())},
        }},
    }};

    const energyDataRaw = {{
        consumption: {{
            baseline: {list(rct_detailed_report.baseline_model_summary["energy_by_end_use"].values())},
            proposed: {list(rct_detailed_report.proposed_model_summary["energy_by_end_use"].values())},
        }},
        eui: {{
            baseline: {list(rct_detailed_report.baseline_model_summary["energy_by_end_use_eui"].values())},
            proposed: {list(rct_detailed_report.proposed_model_summary["energy_by_end_use_eui"].values())},
        }},
    }};

    function makeBarChart(canvasId, title, unit, data) {{
        const ctx = getEl(canvasId);
        if (!ctx) return null;

        return new Chart(ctx, {{
            type: "bar",
            data,
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    title: {{ display: true, text: title }},
                }},
                scales: {{
                    y: {{
                        beginAtZero: true,
                        title: {{ display: true, text: unit }},
                    }},
                }},
            }},
        }});
    }}

    const elecChart = makeBarChart(
        "elecByEndUse",
        "Electricity By End Use",
        "kWh",
        {{
            labels,
            datasets: [
                {{ label: "Baseline", data: elecDataRaw.consumption.baseline, backgroundColor: "rgba(54,162,235,0.7)" }},
                {{ label: "Proposed", data: elecDataRaw.consumption.proposed, backgroundColor: "rgba(75,192,75,0.7)" }},
            ],
        }}
    );

    const gasChart = makeBarChart(
        "gasByEndUse",
        "Natural Gas By End Use",
        "Therms",
        {{
            labels,
            datasets: [
                {{ label: "Baseline", data: gasDataRaw.consumption.baseline, backgroundColor: "rgba(255,180,80,0.6)" }},
                {{ label: "Proposed", data: gasDataRaw.consumption.proposed, backgroundColor: "rgba(255,100,100,0.6)" }},
            ],
        }}
    );

    const energyChart = makeBarChart(
        "energyByEndUse",
        "Total Site Energy By End Use",
        "kBtu",
        {{
            labels,
            datasets: [
                {{ label: "Baseline", data: energyDataRaw.consumption.baseline, backgroundColor: "rgba(128,0,64,0.6)" }},
                {{ label: "Proposed", data: energyDataRaw.consumption.proposed, backgroundColor: "rgba(0,128,128,0.6)" }},
            ],
        }}
    );

    /* -------------------- Chart Controls -------------------- */

    const containers = {{
        elec: getEl("elecChartContainer"),
        gas: getEl("gasChartContainer"),
        energy: getEl("energyChartContainer"),
    }};

    function showChart(type) {{
        Object.entries(containers).forEach(([key, el]) => {{
            if (!el) return;
            el.classList.toggle("d-none", key !== type);
        }});
        [elecChart, gasChart, energyChart].forEach(c => c && c.resize());
        currentChart = type;
        updateTotals();
    }}

    function updateTotals() {{
        const useEUI = getEl("unitToggle")?.checked;
        const unitType = useEUI ? "eui" : "consumption";

        let baseline, proposed, unit;

        if (currentChart === "elec") {{
            baseline = elecDataRaw[unitType].baseline;
            proposed = elecDataRaw[unitType].proposed;
            unit = unitType === "eui" ? "kBtu/ft²" : "kWh";
        }} else if (currentChart === "gas") {{
            baseline = gasDataRaw[unitType].baseline;
            proposed = gasDataRaw[unitType].proposed;
            unit = unitType === "eui" ? "kBtu/ft²" : "Therms";
        }} else {{
            baseline = energyDataRaw[unitType].baseline;
            proposed = energyDataRaw[unitType].proposed;
            unit = unitType === "eui" ? "kBtu/ft²" : "kBtu";
        }}

        const sum = arr => arr.reduce((a, b) => a + b, 0);

        setText("baselineTotal", `Baseline Total: ${{sum(baseline).toLocaleString()}} ${{unit}}`);
        setText("proposedTotal", `Proposed Total: ${{sum(proposed).toLocaleString()}} ${{unit}}`);
    }}

    /* -------------------- Event Wiring -------------------- */

    let currentChart = "elec";

    getEl("btn-elec")?.addEventListener("change", () => showChart("elec"));
    getEl("btn-gas")?.addEventListener("change", () => showChart("gas"));
    getEl("btn-energy")?.addEventListener("change", () => showChart("energy"));

    getEl("unitToggle")?.addEventListener("change", updateTotals);

    showChart("elec");

}});
</script>
"""
    )
