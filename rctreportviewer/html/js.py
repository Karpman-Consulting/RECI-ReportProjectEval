def write_javascript(file, rct_detailed_report):
    file.write(
        f"""
<script>
document.addEventListener("DOMContentLoaded", () => {{

    /* ==================== Utilities ==================== */

    const $ = (id) => document.getElementById(id);
    const sum = (arr) => arr.reduce((a, b) => a + b, 0);

    /* ==================== Back to top ==================== */

    const backToTop = $("back-to-top");
    if (backToTop) {{
        window.addEventListener("scroll", () => {{
            const show = window.scrollY > 100;
            backToTop.style.opacity = show ? "1" : "0";
            backToTop.style.visibility = show ? "visible" : "hidden";
        }});
    }}

    window.scrollToTop = () =>
        window.scrollTo({{ top: 0, behavior: "smooth" }});

    /* ==================== Tooltips ==================== */

    document
        .querySelectorAll('[data-bs-toggle="tooltip"]')
        .forEach(el => new bootstrap.Tooltip(el, {{ container: "body" }}));

    /* ==================== Chart Data ==================== */

    const labels = {[
        k.replace("_", " ").title()
        for k in rct_detailed_report.baseline_model_summary["elec_by_end_use"].keys()
    ]};

    const elecData = {{
        consumption: {{
            baseline: {list(rct_detailed_report.baseline_model_summary["elec_by_end_use"].values())},
            proposed: {list(rct_detailed_report.proposed_model_summary["elec_by_end_use"].values())},
        }},
        eui: {{
            baseline: {list(rct_detailed_report.baseline_model_summary["elec_by_end_use_eui"].values())},
            proposed: {list(rct_detailed_report.proposed_model_summary["elec_by_end_use_eui"].values())},
        }},
    }};

    const gasData = {{
        consumption: {{
            baseline: {list(rct_detailed_report.baseline_model_summary["gas_by_end_use"].values())},
            proposed: {list(rct_detailed_report.proposed_model_summary["gas_by_end_use"].values())},
        }},
        eui: {{
            baseline: {list(rct_detailed_report.baseline_model_summary["gas_by_end_use_eui"].values())},
            proposed: {list(rct_detailed_report.proposed_model_summary["gas_by_end_use_eui"].values())},
        }},
    }};

    const energyData = {{
        consumption: {{
            baseline: {list(rct_detailed_report.baseline_model_summary["energy_by_end_use"].values())},
            proposed: {list(rct_detailed_report.proposed_model_summary["energy_by_end_use"].values())},
        }},
        eui: {{
            baseline: {list(rct_detailed_report.baseline_model_summary["energy_by_end_use_eui"].values())},
            proposed: {list(rct_detailed_report.proposed_model_summary["energy_by_end_use_eui"].values())},
        }},
    }};

    /* ==================== Chart Factory ==================== */

    function makeChart(canvasId, title, unit, data) {{
        const ctx = $(canvasId);
        if (!ctx) return null;

        return new Chart(ctx, {{
            type: "bar",
            data: {{
                labels,
                datasets: [
                    {{
                        label: "Baseline",
                        data: data.baseline,
                        backgroundColor: "rgba(54,162,235,0.7)"
                    }},
                    {{
                        label: "Proposed",
                        data: data.proposed,
                        backgroundColor: "rgba(75,192,75,0.7)"
                    }}
                ]
            }},
            options: {{
                responsive: true,
                interaction: {{ mode: "index", intersect: false }},
                plugins: {{
                    title: {{ display: true, text: title }},
                    tooltip: {{ mode: "index", intersect: false }}
                }},
                scales: {{
                    y: {{
                        beginAtZero: true,
                        title: {{ display: true, text: unit }}
                    }}
                }}
            }}
        }});
    }}

    /* ==================== Charts ==================== */

    let currentChart = "elec";

    const charts = {{
        elec: makeChart("elecByEndUse", "Electricity By End Use", "kWh", elecData.consumption),
        gas: makeChart("gasByEndUse", "Natural Gas By End Use", "Therms", gasData.consumption),
        energy: makeChart("energyByEndUse", "Total Site Energy By End Use", "kBtu", energyData.consumption),
    }};

    const containers = {{
        elec: $("elecChartContainer"),
        gas: $("gasChartContainer"),
        energy: $("energyChartContainer"),
    }};

    /* ==================== Chart Updates ==================== */

    function updateCharts() {{
        const useEUI = $("unitToggle")?.checked;
        const unitType = useEUI ? "eui" : "consumption";

        const map = {{
            elec: elecData,
            gas: gasData,
            energy: energyData,
        }};

        Object.entries(charts).forEach(([key, chart]) => {{
            if (!chart) return;
            chart.data.datasets[0].data = map[key][unitType].baseline;
            chart.data.datasets[1].data = map[key][unitType].proposed;
            chart.options.scales.y.title.text =
                unitType === "eui" ? "kBtu/ft²" :
                key === "elec" ? "kWh" :
                key === "gas" ? "Therms" : "kBtu";
            chart.update();
        }});

        updateTotals();
    }}

    function updateTotals() {{
        const useEUI = $("unitToggle")?.checked;
        const unitType = useEUI ? "eui" : "consumption";

        let data, unit;
        if (currentChart === "elec") {{
            data = elecData[unitType];
            unit = unitType === "eui" ? "kBtu/ft²" : "kWh";
        }} else if (currentChart === "gas") {{
            data = gasData[unitType];
            unit = unitType === "eui" ? "kBtu/ft²" : "Therms";
        }} else {{
            data = energyData[unitType];
            unit = unitType === "eui" ? "kBtu/ft²" : "kBtu";
        }}

        $("baselineTotal").textContent =
            `Baseline Total: ${{sum(data.baseline).toLocaleString()}} ${{unit}}`;
        $("proposedTotal").textContent =
            `Proposed Total: ${{sum(data.proposed).toLocaleString()}} ${{unit}}`;
    }}

    function showChart(type) {{
        currentChart = type;
        Object.entries(containers).forEach(([k, el]) =>
            el?.classList.toggle("d-none", k !== type)
        );
        updateCharts();
    }}

    /* ==================== Event Wiring ==================== */

    $("btn-elec")?.addEventListener("change", () => showChart("elec"));
    $("btn-gas")?.addEventListener("change", () => showChart("gas"));
    $("btn-energy")?.addEventListener("change", () => showChart("energy"));
    $("unitToggle")?.addEventListener("change", updateCharts);

    /* ==================== Init ==================== */

    showChart("elec");

}});
</script>
"""
    )
