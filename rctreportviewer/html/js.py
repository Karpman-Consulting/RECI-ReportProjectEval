def write_javascript(file, rct_detailed_report):
    file.write(
        f"""
<script>
    const inputs = document.querySelectorAll('#energySourceTable input');
    const rows = document.querySelectorAll('#energySourceTable tbody tr');
    const energyPerformanceInputs = document.querySelectorAll('#energySourcePerformanceTable input');
    const energyPerformanceRows = document.querySelectorAll('#energyPerformanceTable tbody tr');

    const parseNumber = (str) => parseFloat(str.replace(/,/g, "")) || 0;
    const getText = (id) => parseNumber(document.getElementById(id).textContent);
    const setText = (id, value) => {{
        document.getElementById(id).textContent = value.toLocaleString();
    }};
    const setRatio = (id, numerator, denominator) => {{
        const ratio = denominator !== 0 ? (numerator / denominator).toFixed(2) : "0.00";
        document.getElementById(id).textContent = ratio;
    }};

    function recalculateEnergyMetrics() {{
        let proposedSourceEnergy = 0;
        let proposedGHGEmissions = 0;
        let baselineUnregulatedSourceEnergy = 0;
        let baselineUnregulatedGHGEmissions = 0;
        let baselineRegulatedSourceEnergy = 0;
        let baselineRegulatedGHGEmissions = 0;

        rows.forEach(row => {{
            const getVal = (cls) => parseNumber(row.querySelector(`.${{cls}}`)?.textContent || "0");
            const getInputVal = (cls) => parseNumber(row.querySelector(`.${{cls}}`)?.value || "0");

            const proposed = getVal('proposedEnergyUse');
            const unreg = getVal('baselineUnregulatedEnergy');
            const reg = getVal('baselineRegulatedEnergy');
            const ssr = getInputVal('siteSourceRatio');
            const ghg = getInputVal('ghgEmissionFactor');

            proposedSourceEnergy += proposed * ssr;
            proposedGHGEmissions += proposed * ghg;
            baselineUnregulatedSourceEnergy += unreg * ssr;
            baselineUnregulatedGHGEmissions += unreg * ghg;
            baselineRegulatedSourceEnergy += reg * ssr;
            baselineRegulatedGHGEmissions += reg * ghg;
        }});

        const baselineSourceEnergy = baselineUnregulatedSourceEnergy + baselineRegulatedSourceEnergy;
        const baselineGHGEmissions = baselineUnregulatedGHGEmissions + baselineRegulatedGHGEmissions;

        const proposedSiteEnergy = getText('pbp_nre_site_energy') - getText('proposed_site_energy_savings');
        const proposedSrcEnergy = proposedSourceEnergy - getText('proposed_source_energy_savings');
        const proposedGHG = proposedGHGEmissions - getText('proposed_ghg_savings');

        const baselineSiteEnergy = getText('bbp_site_energy');
        const baselineUnregulatedSiteEnergy = getText('bbuec_site_energy');
        const baselineRegulatedSiteEnergy = getText('bbrec_site_energy');

        const bpfSite = getText('bpf_site_energy');
        const bpfSource = getText('bpf_source_energy');
        const bpfGHG = getText('bpf_ghg_emissions');

        setText('pbp_nre_source_energy', proposedSourceEnergy.toLocaleString(undefined, {{ minimumFractionDigits: 0, maximumFractionDigits: 0 }}));
        setText('pbp_nre_ghg', proposedGHGEmissions.toLocaleString(undefined, {{ minimumFractionDigits: 0, maximumFractionDigits: 0 }}));
        setText('pbp_site_energy', proposedSiteEnergy.toLocaleString(undefined, {{ minimumFractionDigits: 0, maximumFractionDigits: 0 }}));
        setText('pbp_source_energy', proposedSrcEnergy.toLocaleString(undefined, {{ minimumFractionDigits: 0, maximumFractionDigits: 0 }}));
        setText('pbp_ghg', proposedGHG.toLocaleString(undefined, {{ minimumFractionDigits: 0, maximumFractionDigits: 0 }}));
        
        setText('bbuec_source_energy', baselineUnregulatedSourceEnergy.toLocaleString(undefined, {{ minimumFractionDigits: 0, maximumFractionDigits: 0 }}));
        setText('bbuec_ghg', baselineUnregulatedGHGEmissions.toLocaleString(undefined, {{ minimumFractionDigits: 0, maximumFractionDigits: 0 }}));
        setText('bbrec_source_energy', baselineRegulatedSourceEnergy.toLocaleString(undefined, {{ minimumFractionDigits: 0, maximumFractionDigits: 0 }}));
        setText('bbrec_ghg', baselineRegulatedGHGEmissions.toLocaleString(undefined, {{ minimumFractionDigits: 0, maximumFractionDigits: 0 }}));
        setText('bbp_source_energy', baselineSourceEnergy.toLocaleString(undefined, {{ minimumFractionDigits: 0, maximumFractionDigits: 0 }}));
        setText('bbp_ghg', baselineGHGEmissions.toLocaleString(undefined, {{ minimumFractionDigits: 0, maximumFractionDigits: 0 }}));

        // PCIt ratios
        setRatio('pcit_site_energy', baselineUnregulatedSiteEnergy + bpfSite * baselineRegulatedSiteEnergy, baselineSiteEnergy);
        setRatio('pcit_source_energy', baselineUnregulatedSourceEnergy + bpfSource * baselineRegulatedSourceEnergy, baselineSourceEnergy);
        setRatio('pcit_ghg_emissions', baselineUnregulatedGHGEmissions + bpfGHG * baselineRegulatedGHGEmissions, baselineGHGEmissions);

        setRatio('pci_nre_site_energy', getText('pbp_nre_site_energy'), baselineSiteEnergy);
        setRatio('pci_nre_source_energy', proposedSourceEnergy, baselineSourceEnergy);
        setRatio('pci_nre_ghg', proposedGHGEmissions, baselineGHGEmissions);

        // PCIadjusted calculations
        const capFraction = 0.05;
        const adjustedSiteSavings = Math.min(getText('proposed_site_energy_savings'), capFraction * baselineSiteEnergy);
        const adjustedSourceSavings = Math.min(getText('proposed_source_energy_savings'), capFraction * baselineSourceEnergy);
        const adjustedGHGSavings = Math.min(getText('proposed_ghg_savings'), capFraction * baselineGHGEmissions);

        const adjustedPBPSite = getText('pbp_nre_site_energy') - adjustedSiteSavings;
        const adjustedPBPSource = proposedSourceEnergy - adjustedSourceSavings;
        const adjustedPBPGHG = proposedGHGEmissions - adjustedGHGSavings;

        const pciAdjustedSite = adjustedPBPSite / baselineSiteEnergy;
        const pciAdjustedSource = adjustedPBPSource / baselineSourceEnergy;
        const pciAdjustedGHG = adjustedPBPGHG / baselineGHGEmissions;

        // Update the table
        setText('pci_adjusted_site_energy', pciAdjustedSite.toLocaleString(undefined, {{ minimumFractionDigits: 2, maximumFractionDigits: 2 }}));
        setText('pci_adjusted_source_energy', pciAdjustedSource.toLocaleString(undefined, {{ minimumFractionDigits: 2, maximumFractionDigits: 2 }}));
        setText('pci_adjusted_ghg', pciAdjustedGHG.toLocaleString(undefined, {{ minimumFractionDigits: 2, maximumFractionDigits: 2 }}));

        const getCost = (id) => parseNumber(document.getElementById(id).textContent.replace(/[$,]/g, ''));

        // Get cost values
        const baselineCost = getCost('bbp_cost');
        const proposedCost = getCost('pbp_cost');
        const proposedNRECost = getCost('pbp_nre_cost');

        // % Improvement excluding renewables
        const cost_savings_nre = ((baselineCost - proposedNRECost) / baselineCost) * 100;
        const site_savings_nre = ((baselineSiteEnergy - getText('pbp_nre_site_energy')) / baselineSiteEnergy) * 100;
        const source_savings_nre = ((baselineSourceEnergy - proposedSourceEnergy) / baselineSourceEnergy) * 100;
        const ghg_savings_nre = ((baselineGHGEmissions - proposedGHGEmissions) / baselineGHGEmissions) * 100;

        setText('cost_savings_nre', cost_savings_nre.toFixed(1) + '%');
        setText('site_savings_nre', site_savings_nre.toFixed(1) + '%');
        setText('source_savings_nre', source_savings_nre.toFixed(1) + '%');
        setText('ghg_savings_nre', ghg_savings_nre.toFixed(1) + '%');

        // % Improvement including renewables
        const cost_savings = ((baselineCost - proposedCost) / baselineCost) * 100;
        const site_savings = ((baselineSiteEnergy - getText('pbp_site_energy')) / baselineSiteEnergy) * 100;
        const source_savings = ((baselineSourceEnergy - proposedSrcEnergy) / baselineSourceEnergy) * 100;
        const ghg_savings = ((baselineGHGEmissions - proposedGHG) / baselineGHGEmissions) * 100;

        setText('cost_savings', cost_savings.toFixed(1) + '%');
        setText('site_savings', site_savings.toFixed(1) + '%');
        setText('source_savings', source_savings.toFixed(1) + '%');
        setText('ghg_savings', ghg_savings.toFixed(1) + '%');
    }}

    function recalculateEnergyPerformanceMetrics() {{
        let totProposedSiteEnergy = 0;
        let totBaselineSiteEnergy = 0;
        let totProposedSourceEnergy = 0;
        let totBaselineSourceEnergy = 0;
        let totProposedCost = 0;
        let totBaselineCost = 0;
        let totProposedGHGEmissions = 0;
        let totBaselineGHGEmissions = 0;

        const electricitySiteSourceRatio = parseFloat(document.querySelector('.electricitySiteSourceRatio').value || 2.80);
        const naturalGasSiteSourceRatio = parseFloat(document.querySelector('.naturalGasSiteSourceRatio').value || 1.05);
        const electricityGHGEmissionFactor = parseFloat(document.querySelector('.electricityGhgEmissionFactor').value || 0.37);
        const naturalGasGHGEmissionFactor = parseFloat(document.querySelector('.naturalGasGhgEmissionFactor').value || 0.53);
        energyPerformanceRows.forEach(row => {{
            const getRowText = (id) => parseFloat(row.getElementsByClassName(id)[0].textContent.replace(/[^0-9.\-]/g, ''));
            const setRowText = (id, value, percentile, currency) => {{
                if (percentile) {{
                    row.getElementsByClassName(id)[0].textContent = value.toFixed(1).toLocaleString() + "%";
                }}
                else if (currency) {{
                    row.getElementsByClassName(id)[0].textContent = value.toLocaleString('en-US', {{
                        style: 'currency',
                        currency: 'USD',
                        minimumFractionDigits: 0,
                        maximumFractionDigits: 0
                    }});
                }}
                else {{
                    row.getElementsByClassName(id)[0].textContent = value.toFixed(1).toLocaleString();
                }} 
            }};
            if (row === energyPerformanceRows[energyPerformanceRows.length - 1]){{
                let totSiteSavings = totBaselineSiteEnergy ? (totBaselineSiteEnergy - totProposedSiteEnergy) / totBaselineSiteEnergy * 100 : (0 - totProposedSiteEnergy) * 100;
                let totSourceSavings = totBaselineSourceEnergy ? (totBaselineSourceEnergy - totProposedSourceEnergy) / totBaselineSourceEnergy * 100 : (0 - totProposedSourceEnergy) * 100;
                let totCostSavings = totBaselineCost ? (totBaselineCost - totProposedCost) / totBaselineCost * 100 : (0 - totProposedCost) * 100;
                let totGHGSavings = totBaselineGHGEmissions ? (totBaselineGHGEmissions - totProposedGHGEmissions) / totBaselineGHGEmissions * 100 : (0 - totProposedGHGEmissions) * 100;

                setRowText('totSiteEnergyProposed', totProposedSiteEnergy);
                setRowText('totSiteEnergyBaseline', totBaselineSiteEnergy);
                setRowText('totSiteEnergySavings', totSiteSavings, true);
                setRowText('totSourceEnergyProposed', totProposedSourceEnergy);
                setRowText('totSourceEnergyBaseline', totBaselineSourceEnergy);
                setRowText('totSourceEnergySavings', totSourceSavings, true);
                setRowText('totCostProposed', totProposedCost, false, true);
                setRowText('totCostBaseline', totBaselineCost, false, true);
                setRowText('totCostSavings', totCostSavings, true);
                setRowText('totGhgEmissionsProposed', totProposedGHGEmissions);
                setRowText('totGhgEmissionsBaseline', totBaselineGHGEmissions);
                setRowText('totGhgEmissionsSavings', totGHGSavings, true);
                return;
            }};
            
            const proposedSiteEnergy = getRowText('siteEnergyProposed');
            const baselineSiteEnergy = getRowText('siteEnergyBaseline');
            const proposedCost = getRowText('energyCostProposed');
            const baselineCost = getRowText('energyCostBaseline');

            let proposedSourceEnergy = ((getRowText('electricityProposed') * electricitySiteSourceRatio) + (getRowText('naturalGasProposed') * naturalGasSiteSourceRatio));
            let baselineSourceEnergy = ((getRowText('electricityBaseline') * electricitySiteSourceRatio) + (getRowText('naturalGasBaseline') * naturalGasSiteSourceRatio));
            let sourceEnergySavings = baselineSourceEnergy ? ((baselineSourceEnergy - proposedSourceEnergy) / baselineSourceEnergy * 100) : (0 - proposedSourceEnergy) * 100;
            
            setRowText('sourceEnergyProposed', proposedSourceEnergy);
            setRowText('sourceEnergyBaseline', baselineSourceEnergy);
            setRowText('sourceEnergySavings', sourceEnergySavings);

            let proposedGHGEmissions = ((getRowText('electricityProposed') * electricityGHGEmissionFactor) + (getRowText('naturalGasProposed') * naturalGasGHGEmissionFactor));
            let baselineGHGEmissions = ((getRowText('electricityBaseline') * electricityGHGEmissionFactor) + (getRowText('naturalGasBaseline') * naturalGasGHGEmissionFactor));
            let ghgSavings = baselineGHGEmissions ? ((baselineGHGEmissions - proposedGHGEmissions) / baselineGHGEmissions * 100) : (0 - proposedGHGEmissions) * 100;
            
            setRowText('ghgEmissionsProposed', proposedGHGEmissions);
            setRowText('ghgEmissionsBaseline', baselineGHGEmissions);
            setRowText('ghgEmissionsSavings', ghgSavings);

            totProposedSiteEnergy += proposedSiteEnergy;
            totBaselineSiteEnergy += baselineSiteEnergy;
            totProposedSourceEnergy += proposedSourceEnergy;
            totBaselineSourceEnergy += baselineSourceEnergy;
            totProposedCost += proposedCost;
            totBaselineCost += baselineCost;
            totProposedGHGEmissions += proposedGHGEmissions;
            totBaselineGHGEmissions += baselineGHGEmissions;
        }});
    }}

    function toggleBackToTopButton() {{
        const backToTopButton = document.getElementById("back-to-top");
        if (document.body.scrollTop > 100 || document.documentElement.scrollTop > 100) {{
            backToTopButton.style.opacity = "1";
            backToTopButton.style.visibility = "visible";
        }}  
        else {{
            backToTopButton.style.opacity = "0";
            backToTopButton.style.visibility = "hidden";
        }}
    }}

    function scrollToTop() {{
        window.scrollTo({{
            top: 0,
            behavior: 'smooth'
        }});
    }}

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
                        let cleanedText = td.textContent.replace(/,/g, "").trim();
                        let value = parseFloat(cleanedText) || 0;
                        let decimalPlaces = (cleanedText.split(".")[1] || "").length;
                        columnPrecisions[colIndex] = Math.max(columnPrecisions[colIndex] || 0, decimalPlaces);
                        columnSums[colIndex] = (columnSums[colIndex] || 0) + value;
                    }});
                }}
            }});
        }});
    }}

    window.onscroll = function() {{
        toggleBackToTopButton();
    }};

    inputs.forEach(input => {{
        input.addEventListener('input', recalculateEnergyMetrics);
    }});

    energyPerformanceInputs.forEach(input => {{
        input.addEventListener('input', recalculateEnergyPerformanceMetrics);
    }});

    document.addEventListener("DOMContentLoaded", () => {{
        const tooltipTriggerList = document.querySelectorAll('[data-bs-toggle="tooltip"]');
        const tooltipList = [...tooltipTriggerList].map(tooltipTriggerEl =>
          new bootstrap.Tooltip(tooltipTriggerEl, {{
            container: 'body',
          }})
        );

        if (inputs.length > 0) {{
            recalculateEnergyMetrics();
        }}

        if (energyPerformanceInputs.length > 0) {{
            recalculateEnergyPerformanceMetrics();
        }}

        calculateSubtotals();

        // Chart labels
        const labels = {[label.replace('_', ' ').title() for label in rct_detailed_report.baseline_model_summary["elec_by_end_use"].keys()]};

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

        // Electricity Datasets
        const elecData = {{
            labels: labels,
            datasets: [
                {{
                    label: 'Baseline',
                    data: {list(rct_detailed_report.baseline_model_summary["elec_by_end_use"].values())},
                    backgroundColor: 'rgba(54, 162, 235, 0.7)'
                }},
                {{
                    label: 'Proposed',
                    data: {list(rct_detailed_report.proposed_model_summary["elec_by_end_use"].values())},
                    backgroundColor: 'rgba(75, 192, 75, 0.7)'
                }}
            ]
        }};

        const elecConfig = {{
            type: 'bar',
            data: elecData,
            options: {{
                responsive: true,
                plugins: {{
                    title: {{
                        display: true,
                        text: 'Electricity By End Use'
                    }},
                    tooltip: {{
                        mode: 'index',
                        intersect: false
                    }}
                }},
                interaction: {{
                    mode: 'index',
                    intersect: false
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
                        title: {{
                            display: true,
                            text: 'kWh',
                            font: {{
                                size: 14
                            }}
                        }}
                    }}
                }}
            }}
        }};

        const gasData = {{
            labels: labels,
            datasets: [
                {{
                    label: 'Baseline',
                    data: {list(rct_detailed_report.baseline_model_summary["gas_by_end_use"].values())},
                    backgroundColor: 'rgba(255, 180, 80, 0.5)'
                }},
                {{
                    label: 'Proposed',
                    data: {list(rct_detailed_report.proposed_model_summary["gas_by_end_use"].values())},
                    backgroundColor: 'rgba(255, 100, 100, 0.5)'
                }}
            ]
        }};

        const gasConfig = {{
            type: 'bar',
            data: gasData,
            options: {{
                responsive: true,
                plugins: {{
                    title: {{
                        display: true,
                        text: 'Natural Gas By End Use'
                    }},
                    tooltip: {{
                        mode: 'index',
                        intersect: false
                    }}
                }},
                interaction: {{
                    mode: 'index',
                    intersect: false
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
                        title: {{
                            display: true,
                            text: 'Therms',
                            font: {{
                                size: 14
                            }}
                        }}
                    }}
                }}
            }}
        }};

        const energyData = {{
            labels: labels,
            datasets: [
                {{
                    label: 'Baseline',
                    data: {list(rct_detailed_report.baseline_model_summary["energy_by_end_use"].values())},
                    backgroundColor: 'rgba(128, 0, 64, 0.6)'
                }},
                {{
                    label: 'Proposed',
                    data: {list(rct_detailed_report.proposed_model_summary["energy_by_end_use"].values())},
                    backgroundColor: 'rgba(0, 128, 128, 0.6)'
                }}
            ]
        }};

        const energyConfig = {{
            type: 'bar',
            data: energyData,
            options: {{
                responsive: true,
                plugins: {{
                    title: {{
                        display: true,
                        text: 'Total Site Energy By End Use'
                    }},
                    tooltip: {{
                        mode: 'index',
                        intersect: false
                    }}
                }},
                interaction: {{
                    mode: 'index',
                    intersect: false
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
                        title: {{
                            display: true,
                            text: 'kBtu',
                            font: {{
                                size: 14
                            }}
                        }}
                    }}
                }}
            }}
        }};

        const elecChart = new Chart(document.getElementById('elecByEndUse'), elecConfig);
        const gasChart = new Chart(document.getElementById('gasByEndUse'), gasConfig);
        const energyChart = new Chart(document.getElementById('energyByEndUse'), energyConfig);

        function updateCharts(unitType) {{
          // Update Electricity
          elecChart.data.datasets[0].data = elecDataRaw[unitType].baseline;
          elecChart.data.datasets[1].data = elecDataRaw[unitType].proposed;
          elecChart.options.scales.y.title.text = unitType === 'consumption' ? 'kWh' : 'kBtu/ft²';
          elecChart.update();

          // Update Gas
          gasChart.data.datasets[0].data = gasDataRaw[unitType].baseline;
          gasChart.data.datasets[1].data = gasDataRaw[unitType].proposed;
          gasChart.options.scales.y.title.text = unitType === 'consumption' ? 'Therms' : 'kBtu/ft²';
          gasChart.update();

          // Update Total Energy
          energyChart.data.datasets[0].data = energyDataRaw[unitType].baseline;
          energyChart.data.datasets[1].data = energyDataRaw[unitType].proposed;
          energyChart.options.scales.y.title.text = unitType === 'consumption' ? 'kBtu' : 'kBtu/ft²';
          energyChart.update();
        }}

        function sumArray(arr) {{
          return arr.reduce((acc, val) => acc + val, 0);
        }}

        function getUnitLabel(source, unitType) {{
          if (unitType === 'eui') {{
            return 'kBtu/ft²';
          }} else {{
            return source === 'elec' ? 'kWh' : source === 'gas' ? 'Therms' : 'kBtu';
          }}
        }}

        function updateTotalColors(source) {{
          const baselineEl = document.getElementById('baselineTotal');
          const proposedEl = document.getElementById('proposedTotal');

          if (source === 'elec') {{
            baselineEl.style.color = 'rgb(54, 162, 235)'; // Blue
            proposedEl.style.color = 'rgb(75, 192, 75)';  // Green
          }} else if (source === 'gas') {{
            baselineEl.style.color = 'rgb(255, 180, 80)'; // Orange
            proposedEl.style.color = 'rgb(255, 100, 100)'; // Red
          }} else if (source === 'energy') {{
              baselineEl.style.color = 'rgb(128, 0, 64)';   // Maroon
              proposedEl.style.color = 'rgb(0, 128, 128)';  // Teal
            }}
        }}

        function updateTotals(source, unitType) {{
          let baseline, proposed;

          if (source === 'elec') {{
            baseline = unitType === 'eui'
              ? {list(rct_detailed_report.baseline_model_summary["elec_by_end_use_eui"].values())}
              : {list(rct_detailed_report.baseline_model_summary["elec_by_end_use"].values())};

            proposed = unitType === 'eui'
              ? {list(rct_detailed_report.proposed_model_summary["elec_by_end_use_eui"].values())}
              : {list(rct_detailed_report.proposed_model_summary["elec_by_end_use"].values())};

          }} else if (source === 'gas') {{
            baseline = unitType === 'eui'
              ? {list(rct_detailed_report.baseline_model_summary["gas_by_end_use_eui"].values())}
              : {list(rct_detailed_report.baseline_model_summary["gas_by_end_use"].values())};

            proposed = unitType === 'eui'
              ? {list(rct_detailed_report.proposed_model_summary["gas_by_end_use_eui"].values())}
              : {list(rct_detailed_report.proposed_model_summary["gas_by_end_use"].values())};

          }} else if (source === 'energy') {{
            baseline = unitType === 'eui'
              ? {list(rct_detailed_report.baseline_model_summary["energy_by_end_use_eui"].values())}
              : {list(rct_detailed_report.baseline_model_summary["energy_by_end_use"].values())};

            proposed = unitType === 'eui'
              ? {list(rct_detailed_report.proposed_model_summary["energy_by_end_use_eui"].values())}
              : {list(rct_detailed_report.proposed_model_summary["energy_by_end_use"].values())};
          }}

          const unit = getUnitLabel(source, unitType);
          const baselineSum = sumArray(baseline).toLocaleString(undefined, {{ maximumFractionDigits: 0 }});
          const proposedSum = sumArray(proposed).toLocaleString(undefined, {{ maximumFractionDigits: 0 }});

          document.getElementById('baselineTotal').textContent = `Baseline Total: ${{baselineSum}} ${{unit}}`;
          document.getElementById('proposedTotal').textContent = `Proposed Total: ${{proposedSum}} ${{unit}}`;
        }}

        let currentChart = 'elec';

        window.toggleUnits = function() {{
          const useEUI = document.getElementById('unitToggle').checked;
          const unitType = useEUI ? 'eui' : 'consumption';
          updateCharts(unitType);
          updateTotals(currentChart, unitType);
        }};

        window.showChart = function(type) {{
          const elecContainer = document.getElementById('elecChartContainer');
          const gasContainer = document.getElementById('gasChartContainer');
          const energyContainer = document.getElementById('energyChartContainer');
          elecContainer.style.display = type === 'elec' ? 'block' : 'none';
          gasContainer.style.display = type === 'gas' ? 'block' : 'none';
          energyContainer.style.display = type === 'energy' ? 'block' : 'none';
          currentChart = type;
          const useEUI = document.getElementById('unitToggle').checked;
          const unitType = useEUI ? 'eui' : 'consumption';
          updateTotals(type, unitType);
          updateTotalColors(type);
        }};

        // Initial total update
        updateTotals(currentChart, 'consumption');
        updateTotalColors(currentChart);

    }});
</script>
    """
    )
