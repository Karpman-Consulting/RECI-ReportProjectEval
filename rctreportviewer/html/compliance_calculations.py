def write_compliance_calculations(file, rct_detailed_report):
    # ----------------------- Compliance Calculations -----------------------
    file.write(
        f"""
        <div class="mb-3 me-4">
            <button class="btn btn-info collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapse-compliance-calcs" aria-expanded="false">
                Compliance Calculations
            </button>

            <div id="collapse-compliance-calcs" class="accordion-collapse collapse">
                <div class="accordion-body">
                    <h3>Compliance Calculations</h3>
                    <table class="table table-sm table-borderless" style="width: 900px;" id="energySourceTable">
                        <thead>
                            <tr class="text-center">
                                <th style="border: 2px solid black; width: 100px;">Energy Source</th>
                                <th style="border: 2px solid black;">Baseline Unregulated Energy<br>(MMBtu)</th>
                                <th style="border: 2px solid black;">Baseline Regulated Energy<br>(MMBtu)</th>
                                <th style="border: 2px solid black;">Total Baseline Energy<br>(MMBtu)</th>
                                <th style="border: 2px solid black;">Total Proposed Energy<br>(MMBtu)</th>
                                <th style="border: 2px solid black; width: 110px;">Source-Site<br>Ratio</th>
                                <th style="border: 2px solid black;">GHG Emission Factor<br>(Metric Ton CO<sub>2</sub>/MMBtu)</th>
                            </tr>
                        </thead>
                        <tbody style="border: 2px solid black;">
    """
    )

    row_number = 0
    for (
        energy_source,
        proposed_energy_use,
    ) in rct_detailed_report.proposed_model_summary.get(
        "energy_by_fuel_type", {}
    ).items():
        baseline_energy_use = rct_detailed_report.baseline_model_summary.get(
            "energy_by_fuel_type", {}
        ).get(energy_source, 0)
        baseline_unregulated_energy = (
            rct_detailed_report.baseline_model_summary.get(
                "compliance_calcs_by_parameter", {}
            )
            .get("bbuec", {})
            .get(energy_source, 0)
        )
        baseline_regulated_energy = (
            rct_detailed_report.baseline_model_summary.get(
                "compliance_calcs_by_parameter", {}
            )
            .get("bbrec", {})
            .get(energy_source, 0)
        )

        if energy_source == "ELECTRICITY":
            file.write(
                f"""
                            <tr style="font-size: 12px;" class="lh-1 text-center">
                                <td class="align-middle">Electricity</td>
                                <td class="align-middle baselineUnregulatedEnergy">{round(baseline_unregulated_energy, 1):,}</td>
                                <td class="align-middle baselineRegulatedEnergy">{round(baseline_regulated_energy, 1):,}</td>
                                <td class="align-middle baselineEnergyUse">{round(baseline_energy_use, 1):,}</td>
                                <td class="align-middle proposedEnergyUse">{round(proposed_energy_use, 1):,}</td>
                                <td><input type="number" class="siteSourceRatio" value="2.80" step="0.01" style="width: 60px;"></td>
                                <td><input type="number" class="ghgEmissionFactor" value="0.037" step="0.001" style="width: 60px;"></td>
                            </tr>
            """
            )

        elif energy_source == "NATURAL_GAS":
            file.write(
                f"""
                            <tr style="font-size: 12px;" class="lh-1 text-center">
                                <td class="align-middle">Natural Gas</td>
                                <td class="align-middle baselineUnregulatedEnergy">{round(baseline_unregulated_energy, 1):,}</td>
                                <td class="align-middle baselineRegulatedEnergy">{round(baseline_regulated_energy, 1):,}</td>
                                <td class="align-middle baselineEnergyUse">{round(baseline_energy_use, 1):,}</td>
                                <td class="align-middle proposedEnergyUse">{round(proposed_energy_use, 1):,}</td>
                                <td><input type="number" class="siteSourceRatio" value="1.05" step="0.01" style="width: 60px;"></td>
                                <td><input type="number" class="ghgEmissionFactor" value="0.053" step="0.001" style="width: 60px;"></td>
                            </tr>
            """
            )

        else:
            file.write(
                f"""
                            <tr style="font-size: 12px;" class="lh-1 text-center">
                                    <td class="align-middle">{energy_source}</td>
                                    <td class="align-middle baselineUnregulatedEnergy">{round(baseline_unregulated_energy, 1):,}</td>
                                    <td class="align-middle baselineRegulatedEnergy">{round(baseline_regulated_energy, 1):,}</td>
                                    <td class="align-middle baselineEnergyUse">{round(baseline_energy_use, 1):,}</td>
                                    <td class="align-middle proposedEnergyUse">{round(proposed_energy_use, 1):,}</td>
                                    <td><input type="number" class="siteSourceRatio" value="0.0" step="0.01" style="width: 60px;"></td>
                                    <td><input type="number" class="ghgEmissionFactor" value="0.0" step="0.001" style="width: 60px;"></td>
                            </tr>
            """
            )

        row_number += 1

    output = rct_detailed_report.rpd_data.get("output", {})
    baseline_compliance_calcs = rct_detailed_report.baseline_model_summary.get(
        "compliance_calcs_by_parameter", {}
    )
    proposed_compliance_calcs = rct_detailed_report.proposed_model_summary.get(
        "compliance_calcs_by_parameter", {}
    )

    file.write(
        f"""
                        </tbody>
                    </table>
                    
                    <table class="table table-sm table-borderless" style="width: 1300px;" id="complianceCalcsTable">
                        <thead>
                            <tr class="text-center">
                                <th colspan="2" class="col-4"></th>
                                <th colspan="4" class="col-4" style="border: 2px solid black;">Performance Metric</th>
                            </tr>
                            <tr class="text-center">
                                <th style="border: 2px solid black;">Parameter</th>
                                <th style="border: 2px solid black;">Symbol</th>
                                <th style="border: 2px solid black;">Cost ($)</th>
                                <th style="border: 2px solid black;">Site Energy (MMBtu)</th>
                                <th style="border: 2px solid black;">Source Energy (MMBtu)</th>
                                <th style="border: 2px solid black;">GHG Emissions (Mt CO<sub>2</sub>e)</th>
                            </tr>
                        </thead>
                        <tbody style="border: 2px solid black;">
                            <tr style="font-size: 12px;" class="lh-1 text-center">
                                <td style="border-right: 2px solid black;">Proposed building performance before site-generated renewable energy</td>
                                <td style="border-right: 2px solid black; font-weight: bold;">PBP<sub>nre</sub></td>
                                <td id="pbp_nre_cost">${round(output.get("total_proposed_building_energy_cost_excluding_renewable_energy", 0)):,}</td>
                                <td id="pbp_nre_site_energy">{round(proposed_compliance_calcs.get("pbp_nre", {}).get("site_energy", 0)):,}</td>
                                <td id="pbp_nre_source_energy">-</td>
                                <td id="pbp_nre_ghg">-</td>
                            </tr>
                            <tr style="font-size: 12px;" class="lh-1 text-center">
                                <td style="border-right: 2px solid black;">Proposed design on-site renewable savings</td>
                                <td style="border-right: 2px solid black; font-weight: bold;">-</td>
                                <td id="proposed_cost_savings">$0</td>
                                <td id="proposed_site_energy_savings">0</td>
                                <td id="proposed_source_energy_savings">0</td>
                                <td id="proposed_ghg_savings">0</td>
                            </tr>
    """
    )

    # if "ASHRAE 90.1-2022" in rct_detailed_report.ruleset:
    #     file.write(
    #         """
    #                         <tr style="font-size: 12px;" class="lh-1 text-center">
    #                             <td style="border-right: 2px solid black;">Prescriptive renewable savings</td>
    #                             <td style="border-right: 2px solid black; font-weight: bold;">PRE</td>
    #                             <td>$0</td>
    #                             <td>0</td>
    #                             <td>0</td>
    #                             <td>0</td>
    #                         </tr>
    #     """
    #     )

    file.write(
        f"""
                            <tr style="font-size: 12px;" class="lh-1 text-center">
                                <td style="border-right: 2px solid black;">Proposed building performance including on-site renewable energy</td>
                                <td style="border-right: 2px solid black; font-weight: bold;">PBP</td>
                                <td id="pbp_cost">${round(output.get("total_proposed_building_energy_cost_including_renewable_energy", 0)):,}</td>
                                <td id="pbp_site_energy">-</td>
                                <td id="pbp_source_energy">-</td>
                                <td id="pbp_ghg">-</td>
                            </tr>
                            <tr style="font-size: 12px;" class="lh-1 text-center">
                                <td style="border-right: 2px solid black;">Baseline building unregulated energy, GHG emissions, and/or energy cost</td>
                                <td style="border-right: 2px solid black; font-weight: bold;">BBUEC</td>
                                <td id="bbuec_cost">${round(output.get("baseline_building_unregulated_energy_cost", 0)):,}</td>
                                <td id="bbuec_site_energy">{round(baseline_compliance_calcs.get("bbuec", {}).get("site_energy", 0)):,}</td>
                                <td id="bbuec_source_energy">-</td>
                                <td id="bbuec_ghg">-</td>
                            </tr>
                            <tr style="font-size: 12px;" class="lh-1 text-center">
                                <td style="border-right: 2px solid black;">Baseline building regulated energy, GHG memissions, and/or energy cost</td>
                                <td style="border-right: 2px solid black; font-weight: bold;">BBREC</td>
                                <td id="bbrec_cost">${round(output.get("baseline_building_regulated_energy_cost", 0)):,}</td>
                                <td id="bbrec_site_energy">{round(baseline_compliance_calcs.get("bbrec", {}).get("site_energy", 0)):,}</td>
                                <td id="bbrec_source_energy"></td>
                                <td id="bbrec_ghg">-</td>
                            </tr>
                            <tr style="font-size: 12px;" class="lh-1 text-center">
                                <td style="border-right: 2px solid black;">Baseline buidling performance</td>
                                <td style="border-right: 2px solid black; font-weight: bold;">BBP</td>
                                <td id="bbp_cost">${round(output.get("baseline_building_performance_energy_cost", 0)):,}</td>
                                <td id="bbp_site_energy">{round(baseline_compliance_calcs.get("bbp", {}).get("site_energy", 0)):,}</td>
                                <td id="bbp_source_energy">-</td>
                                <td id="bbp_ghg">-</td>
                            </tr>
                            <tr style="font-size: 12px;" class="lh-1 text-center">
                                <td style="border-right: 2px solid black;">Building Performance Factor</td>
                                <td style="border-right: 2px solid black; font-weight: bold;">BPF</td>
                                <td id="bpf_energy_cost">{round(output.get("total_area_weighted_building_performance_factor", rct_detailed_report.bpfs_by_metric.get("Cost", 0)), 2)}</td>
                                <td id="bpf_site_energy">{round(rct_detailed_report.bpfs_by_metric.get("Site Energy", 0), 2)}</td>
                                <td id="bpf_source_energy">{round(rct_detailed_report.bpfs_by_metric.get("Source Energy", 0), 2)}</td>
                                <td id="bpf_ghg_emissions">{round(rct_detailed_report.bpfs_by_metric.get("GHG Emissions", 0), 2)}</td>
                            </tr>
                            <tr style="font-size: 12px;" class="lh-1 text-center">
                                <td style="border-right: 2px solid black;">Performance Index Target</td>
                                <td style="border-right: 2px solid black; font-weight: bold;">PCI<sub>t</sub></td>
                                <td id="pcit_cost">{round(output.get("performance_cost_index_target", 0), 2):,}</td>
                                <td id="pcit_site_energy">-</td>
                                <td id="pcit_source_energy">-</td>
                                <td id="pcit_ghg_emissions">-</td>
                            </tr>
                            <tr style="font-size: 12px;" class="lh-1 text-center">
                                <td style="border-right: 2px solid black;">Performance index without on-site renewable energy</td>
                                <td style="border-right: 2px solid black; font-weight: bold;">PCI<sub>nre</sub></td>
                                <td id="pci_nre_cost">{round(output.get("total_proposed_building_energy_cost_excluding_renewable_energy", 0) / output.get("baseline_building_performance_energy_cost", 0.0000001), 2):,}</td>
                                <td id="pci_nre_site_energy">-</td>
                                <td id="pci_nre_source_energy">-</td>
                                <td id="pci_nre_ghg">-</td>
                            </tr>
                            <tr style="font-size: 12px;" class="lh-1 text-center">
                                <td style="border-right: 2px solid black;">Performance Index adjusted based upon ASHRAE 90.1-2019 Section 4.2.1.1</td>
                                <td style="border-right: 2px solid black; font-weight: bold;">PCI</td>
                                <td id="pci_adjusted_cost">{round(output.get("performance_cost_index", 0), 2):,}</td>
                                <td id="pci_adjusted_site_energy">-</td>
                                <td id="pci_adjusted_source_energy">-</td>
                                <td id="pci_adjusted_ghg">-</td>
                            </tr>
                            <tr style="font-size: 12px;" class="lh-1 text-center">
                                <td style="border-right: 2px solid black; font-weight: bold;">% improvement beyond ASHRAE 90.1-2019, excluding proposed design on-site renewable energy</td>
                                <td style="border-right: 2px solid black; font-weight: bold;">-</td>
                                <td id="cost_savings_nre" style="font-weight: bold;">-</td>
                                <td id="site_savings_nre" style="font-weight: bold;">-</td>
                                <td id="source_savings_nre" style="font-weight: bold;">-</td>
                                <td id="ghg_savings_nre" style="font-weight: bold;">-</td>
                            </tr>
                            <tr style="font-size: 12px;" class="lh-1 text-center">
                                <td style="border-right: 2px solid black; font-weight: bold;">% improvement beyond ASHRAE 90.1-2019, including proposed design on-site renewable energy</td>
                                <td style="border-right: 2px solid black; font-weight: bold;">-</td>
                                <td id="cost_savings" style="font-weight: bold;">-</td>
                                <td id="site_savings" style="font-weight: bold;">-</td>
                                <td id="source_savings" style="font-weight: bold;">-</td>
                                <td id="ghg_savings" style="font-weight: bold;">-</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
"""
    )
