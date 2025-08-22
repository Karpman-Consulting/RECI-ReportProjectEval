from rctreportviewer.constants import efficiency_display_map


def write_hvac_summary(file, rct_detailed_report):
    file.write(
        f"""

                    <div class="mb-3 me-4">
                        <button class="btn btn-info collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapse-hvac-summary" aria-expanded="false">
                            HVAC Summary
                        </button>

                        <div id="collapse-hvac-summary" class="accordion-collapse collapse">
                            <div class="accordion-body">
            """
    )

    if (
        rct_detailed_report.proposed_model_summary["chiller_count"]
        + rct_detailed_report.baseline_model_summary["chiller_count"]
    ) > 0:
        # -------------------------- Cooling Plant Summary Table-------------------------
        file.write(
            f"""   
                                    <h3> Cooling Plant Summary</h3>
                                    <table class="table table-sm table-borderless fan-summary" style="width: 1150px;">
                                        <thead>
                                            <tr class="text-center">
                                                <th style="border: 2px solid black; width: 12%;" rowspan="2">Fuel Type</th>
                                                <th style="border: 2px solid black; width: 14%;" colspan="4">Baseline Design</th>
                                                <th style="border: 2px solid black; width: 14%;" colspan="4">Proposed Design</th>
                                            </tr>
                                            <tr class="text-center">
                                                <th style="border: 2px solid black;">Total Quantity of Chillers</th>
                                                <th style="border: 2px solid black;">Total Chiller Plant Capacity [ton]</th>
                                                <th style="border: 2px solid black;">Total Cooling Tower GPM</th>
                                                <th style="border: 2px solid black;">Total Cooling Tower HP</th>
                                                <th style="border: 2px solid black;">Total Quantity of Chillers</th>
                                                <th style="border: 2px solid black;">Total Chiller Plant Capacity [ton]</th>
                                                <th style="border: 2px solid black;">Total Cooling Tower GPM</th>
                                                <th style="border: 2px solid black;">Total Cooling Tower HP</th>
                                            </tr>
                                        </thead>
                                        <tbody style="border: 2px solid black;">
        """
        )
        # Check if there is any chiller plant info for electricity in the baseline model
        write_row = False
        for val in [
            rct_detailed_report.baseline_model_summary.get("electric_chiller_count", 0),
            rct_detailed_report.baseline_model_summary.get(
                "electric_chiller_plant_capacity", 0
            ),
            rct_detailed_report.baseline_model_summary.get("cooling_tower_gpm", 0),
            rct_detailed_report.baseline_model_summary.get("cooling_tower_hp", 0),
            rct_detailed_report.proposed_model_summary.get("electric_chiller_count", 0),
            rct_detailed_report.proposed_model_summary.get(
                "electric_chiller_plant_capacity", 0
            ),
            rct_detailed_report.proposed_model_summary.get("cooling_tower_gpm", 0),
            rct_detailed_report.proposed_model_summary.get("cooling_tower_hp", 0),
        ]:
            if val > 0:
                write_row = True
                break
        if write_row:
            file.write(
                f"""
                                            <tr style="font-size: 12px;" class="text-center">
                                                <td style="border-right: 2px solid black;">Electricity</td>
                                                <td>{round(rct_detailed_report.baseline_model_summary.get("electric_chiller_count", 0)):,}</td>
                                                <td>{round(rct_detailed_report.baseline_model_summary.get("electric_chiller_plant_capacity", 0), 1):,}</td>
                                                <td>{round(rct_detailed_report.baseline_model_summary.get("cooling_tower_gpm", 0), 1):,}</td>
                                                <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary.get("cooling_tower_hp", 0), 1):,}</td>
                                                <td>{round(rct_detailed_report.proposed_model_summary.get("electric_chiller_count", 0)):,}</td>
                                                <td>{round(rct_detailed_report.proposed_model_summary.get("electric_chiller_plant_capacity", 0), 1):,}</td>
                                                <td>{round(rct_detailed_report.proposed_model_summary.get("cooling_tower_gpm", 0), 1):,}</td>
                                                <td>{round(rct_detailed_report.proposed_model_summary.get("cooling_tower_hp", 0), 1):,}</td>
                                            </tr>
            """
            )
        write_row = False
        for val in [
            rct_detailed_report.proposed_model_summary.get(
                "fossil_fuel_chiller_count", 0
            ),
            rct_detailed_report.proposed_model_summary.get(
                "fossil_fuel_chiller_plant_capacity", 0.0
            ),
        ]:
            if val > 0:
                write_row = True
                break
        if write_row:
            file.write(
                f"""
                                            <tr style="font-size: 12px;" class="text-center">
                                                <td style="border-right: 2px solid black;">Fossil Fuel</td>
                                                <td style="background: black;"></td>
                                                <td style="background: black;"></td>
                                                <td style="background: black;"></td>
                                                <td style="background: black;"></td>
                                                <td>{round(rct_detailed_report.proposed_model_summary.get("fossil_fuel_chiller_count", 0)):,}</td>
                                                <td>{round(rct_detailed_report.proposed_model_summary.get("fossil_fuel_chiller_plant_capacity", 0.0), 1):,}</td>
                                                <td style="background: black;"></td>
                                                <td style="background: black;"></td>
                                            </tr>
            """
            )
        file.write(
            f"""
                                            <tr style="font-size: 12px; border-top: 1px solid black;" class="fw-bold text-center subtotal">
                                                <td style="border-right: 2px solid black;">Total</td>
                                                <td>{round(rct_detailed_report.baseline_model_summary.get("electric_chiller_count", 0)):,}</td>
                                                <td>{round(rct_detailed_report.baseline_model_summary.get("electric_chiller_plant_capacity", 0), 1):,}</td>
                                                <td>{round(rct_detailed_report.baseline_model_summary.get("cooling_tower_gpm", 0), 1):,}</td>
                                                <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary.get("cooling_tower_hp", 0), 1):,}</td>
                                                <td>{round(rct_detailed_report.proposed_model_summary.get("chiller_count", 0)):,}</td>
                                                <td>{round((rct_detailed_report.proposed_model_summary.get("electric_chiller_plant_capacity", 0) +
                                                            rct_detailed_report.proposed_model_summary.get("fossil_fuel_chiller_plant_capacity", 0)), 1):,}</td>
                                                <td>{round(rct_detailed_report.proposed_model_summary.get("cooling_tower_gpm", 0), 1):,}</td>
                                                <td>{round(rct_detailed_report.proposed_model_summary.get("cooling_tower_hp", 0), 1):,}</td>
                                            </tr>
                                        </tbody>
                                    </table>
        """
        )

    # -------------------------- Heating Plant Summary Table-------------------------
    if (
        rct_detailed_report.proposed_model_summary["boiler_count"]
        + rct_detailed_report.baseline_model_summary["boiler_count"]
    ) > 0:
        file.write(
            f"""   
                                            <h3> Heating Plant Summary</h3>
                                            <table class="table table-sm table-borderless fan-summary" style="width: 800px;">
                                                <thead>
                                                    <tr class="text-center">
                                                        <th style="border: 2px solid black; width: 12%;" rowspan="2">Fuel Type</th>
                                                        <th style="border: 2px solid black; width: 14%;" colspan="2">Baseline Design</th>
                                                        <th style="border: 2px solid black; width: 14%;" colspan="2">Proposed Design</th>
                                                    </tr>
                                                    <tr class="text-center">
                                                        <th style="border: 2px solid black;">Total Quantity of Boilers</th>
                                                        <th style="border: 2px solid black;">Total Boiler Plant Capacity [kBtu/hr]</th>
                                                        <th style="border: 2px solid black;">Total Quantity of Boilers</th>
                                                        <th style="border: 2px solid black;">Total Boiler Plant Capacity [kBtu/hr]</th>
                                                    </tr>
                                                </thead>
                                                <tbody style="border: 2px solid black;">
        """
        )

        # Check if there is any boiler plant info for electricity in the proposed model
        write_row = False
        for val in [
            rct_detailed_report.proposed_model_summary.get("electric_boiler_count", 0),
            rct_detailed_report.proposed_model_summary.get(
                "electric_boiler_plant_capacity", 0.0
            ),
        ]:
            if val > 0:
                write_row = True
                break
        if write_row:
            file.write(
                f"""
                                                    <tr style="font-size: 12px;" class="text-center">
                                                        <td style="border-right: 2px solid black;">Electricity</td>
                                                        <td style="background: black;"></td>
                                                        <td style="border-right: 2px solid black; background: black;"></td>
                                                        <td>{round(rct_detailed_report.proposed_model_summary.get("electric_boiler_count", 0)):,}</td>
                                                        <td>{round(rct_detailed_report.proposed_model_summary.get("electric_boiler_plant_capacity", 0)):,}</td>
                                                    </tr>
            """
            )
        # Check if there is any boiler plant info for fossil fuel
        write_row = False
        for val in [
            rct_detailed_report.baseline_model_summary.get(
                "fossil_fuel_boiler_count", 0
            ),
            rct_detailed_report.baseline_model_summary.get(
                "fossil_fuel_boiler_plant_capacity", 0.0
            ),
            rct_detailed_report.proposed_model_summary.get(
                "fossil_fuel_boiler_count", 0
            ),
            rct_detailed_report.proposed_model_summary.get(
                "fossil_fuel_boiler_plant_capacity", 0.0
            ),
        ]:
            if val > 0:
                write_row = True
                break
        if write_row:
            file.write(
                f"""
                                                    <tr style="font-size: 12px;" class="text-center">
                                                        <td style="border-right: 2px solid black;">Fossil Fuel</td>
                                                        <td>{round(rct_detailed_report.baseline_model_summary.get("fossil_fuel_boiler_count", 0)):,}</td>
                                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary.get("fossil_fuel_boiler_plant_capacity", 0), 1):,}</td>
                                                        <td>{round(rct_detailed_report.proposed_model_summary.get("fossil_fuel_boiler_count", 0)):,}</td>
                                                        <td>{round(rct_detailed_report.proposed_model_summary.get("fossil_fuel_boiler_plant_capacity", 0), 1):,}</td>
                                                    </tr>
            """
            )
        file.write(
            f"""
                                                <tr style="font-size: 12px; border-top: 1px solid black;" class="fw-bold text-center subtotal">
                                                    <td style="border-right: 2px solid black;">Total</td>
                                                    <td>{round(rct_detailed_report.baseline_model_summary.get("fossil_fuel_boiler_count", 0)):,}</td>
                                                    <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary.get("fossil_fuel_boiler_plant_capacity", 0), 1):,}</td>
                                                    <td>{round(rct_detailed_report.proposed_model_summary.get("boiler_count", 0)):,}</td>
                                                    <td>{round((rct_detailed_report.proposed_model_summary.get("electric_boiler_plant_capacity", 0) +
                                                                rct_detailed_report.proposed_model_summary.get("fossil_fuel_boiler_plant_capacity", 0)), 1):,}</td>
                                                </tr>
                                            </tbody>
                                        </table>
        """
        )

    # -------------------------- Air-Side HVAC Capacity Summary Table-------------------------
    file.write(
        f"""   
                                <h3> Air-side HVAC Capacity Summary</h3>
                                <table class="table table-sm table-borderless fan-summary" style="width: 750px;">
                                    <thead>
                                        <tr class="text-center">
                                            <th style="border: 2px solid black; width: 12%;" rowspan="2">Fuel Type</th>
                                            <th style="border: 2px solid black; width: 14%;" colspan="2">Baseline Design</th>
                                            <th style="border: 2px solid black; width: 14%;" colspan="2">Proposed Design</th>
                                        </tr>
                                        <tr class="text-center">
                                            <th style="border: 2px solid black;">Heating Capacity [kBtu/hr]</th>
                                            <th style="border: 2px solid black;">Cooling Capacity [kBtu/hr]</th>
                                            <th style="border: 2px solid black;">Heating Capacity [kBtu/hr]</th>
                                            <th style="border: 2px solid black;">Cooling Capacity [kBtu/hr]</th>
                                        </tr>
                                    </thead>
                                    <tbody style="border: 2px solid black;">
    """
    )
    # Check if there are any electricity heating or cooling capacities in the baseline or proposed models
    write_row = False
    for val in [
        rct_detailed_report.baseline_model_summary["heating_capacity_by_fuel_type"].get(
            "Electricity", 0.0
        ),
        rct_detailed_report.baseline_model_summary["cooling_capacity_by_fuel_type"].get(
            "Electricity", 0.0
        ),
        rct_detailed_report.proposed_model_summary["heating_capacity_by_fuel_type"].get(
            "Electricity", 0.0
        ),
        rct_detailed_report.proposed_model_summary["cooling_capacity_by_fuel_type"].get(
            "Electricity", 0.0
        ),
    ]:
        if val > 0:
            write_row = True
            break
    if write_row:
        file.write(
            f"""
                                        <tr style="font-size: 12px;" class="text-center">
                                            <td style="border-right: 2px solid black;">Electricity</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['heating_capacity_by_fuel_type'].get("Electricity", 0.0)):,}</td>
                                            <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['cooling_capacity_by_fuel_type'].get("Electricity", 0.0)):,}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['heating_capacity_by_fuel_type'].get("Electricity", 0.0)):,}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['cooling_capacity_by_fuel_type'].get("Electricity", 0.0)):,}</td>
                                        </tr>
        """
        )
    # Check if there are any fossil fuel heating capacities in the baseline or proposed models
    write_row = False
    for val in [
        rct_detailed_report.baseline_model_summary["heating_capacity_by_fuel_type"].get(
            "Fossil Fuel", 0.0
        ),
        rct_detailed_report.proposed_model_summary["heating_capacity_by_fuel_type"].get(
            "Fossil Fuel", 0.0
        ),
    ]:
        if val > 0:
            write_row = True
            break
    if write_row:
        file.write(
            f"""
                                        <tr style="font-size: 12px;" class="text-center">
                                            <td style="border-right: 2px solid black;">Fossil Fuel</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['heating_capacity_by_fuel_type'].get("Fossil Fuel", 0.0)):,}</td>
                                            <td style="background: black; border-right: 2px solid black;"></td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['heating_capacity_by_fuel_type'].get("Fossil Fuel", 0.0)):,}</td>
                                            <td style="background: black;"></td>
                                        </tr>
        """
        )
    # Check if there are any On-site Boiler Plant heating capacities in the baseline or proposed models
    write_row = False
    for val in [
        rct_detailed_report.baseline_model_summary["heating_capacity_by_fuel_type"].get(
            "On-site Boiler Plant", 0.0
        ),
        rct_detailed_report.proposed_model_summary["heating_capacity_by_fuel_type"].get(
            "On-site Boiler Plant", 0.0
        ),
    ]:
        if val > 0:
            write_row = True
            break
    if write_row:
        file.write(
            f"""
                                        <tr style="font-size: 12px;" class="text-center">
                                            <td style="border-right: 2px solid black;">On-site Boiler Plant</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['heating_capacity_by_fuel_type'].get("On-site Boiler Plant", 0.0)):,}</td>
                                            <td style="background: black; border-right: 2px solid black;"></td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['heating_capacity_by_fuel_type'].get("On-site Boiler Plant", 0.0)):,}</td>
                                            <td style="background: black;"></td>
                                        </tr>
        """
        )
    # Check if there are any Purchased Heat heating capacities in the baseline or proposed models
    write_row = False
    for val in [
        rct_detailed_report.baseline_model_summary["heating_capacity_by_fuel_type"].get(
            "Purchased Heat", 0.0
        ),
        rct_detailed_report.proposed_model_summary["heating_capacity_by_fuel_type"].get(
            "Purchased Heat", 0.0
        ),
    ]:
        if val > 0:
            write_row = True
            break
    if write_row:
        file.write(
            f"""
                                        <tr style="font-size: 12px;" class="text-center">
                                            <td style="border-right: 2px solid black;">Purchased Heat</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['heating_capacity_by_fuel_type'].get("Purchased Heat", 0.0)):,}</td>
                                            <td style="background: black; border-right: 2px solid black;"></td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['heating_capacity_by_fuel_type'].get("Purchased Heat", 0.0)):,}</td>
                                            <td style="background: black;"></td>
                                        </tr>
        """
        )
    # Check if there are any On-site Chiller Plant cooling capacities in the baseline or proposed models
    write_row = False
    for val in [
        rct_detailed_report.baseline_model_summary["cooling_capacity_by_fuel_type"].get(
            "On-site Chiller Plant", 0.0
        ),
        rct_detailed_report.proposed_model_summary["cooling_capacity_by_fuel_type"].get(
            "On-site Chiller Plant", 0.0
        ),
    ]:
        if val > 0:
            write_row = True
            break
    if write_row:
        file.write(
            f"""
                                        <tr style="font-size: 12px;" class="text-center">
                                            <td style="border-right: 2px solid black;">On-site Chiller Plant</td>
                                            <td style="background: black;"></td>
                                            <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['cooling_capacity_by_fuel_type'].get("On-site Chiller Plant", 0.0)):,}</td>
                                            <td style="background: black;"></td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['cooling_capacity_by_fuel_type'].get("On-site Chiller Plant", 0.0)):,}</td>
                                        </tr>
        """
        )
    # Check if there are any Purchased CHW cooling capacities in the baseline or proposed models
    write_row = False
    for val in [
        rct_detailed_report.baseline_model_summary["cooling_capacity_by_fuel_type"].get(
            "Purchased CHW", 0.0
        ),
        rct_detailed_report.proposed_model_summary["cooling_capacity_by_fuel_type"].get(
            "Purchased CHW", 0.0
        ),
    ]:
        if val > 0:
            write_row = True
            break
    if write_row:
        file.write(
            f"""
                                        <tr style="font-size: 12px;" class="text-center">
                                            <td style="border-right: 2px solid black;">Purchased CHW</td>
                                            <td style="background: black;"></td>
                                            <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['cooling_capacity_by_fuel_type'].get("Purchased CHW", 0.0)):,}</td>
                                            <td style="background: black;"></td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['cooling_capacity_by_fuel_type'].get("Purchased CHW", 0.0)):,}</td>
                                        </tr>
        """
        )
    file.write(
        f"""
                                    <tr style="font-size: 12px; border-top: 1px solid black;" class="fw-bold text-center subtotal">
                                        <td style="border-right: 2px solid black;">Total</td>
                                        <td>{round(rct_detailed_report.baseline_model_summary['heating_capacity_by_fuel_type'].get("Total", 0.0)):,}</td>
                                        <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['cooling_capacity_by_fuel_type'].get("Total", 0.0)):,}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['heating_capacity_by_fuel_type'].get("Total", 0.0)):,}</td>
                                        <td>{round(rct_detailed_report.proposed_model_summary['cooling_capacity_by_fuel_type'].get("Total", 0.0)):,}</td>
                                    </tr>
                                </tbody>
                            </table>
    """
    )

    # ----------------------- HVAC Fan Summary Table -----------------------
    file.write(
        f"""
                                <h3>Baseline HVAC Fan Summary</h3>
                                <p><strong>Outdoor Airflow:</strong> {round(rct_detailed_report.baseline_model_summary['total_zone_minimum_oa_flow']):,} CFM</p>
                                <table class="table table-sm table-borderless fan-summary" style="width: 1250px;">
                                    <thead>
                                        <tr class="text-center">
                                            <th style="border: 2px solid black; width: 12%;" rowspan="2">Fan Type</th>
                                            <th style="border: 2px solid black; width: 14%;" colspan="3">Constant Volume</th>
                                            <th style="border: 2px solid black; width: 14%;" colspan="3">Variable Volume</th>
                                            <th style="border: 2px solid black; width: 14%;" colspan="3">Multispeed</th>
                                            <th style="border: 2px solid black; width: 14%;" colspan="3">Constant Volume, Cycling</th>
                                            <th style="border: 2px solid black; width: 14%;" colspan="3">Other</th>
                                            <th style="border: 2px solid black; width: 18%;" colspan="4">Total</th>
                                        </tr>
                                        <tr class="text-center">
                                            <th style="border: 2px solid black;">CFM</th>
                                            <th style="border: 2px solid black;">kW</th>
                                            <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                            <th style="border: 2px solid black;">CFM</th>
                                            <th style="border: 2px solid black;">kW</th>
                                            <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                            <th style="border: 2px solid black;">CFM</th>
                                            <th style="border: 2px solid black;">kW</th>
                                            <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                            <th style="border: 2px solid black;">CFM</th>
                                            <th style="border: 2px solid black;">kW</th>
                                            <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                            <th style="border: 2px solid black;">CFM</th>
                                            <th style="border: 2px solid black;">kW</th>
                                            <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                            <th style="border: 2px solid black;">CFM</th>
                                            <th style="border: 2px solid black;">kW</th>
                                            <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                            <th style="border: 2px solid black;">% of Subtotal kW</th>
                                        </tr>
                                    </thead>
                                    <tbody style="border: 2px solid black;">
    """
    )

    for fan_type in ["Supply", "Return/Relief", "Exhaust", "Zonal Exhaust"]:
        file.write(
            f"""
                                        <tr style="font-size: 12px;" class="text-center">
                                            <td style="border-right: 2px solid black;">{fan_type}</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("CONSTANT", {}).get(fan_type, 0)):,}</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("CONSTANT", {}).get(fan_type, 0) / 1000, 2):,}</td>
                                            <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("CONSTANT", {}).get(fan_type, 0) / (rct_detailed_report.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("CONSTANT", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get(fan_type, 0)):,}</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get(fan_type, 0) / 1000, 2):,}</td>
                                            <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get(fan_type, 0) / (rct_detailed_report.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get(fan_type, 0)):,}</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get(fan_type, 0) / 1000, 2):,}</td>
                                            <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get(fan_type, 0) / (rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get(fan_type, 0)):,}</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get(fan_type, 0) / 1000, 2):,}</td>
                                            <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get(fan_type, 0) / (rct_detailed_report.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['other_air_flow_by_fan_type'].get(fan_type, 0)):,}</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['other_fan_power_by_fan_type'].get(fan_type, 0) / 1000, 2):,}</td>
                                            <td style="border-right: 2px solid black;">{round(rct_detailed_report.baseline_model_summary['other_fan_power_by_fan_type'].get(fan_type, 0) / (rct_detailed_report.baseline_model_summary['other_air_flow_by_fan_type'].get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['total_air_flow_by_fan_type'].get(fan_type, 0)):,}</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_type'].get(fan_type, 0) / 1000, 2):,}</td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_type'].get(fan_type, 0) / (rct_detailed_report.baseline_model_summary['total_air_flow_by_fan_type'].get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(100 * rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_type'].get(fan_type, 0) / sum(rct_detailed_report.baseline_model_summary["total_fan_power_by_fan_type"].values()))}</td>
                                        </tr>
        """
        )
    # --------- Subtotal Row -------------
    file.write(
        f"""
                                        <tr style="font-size: 12px; border-top: 1px solid black;" class="fw-bold text-center subtotal">
                                            <td style="border-right: 2px solid black;">Subtotal</td>
                                            <td></td>
                                            <td></td>
                                            <td style="border-right: 2px solid black;"></td>
                                            <td></td>
                                            <td></td>
                                            <td style="border-right: 2px solid black;"></td>
                                            <td></td>
                                            <td></td>
                                            <td style="border-right: 2px solid black;"></td>
                                            <td></td>
                                            <td></td>
                                            <td style="border-right: 2px solid black;"></td>
                                            <td></td>
                                            <td></td>
                                            <td style="border-right: 2px solid black;"></td>
                                            <td></td>
                                            <td></td>
                                            <td></td>
                                            <td>0</td>
                                        </tr>
    """
    )

    # --------- Terminal Units Row ------------
    file.write(
        f"""
                                        <tr style="font-size: 12px; border-top: 1px solid black;" class="text-center">
                                            <td style="border-right: 2px solid black;">Terminal Units</td>
                                            <td style="background: black;"></td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("CONSTANT", {}).get("Terminal Unit", 0)):,}</td>
                                            <td style="border-right: 2px solid black; background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get("Terminal Unit", 0)):,}</td>
                                            <td style="border-right: 2px solid black; background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get("Terminal Unit", 0)):,}</td>
                                            <td style="border-right: 2px solid black; background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get("Terminal Unit", 0)):,}</td>
                                            <td style="border-right: 2px solid black; background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['other_fan_power_by_fan_type'].get("Terminal Unit", 0)):,}</td>
                                            <td style="border-right: 2px solid black; background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(rct_detailed_report.baseline_model_summary['total_fan_power_by_fan_type'].get("Terminal Unit", 0)):,}</td>
                                            <td style="background: black;"></td>
                                            <td style="background: black;"></td>
                                        </tr>
    """
    )
    file.write(
        f"""
                                    </tbody>
                                </table>

                                <h3>Proposed HVAC Fan Summary</h3>
                                <p><strong>Outdoor Airflow:</strong> {round(rct_detailed_report.baseline_model_summary['total_zone_minimum_oa_flow']):,} CFM</p>
                                <table class="table table-sm table-borderless fan-summary" style="width: 1250px;">
                                    <thead>
                                        <tr class="text-center">
                                            <th style="border: 2px solid black; width: 12%;" rowspan="2">Fan Type</th>
                                            <th style="border: 2px solid black; width: 14%;" colspan="3">Constant Volume</th>
                                            <th style="border: 2px solid black; width: 14%;" colspan="3">Variable Volume</th>
                                            <th style="border: 2px solid black; width: 14%;" colspan="3">Multispeed</th>
                                            <th style="border: 2px solid black; width: 14%;" colspan="3">Constant Volume, Cycling</th>
                                            <th style="border: 2px solid black; width: 14%;" colspan="3">Other</th>
                                            <th style="border: 2px solid black; width: 18%;" colspan="4">Total</th>
                                        </tr>
                                        <tr class="text-center">
                                            <th style="border: 2px solid black;">CFM</th>
                                            <th style="border: 2px solid black;">kW</th>
                                            <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                            <th style="border: 2px solid black;">CFM</th>
                                            <th style="border: 2px solid black;">kW</th>
                                            <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                            <th style="border: 2px solid black;">CFM</th>
                                            <th style="border: 2px solid black;">kW</th>
                                            <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                            <th style="border: 2px solid black;">CFM</th>
                                            <th style="border: 2px solid black;">kW</th>
                                            <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                            <th style="border: 2px solid black;">CFM</th>
                                            <th style="border: 2px solid black;">kW</th>
                                            <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                            <th style="border: 2px solid black;">CFM</th>
                                            <th style="border: 2px solid black;">kW</th>
                                            <th style="border: 2px solid black;">W/CFM<sub>s</sub></th>
                                            <th style="border: 2px solid black;">% of Subtotal kW</th>
                                        </tr>
                                    </thead>
                                    <tbody style="border: 2px solid black;">
    """
    )

    for fan_type in ["Supply", "Return/Relief", "Exhaust", "Zonal Exhaust"]:
        file.write(
            f"""
                                        <tr style="font-size: 12px;" class="text-center">
                                            <td style="border-right: 2px solid black;">{fan_type}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("CONSTANT", {}).get(fan_type, 0)):,}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("CONSTANT", {}).get(fan_type, 0) / 1000, 2):,}</td>
                                            <td style="border-right: 2px solid black;">{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("CONSTANT", {}).get(fan_type, 0) / (rct_detailed_report.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("CONSTANT", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get(fan_type, 0)):,}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get(fan_type, 0) / 1000, 2):,}</td>
                                            <td style="border-right: 2px solid black;">{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get(fan_type, 0) / (rct_detailed_report.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get(fan_type, 0)):,}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get(fan_type, 0) / 1000, 2):,}</td>
                                            <td style="border-right: 2px solid black;">{round(rct_detailed_report.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get(fan_type, 0) / (rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get(fan_type, 0)):,}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get(fan_type, 0) / 1000, 2):,}</td>
                                            <td style="border-right: 2px solid black;">{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get(fan_type, 0) / (rct_detailed_report.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['other_air_flow_by_fan_type'].get(fan_type, 0)):,}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['other_fan_power_by_fan_type'].get(fan_type, 0) / 1000, 2):,}</td>
                                            <td style="border-right: 2px solid black;">{round(rct_detailed_report.proposed_model_summary['other_fan_power_by_fan_type'].get(fan_type, 0) / (rct_detailed_report.proposed_model_summary['other_air_flow_by_fan_type'].get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['total_air_flow_by_fan_type'].get(fan_type, 0)):,}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_type'].get(fan_type, 0) / 1000, 2):,}</td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_type'].get(fan_type, 0) / (rct_detailed_report.proposed_model_summary['total_air_flow_by_fan_type'].get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(100 * rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_type'].get(fan_type, 0) / sum(rct_detailed_report.proposed_model_summary["total_fan_power_by_fan_type"].values()))}</td>
                                        </tr>
        """
        )
    # ---------- Subtotal Row -------------
    file.write(
        f"""
                                        <tr style="font-size: 12px; border-top: 1px solid black;" class="fw-bold text-center subtotal">
                                            <td style="border-right: 2px solid black;">Subtotal</td>
                                            <td></td>
                                            <td></td>
                                            <td style="border-right: 2px solid black;"></td>
                                            <td></td>
                                            <td></td>
                                            <td style="border-right: 2px solid black;"></td>
                                            <td></td>
                                            <td></td>
                                            <td style="border-right: 2px solid black;"></td>
                                            <td></td>
                                            <td></td>
                                            <td style="border-right: 2px solid black;"></td>
                                            <td></td>
                                            <td></td>
                                            <td style="border-right: 2px solid black;"></td>
                                            <td></td>
                                            <td></td>
                                            <td></td>
                                            <td>0</td>
                                        </tr>
    """
    )
    # --------- Terminal Units Row -----------
    file.write(
        f"""
                                        <tr style="font-size: 12px; border-top: 1px solid black;" class="text-center">
                                            <td>Terminal Units</td>
                                            <td style="background: black;"></td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("CONSTANT", {}).get("Terminal Unit", 0)):,}</td>
                                            <td style="border-right: 2px solid black; background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get("Terminal Unit", 0)):,}</td>
                                            <td style="border-right: 2px solid black; background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get("Terminal Unit", 0)):,}</td>
                                            <td style="border-right: 2px solid black; background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get("Terminal Unit", 0)):,}</td>
                                            <td style="border-right: 2px solid black; background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['other_fan_power_by_fan_type'].get("Terminal Unit", 0)):,}</td>
                                            <td style="border-right: 2px solid black; background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(rct_detailed_report.proposed_model_summary['total_fan_power_by_fan_type'].get("Terminal Unit", 0)):,}</td>
                                            <td style="background: black;"></td>
                                            <td style="background: black;"></td>
                                        </tr>
                                    </tbody>
                                </table>
    """
    )

    file.write(
        f""" 
                        </div>
                    </div>
                </div>

                <div class="mb-3 me-4">
                    <button class="btn btn-info collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapse-hvac-details" aria-expanded="false">
                            HVAC Details
                    </button>

                    <div id="collapse-hvac-details" class="accordion-collapse collapse">
                        <div class="accordion-body">
    """
    )

    # -------------------------- Air-Side HVAC System Type, Capacity, and Efficiency Summary Table-------------------------
    file.write(
        f"""   
                                                    <h3> Baseline Air-Side HVAC System Type, Capacity, and Efficiency</h3>
                                                    <table class="table table-sm table-borderless fan-summary">
                                                        <thead>
                                                            <tr class="text-center">
                                                                <th style="border: 2px solid black;" rowspan="2">Modeled System Name</th>
                                                                <th style="border: 2px solid black;" rowspan="2">System Type</th>
                                                                <th style="border: 2px solid black;" rowspan="2">Zone Qty.</th>
                                                                <th style="border: 2px solid black;" colspan="6">Heating</th>
                                                                <th style="border: 2px solid black;" colspan="5">Cooling</th>
                                                            </tr>
                                                            <tr class="text-center">
                                                                <th style="border: 2px solid black;">Equipment Type</th>
                                                                <th style="border: 2px solid black;">Fuel Type/Heating Source</th>
                                                                <th style="border: 2px solid black;">Total Capacity</th>
                                                                <th style="border: 2px solid black;">Cap. Units</th>
                                                                <th style="border: 2px solid black;">Unitary Eff.</th>
                                                                <th style="border: 2px solid black;">Eff. Units</th>
                                                                <th style="border: 2px solid black;">Equipment Type</th>
                                                                <th style="border: 2px solid black;">Total Capacity</th>
                                                                <th style="border: 2px solid black;">Cap. Units</th>
                                                                <th style="border: 2px solid black;">Unitary Eff.</th>
                                                                <th style="border: 2px solid black;">Eff. Units</th>
                                                            </tr>
                                                        </thead>
                                                        <tbody style="border: 2px solid black;">
    """
    )
    # A row for every system
    for system_summary in rct_detailed_report.baseline_model_summary[
        "hvac_system_summaries"
    ]:
        file.write(
            f"""
                                                            <tr style="font-size: 12px;" class="text-center">
                                                                <td>{system_summary.get("name", "-")}</td>
                                                                <td>{system_summary.get("type", "-")}</td>
                                                                <td style="border-right: 2px solid black;">{system_summary.get("zone_qty", 0)}</td>
                                                                <td>{system_summary.get("heating_equipment_type", "-").replace("_", " ").title()}</td>
                                                                <td>{system_summary.get("heating_energy_source", "-").replace("_", " ").title()}</td>
                                                                <td>{round(system_summary.get("heating_capacity", 0)):,}</td>
                                                                <td>{system_summary.get("heating_capacity_units", "-")}</td>
        """
        )
        if system_summary.get(
            "heating_efficiency_metric_values"
        ) and system_summary.get("heating_efficiency_metric_types"):
            efficiency_values = ", ".join(
                str(round(x, 3))
                for x in system_summary["heating_efficiency_metric_values"]
            )
            efficiency_types = ", ".join(
                efficiency_display_map.get(metric, metric)
                for metric in system_summary["heating_efficiency_metric_types"]
            )
            file.write(
                f"""
                                                                <td>{efficiency_values}</td>
                                                                <td style="border-right: 2px solid black;">{efficiency_types}</td>
                                                                <td>{system_summary.get("cooling_equipment_type", "-").replace("_", " ").title()}</td>
                                                                <td>{round(system_summary.get("cooling_capacity", 0)):,}</td>
                                                                <td>{system_summary.get("cooling_capacity_units", "-")}</td>
            """
            )
        else:
            file.write(
                f"""
                                                                <td>-</td>
                                                                <td style="border-right: 2px solid black;">-</td>
                                                                <td>{system_summary.get("cooling_equipment_type", "-").replace("_", " ").title()}</td>
                                                                <td>{round(system_summary.get("cooling_capacity", 0)):,}</td>
                                                                <td>{system_summary.get("cooling_capacity_units", "-")}</td>
            """
            )
        if system_summary.get(
            "cooling_efficiency_metric_values"
        ) and system_summary.get("cooling_efficiency_metric_types"):
            efficiency_values = ", ".join(
                str(round(x, 3))
                for x in system_summary["cooling_efficiency_metric_values"]
            )
            efficiency_types = ", ".join(
                efficiency_display_map.get(metric, metric)
                for metric in system_summary["cooling_efficiency_metric_types"]
            )
            file.write(
                f"""
                                                            <td>{efficiency_values}</td>
                                                            <td style="border-right: 2px solid black;">{efficiency_types}</td>
            """
            )
        else:
            file.write(
                f"""
                                                            <td>-</td>
                                                            <td style="border-right: 2px solid black;">-</td>
                    """
            )

        file.write(
            f"""
                                                        </tr>             
        """
        )

    file.write(
        f"""
                                                    </tbody>
                                                </table>
                                            </div>
                                        </div>
                                    </div>
    """
    )
