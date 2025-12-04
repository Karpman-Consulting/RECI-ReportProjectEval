def determine_fan_power(fan):
    if "design_electric_power" in fan:
        return fan["design_electric_power"]
    elif "shaft_power" in fan and "motor_efficiency" in fan:
        return fan["shaft_power"] / fan["motor_efficiency"]
    elif (
        "total_efficiency" in fan
        and "design_pressure_rise" in fan
        and "design_airflow" in fan
    ):
        return (
            fan["design_airflow"]
            * fan["design_pressure_rise"]
            / fan["total_efficiency"]
        )


def determine_pump_power(pump):
    if "design_electric_power" in pump:
        return pump["design_electric_power"]
    elif (
        "design_flow" in pump
        and "design_head" in pump
        and "impeller_efficiency" in pump
        and "motor_efficiency" in pump
    ):
        return (
            pump["design_flow"]
            * pump["design_head"]
            / (pump["impeller_efficiency"] * pump["motor_efficiency"])
        )
