import os


path_to_bpf_data = os.path.join(
    os.path.dirname(__file__),
    "BPFs.json",
)
path_to_lpd_data = os.path.join(
    os.path.dirname(__file__),
    "LPDs.json",
)
path_to_ureg = os.path.join(
    os.path.dirname(__file__),
    "unit_registry.txt",
)

model_type_disp_map = {
    "USER": "Design",
    "PROPOSED": "Proposed",
    "BASELINE_0": "Baseline",
    "BASELINE_90": "Baseline (90 deg)",
    "BASELINE_180": "Baseline (180 deg)",
    "BASELINE_270": "Baseline (270 deg)",
}
outcome_disp_map = {
    "PASS": "Passing",
    "FAILED": "Failing",
    "NOT_APPLICABLE": "N/A",
    "UNDETERMINED": "Undetermined",
}
efficiency_display_map = {
    "FULL_LOAD_COEFFICIENT_OF_PERFORMANCE": "Full Load COP",
    "FULL_LOAD_COEFFICIENT_OF_PERFORMANCE_NO_FAN": "Full Load COP<sub>nf</sub>",
    "ENERGY_EFFICIENCY_RATIO": "EER",
    "SEASONAL_ENERGY_EFFICIENCY_RATIO": "SEER",
    "SEASONAL_ENERGY_EFFICIENCY_RATIO_2": "SEER2",
    "INTEGRATED_ENERGY_EFFICIENCY_RATIO": "IEER",
    "INTEGRATED_PART_LOAD_VALUE": "IPLV",
    "COMBINED_ENERGY_EFFICIENCY_RATIO": "CEER",
    "COEFFICIENT_OF_PERFORMANCE_WATER_TO_AIR_WATER_LOOP": "COP<sub>water-to-air</sub>",
    "COEFFICIENT_OF_PERFORMANCE_WATER_TO_AIR_WATER_LOOP_NO_FAN": "COP<sub>water-to-air-nf</sub>",
    "COEFFICIENT_OF_PERFORMANCE_WATER_TO_AIR_GROUND_WATER": "COP<sub>ground-water-to-air</sub>",
    "COEFFICIENT_OF_PERFORMANCE_WATER_TO_AIR_GROUND_WATER_NO_FAN": "COP<sub>ground-water-to-air-nf</sub>",
    "COEFFICIENT_OF_PERFORMANCE_BRINE_TO_AIR_GROUND_LOOP": "COP<sub>ground-loop-to-air</sub>",
    "COEFFICIENT_OF_PERFORMANCE_BRINE_TO_AIR_GROUND_LOOP_NO_FAN": "COP<sub>ground-loop-to-air-nf</sub>",
    "COEFFICIENT_OF_PERFORMANCE_WATER_TO_WATER_WATER_LOOP": "COP<sub>water-to-water</sub>",
    "COEFFICIENT_OF_PERFORMANCE_WATER_TO_WATER_GROUND_WATER": "COP<sub>ground-water-to-water</sub>",
    "COEFFICIENT_OF_PERFORMANCE_BRINE_TO_WATER_GROUND_LOOP": "COP<sub>ground-loop-to-water</sub>",
    "NONE": "None",
    "OTHER": "Other",
    "HEAT_PUMP_COEFFICIENT_OF_PERFORMANCE_HIGH_TEMPERATURE": "Heat Pump COP<sub>high</sub>",
    "HEAT_PUMP_COEFFICIENT_OF_PERFORMANCE_LOW_TEMPERATURE": "Heat Pump COP<sub>low</sub>",
    "HEAT_PUMP_COEFFICIENT_OF_PERFORMANCE_HIGH_TEMPERATURE_NO_FAN": "Heat Pump COP<sub>high-nf</sub>",
    "HEAT_PUMP_COEFFICIENT_OF_PERFORMANCE_LOW_TEMPERATURE_NO_FAN": "Heat Pump COP<sub>low-nf</sub>",
    "THERMAL_EFFICIENCY": "E<sub>t</sub>",
    "COMBUSTION_EFFICIENCY": "Combustion E<sub>c</sub>",
    "ANNUAL_FUEL_UTILIZATION_EFFICIENCY": "AFUE",
    "HEATING_SEASONAL_PERFORMANCE_FACTOR": "HSPF",
    "HEATING_SEASONAL_PERFORMANCE_FACTOR_2": "HSPF2",
}
fuel_type_map = {
    "ELECTRICITY": "Electricity",
    "NATURAL_GAS": "Fossil Fuel",
    "PROPANE": "Fossil Fuel",
    "FUEL_OIL": "Fossil Fuel",
    "STEAM": "Fossil Fuel",
    "OTHER": "Other",
}
bpf_area_type_map = {
    "AUTOMOTIVE_FACILITY": "All Others",
    "CONVENTION_CENTER": "All Others",
    "COURTHOUSE": "All Others",
    "DINING_BAR_LOUNGE_LEISURE": "Restaurant",
    "DINING_CAFETERIA_FAST_FOOD": "Restaurant",
    "DINING_FAMILY": "Restaurant",
    "DORMITORY": "Multifamily",
    "EXERCISE_CENTER": "All Others",
    "FIRE_STATION": "All Others",
    "GYMNASIUM": "All Others",
    "HEALTH_CARE_CLINIC": "Healthcare/hospital",
    "HOSPITAL": "Healthcare/hospital",
    "HOTEL_MOTEL": "Hotel/motel",
    "LIBRARY": "All Others",
    "MANUFACTURING_FACILITY": "All Others",
    "MOTION_PICTURE_THEATER": "All Others",
    "MULTIFAMILY": "Multifamily",
    "MUSEUM": "All Others",
    "OFFICE": "Office",
    "PARKING_GARAGE": "All Others",
    "PENITENTIARY": "All Others",
    "PERFORMING_ARTS_THEATER": "All Others",
    "POLICE_STATION": "All Others",
    "POST_OFFICE": "All Others",
    "RELIGIOUS_FACILITY": "All Others",
    "RETAIL": "Retail",
    "SCHOOL_UNIVERSITY": "School",
    "SPORTS_ARENA": "All Others",
    "TOWN_HALL": "Office",
    "TRANSPORTATION": "All Others",
    "WAREHOUSE": "Warehouse",
    "WORKSHOP": "All Others",
}
