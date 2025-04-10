import json
import pint
import os

from rctreportviewer.write_html import write_html_file

path_to_ureg = os.path.join(
    os.path.dirname(__file__),
    "unit_registry.txt",
)
ureg = pint.UnitRegistry(path_to_ureg, autoconvert_offset_to_baseunit=True)


class RCTDetailedReport:
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

    def __init__(
            self,
            detailed_evaluation_report_file_path: str,
            rpd_file_paths: list[str],
            output_file_path: str = "report.html",
    ):
        """
        Args:
            detailed_evaluation_report_file_path (str): Path to the JSON file.
            rpd_file_paths (List[str]): List of paths to the RPD file(s).
            output_file_path (str): Path to the output HTML file.
        """
        self.detailed_evaluation_report_file_path = detailed_evaluation_report_file_path
        self.rpd_file_paths = rpd_file_paths
        self.output_file_path = output_file_path
        self.rpd_data = None
        self.evaluation_data = None

        self.model_types = set()
        self.space_areas = {}
        self.baseline_space_space_types = {}
        self.space_lpd_allowances = {}
        self.baseline_total_lighting_power_allowance = 0
        self.baseline_lighting_power_allowance_by_space_type = {}
        self.proposed_model_summary = {}
        self.baseline_model_summary = {}

        self.rules_passed = []  # ALL outcomes are PASS or N/A
        self.rules_failed = []  # ANY outcome is FAIL
        self.full_eval_rules_undetermined = []  # ANY outcome is UNDETERMINED
        self.appl_eval_rules_undetermined = []  # ANY outcome is UNDETERMINED
        self.rules_not_applicable = []  # ALL outcomes are N/A
        self.rule_evaluation_outcome_counts = {}
        self.rule_evaluation_message_counts = {}

    @staticmethod
    def load_file(file_path):
        """
        Reads a JSON file and returns the python equivalent data structure.
        """
        # Verify the file path is to a JSON file extension
        if not file_path.endswith(".json"):
            raise ValueError("Invalid file type. Please provide a JSON file.")

        with open(file_path, "r") as file:
            data = json.load(file)

        return data

    @staticmethod
    def convert_unit(value, from_unit, to_unit):
        """Convert a numerical value from one unit to another and return the magnitude."""
        quantity = value * ureg[from_unit]
        return quantity.to(ureg[to_unit]).magnitude

    @staticmethod
    def determine_fan_power(fan):
        if "design_electric_power" in fan:
            return fan["design_electric_power"]
        elif (
                "shaft_power" in fan
                and "motor_efficiency" in fan
        ):
            return (
                    fan["shaft_power"]
                    / fan["motor_efficiency"]
            )
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

    @staticmethod
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
                    / (
                            pump["impeller_efficiency"]
                            * pump["motor_efficiency"]
                    )
            )

    def summarize_rmd_data(self, rmd_data, model_type):
        rmd_building_summary = {
            "rmd_type": model_type,
            "building_count": len(rmd_data.get("buildings", [])),
            "building_segment_count": 0,
            "zone_count": 0,
            "space_count": 0,
            "system_count": 0,
            "boiler_count": len(rmd_data.get("boilers", [])),
            "chiller_count": len(rmd_data.get("chillers", [])),
            "heat_rejection_count": len(rmd_data.get("heat_rejections", [])),
            "pump_count": len(rmd_data.get("pumps", [])),
            "fluid_loop_types": {proposed_fluid_loop.get("type") for proposed_fluid_loop in
                                 rmd_data.get("fluid_loops", [])},
            "overall_wall_ua_by_building_segment": {},
            "overall_wall_u_factor_by_building_segment": {},
            "overall_roof_ua_by_building_segment": {},
            "overall_roof_u_factor_by_building_segment": {},
            "overall_window_ua_by_building_segment": {},
            "overall_window_u_factor_by_building_segment": {},
            "overall_skylight_ua_by_building_segment": {},
            "overall_skylight_u_factor_by_building_segment": {},
            "total_floor_area_by_building_segment": {},
            "total_wall_area_by_building_segment": {},
            "total_roof_area_by_building_segment": {},
            "total_window_area_by_building_segment": {},
            "total_skylight_area_by_building_segment": {},
            "total_floor_area_by_space_type": {},
            "total_occupants_by_space_type": {},
            "total_lighting_power_by_space_type": {},
            "total_miscellaneous_equipment_power_by_space_type": {},
            "average_occupancy_by_space_type": {},
            "average_lighting_power_by_space_type": {},
            "average_miscellaneous_equipment_power_by_space_type": {},
            "total_fan_power_by_fan_control_by_fan_type": {},
            "total_air_flow_by_fan_control_by_fan_type": {},
            "other_fan_power_by_fan_type": {},
            "other_air_flow_by_fan_type": {},
            "total_fan_power_by_fan_type": {},
            "total_air_flow_by_fan_type": {},
            "energy_by_fuel_type": {},
            "cost_by_fuel_type": {},
            "energy_by_end_use": {},
            "elec_by_end_use": {},
            "gas_by_end_use": {},
            "energy_by_end_use_eui": {},
            "elec_by_end_use_eui": {},
            "gas_by_end_use_eui": {},
            "total_floor_area": 0,
            "total_exterior_wall_area": 0,
            "total_roof_area": 0,
            "total_window_area": 0,
            "total_skylight_area": 0,
            "total_occupants": 0,
            "total_lighting_power": 0,
            "total_equipment_power": 0,
            "total_pump_power": 0,
            "total_fan_power": 0,
            "total_zone_minimum_oa_flow": 0,
            "total_infiltration": 0,
            "unmet_heating_hours": 0,
            "unmet_cooling_hours": 0,
            "total_energy": 0,
            "total_cost": 0,
        }

        output = rmd_data.get("output")
        if output is not None:
            self.summarize_output_data(output, rmd_building_summary)

        for building in rmd_data.get("buildings", []):
            rmd_building_summary["building_segment_count"] += len(
                building.get("building_segments", [])
            )

            self.summarize_building_segment_data(building, rmd_building_summary)

        for pump in rmd_data.get("pumps", []):
            pump_power = self.determine_pump_power(pump)
            if pump_power:
                rmd_building_summary["total_pump_power"] += pump_power

        return rmd_building_summary

    @staticmethod
    def summarize_output_data(output, rmd_building_summary):
        output_instance = output.get("output_instance")
        if output_instance is not None:
            rmd_building_summary["unmet_heating_hours"] += output_instance.get(
                "unmet_heating_hours", 0
            )
            rmd_building_summary["unmet_cooling_hours"] += output_instance.get(
                "unmet_cooling_hours", 0
            )

            source_results = output_instance.get("annual_source_results", [])
            for source_result in source_results:
                source = source_result.get("energy_source")

                rmd_building_summary["total_energy"] += source_result.get("annual_consumption", 0)
                rmd_building_summary["total_cost"] += source_result.get("annual_cost", 0)
                rmd_building_summary["energy_by_fuel_type"][source] = (
                    rmd_building_summary["energy_by_fuel_type"].get(source, 0)
                    + source_result.get("annual_consumption", 0)
                )
                rmd_building_summary["cost_by_fuel_type"][source] = (
                    rmd_building_summary["cost_by_fuel_type"].get(source, 0)
                    + source_result.get("annual_cost", 0)
                )

            end_use_results = output_instance.get("annual_end_use_results", [])
            for end_use in end_use_results:
                end_use_name = end_use.get("type")

                rmd_building_summary["total_energy"] += end_use.get("annual_site_energy_use", 0)
                rmd_building_summary["energy_by_end_use"][end_use_name] = (
                    rmd_building_summary["energy_by_end_use"].get(end_use_name, 0)
                    + end_use.get("annual_site_energy_use", 0)
                )

                source = end_use.get("energy_source")
                if source == "ELECTRICITY":
                    rmd_building_summary["elec_by_end_use"][end_use_name] = (
                        rmd_building_summary["elec_by_end_use"].get(end_use_name, 0)
                        + end_use.get("annual_site_energy_use", 0)
                    )
                elif source == "NATURAL_GAS":
                    rmd_building_summary["gas_by_end_use"][end_use_name] = (
                        rmd_building_summary["gas_by_end_use"].get(end_use_name, 0)
                        + end_use.get("annual_site_energy_use", 0)
                    )

    def summarize_building_segment_data(self, building, rmd_building_summary):
        for building_segment in building.get(
                "building_segments", []
        ):
            rmd_building_summary["zone_count"] += len(
                building_segment.get("zones", [])
            )
            rmd_building_summary["system_count"] += len(
                building_segment.get(
                    "heating_ventilating_air_conditioning_systems", []
                )
            )

            self.summarize_rmd_zone_data(building_segment, rmd_building_summary)

            self.summarize_rmd_system_data(building_segment, rmd_building_summary)

    def summarize_rmd_zone_data(self, building_segment, rmd_building_summary):
        for zone in building_segment.get("zones", []):
            rmd_building_summary["space_count"] += len(
                zone.get("spaces", [])
            )

            infiltration = zone.get("infiltration")
            if infiltration and "flow_rate" in infiltration:
                rmd_building_summary[
                    "total_infiltration"
                ] += infiltration["flow_rate"]

            zonal_exhaust_fan = zone.get("zonal_exhaust_fan")
            if zonal_exhaust_fan:
                fan_power = self.determine_fan_power(zonal_exhaust_fan)
                if fan_power:
                    if "Undefined" not in rmd_building_summary[
                        "total_fan_power_by_fan_control_by_fan_type"
                    ]:
                        rmd_building_summary["total_fan_power_by_fan_control_by_fan_type"][
                            "Undefined"
                        ] = {
                            "Supply": 0,
                            "Return/Relief": 0,
                            "Exhaust": 0,
                            "Zonal Exhaust": 0,
                            "Terminal Unit": 0,
                        }

                    rmd_building_summary["total_fan_power_by_fan_control_by_fan_type"]["Undefined"]["Zonal Exhaust"] += fan_power
                    rmd_building_summary["total_fan_power"] += fan_power

            self.summarize_rmd_space_data(building_segment, zone, rmd_building_summary)

            self.summarize_rmd_surface_data(building_segment, zone, rmd_building_summary)

            self.summarize_rmd_terminal_data(zone, rmd_building_summary)

    def summarize_rmd_space_data(self, building_segment, zone, rmd_building_summary):
        for space in zone.get("spaces", []):
            if "floor_area" in space:
                rmd_building_summary[
                    "total_floor_area"
                ] += space["floor_area"]
                rmd_building_summary[
                    "total_floor_area_by_building_segment"
                ][building_segment["id"]] = (
                        rmd_building_summary[
                            "total_floor_area_by_building_segment"
                        ].get(building_segment["id"], 0)
                        + space["floor_area"]
                )
                self.space_areas[space["id"]] = space[
                    "floor_area"
                ]
            if "lighting_space_type" in space and rmd_building_summary["rmd_type"] == "Baseline":
                self.baseline_space_space_types[space["id"]] = space[
                    "lighting_space_type"
                ]
            if "number_of_occupants" in space:
                rmd_building_summary[
                    "total_occupants"
                ] += space["number_of_occupants"]
                if "lighting_space_type" in space:
                    rmd_building_summary[
                        "total_occupants_by_space_type"
                    ][space["lighting_space_type"]] = (
                            rmd_building_summary[
                                "total_occupants_by_space_type"
                            ].get(space["lighting_space_type"], 0)
                            + space.get("number_of_occupants", 0)
                    )

            for interior_lighting in space.get(
                    "interior_lighting", [{"power_per_area": 0}]
            ):
                if (
                        "power_per_area" in interior_lighting
                        and "floor_area" in space
                ):
                    rmd_building_summary["total_lighting_power"] += (
                            interior_lighting["power_per_area"]
                            * space["floor_area"]
                    )
                    if "lighting_space_type" in space:
                        rmd_building_summary[
                            "total_floor_area_by_space_type"
                        ][space["lighting_space_type"]] = (
                                rmd_building_summary[
                                    "total_floor_area_by_space_type"
                                ].get(space["lighting_space_type"], 0)
                                + space["floor_area"]
                        )
                        rmd_building_summary[
                            "total_lighting_power_by_space_type"
                        ][space["lighting_space_type"]] = (
                                rmd_building_summary[
                                    "total_lighting_power_by_space_type"
                                ].get(space["lighting_space_type"], 0)
                                + interior_lighting["power_per_area"]
                                * space["floor_area"]
                        )

            for miscellaneous_equipment in space.get(
                    "miscellaneous_equipment", [{"power": 0}]
            ):
                if (
                        "power" in miscellaneous_equipment
                        and "floor_area" in space
                ):
                    rmd_building_summary[
                        "total_equipment_power"
                    ] += miscellaneous_equipment["power"]
                    if "lighting_space_type" in space:
                        rmd_building_summary[
                            "total_miscellaneous_equipment_power_by_space_type"
                        ][space["lighting_space_type"]] = (
                                rmd_building_summary[
                                    "total_miscellaneous_equipment_power_by_space_type"
                                ].get(space["lighting_space_type"], 0)
                                + miscellaneous_equipment["power"]
                        )

    @staticmethod
    def summarize_rmd_surface_data(building_segment, zone, rmd_building_summary):
        for surface in zone.get("surfaces", []):
            if (
                    surface.get("classification") == "WALL"
                    and surface.get("adjacent_to") == "EXTERIOR"
            ):
                if "area" in surface:
                    rmd_building_summary[
                        "total_exterior_wall_area"
                    ] += surface["area"]
                    rmd_building_summary[
                        "total_wall_area_by_building_segment"
                    ][building_segment["id"]] = (
                            rmd_building_summary[
                                "total_wall_area_by_building_segment"
                            ].get(building_segment["id"], 0)
                            + surface["area"]
                    )
                construction = surface.get("construction")
                if (
                        construction
                        and "u_factor" in construction
                ):
                    rmd_building_summary[
                        "overall_wall_ua_by_building_segment"
                    ][building_segment["id"]] = (
                            rmd_building_summary[
                                "overall_wall_ua_by_building_segment"
                            ].get(building_segment["id"], 0)
                            + construction["u_factor"]
                            * surface["area"]
                    )
            if (
                    surface.get("classification") == "CEILING"
                    and surface.get("adjacent_to") == "EXTERIOR"
            ):
                if "area" in surface:
                    rmd_building_summary[
                        "total_roof_area"
                    ] += surface["area"]
                    rmd_building_summary[
                        "total_roof_area_by_building_segment"
                    ][building_segment["id"]] = (
                            rmd_building_summary[
                                "total_roof_area_by_building_segment"
                            ].get(building_segment["id"], 0)
                            + surface["area"]
                    )
                construction = surface.get("construction")
                if (
                        construction
                        and "u_factor" in construction
                ):
                    rmd_building_summary[
                        "overall_roof_ua_by_building_segment"
                    ][building_segment["id"]] = (
                            rmd_building_summary[
                                "overall_roof_ua_by_building_segment"
                            ].get(building_segment["id"], 0)
                            + construction["u_factor"]
                            * surface["area"]
                    )

            for subsurface in surface.get(
                    "subsurfaces", []
            ):
                if (
                        surface.get("adjacent_to") == "EXTERIOR"
                        and subsurface.get("classification")
                        == "WINDOW"
                ):
                    if "glazed_area" in subsurface:
                        rmd_building_summary[
                            "total_window_area"
                        ] += subsurface["glazed_area"]
                        rmd_building_summary[
                            "total_window_area_by_building_segment"
                        ][building_segment["id"]] = (
                                rmd_building_summary[
                                    "total_window_area_by_building_segment"
                                ].get(building_segment["id"], 0)
                                + subsurface["glazed_area"]
                        )
                    if "u_factor" in subsurface:
                        rmd_building_summary[
                            "overall_window_ua_by_building_segment"
                        ][building_segment["id"]] = (
                                rmd_building_summary[
                                    "overall_window_ua_by_building_segment"
                                ].get(building_segment["id"], 0)
                                + subsurface["u_factor"]
                                * subsurface["glazed_area"]
                        )
                elif (
                        surface.get("adjacent_to") == "EXTERIOR"
                        and subsurface.get("classification")
                        == "SKYLIGHT"
                ):
                    if "glazed_area" in subsurface:
                        rmd_building_summary[
                            "total_skylight_area"
                        ] += subsurface["glazed_area"]
                        rmd_building_summary[
                            "total_skylight_area_by_building_segment"
                        ][building_segment["id"]] = (
                                rmd_building_summary[
                                    "total_skylight_area_by_building_segment"
                                ].get(building_segment["id"], 0)
                                + subsurface["glazed_area"]
                        )
                    if "u_factor" in subsurface:
                        rmd_building_summary[
                            "overall_skylight_ua_by_building_segment"
                        ][building_segment["id"]] = (
                                rmd_building_summary[
                                    "overall_skylight_ua_by_building_segment"
                                ].get(building_segment["id"], 0)
                                + subsurface["u_factor"]
                                * subsurface["glazed_area"]
                        )

    def summarize_rmd_terminal_data(self, zone, rmd_building_summary):
        for terminal in zone["terminals"]:
            if "minimum_outdoor_airflow" in terminal:
                rmd_building_summary[
                    "total_zone_minimum_oa_flow"
                ] += terminal["minimum_outdoor_airflow"]

            if "fan" in terminal:
                fan_power = self.determine_fan_power(terminal["fan"])
                if fan_power:
                    if "Undefined" not in rmd_building_summary[
                        "total_fan_power_by_fan_control_by_fan_type"
                    ]:
                        rmd_building_summary["total_fan_power_by_fan_control_by_fan_type"][
                            "Undefined"
                        ] = {
                            "Supply": 0,
                            "Return/Relief": 0,
                            "Exhaust": 0,
                            "Zonal Exhaust": 0,
                            "Terminal Unit": 0,
                        }

                    rmd_building_summary["total_fan_power_by_fan_control_by_fan_type"]["Undefined"]["Terminal Unit"] += fan_power
                    rmd_building_summary["total_fan_power"] += fan_power

    def summarize_rmd_system_data(self, building_segment, rmd_building_summary):
        for hvac_system in building_segment.get(
                "heating_ventilating_air_conditioning_systems", []
        ):
            hvac_fan_system = hvac_system.get("fan_system")
            if hvac_fan_system:
                supply_fan_controls = hvac_fan_system.get(
                    "fan_control",
                    "Undefined"
                )
                if supply_fan_controls == "CONSTANT":
                    occupied_operation = hvac_fan_system.get(
                        "operation_during_occupied", "Undefined"
                    )
                    if occupied_operation == "CYCLING":
                        supply_fan_controls = "Constant Cycling"

                if supply_fan_controls not in rmd_building_summary[
                    "total_fan_power_by_fan_control_by_fan_type"
                ]:
                    rmd_building_summary["total_fan_power_by_fan_control_by_fan_type"][
                        supply_fan_controls
                    ] = {
                        "Supply": 0,
                        "Return/Relief": 0,
                        "Exhaust": 0,
                        "Zonal Exhaust": 0,
                        "Terminal Unit": 0,
                    }
                if supply_fan_controls not in rmd_building_summary[
                    "total_air_flow_by_fan_control_by_fan_type"
                ]:
                    rmd_building_summary["total_air_flow_by_fan_control_by_fan_type"][
                        supply_fan_controls
                    ] = {
                        "Supply": 0,
                        "Return/Relief": 0,
                        "Exhaust": 0,
                        "Zonal Exhaust": 0,
                        "Terminal Unit": 0,
                    }

                for fan in hvac_fan_system.get("supply_fans", []):
                    fan_power = self.determine_fan_power(fan)
                    if fan_power:
                        rmd_building_summary["total_fan_power_by_fan_control_by_fan_type"][
                            supply_fan_controls]["Supply"] += fan_power
                        rmd_building_summary["total_fan_power"] += fan_power
                    if "design_airflow" in fan:
                        rmd_building_summary["total_air_flow_by_fan_control_by_fan_type"][
                            supply_fan_controls
                        ]["Supply"] += fan["design_airflow"]

                for fan in hvac_fan_system.get("return_fans", []) + hvac_fan_system.get(
                        "relief_fans", []):
                    fan_power = self.determine_fan_power(fan)
                    if fan_power:
                        rmd_building_summary["total_fan_power_by_fan_control_by_fan_type"][
                            supply_fan_controls]["Return/Relief"] += fan_power
                        rmd_building_summary["total_fan_power"] += fan_power
                    if "design_airflow" in fan:
                        rmd_building_summary["total_air_flow_by_fan_control_by_fan_type"][
                            supply_fan_controls
                        ]["Return/Relief"] += fan["design_airflow"]

                for fan in hvac_fan_system.get("exhaust_fans", []):
                    fan_power = self.determine_fan_power(fan)
                    if fan_power:
                        rmd_building_summary["total_fan_power_by_fan_control_by_fan_type"][
                            supply_fan_controls]["Exhaust"] += fan_power
                        rmd_building_summary["total_fan_power"] += fan_power
                    if "design_airflow" in fan:
                        rmd_building_summary["total_air_flow_by_fan_control_by_fan_type"][
                            supply_fan_controls
                        ]["Exhaust"] += fan["design_airflow"]

    def load_files(self):
        """
        Loads the JSON files into memory that are needed to produce the HTML report.
        """
        self.evaluation_data = self.load_file(self.detailed_evaluation_report_file_path)
        self.rpd_data = [self.load_file(file_path) for file_path in self.rpd_file_paths]

    def extract_evaluation_data(self):
        """
        Extracts select evaluation data from the overall data structure for reformatting and easy presentation.
        """
        for rpd_file in self.evaluation_data["rpd_files"]:
            self.model_types.add(
                self.model_type_disp_map.get(rpd_file["ruleset_model_type"])
            )

        for rule in self.evaluation_data["rules"]:
            rule_id = rule["rule_id"]
            eval_type = rule["evaluation_type"]
            outcomes = set()
            messages = set()

            # Initialize the nested dictionaries if the rule_id is new
            if rule_id not in self.rule_evaluation_outcome_counts:
                self.rule_evaluation_outcome_counts[rule_id] = {}
            if rule_id not in self.rule_evaluation_message_counts:
                self.rule_evaluation_message_counts[rule_id] = {}

            for evaluation in rule["evaluations"]:
                outcome = self.outcome_disp_map.get(evaluation["outcome"])
                outcomes.add(outcome)

                if isinstance(evaluation["messages"], str):
                    messages.add(evaluation["messages"])
                if isinstance(evaluation["messages"], dict):
                    for key, message in evaluation["messages"].items():
                        messages.add(f"{key}: {message}")
                if isinstance(evaluation["messages"], list):
                    for message in evaluation["messages"]:
                        messages.add(message)

                # Update outcome counts
                if outcome in self.rule_evaluation_outcome_counts[rule_id]:
                    self.rule_evaluation_outcome_counts[rule_id][outcome] += 1
                else:
                    self.rule_evaluation_outcome_counts[rule_id][outcome] = 1

                # Update message counts
                for message in evaluation["messages"]:
                    if message in self.rule_evaluation_message_counts[rule_id]:
                        self.rule_evaluation_message_counts[rule_id][message] += 1
                    else:
                        self.rule_evaluation_message_counts[rule_id][message] = 1

                if rule_id == "6-4" and "calculated_values" in evaluation:
                    lpd_allowance_calc_value = next(
                        calc_value
                        for calc_value in evaluation["calculated_values"]
                        if calc_value["variable"] == "lpd_allowance_b"
                    )
                    if lpd_allowance_calc_value:
                        self.space_lpd_allowances[evaluation["data_group_id"]] = float(
                            lpd_allowance_calc_value["value"]
                        )

            # Determine rule status
            if outcomes == {"Failing"} and messages == {" ::TOLERANCE::"}:
                self.rules_passed.append(rule_id)
            elif "Failing" in outcomes:
                self.rules_failed.append(rule_id)
            elif "Undetermined" in outcomes and eval_type == "FULL":
                self.full_eval_rules_undetermined.append(rule_id)
            elif "Undetermined" in outcomes and eval_type == "APPLICABILITY":
                self.appl_eval_rules_undetermined.append(rule_id)
            elif outcomes == {"Passing"} or outcomes == {"Passing", "N/A"}:
                self.rules_passed.append(rule_id)
            elif outcomes == {"N/A"}:
                self.rules_not_applicable.append(rule_id)

    def extract_model_data(self):
        if len(self.rpd_data) == 1:
            self.rpd_data = self.rpd_data[0]
        else:
            rpd_data = self.rpd_data[0]
            for rpd in self.rpd_data[1:]:
                rpd_data["ruleset_model_descriptions"].extend(
                    rpd["ruleset_model_descriptions"]
                )
            self.rpd_data = rpd_data

        proposed_rmd = next(
            rmd
            for rmd in self.rpd_data["ruleset_model_descriptions"]
            if rmd["type"] == "PROPOSED"
        )
        self.proposed_model_summary = self.summarize_rmd_data(proposed_rmd, model_type="Proposed")

        baseline_rmd = next(
            rmd
            for rmd in self.rpd_data["ruleset_model_descriptions"]
            if rmd["type"] == "BASELINE_0"
        )
        self.baseline_model_summary = self.summarize_rmd_data(baseline_rmd, model_type="Baseline")

    def perform_analytic_calculations(self):
        """
        Perform calculations on the model data to extract additional information.
        """
        # Calculate the LPD allowance based on evaluation data + RPD data combined
        for space_id in self.space_areas:
            self.baseline_total_lighting_power_allowance += (
                    self.space_lpd_allowances.get(space_id, 0)
                    * self.convert_unit(self.space_areas[space_id], "m2", "ft2")
            )
            space_type = self.baseline_space_space_types.get(space_id)
            if space_type:
                self.baseline_lighting_power_allowance_by_space_type[space_type] = (
                        self.baseline_lighting_power_allowance_by_space_type.get(space_type, 0)
                        + self.space_lpd_allowances.get(space_id, 0)
                        * self.convert_unit(self.space_areas[space_id], "m2", "ft2")
                )

        # Calculate the average U-factors by building segment
        for building_segment_id in self.baseline_model_summary[
            "overall_wall_ua_by_building_segment"
        ]:
            self.baseline_model_summary["overall_wall_u_factor_by_building_segment"][
                building_segment_id
            ] = (
                    self.baseline_model_summary["overall_wall_ua_by_building_segment"][
                        building_segment_id
                    ]
                    / self.baseline_model_summary["total_wall_area_by_building_segment"][
                        building_segment_id
                    ]
            )
        for building_segment_id in self.baseline_model_summary[
            "overall_roof_ua_by_building_segment"
        ]:
            self.baseline_model_summary["overall_roof_u_factor_by_building_segment"][
                building_segment_id
            ] = (
                    self.baseline_model_summary["overall_roof_ua_by_building_segment"][
                        building_segment_id
                    ]
                    / self.baseline_model_summary["total_roof_area_by_building_segment"][
                        building_segment_id
                    ]
            )
        for building_segment_id in self.baseline_model_summary[
            "overall_window_ua_by_building_segment"
        ]:
            self.baseline_model_summary["overall_window_u_factor_by_building_segment"][
                building_segment_id
            ] = (
                    self.baseline_model_summary["overall_window_ua_by_building_segment"][
                        building_segment_id
                    ]
                    / self.baseline_model_summary["total_window_area_by_building_segment"][
                        building_segment_id
                    ]
            )
        for building_segment_id in self.baseline_model_summary[
            "overall_skylight_ua_by_building_segment"
        ]:
            self.baseline_model_summary["overall_skylight_u_factor_by_building_segment"][
                building_segment_id
            ] = (
                    self.baseline_model_summary["overall_skylight_ua_by_building_segment"][
                        building_segment_id
                    ]
                    / self.baseline_model_summary["total_skylight_area_by_building_segment"][
                        building_segment_id
                    ]
            )
        for building_segment_id in self.proposed_model_summary[
            "overall_wall_ua_by_building_segment"
        ]:
            self.proposed_model_summary["overall_wall_u_factor_by_building_segment"][
                building_segment_id
            ] = (
                    self.proposed_model_summary["overall_wall_ua_by_building_segment"][
                        building_segment_id
                    ]
                    / self.proposed_model_summary["total_wall_area_by_building_segment"][
                        building_segment_id
                    ]
            )
        for building_segment_id in self.proposed_model_summary[
            "overall_roof_ua_by_building_segment"
        ]:
            self.proposed_model_summary["overall_roof_u_factor_by_building_segment"][
                building_segment_id
            ] = (
                    self.proposed_model_summary["overall_roof_ua_by_building_segment"][
                        building_segment_id
                    ]
                    / self.proposed_model_summary["total_roof_area_by_building_segment"][
                        building_segment_id
                    ]
            )
        for building_segment_id in self.proposed_model_summary[
            "overall_window_ua_by_building_segment"
        ]:
            self.proposed_model_summary["overall_window_u_factor_by_building_segment"][
                building_segment_id
            ] = (
                    self.proposed_model_summary["overall_window_ua_by_building_segment"][
                        building_segment_id
                    ]
                    / self.proposed_model_summary["total_window_area_by_building_segment"][
                        building_segment_id
                    ]
            )
        for building_segment_id in self.proposed_model_summary[
            "overall_skylight_ua_by_building_segment"
        ]:
            self.proposed_model_summary["overall_skylight_u_factor_by_building_segment"][
                building_segment_id
            ] = (
                    self.proposed_model_summary["overall_skylight_ua_by_building_segment"][
                        building_segment_id
                    ]
                    / self.proposed_model_summary["total_skylight_area_by_building_segment"][
                        building_segment_id
                    ]
            )

        # Calculate the average LPD by lighting space type
        for lighting_space_type in self.baseline_model_summary[
            "total_lighting_power_by_space_type"
        ]:
            self.baseline_model_summary[
                "average_lighting_power_by_space_type"
            ][lighting_space_type] = (
                    self.baseline_model_summary[
                        "total_lighting_power_by_space_type"
                    ][lighting_space_type]
                    / self.baseline_model_summary[
                        "total_floor_area_by_space_type"
                    ][lighting_space_type]
            )
        for lighting_space_type in self.proposed_model_summary[
            "total_lighting_power_by_space_type"
        ]:
            self.proposed_model_summary[
                "average_lighting_power_by_space_type"
            ][lighting_space_type] = (
                    self.proposed_model_summary[
                        "total_lighting_power_by_space_type"
                    ][lighting_space_type]
                    / self.proposed_model_summary[
                        "total_floor_area_by_space_type"
                    ][lighting_space_type]
            )

        # Calculate the Other and Total fan summary details
        self.baseline_model_summary["total_fan_power_by_fan_type"] = {
            "Supply": 0,
            "Return/Relief": 0,
            "Exhaust": 0,
            "Zonal Exhaust": 0,
            "Terminal Unit": 0,
        }
        self.baseline_model_summary["total_air_flow_by_fan_type"] = {
            "Supply": 0,
            "Return/Relief": 0,
            "Exhaust": 0,
            "Zonal Exhaust": 0,
            "Terminal Unit": 0,
        }
        self.proposed_model_summary["total_fan_power_by_fan_type"] = {
            "Supply": 0,
            "Return/Relief": 0,
            "Exhaust": 0,
            "Zonal Exhaust": 0,
            "Terminal Unit": 0,
        }
        self.proposed_model_summary["total_air_flow_by_fan_type"] = {
            "Supply": 0,
            "Return/Relief": 0,
            "Exhaust": 0,
            "Zonal Exhaust": 0,
            "Terminal Unit": 0,
        }

        for fan_control in self.baseline_model_summary[
            "total_fan_power_by_fan_control_by_fan_type"
        ]:
            for fan_type in self.baseline_model_summary[
                "total_fan_power_by_fan_control_by_fan_type"
            ][fan_control]:
                self.baseline_model_summary["total_fan_power_by_fan_type"][
                    fan_type
                ] += self.baseline_model_summary[
                    "total_fan_power_by_fan_control_by_fan_type"
                ][fan_control][fan_type]
                self.baseline_model_summary["total_fan_power"] += self.baseline_model_summary[
                    "total_fan_power_by_fan_control_by_fan_type"
                ][fan_control][fan_type]

        for fan_control in self.baseline_model_summary[
            "total_air_flow_by_fan_control_by_fan_type"
        ]:
            for fan_type in self.baseline_model_summary[
                "total_air_flow_by_fan_control_by_fan_type"
            ][fan_control]:
                self.baseline_model_summary["total_air_flow_by_fan_type"][
                    fan_type
                ] += self.baseline_model_summary[
                    "total_air_flow_by_fan_control_by_fan_type"
                ][fan_control][fan_type]

        for fan_control in self.proposed_model_summary[
            "total_fan_power_by_fan_control_by_fan_type"
        ]:
            for fan_type in self.proposed_model_summary[
                "total_fan_power_by_fan_control_by_fan_type"
            ][fan_control]:
                self.proposed_model_summary["total_fan_power_by_fan_type"][
                    fan_type
                ] += self.proposed_model_summary[
                    "total_fan_power_by_fan_control_by_fan_type"
                ][fan_control][fan_type]
                self.proposed_model_summary["total_fan_power"] += self.proposed_model_summary[
                    "total_fan_power_by_fan_control_by_fan_type"
                ][fan_control][fan_type]

        for fan_control in self.proposed_model_summary[
            "total_air_flow_by_fan_control_by_fan_type"
        ]:
            for fan_type in self.proposed_model_summary[
                "total_air_flow_by_fan_control_by_fan_type"
            ][fan_control]:
                self.proposed_model_summary["total_air_flow_by_fan_type"][
                    fan_type
                ] += self.proposed_model_summary[
                    "total_air_flow_by_fan_control_by_fan_type"
                ][fan_control][fan_type]

        self.baseline_model_summary["other_fan_power_by_fan_type"] = {
            "Supply": 0,
            "Return/Relief": 0,
            "Exhaust": 0,
            "Zonal Exhaust": 0,
            "Terminal Unit": 0,
        }

        self.baseline_model_summary["other_air_flow_by_fan_type"] = {
            "Supply": 0,
            "Return/Relief": 0,
            "Exhaust": 0,
            "Zonal Exhaust": 0,
            "Terminal Unit": 0,
        }

        self.proposed_model_summary["other_fan_power_by_fan_type"] = {
            "Supply": 0,
            "Return/Relief": 0,
            "Exhaust": 0,
            "Zonal Exhaust": 0,
            "Terminal Unit": 0,
        }

        self.proposed_model_summary["other_air_flow_by_fan_type"] = {
            "Supply": 0,
            "Return/Relief": 0,
            "Exhaust": 0,
            "Zonal Exhaust": 0,
            "Terminal Unit": 0,
        }

        for fan_control in self.baseline_model_summary[
            "total_fan_power_by_fan_control_by_fan_type"
        ]:
            if fan_control in ["Undefined", "INLET_VANE", "DISCHARGE_DAMPER", "OTHER"]:
                for fan_type in self.baseline_model_summary[
                    "total_fan_power_by_fan_control_by_fan_type"
                ][fan_control]:
                    self.baseline_model_summary["other_fan_power_by_fan_type"][
                        fan_type
                    ] += self.baseline_model_summary[
                        "total_fan_power_by_fan_control_by_fan_type"
                    ][fan_control][fan_type]

        for fan_control in self.baseline_model_summary[
            "total_air_flow_by_fan_control_by_fan_type"
        ]:
            if fan_control in ["Undefined", "INLET_VANE", "DISCHARGE_DAMPER", "OTHER"]:
                for fan_type in self.baseline_model_summary[
                    "total_air_flow_by_fan_control_by_fan_type"
                ][fan_control]:
                    self.baseline_model_summary["other_air_flow_by_fan_type"][
                        fan_type
                    ] += self.baseline_model_summary[
                        "total_air_flow_by_fan_control_by_fan_type"
                    ][fan_control][fan_type]

        for fan_control in self.proposed_model_summary[
            "total_fan_power_by_fan_control_by_fan_type"
        ]:
            if fan_control in ["Undefined", "INLET_VANE", "DISCHARGE_DAMPER", "OTHER"]:
                for fan_type in self.proposed_model_summary[
                    "total_fan_power_by_fan_control_by_fan_type"
                ][fan_control]:
                    self.proposed_model_summary["other_fan_power_by_fan_type"][
                        fan_type
                    ] += self.proposed_model_summary[
                        "total_fan_power_by_fan_control_by_fan_type"
                    ][fan_control][fan_type]

        for fan_control in self.proposed_model_summary[
            "total_air_flow_by_fan_control_by_fan_type"
        ]:
            if fan_control in ["Undefined", "INLET_VANE", "DISCHARGE_DAMPER", "OTHER"]:
                for fan_type in self.proposed_model_summary[
                    "total_air_flow_by_fan_control_by_fan_type"
                ][fan_control]:
                    self.proposed_model_summary["other_air_flow_by_fan_type"][
                        fan_type
                    ] += self.proposed_model_summary[
                        "total_air_flow_by_fan_control_by_fan_type"
                    ][fan_control][fan_type]

    def convert_model_data_units(self):
        """
        Converts the model data from the JSON files to the desired units.
        """
        units_dict = {
            "overall_wall_ua_by_building_segment": ("W / K", "Btu / h / degR"),
            "overall_wall_u_factor_by_building_segment": (
                "W / m2 / K",
                "Btu / h / ft2 / degR",
            ),
            "overall_roof_ua_by_building_segment": ("W / K", "Btu / h / degR"),
            "overall_roof_u_factor_by_building_segment": (
                "W / m2 / K",
                "Btu / h / ft2 / degR",
            ),
            "overall_window_ua_by_building_segment": ("W / K", "Btu / h / degR"),
            "overall_window_u_factor_by_building_segment": (
                "W / m2 / K",
                "Btu / h / ft2 / degR",
            ),
            "overall_skylight_ua_by_building_segment": ("W / K", "Btu / h / degR"),
            "overall_skylight_u_factor_by_building_segment": (
                "W / m2 / K",
                "Btu / h / ft2 / degR",
            ),
            "average_lighting_power_by_space_type": ("W / m2", "W / ft2"),
            "total_floor_area_by_building_segment": ("m2", "ft2"),
            "total_wall_area_by_building_segment": ("m2", "ft2"),
            "total_roof_area_by_building_segment": ("m2", "ft2"),
            "total_window_area_by_building_segment": ("m2", "ft2"),
            "total_floor_area_by_space_type": ("m2", "ft2"),
            "total_floor_area": ("m2", "ft2"),
            "total_exterior_wall_area": ("m2", "ft2"),
            "total_roof_area": ("m2", "ft2"),
            "total_window_area": ("m2", "ft2"),
            "total_zone_minimum_oa_flow": ("L / s", "cfm"),
            "total_infiltration": ("L / s", "cfm"),
            "total_air_flow_by_fan_control_by_fan_type": ("L / s", "cfm"),
            "total_air_flow_by_fan_type": ("L / s", "cfm"),
            "total_energy": ("Btu", "kBtu"),
            "energy_by_fuel_type": ("Btu", "kBtu"),
            "energy_by_end_use": ("Btu", "kBtu"),
            "elec_by_end_use": ("Btu", "kWh"),
            "gas_by_end_use": ("Btu", "therm"),
        }

        # Convert baseline model summary values
        for key in self.baseline_model_summary:
            if key in units_dict:
                if isinstance(self.baseline_model_summary[key], dict):
                    for sub_key in self.baseline_model_summary[key]:
                        if isinstance(self.baseline_model_summary[key][sub_key], dict):
                            for sub_sub_key in self.baseline_model_summary[key][sub_key]:
                                self.baseline_model_summary[key][sub_key][sub_sub_key] = self.convert_unit(
                                    self.baseline_model_summary[key][sub_key][sub_sub_key],
                                    units_dict[key][0],
                                    units_dict[key][1],
                                )
                        else:
                            self.baseline_model_summary[key][sub_key] = self.convert_unit(
                                self.baseline_model_summary[key][sub_key],
                                units_dict[key][0],
                                units_dict[key][1],
                            )
                else:
                    self.baseline_model_summary[key] = self.convert_unit(
                        self.baseline_model_summary[key],
                        units_dict[key][0],
                        units_dict[key][1],
                    )

        # Convert proposed model summary values
        for key in self.proposed_model_summary:
            if key in units_dict:
                if isinstance(self.proposed_model_summary[key], dict):
                    for sub_key in self.proposed_model_summary[key]:
                        if isinstance(self.proposed_model_summary[key][sub_key], dict):
                            for sub_sub_key in self.proposed_model_summary[key][sub_key]:
                                self.proposed_model_summary[key][sub_key][sub_sub_key] = self.convert_unit(
                                    self.proposed_model_summary[key][sub_key][sub_sub_key],
                                    units_dict[key][0],
                                    units_dict[key][1],
                                )
                        else:
                            self.proposed_model_summary[key][sub_key] = self.convert_unit(
                                self.proposed_model_summary[key][sub_key],
                                units_dict[key][0],
                                units_dict[key][1],
                            )
                else:
                    self.proposed_model_summary[key] = self.convert_unit(
                        self.proposed_model_summary[key],
                        units_dict[key][0],
                        units_dict[key][1],
                    )

        # Convert each end use by fuel type to EUI
        for end_use in self.baseline_model_summary["elec_by_end_use"]:
            self.baseline_model_summary["elec_by_end_use_eui"][end_use] = self.baseline_model_summary["elec_by_end_use"][end_use] * 3.412 / self.baseline_model_summary["total_floor_area"]
        for end_use in self.baseline_model_summary["gas_by_end_use"]:
            self.baseline_model_summary["gas_by_end_use_eui"][end_use] = self.baseline_model_summary["gas_by_end_use"][end_use] * 100 / self.baseline_model_summary["total_floor_area"]
        for end_use in self.baseline_model_summary["energy_by_end_use"]:
            self.baseline_model_summary["energy_by_end_use_eui"][end_use] = self.baseline_model_summary["energy_by_end_use"][end_use] / self.baseline_model_summary["total_floor_area"]
        for end_use in self.proposed_model_summary["elec_by_end_use"]:
            self.proposed_model_summary["elec_by_end_use_eui"][end_use] = self.proposed_model_summary["elec_by_end_use"][end_use] * 3.412 / self.proposed_model_summary["total_floor_area"]
        for end_use in self.proposed_model_summary["gas_by_end_use"]:
            self.proposed_model_summary["gas_by_end_use_eui"][end_use] = self.proposed_model_summary["gas_by_end_use"][end_use] * 100 / self.proposed_model_summary["total_floor_area"]
        for end_use in self.proposed_model_summary["energy_by_end_use"]:
            self.proposed_model_summary["energy_by_end_use_eui"][end_use] = self.proposed_model_summary["energy_by_end_use"][end_use] / self.proposed_model_summary["total_floor_area"]

    def heating_cooling_capacity_summary(self):
        """Summarizes the heating and cooling capacities from the model data."""
        proposed_rmd = next(
            rmd
            for rmd in self.rpd_data["ruleset_model_descriptions"]
            if rmd["type"] == "PROPOSED"
        )
        proposed_heating_capacity = {
            "Electricity": 0,
            "Fossil Fuel": 0,
            "On-site Boiler Plant": 0,
            "Purchased Heat": 0,
        }
        proposed_cooling_capacity = {
            "Electricity": 0,
            "On-site Chiller Plant": 0,
            "Purchased CHW": 0,
        }
        for building in proposed_rmd.get("buildings", []):
            for building_segment in building.get("building_segments", []):
                for hvac_system in building_segment.get("heating_ventilating_air_conditioning_systems", []):
                    energy_source = None
                    # Heating systems
                    heating_system = hvac_system.get("heating_system")
                    if heating_system:
                        energy_source = heating_system.get("energy_source_type")
                        fuel = self.get_fuel_type(energy_source, heating_system.get("type"))
                        if fuel in proposed_heating_capacity:
                            proposed_heating_capacity[fuel] += heating_system.get("design_capacity", 0.0)
                    # Cooling systems
                    cooling_system = hvac_system.get("cooling_system")
                    if cooling_system and energy_source:
                        fuel = self.get_fuel_type(energy_source, cooling_system.get("type"))
                        if fuel in proposed_cooling_capacity:
                            proposed_cooling_capacity[fuel] += cooling_system.get("design_total_cool_capacity", 0.0)
                # Terminal capacities only populated by our RPD generation software when appropriate.
                zones = building_segment.get("zones", [])
                for zone in zones:
                    for terminal in zone.get("terminals", []):
                        # TODO: not necessarily a 1:1 mapping with hvac_systems, which has energy type info. Discuss
                        proposed_heating_capacity["Electricity"] += terminal.get("heating_capacity", 0.0)
                        proposed_cooling_capacity["Electricity"] += terminal.get("cooling_capacity", 0.0)

        baseline_rmd = next(
            rmd
            for rmd in self.rpd_data["ruleset_model_descriptions"]
            if rmd["type"] == "BASELINE_0"
        )
        baseline_heating_capacity = {
            "Electricity": 0,
            "Fossil Fuel": 0,
            "On-site Boiler Plant": 0,
            "Purchased Heat": 0,
        }
        baseline_cooling_capacity = {
            "Electricity": 0,
            "On-site Chiller Plant": 0,
            "Purchased CHW": 0,
        }
        for building in baseline_rmd.get("buildings", []):
            for building_segment in building.get("building_segments", []):
                for hvac_system in building_segment.get("heating_ventilating_air_conditioning_systems", []):
                    energy_source = None
                    # Heating systems
                    heating_system = hvac_system.get("heating_system")
                    if heating_system:
                        energy_source = heating_system.get("energy_source_type")
                        fuel = self.get_fuel_type(energy_source, heating_system.get("type"))
                        if fuel in baseline_heating_capacity:
                            baseline_heating_capacity[fuel] += heating_system.get("design_capacity", 0.0)
                    # Cooling systems
                    cooling_system = hvac_system.get("cooling_system")
                    if cooling_system and energy_source:
                        fuel = self.get_fuel_type(energy_source, cooling_system.get("type"))
                        if fuel in baseline_cooling_capacity:
                            baseline_cooling_capacity[fuel] += cooling_system.get("design_total_cool_capacity", 0.0)
                # Terminal capacities only populated by our RPD generation software when appropriate.
                zones = building_segment.get("zones", [])
                for zone in zones:
                    for terminal in zone.get("terminals", []):
                        # TODO: not necessarily a 1:1 mapping with hvac_systems, which has energy type info. Discuss
                        proposed_heating_capacity["Electricity"] += terminal.get("heating_capacity", 0.0)
                        proposed_cooling_capacity["Electricity"] += terminal.get("cooling_capacity", 0.0)

        # Sum capacities in each capacity column (dict) and add the total as a key:value pair
        self.convert_and_total_capacities(proposed_heating_capacity)
        self.convert_and_total_capacities(proposed_cooling_capacity)
        self.convert_and_total_capacities(baseline_heating_capacity)
        self.convert_and_total_capacities(baseline_cooling_capacity)
        # TODO: Capacities here are in "columns." Use these to recreate table 1B in html report

    # TODO: Check this mapping
    def get_fuel_type(self, energy_source, hvac_type):
        if hvac_type == "DIRECT_EXPANSION":
            return "Electricity"
        elif energy_source == "ELECTRICITY":
            return "Electricity"
        elif energy_source in ["NATURAL_GAS", "PROPANE", "FUEL_OIL"]:
            return "Fossil Fuel"
        elif energy_source == "STEAM" and hvac_type in ["BOILER", "FURNACE"]:
            return "On-site Boiler Plant"
        elif energy_source == "STEAM":
            return "On-site Chiller Plant"
        elif energy_source == "PURCHASED_HOT_WATER":
            return "Purchased Heat"
        elif energy_source == "PURCHASED_CHILLED_WATER":
            return "Purchased CHW"
        return "Unknown"

    def convert_and_total_capacities(self, capacity_dict):
        """Sum all the values in a capacity dictionary and add a total key.
        Convert values from Btu/h to kBtu/h for heating and cooling capacities."""
        total = 0
        for key in capacity_dict:
            capacity_dict[key] = self.convert_unit(capacity_dict[key], "Btu", "kBtu")
            total += capacity_dict[key]
        capacity_dict["Total"] = total
        return capacity_dict

    def run(self):
        self.load_files()
        self.extract_evaluation_data()
        self.extract_model_data()
        self.perform_analytic_calculations()
        self.convert_model_data_units()
        self.heating_cooling_capacity_summary()
        write_html_file(self)
