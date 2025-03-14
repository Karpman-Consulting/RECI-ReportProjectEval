import json
import pint
import os
import math

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
        self.rules_passed = []  # ALL outcomes are PASS or N/A
        self.rules_failed = []  # ANY outcome is FAIL
        self.full_eval_rules_undetermined = []  # ANY outcome is UNDETERMINED
        self.appl_eval_rules_undetermined = []  # ANY outcome is UNDETERMINED
        self.rules_not_applicable = []  # ALL outcomes are N/A
        self.space_areas = {}
        self.baseline_space_space_types = {}
        self.space_lpd_allowances = {}

        self.rule_evaluation_outcome_counts = {}
        self.rule_evaluation_message_counts = {}

        self.baseline_total_lighting_power_allowance = 0
        self.baseline_lighting_power_allowance_by_space_type = {}
        self.proposed_model_summary = {}
        self.baseline_model_summary = {}

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
            "total_infiltration": 0
        }

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

    def write_html_file(self):
        """
        Writes the extracted data to an HTML file for easy viewing with Bootstrap styling.
        """
        section_titles_with_colors = {
            1: ("Design Model and Compliance Calculations", "#D8BFD8"),
            2: ("Additions and Alterations", "#66b3ff"),
            3: ("Space Use Classification", "#99ff99"),
            4: ("Schedules", "#ffcc99"),
            5: ("Envelope", "#f4a460"),
            6: ("Lighting", "#ffd700"),
            7: ("Thermal Blocks - HVAC Zones Designed", "#c2f0c2"),
            8: ("Thermal Blocks - HVAC Zones Not Designed", "#f0c2c2"),
            9: ("Thermal Blocks - Multifamily Residential Buildings", "#f0e68c"),
            10: ("HVAC Systems", "#4682b4"),
            11: ("Service Water Heating Systems", "#E97451"),
            12: ("Receptacles and Other Loads", "#d3d3d3"),
            13: ("Modeling Limitations to the Simulation Program", "#f4cccc"),
            14: ("Exterior Conditions", "#87ceeb"),
            15: ("Distribution Transformers", "#d9ead3"),
            16: ("Elevators", "#c0c0c0"),
            17: ("Refrigeration", "#5f9ea0"),
            18: ("Baseline HVAC Selection", "#ead1dc"),
            19: ("General Baseline HVAC System Requirements", "#778899"),
            20: ("System-Specific Baseline HVAC System Requirements", "#ffdab9"),
            21: ("Baseline HVAC - Water Side Requirements: Hot Water", "#ff6347"),
            22: ("Baseline HVAC - Water Side Requirements: Chilled Water", "#6495ED"),
            23: ("Baseline HVAC - Air Side Requirements", "#F0FFFF"),
        }

        with open(self.output_file_path, "w", encoding="utf-8") as file:
            file.write(
                """
            <html style="scrollbar-gutter: stable;">
            <head>
                <meta charset="UTF-8">
                <title>SIMcheck Detailed Evaluation Report</title>
                <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
                <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
                <style>
                    td.rule-id { white-space: nowrap; }
                    td.outcome-summary { white-space: pre-wrap; }
                    .sticky-top-2 {
                        top: 37px;
                        z-index: 1029;
                    }
                </style>
            </head>
            """
            )
            file.write(
                f"""
            <body class="mt-2 ms-2">
                <div class="d-flex flex-nowrap">
                    
                    <div class="flex-grow-1">
                        <h1 class="text-center mb-4">RECI - Project Evaluation Report</h1>
                        <div class="mb-3">
                            <p><strong>Ruleset:</strong> {self.evaluation_data["ruleset"]}</p>
                            <p><strong>Generated on:</strong> {self.evaluation_data["date_run"]}</p>
                            <p><strong>Models Analyzed:</strong> {", ".join(self.model_types)}</p>
                        </div>
                    
            """
            )

            rule_categories = {
                "Failing": self.rules_failed,
                "Passing": self.rules_passed,
                "Undetermined": self.full_eval_rules_undetermined + self.appl_eval_rules_undetermined,
                "N/A": self.rules_not_applicable,
            }

            file.write(
                f"""
                    <div class="mb-3 me-4">
                        <button class="btn btn-info collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapse-model-component-summary" aria-expanded="false">
                            Model Component Summary
                        </button>

                        <div id="collapse-model-component-summary" class="accordion-collapse collapse">
                            <div class="accordion-body">
                                <table class="table table-sm table-borderless" style="width: 400px;">
                                    <thead>
                                        <tr style="border-bottom: 2px solid black;"><th class="col-4 text-end"></th><th class="col-4 text-center">Baseline</th><th class="col-4 text-center">Proposed</th></tr>
                                    </thead>
                                    <tbody>
                                        <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Building Qty</td><td class="col-4 text-center">{self.baseline_model_summary["building_count"]}</td><td class="col-4 text-center">{self.proposed_model_summary["building_count"]}</td></tr>
                                        <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Total Floor Area</td><td class="col-4 text-center">{round(self.baseline_model_summary['total_floor_area']):,}</td><td class="col-4 text-center">{round(self.proposed_model_summary["total_floor_area"]):,}</td></tr>
                                        <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Building Area Qty</td><td class="col-4 text-center">{self.baseline_model_summary["building_segment_count"]}</td><td class="col-4 text-center">{self.proposed_model_summary["building_segment_count"]}</td></tr>
                                        <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">System Qty</td><td class="col-4 text-center">{self.baseline_model_summary["system_count"]}</td><td class="col-4 text-center">{self.proposed_model_summary["system_count"]}</td></tr>
                                        <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Zone Qty</td><td class="col-4 text-center">{self.baseline_model_summary["zone_count"]}</td><td class="col-4 text-center">{self.proposed_model_summary["zone_count"]}</td></tr>
                                        <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Space Qty</td><td class="col-4 text-center">{self.baseline_model_summary["space_count"]}</td><td class="col-4 text-center">{self.proposed_model_summary["space_count"]}</td></tr>
                                        <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Fluid Loops</td><td class="col-4 text-center">{", ".join(s.title() for s in self.baseline_model_summary["fluid_loop_types"])}</td><td class="col-4 text-center">{", ".join(s.title() for s in self.proposed_model_summary["fluid_loop_types"])}</td></tr>
                                        <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Pump Qty</td><td class="col-4 text-center">{self.baseline_model_summary["pump_count"]}</td><td class="col-4 text-center">{self.proposed_model_summary["pump_count"]}</td></tr>
                                        <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Boiler Qty</td><td class="col-4 text-center">{self.baseline_model_summary["boiler_count"]}</td><td class="col-4 text-center">{self.proposed_model_summary["boiler_count"]}</td></tr>
                                        <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Chiller Qty</td><td class="col-4 text-center">{self.baseline_model_summary["chiller_count"]}</td><td class="col-4 text-center">{self.proposed_model_summary["chiller_count"]}</td></tr>
                                        <tr style="font-size: 12px;" class="lh-1"><td class="col-3 text-end">Heat Rejection Qty</td><td class="col-4 text-center">{self.baseline_model_summary["heat_rejection_count"]}</td><td class="col-4 text-center">{self.proposed_model_summary["heat_rejection_count"]}</td></tr>
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <div class="mb-3 me-4">
                        <button class="btn btn-info collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapse-envelope-summary" aria-expanded="false">
                            Envelope Summary
                        </button>

                        <div id="collapse-envelope-summary" class="accordion-collapse collapse">
                            <div class="accordion-body">
                                <table class="table table-sm table-borderless" style="width: 1200px;">
                                    <thead>
                                        <tr class="text-center">
                                            <th colspan="2"></th>
                                            <th colspan="6" style="border: 2px solid black;">Baseline</th>
                                            <th colspan="6" style="border: 2px solid black;">Proposed</th>
                                        </tr>
                                        <tr class="text-center">
                                            <th rowspan="2" style="border: 2px solid black;">Building Area</th>
                                            <th rowspan="2" style="border: 2px solid black;">Surface Type</th>
                                            <th colspan="3" style="border: 2px solid black;">Opaque Surface</th>
                                            <th colspan="3" style="border: 2px solid black;">Fenestration</th>
                                            <th colspan="3" style="border: 2px solid black;">Opaque Surface</th>
                                            <th colspan="3" style="border: 2px solid black;">Fenestration</th>
                                        </tr>
                                        <tr class="text-center">
                                            <th style="border: 2px solid black;">Area (ft<sup>2</sup>)</th>
                                            <th style="border: 2px solid black;"> % </th>
                                            <th style="border: 2px solid black;"> U-Factor </th>
                                            <th style="border: 2px solid black;">Area (ft<sup>2</sup>)</th>
                                            <th style="border: 2px solid black;"> % </th>
                                            <th style="border: 2px solid black;"> U-Factor </th>
                                            <th style="border: 2px solid black;">Area (ft<sup>2</sup>)</th>
                                            <th style="border: 2px solid black;"> % </th>
                                            <th style="border: 2px solid black;"> U-Factor </th>
                                            <th style="border: 2px solid black;">Area (ft<sup>2</sup>)</th>
                                            <th style="border: 2px solid black;"> % </th>
                                            <th style="border: 2px solid black;"> U-Factor </th>
                                        </tr>
                                    </thead>
                                    <tbody style="border: 2px solid black;">
                """)

            for building_segment_id in self.baseline_model_summary["total_floor_area_by_building_segment"]:
                if building_segment_id in self.baseline_model_summary["total_roof_area_by_building_segment"]:
                    file.write(
                        f"""
                                        <tr style="font-size: 12px;" class="lh-1 text-center">
                                            <td>{building_segment_id}</td>
                                            <td style="border-right: 2px solid black;">Roof</td>
                                            <td>{round(self.baseline_model_summary['total_roof_area_by_building_segment'].get(building_segment_id, 0) - self.baseline_model_summary['total_skylight_area_by_building_segment'].get(building_segment_id, 0)):,}</td>
                                            <td>{round((self.baseline_model_summary['total_roof_area_by_building_segment'].get(building_segment_id, 0) - self.baseline_model_summary['total_skylight_area_by_building_segment'].get(building_segment_id, 0)) / self.baseline_model_summary['total_roof_area_by_building_segment'][building_segment_id] * 100, 1)}</td>
                                            <td>{round(self.baseline_model_summary["overall_roof_u_factor_by_building_segment"].get(building_segment_id, 0), 3)}</td>
                                            <td>{round(self.baseline_model_summary["total_skylight_area_by_building_segment"].get(building_segment_id, 0)):,}</td>
                                            <td>{round(self.baseline_model_summary["total_skylight_area_by_building_segment"].get(building_segment_id, 0) / self.baseline_model_summary['total_roof_area_by_building_segment'][building_segment_id] * 100, 1)}</td>
                                            <td style="border-right: 2px solid black;">{round(self.baseline_model_summary["overall_skylight_u_factor_by_building_segment"].get(building_segment_id, 0), 3)}</td>
                                            <td>{round(self.proposed_model_summary["total_roof_area_by_building_segment"].get(building_segment_id, 0) - self.proposed_model_summary['total_skylight_area_by_building_segment'].get(building_segment_id, 0)):,}</td>
                                            <td>{round((self.proposed_model_summary["total_roof_area_by_building_segment"].get(building_segment_id, 0) - self.proposed_model_summary['total_skylight_area_by_building_segment'].get(building_segment_id, 0)) / self.proposed_model_summary['total_roof_area_by_building_segment'][building_segment_id] * 100, 1)}</td>
                                            <td>{round(self.proposed_model_summary["overall_roof_u_factor_by_building_segment"].get(building_segment_id, 0), 3)}</td>
                                            <td>{round(self.proposed_model_summary["total_skylight_area_by_building_segment"].get(building_segment_id, 0)):,}</td>
                                            <td>{round(self.proposed_model_summary["total_skylight_area_by_building_segment"].get(building_segment_id, 0) / self.proposed_model_summary['total_roof_area_by_building_segment'][building_segment_id] * 100, 1)}</td>
                                            <td>{round(self.proposed_model_summary["overall_skylight_u_factor_by_building_segment"].get(building_segment_id, 0), 3)}</td>
                                        </tr>
                        """
                    )
                if building_segment_id in self.baseline_model_summary["total_wall_area_by_building_segment"]:
                    file.write(
                        f"""
                                        <tr style="font-size: 12px;" class="lh-1 text-center">
                                            <td>{building_segment_id}</td>
                                            <td style="border-right: 2px solid black;">Ext. Wall</td>
                                            <td>{round(self.baseline_model_summary['total_wall_area_by_building_segment'].get(building_segment_id, 0) - self.baseline_model_summary["total_window_area_by_building_segment"].get(building_segment_id, 0)):,}</td>
                                            <td>{round((self.baseline_model_summary['total_wall_area_by_building_segment'].get(building_segment_id, 0) - self.baseline_model_summary["total_window_area_by_building_segment"].get(building_segment_id, 0)) / self.baseline_model_summary['total_wall_area_by_building_segment'][building_segment_id] * 100, 1)}</td>
                                            <td>{round(self.baseline_model_summary["overall_wall_u_factor_by_building_segment"].get(building_segment_id, 0), 3)}</td>
                                            <td>{round(self.baseline_model_summary["total_window_area_by_building_segment"].get(building_segment_id, 0)):,}</td>
                                            <td>{round(self.baseline_model_summary["total_window_area_by_building_segment"].get(building_segment_id, 0) / self.baseline_model_summary['total_wall_area_by_building_segment'][building_segment_id] * 100, 1)}</td>
                                            <td style="border-right: 2px solid black;">{round(self.baseline_model_summary["overall_window_u_factor_by_building_segment"].get(building_segment_id, 0), 3)}</td>
                                            <td>{round(self.proposed_model_summary["total_wall_area_by_building_segment"].get(building_segment_id, 0) - self.proposed_model_summary["total_window_area_by_building_segment"].get(building_segment_id, 0)):,}</td>
                                            <td>{round((self.proposed_model_summary["total_wall_area_by_building_segment"].get(building_segment_id, 0) - self.proposed_model_summary["total_window_area_by_building_segment"].get(building_segment_id, 0)) / self.proposed_model_summary['total_wall_area_by_building_segment'][building_segment_id] * 100, 1)}</td>
                                            <td>{round(self.proposed_model_summary["overall_wall_u_factor_by_building_segment"].get(building_segment_id, 0), 3)}</td>
                                            <td>{round(self.proposed_model_summary["total_window_area_by_building_segment"].get(building_segment_id, 0)):,}</td>
                                            <td>{round(self.proposed_model_summary["total_window_area_by_building_segment"].get(building_segment_id, 0) / self.proposed_model_summary['total_wall_area_by_building_segment'][building_segment_id] * 100, 1)}</td>
                                            <td>{round(self.proposed_model_summary["overall_window_u_factor_by_building_segment"].get(building_segment_id, 0), 3)}</td>
                                        </tr>
                        """
                    )

            file.write("""          </tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <div class="mb-3 me-4">
                        <button class="btn btn-info collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapse-internal-loads-summary" aria-expanded="false">
                            Internal Loads Summary
                        </button>

                        <div id="collapse-internal-loads-summary" class="accordion-collapse collapse">
                            <div class="accordion-body">
                                <table class="table table-sm table-borderless" style="width: 900px;">
                                    <thead>
                                        <tr class="text-center">
                                            <th colspan="2" class="col-4"></th>
                                            <th colspan="4" class="col-4" style="border: 2px solid black;">Baseline</th>
                                            <th colspan="3" class="col-4" style="border: 2px solid black;">Proposed</th>
                                        </tr>
                                        <tr class="text-center">
                                            <th style="border: 2px solid black;">Space Type</th>
                                            <th style="border: 2px solid black;">Area (ft<sup>2</sup>)</th>
                                            <th style="border: 2px solid black;">Occupancy Density (ft<sup>2</sup>/person)</th>
                                            <th style="border: 2px solid black;">Equipment Power Density (W/ft<sup>2</sup>)</th>
                                            <th style="border: 2px solid black;">Allowed Lighting Power Density (W/ft<sup>2</sup>)</th>
                                            <th style="border: 2px solid black;">Lighting Power Density (W/ft<sup>2</sup>)</th>
                                            <th style="border: 2px solid black;">Lighting Power Density (W/ft<sup>2</sup>)</th>
                                            <th style="border: 2px solid black;">Equipment Power Density (W/ft<sup>2</sup>)</th>
                                            <th style="border: 2px solid black;">Occupancy Density (ft<sup>2</sup>/person)</th>
                                        </tr>
                                    </thead>
                                    <tbody style="border: 2px solid black;">
            """)

            for space_type in self.baseline_model_summary["total_floor_area_by_space_type"]:
                file.write(
                    f"""
                                        <tr style="font-size: 12px;" class="lh-1 text-center">
                                            <td>{space_type.replace("_", " ").title()}</td>
                                            <td style="border-right: 2px solid black;">{round(self.baseline_model_summary['total_floor_area_by_space_type'].get(space_type, 0)):,}</td>
                                            <td>{round(self.baseline_model_summary['total_floor_area_by_space_type'][space_type] / self.baseline_model_summary['total_occupants_by_space_type'].get(space_type, math.inf))}</td>
                                            <td>{round(self.baseline_model_summary['total_miscellaneous_equipment_power_by_space_type'].get(space_type, 0) / self.baseline_model_summary['total_floor_area_by_space_type'][space_type], 2)}</td>
                                            <td>{round(self.baseline_lighting_power_allowance_by_space_type.get(space_type, 0) / self.baseline_model_summary['total_floor_area_by_space_type'][space_type], 2)}</td>
                                            <td style="border-right: 2px solid black;">{round(self.baseline_model_summary['total_lighting_power_by_space_type'].get(space_type, 0) / self.baseline_model_summary['total_floor_area_by_space_type'][space_type], 2)}</td>
                                            <td>{round(self.proposed_model_summary['total_lighting_power_by_space_type'].get(space_type, 0) / self.proposed_model_summary['total_floor_area_by_space_type'][space_type], 2)}</td>
                                            <td>{round(self.proposed_model_summary['total_miscellaneous_equipment_power_by_space_type'].get(space_type, 0) / self.proposed_model_summary['total_floor_area_by_space_type'][space_type], 2)}</td>
                                            <td>{round(self.proposed_model_summary['total_floor_area_by_space_type'][space_type] / self.proposed_model_summary['total_occupants_by_space_type'].get(space_type, math.inf))}</td>
                                        </tr>
                    """
                )
            file.write(f"""
                                        <tr  style="font-size: 12px; border-top: 1px solid black;" class="lh-1 fw-bold text-center">
                                            <td>Total</td>
                                            <td style="border-right: 2px solid black;">{round(self.baseline_model_summary['total_floor_area']):,}</td>
                                            <td>{round(self.baseline_model_summary['total_floor_area'] / self.baseline_model_summary['total_occupants'], 2)}</td>
                                            <td>{round(self.baseline_model_summary['total_equipment_power'] / self.baseline_model_summary['total_floor_area'], 2)}</td>
                                            <td>{round(self.baseline_total_lighting_power_allowance / self.baseline_model_summary['total_floor_area'], 2)}</td>
                                            <td style="border-right: 2px solid black;">{round(self.baseline_model_summary['total_lighting_power'] / self.baseline_model_summary['total_floor_area'], 2)}</td>
                                            <td>{round(self.proposed_model_summary['total_lighting_power'] / self.proposed_model_summary['total_floor_area'], 2)}</td>
                                            <td>{round(self.proposed_model_summary['total_equipment_power'] / self.proposed_model_summary['total_floor_area'], 2)}</td>
                                            <td>{round(self.proposed_model_summary['total_floor_area'] / self.proposed_model_summary['total_occupants'], 2)}</td>
                                        </tr>
            """)
            file.write(f"""
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <div class="mb-3 me-4">
                        <button class="btn btn-info collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapse-hvac-summary" aria-expanded="false">
                            HVAC Summary
                        </button>

                        <div id="collapse-hvac-summary" class="accordion-collapse collapse">
                            <div class="accordion-body">
                                <h3>Baseline HVAC Fan Summary</h3>
                                <p><strong>Outdoor Airflow:</strong> {round(self.baseline_model_summary['total_zone_minimum_oa_flow']):,} CFM</p>
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
            """)

            for fan_type in ["Supply", "Return/Relief", "Exhaust", "Zonal Exhaust"]:
                file.write(
                    f"""
                                        <tr style="font-size: 12px;" class="text-center">
                                            <td style="border-right: 2px solid black;">{fan_type}</td>
                                            <td>{round(self.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("CONSTANT", {}).get(fan_type, 0)):,}</td>
                                            <td>{round(self.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("CONSTANT", {}).get(fan_type, 0)/1000, 2):,}</td>
                                            <td style="border-right: 2px solid black;">{round(self.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("CONSTANT", {}).get(fan_type, 0) / (self.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("CONSTANT", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(self.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get(fan_type, 0)):,}</td>
                                            <td>{round(self.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get(fan_type, 0)/1000, 2):,}</td>
                                            <td style="border-right: 2px solid black;">{round(self.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get(fan_type, 0) / (self.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(self.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get(fan_type, 0)):,}</td>
                                            <td>{round(self.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get(fan_type, 0)/1000, 2):,}</td>
                                            <td style="border-right: 2px solid black;">{round(self.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get(fan_type, 0) / (self.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(self.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get(fan_type, 0)):,}</td>
                                            <td>{round(self.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get(fan_type, 0)/1000, 2):,}</td>
                                            <td style="border-right: 2px solid black;">{round(self.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get(fan_type, 0) / (self.baseline_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(self.baseline_model_summary['other_air_flow_by_fan_type'].get(fan_type, 0)):,}</td>
                                            <td>{round(self.baseline_model_summary['other_fan_power_by_fan_type'].get(fan_type, 0)/1000, 2):,}</td>
                                            <td style="border-right: 2px solid black;">{round(self.baseline_model_summary['other_fan_power_by_fan_type'].get(fan_type, 0) / (self.baseline_model_summary['other_air_flow_by_fan_type'].get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(self.baseline_model_summary['total_air_flow_by_fan_type'].get(fan_type, 0)):,}</td>
                                            <td>{round(self.baseline_model_summary['total_fan_power_by_fan_type'].get(fan_type, 0)/1000, 2):,}</td>
                                            <td>{round(self.baseline_model_summary['total_fan_power_by_fan_type'].get(fan_type, 0) / (self.baseline_model_summary['total_air_flow_by_fan_type'].get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(100 * self.baseline_model_summary['total_fan_power_by_fan_type'].get(fan_type, 0) / sum(self.baseline_model_summary["total_fan_power_by_fan_type"].values()))}</td>
                                        </tr>
                    """
                )
            # ------------------------- Subtotal Row --------------------------------
            file.write(f"""
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
                        """)
            # ------------------------- Terminal Units Row --------------------------------
            file.write(f"""
                                        <tr style="font-size: 12px; border-top: 1px solid black;" class="text-center">
                                            <td style="border-right: 2px solid black;">Terminal Units</td>
                                            <td style="background: black;"></td>
                                            <td>{round(self.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("CONSTANT", {}).get("Terminal Unit", 0)):,}</td>
                                            <td style="border-right: 2px solid black; background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(self.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get("Terminal Unit", 0)):,}</td>
                                            <td style="border-right: 2px solid black; background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(self.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get("Terminal Unit", 0)):,}</td>
                                            <td style="border-right: 2px solid black; background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(self.baseline_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get("Terminal Unit", 0)):,}</td>
                                            <td style="border-right: 2px solid black; background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(self.baseline_model_summary['other_fan_power_by_fan_type'].get("Terminal Unit", 0)):,}</td>
                                            <td style="border-right: 2px solid black; background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(self.baseline_model_summary['total_fan_power_by_fan_type'].get("Terminal Unit", 0)):,}</td>
                                            <td style="background: black;"></td>
                                            <td style="background: black;"></td>
                                        </tr>
            """)
            file.write(f"""
                                    </tbody>
                                </table>
                                
                                <h3>Proposed HVAC Fan Summary</h3>
                                <p><strong>Outdoor Airflow:</strong> {round(self.baseline_model_summary['total_zone_minimum_oa_flow']):,} CFM</p>
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
            """)

            for fan_type in ["Supply", "Return/Relief", "Exhaust", "Zonal Exhaust"]:
                file.write(
                    f"""
                                        <tr style="font-size: 12px;" class="text-center">
                                            <td style="border-right: 2px solid black;">{fan_type}</td>
                                            <td>{round(self.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("CONSTANT", {}).get(fan_type, 0)):,}</td>
                                            <td>{round(self.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("CONSTANT", {}).get(fan_type, 0)/1000, 2):,}</td>
                                            <td style="border-right: 2px solid black;">{round(self.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("CONSTANT", {}).get(fan_type, 0) / (self.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("CONSTANT", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(self.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get(fan_type, 0)):,}</td>
                                            <td>{round(self.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get(fan_type, 0)/1000, 2):,}</td>
                                            <td style="border-right: 2px solid black;">{round(self.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get(fan_type, 0) / (self.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(self.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get(fan_type, 0)):,}</td>
                                            <td>{round(self.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get(fan_type, 0)/1000, 2):,}</td>
                                            <td style="border-right: 2px solid black;">{round(self.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get(fan_type, 0) / (self.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(self.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get(fan_type, 0)):,}</td>
                                            <td>{round(self.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get(fan_type, 0)/1000, 2):,}</td>
                                            <td style="border-right: 2px solid black;">{round(self.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get(fan_type, 0) / (self.proposed_model_summary['total_air_flow_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(self.proposed_model_summary['other_air_flow_by_fan_type'].get(fan_type, 0)):,}</td>
                                            <td>{round(self.proposed_model_summary['other_fan_power_by_fan_type'].get(fan_type, 0)/1000, 2):,}</td>
                                            <td style="border-right: 2px solid black;">{round(self.proposed_model_summary['other_fan_power_by_fan_type'].get(fan_type, 0) / (self.proposed_model_summary['other_air_flow_by_fan_type'].get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(self.proposed_model_summary['total_air_flow_by_fan_type'].get(fan_type, 0)):,}</td>
                                            <td>{round(self.proposed_model_summary['total_fan_power_by_fan_type'].get(fan_type, 0)/1000, 2):,}</td>
                                            <td>{round(self.proposed_model_summary['total_fan_power_by_fan_type'].get(fan_type, 0) / (self.proposed_model_summary['total_air_flow_by_fan_type'].get("Supply", 99999999) or 99999999), 4)}</td>
                                            <td>{round(100 * self.proposed_model_summary['total_fan_power_by_fan_type'].get(fan_type, 0) / sum(self.proposed_model_summary["total_fan_power_by_fan_type"].values()))}</td>
                                        </tr>
                    """
                )
            # ------------------------- Subtotal Row --------------------------------
            file.write(f"""
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
            """)
            # ------------------------- Terminal Units Row --------------------------------
            file.write(f"""
                                        <tr style="font-size: 12px; border-top: 1px solid black;" class="text-center">
                                            <td>Terminal Units</td>
                                            <td style="background: black;"></td>
                                            <td>{round(self.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("CONSTANT", {}).get("Terminal Unit", 0)):,}</td>
                                            <td style="border-right: 2px solid black; background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(self.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("VARIABLE_SPEED_DRIVE", {}).get("Terminal Unit", 0)):,}</td>
                                            <td style="border-right: 2px solid black; background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(self.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("MULTISPEED", {}).get("Terminal Unit", 0)):,}</td>
                                            <td style="border-right: 2px solid black; background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(self.proposed_model_summary['total_fan_power_by_fan_control_by_fan_type'].get("Constant Cycling", {}).get("Terminal Unit", 0)):,}</td>
                                            <td style="border-right: 2px solid black; background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(self.proposed_model_summary['other_fan_power_by_fan_type'].get("Terminal Unit", 0)):,}</td>
                                            <td style="border-right: 2px solid black; background: black;"></td>
                                            <td style="background: black;"></td>
                                            <td>{round(self.proposed_model_summary['total_fan_power_by_fan_type'].get("Terminal Unit", 0)):,}</td>
                                            <td style="background: black;"></td>
                                            <td style="background: black;"></td>
                                        </tr>
                        """)
            file.write(f"""
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                """
                       )

            for category, rules in rule_categories.items():
                btn_class = (
                    "btn-danger"
                    if category == "Failing"
                    else "btn-warning"
                    if category == "Undetermined"
                    else "btn-success"
                    if category == "Passing"
                    else "btn-secondary"
                )
                file.write(
                    f"""
                        <div class="mb-3 me-4">
                            <button class="btn {btn_class} w-100 text-start sticky-top" 
                                type="button" data-bs-toggle="collapse" data-bs-target="#collapse_fully_{category.replace(' ', '_')}">
                                <strong>{category} Rules ({len(rules)})</strong>
                            </button>
                            <div class="collapse mx-4" id="collapse_fully_{category.replace(' ', '_')}">
                                <h3 class="mt-4">Rules Fully Evaluated</h3>
                                <table class="table table-bordered table-striped mt-2">
                                    <thead class="table-dark">
                                        <tr>
                                            <th rowspan='2'>Rule ID</th>
                                            <th>Description</th>
                                            <th>Standard Section</th>
                                            <th>Outcome Counts</th>
                                        </tr>
                                        <tr><th colspan='3'>Evaluations</th></tr>
                                    </thead>
                                    <tbody>
                        """
                )

                if category == "Undetermined":
                    sections_seen = set()
                    for rule_id in self.full_eval_rules_undetermined:
                        rule_data = next(
                            rule
                            for rule in self.evaluation_data["rules"]
                            if rule["rule_id"] == rule_id
                        )
                        section = rule_id.split("-")[0]
                        if section not in sections_seen:
                            sections_seen.add(section)
                            section_title = section_titles_with_colors.get(
                                int(section)
                            )[0]
                            section_color = section_titles_with_colors.get(
                                int(section)
                            )[1]
                            file.write(
                                f"""
                                </tbody>
                                    <thead class="table-group-divider">
                                        <tr>
                                            <td colspan="4" class="section-title sticky-top sticky-top-2" style="background-color: {section_color} !important;">{section_title}</td>
                                        </tr>
                                    </thead>
                                <tbody>
                                """
                            )

                        description = rule_data.get("description", "N/A")
                        standard_section = rule_data.get("standard_section", "N/A")
                        outcome_summary = " | ".join([f"{k}: {v}" for k, v in self.rule_evaluation_outcome_counts[rule_id].items()])

                        file.write(
                            f"""
                                <tr>
                                    <td class="rule-id" rowspan='2'>{rule_id}</td>
                                    <td>{description}</td>
                                    <td>{standard_section}</td>
                                    <td class="outcome-summary">{outcome_summary}</td>
                                </tr>
                                <tr>
                                    <td colspan='3'>
                                        <button class="btn btn-primary" type="button" data-bs-toggle="collapse" data-bs-target="#eval_{rule_id}">
                                            View Evaluations
                                        </button>
                                        <div class="collapse" id="eval_{rule_id}">
                                            <ul>
                            """
                        )
                        outcome_order = {
                            "FAILED": 0,
                            "UNDETERMINED": 1,
                            "PASS": 2,
                            "NOT_APPLICABLE": 3,
                        }

                        # Sort evaluations based on outcome priority
                        sorted_evaluations = sorted(
                            rule_data["evaluations"],
                            key=lambda e: outcome_order.get(e["outcome"], 3),
                        )

                        for evaluation in sorted_evaluations:
                            has_any_units = False
                            styles = {
                                "FAILED": "background-color: #ffcccc; color: black; font-weight: bold; padding-left: 10px; border-radius: 8px; border: 2px solid #ff0000;",
                                "PASS": "background-color: #ccffcc; color: black; font-weight: bold; padding-left: 10px; border-radius: 8px; border: 2px solid #008000;",
                                "UNDETERMINED": "background-color: #ffffcc; color: black; font-weight: bold; padding-left: 10px; border-radius: 8px; border: 2px solid #ffcc00;",
                                "DEFAULT": "padding-left: 10px; border: 2px solid #ccc; border-radius: 8px;",
                            }

                            # Select the appropriate style based on outcome
                            li_style = styles.get(
                                evaluation["outcome"], styles["DEFAULT"]
                            )
                            file.write(
                                f"""
                                    <li style=\"{li_style}\"  class=\"p-2 m-1\">{evaluation['data_group_id']}
                                        <ul>
                                            <li><strong>Outcome:</strong> {evaluation['outcome']}</li>
                                    """
                            )
                            if evaluation["messages"]:
                                messages = set()
                                if isinstance(evaluation["messages"], str):
                                    messages.add(evaluation["messages"])
                                if isinstance(evaluation["messages"], dict):
                                    for key, message in evaluation["messages"].items():
                                        messages.add(f"{key}: {message}")
                                if isinstance(evaluation["messages"], list):
                                    for message in evaluation["messages"]:
                                        messages.add(message)
                                file.write(
                                    f"<li><strong>Messages:</strong> {', '.join(messages)}</li>"
                                )
                            if evaluation["calculated_values"]:
                                file.write(
                                    """
                                        <li><strong>Calculated Values:</strong>
                                            <table class="mb-2 me-2 table table-sm table-bordered">
                                                <thead>
                                                    <tr><th>Variable</th><th>Value</th>
                                    """
                                )
                                if any(
                                        cv.get("unit")
                                        for cv in evaluation["calculated_values"]
                                ):
                                    has_any_units = True
                                    file.write("<th>Unit</th>")
                                file.write("</tr></thead><tbody>")

                                for calculated_value in evaluation["calculated_values"]:
                                    file.write(
                                        f"""
                                        <tr>
                                        <td>{calculated_value['variable']}</td>
                                        <td>{calculated_value['value'][0] if len(calculated_value['value']) == 1
                                        else calculated_value['value']}
                                        </td>
                                        """
                                    )
                                    if calculated_value.get("unit"):
                                        file.write(
                                            f"<td>{calculated_value['unit']}</td>"
                                        )
                                    elif has_any_units:
                                        file.write("<td></td>")
                                    file.write("</tr>")
                                file.write("</tbody></table></li>")
                            file.write("</ul></li>")
                        file.write("</ul></div></td></tr>")
                    file.write(
                        f"""
                            </tbody>
                            </table>
                            <h3 class="mt-4">Rules Evaluated for Applicability Only</h3>
                            <table class="table table-bordered table-striped mt-2">
                                <thead class="table-dark">
                                    <tr>
                                        <th rowspan='2'>Rule ID</th>
                                        <th>Description</th>
                                        <th>Standard Section</th>
                                        <th>Outcome Counts</th>
                                    </tr>
                                    <tr><th colspan='3'>Evaluations</th></tr>
                                </thead>
                                <tbody>
                        """
                    )
                    sections_seen = set()
                    for rule_id in self.appl_eval_rules_undetermined:
                        rule_data = next(
                            rule
                            for rule in self.evaluation_data["rules"]
                            if rule["rule_id"] == rule_id
                        )
                        section = rule_id.split("-")[0]
                        if section not in sections_seen:
                            sections_seen.add(section)
                            section_title = section_titles_with_colors.get(
                                int(section)
                            )[0]
                            section_color = section_titles_with_colors.get(
                                int(section)
                            )[1]
                            file.write(
                                f"""
                                </tbody>
                                    <thead class="table-group-divider">
                                        <tr>
                                            <td colspan="4" class="section-title sticky-top sticky-top-2" style="background-color: {section_color} !important;">{section_title}</td>
                                        </tr>
                                    </thead>
                                <tbody>
                                """
                            )

                        description = rule_data.get("description", "N/A")
                        standard_section = rule_data.get("standard_section", "N/A")
                        outcome_summary = " | ".join([f"{k}: {v}" for k, v in self.rule_evaluation_outcome_counts[rule_id].items()])

                        file.write(
                            f"""
                                <tr>
                                    <td class="rule-id" rowspan='2'>{rule_id}</td>
                                    <td>{description}</td>
                                    <td>{standard_section}</td>
                                    <td class="outcome-summary">{outcome_summary}</td>
                                </tr>
                                <tr>
                                    <td colspan='3'>
                                        <button class="btn btn-primary" type="button" data-bs-toggle="collapse" data-bs-target="#eval_{rule_id}">
                                            View Evaluations
                                        </button>
                                        <div class="collapse" id="eval_{rule_id}">
                                            <ul>
                                """
                        )
                        outcome_order = {
                            "FAILED": 0,
                            "UNDETERMINED": 1,
                            "PASS": 2,
                            "NOT_APPLICABLE": 3,
                        }

                        # Sort evaluations based on outcome priority
                        sorted_evaluations = sorted(
                            rule_data["evaluations"],
                            key=lambda e: outcome_order.get(e["outcome"], 3),
                        )

                        for evaluation in sorted_evaluations:
                            has_any_units = False
                            styles = {
                                "FAILED": "background-color: #ffcccc; color: black; font-weight: bold; padding-left: 10px; border-radius: 8px; border: 2px solid #ff0000;",
                                "PASS": "background-color: #ccffcc; color: black; font-weight: bold; padding-left: 10px; border-radius: 8px; border: 2px solid #008000;",
                                "UNDETERMINED": "background-color: #ffffcc; color: black; font-weight: bold; padding-left: 10px; border-radius: 8px; border: 2px solid #ffcc00;",
                                "DEFAULT": "padding-left: 10px; border: 2px solid #ccc; border-radius: 8px;",
                            }

                            # Select the appropriate style based on outcome
                            li_style = styles.get(
                                evaluation["outcome"], styles["DEFAULT"]
                            )
                            file.write(
                                f"""
                                    <li style=\"{li_style}\"  class=\"p-2 m-1\">{evaluation['data_group_id']}
                                        <ul>
                                            <li><strong>Outcome:</strong> {evaluation['outcome']}</li>
                                    """
                            )
                            if evaluation["messages"]:
                                messages = set()
                                if isinstance(evaluation["messages"], str):
                                    messages.add(evaluation["messages"])
                                if isinstance(evaluation["messages"], dict):
                                    for key, message in evaluation["messages"].items():
                                        messages.add(f"{key}: {message}")
                                if isinstance(evaluation["messages"], list):
                                    for message in evaluation["messages"]:
                                        messages.add(message)
                                file.write(
                                    f"<li><strong>Messages:</strong> {', '.join(messages)}</li>"
                                )
                            if evaluation["calculated_values"]:
                                file.write(
                                    """
                                        <li><strong>Calculated Values:</strong>
                                            <table class="mb-2 me-2 table table-sm table-bordered">
                                                <thead>
                                                    <tr><th>Variable</th><th>Value</th>
                                    """
                                )
                                if any(
                                        cv.get("unit")
                                        for cv in evaluation["calculated_values"]
                                ):
                                    has_any_units = True
                                    file.write("<th>Unit</th>")
                                file.write("</tr></thead><tbody>")

                                for calculated_value in evaluation["calculated_values"]:
                                    file.write(
                                        f"""
                                        <tr>
                                        <td>{calculated_value['variable']}</td>
                                        <td>{calculated_value['value'][0] if len(calculated_value['value']) == 1
                                        else calculated_value['value']}
                                        </td>
                                        """
                                    )
                                    if calculated_value.get("unit"):
                                        file.write(
                                            f"<td>{calculated_value['unit']}</td>"
                                        )
                                    elif has_any_units:
                                        file.write("<td></td>")
                                    file.write("</tr>")
                                file.write("</tbody></table></li>")
                            file.write("</ul></li>")
                        file.write("</ul></div></td></tr>")
                else:
                    sections_seen = set()
                    for rule_id in rules:
                        rule_data = next(
                            rule
                            for rule in self.evaluation_data["rules"]
                            if rule["rule_id"] == rule_id
                        )
                        section = rule_id.split("-")[0]
                        if section not in sections_seen:
                            sections_seen.add(section)
                            section_title = section_titles_with_colors.get(
                                int(section)
                            )[0]
                            section_color = section_titles_with_colors.get(
                                int(section)
                            )[1]
                            file.write(
                                f"""
                                </tbody>
                                    <thead class="table-group-divider">
                                        <tr>
                                            <th colspan="4" class="section-title sticky-top sticky-top-2" style="background-color: {section_color} !important;">{section_title}</th>
                                        </tr>
                                    </thead>
                                <tbody>
                                """
                            )

                        description = rule_data.get("description", "N/A")
                        standard_section = rule_data.get("standard_section", "N/A")
                        outcome_summary = " | ".join([f"{k}: {v}" for k, v in self.rule_evaluation_outcome_counts[rule_id].items()])

                        file.write(
                            f"""
                                <tr>
                                    <td class="rule-id" rowspan='2'>{rule_id}</td>
                                    <td>{description}</td>
                                    <td>{standard_section}</td>
                                    <td class="outcome-summary">{outcome_summary}</td>
                                </tr>
                                <tr>
                                    <td colspan='3'>
                                        <button class="btn btn-primary" type="button" data-bs-toggle="collapse" data-bs-target="#eval_{rule_id}">
                                            View Evaluations
                                        </button>
                                        <div class="collapse" id="eval_{rule_id}">
                                            <ul>
                            """
                        )
                        outcome_order = {
                            "FAILED": 0,
                            "UNDETERMINED": 1,
                            "PASS": 2,
                            "NOT_APPLICABLE": 3,
                        }

                        # Sort evaluations based on outcome priority
                        sorted_evaluations = sorted(
                            rule_data["evaluations"],
                            key=lambda e: outcome_order.get(e["outcome"], 3),
                        )

                        for evaluation in sorted_evaluations:
                            has_any_units = False
                            styles = {
                                "FAILED": "background-color: #ffcccc; color: black; font-weight: bold; padding-left: 10px; border-radius: 8px; border: 2px solid #ff0000;",
                                "PASS": "background-color: #ccffcc; color: black; font-weight: bold; padding-left: 10px; border-radius: 8px; border: 2px solid #008000;",
                                "UNDETERMINED": "background-color: #ffffcc; color: black; font-weight: bold; padding-left: 10px; border-radius: 8px; border: 2px solid #ffcc00;",
                                "DEFAULT": "padding-left: 10px; border: 2px solid #ccc; border-radius: 8px;",
                            }

                            # Select the appropriate style based on outcome
                            li_style = styles.get(
                                evaluation["outcome"], styles["DEFAULT"]
                            )
                            file.write(
                                f"""
                                    <li style=\"{li_style}\"  class=\"p-2 m-1\">{evaluation['data_group_id']}
                                        <ul>
                                            <li><strong>Outcome:</strong> {evaluation['outcome']}</li>
                                    """
                            )
                            if evaluation["messages"]:
                                messages = set()
                                if isinstance(evaluation["messages"], str):
                                    messages.add(evaluation["messages"])
                                if isinstance(evaluation["messages"], dict):
                                    for key, message in evaluation["messages"].items():
                                        messages.add(f"{key}: {message}")
                                if isinstance(evaluation["messages"], list):
                                    for message in evaluation["messages"]:
                                        messages.add(message)
                                file.write(
                                    f"<li><strong>Messages:</strong> {', '.join(messages)}</li>"
                                )
                            if evaluation["calculated_values"]:
                                file.write(
                                    """
                                        <li><strong>Calculated Values:</strong>
                                            <table class="mb-2 me-2 table table-sm table-bordered">
                                                <thead>
                                                    <tr><th>Variable</th><th>Value</th>
                                    """
                                )
                                if any(
                                        cv.get("unit")
                                        for cv in evaluation["calculated_values"]
                                ):
                                    has_any_units = True
                                    file.write("<th>Unit</th>")
                                file.write("</tr></thead><tbody>")

                                for calculated_value in evaluation["calculated_values"]:
                                    file.write(
                                        f"""
                                        <tr>
                                        <td>{calculated_value['variable']}</td>
                                        <td>{calculated_value['value'][0] if len(calculated_value['value']) == 1
                                        else calculated_value['value']}
                                        </td>
                                        """
                                    )
                                    if calculated_value.get("unit"):
                                        file.write(
                                            f"<td>{calculated_value['unit']}</td>"
                                        )
                                    elif has_any_units:
                                        file.write("<td></td>")
                                    file.write("</tr>")
                                file.write("</tbody></table></li>")
                            file.write("</ul></li>")
                        file.write("</ul></div></td></tr>")

                file.write("</tbody></table></div></div>")

            file.write("</div></div>")
            file.write(
                """
            <div class="position-fixed bottom-0 end-0 mb-2 me-2" style="z-index: 1050;">
                <button id="back-to-top" class="btn btn-primary" onclick="scrollToTop()" style="opacity: 0; visibility: hidden;"> ↑ </button>
            </div>
            """
            )
            file.write("</body>")
            file.write(
                """
            <script>
            window.onscroll = function() {
                toggleBackToTopButton();
            };
            
            function toggleBackToTopButton() {
                const backToTopButton = document.getElementById("back-to-top");
                if (document.body.scrollTop > 100 || document.documentElement.scrollTop > 100) {
                    backToTopButton.style.opacity = "1";
                    backToTopButton.style.visibility = "visible";
                }    
                else {
                    backToTopButton.style.opacity = "0";
                    backToTopButton.style.visibility = "hidden";
                }
            }
            
            function scrollToTop() {
                window.scrollTo({
                    top: 0,
                    behavior: 'smooth'
                });
            }
            
            function calculateSubtotals() {
                document.querySelectorAll(".fan-summary").forEach(table => {
                    let columnSums = [];
                    let columnPrecisions = [];
            
                    table.querySelectorAll("tr").forEach(row => {
                        if (row.classList.contains("subtotal")) {
                            // Populate the subtotal row with column sums
                            row.querySelectorAll("td").forEach((td, colIndex) => {
                                if (colIndex === 0) return;
                                let sum = columnSums[colIndex] || 0;
                                let precision = columnPrecisions[colIndex] || 0;
                                td.textContent = sum.toLocaleString(undefined, { minimumFractionDigits: precision, maximumFractionDigits: precision });
                            });
            
                            // Reset the column sums and precisions after each subtotal row
                            columnSums = [];
                            columnPrecisions = [];
                        } else {
                            // Sum values in the current row and track precision
                            row.querySelectorAll("td").forEach((td, colIndex) => {
                                let cleanedText = td.textContent.replace(/,/g, "").trim();
                                let value = parseFloat(cleanedText) || 0;
            
                                // Determine decimal precision
                                let decimalPlaces = (cleanedText.split(".")[1] || "").length;
                                columnPrecisions[colIndex] = Math.max(columnPrecisions[colIndex] || 0, decimalPlaces);
            
                                // Sum values
                                columnSums[colIndex] = (columnSums[colIndex] || 0) + value;
                            });
                        }
                    });
                });
            }
            
            document.addEventListener("DOMContentLoaded", () => {
                calculateSubtotals();
            });
            </script>
            """
            )
            file.write("</html>")

    def run(self):
        self.load_files()
        self.extract_evaluation_data()
        self.extract_model_data()
        self.perform_analytic_calculations()
        self.convert_model_data_units()
        self.write_html_file()
