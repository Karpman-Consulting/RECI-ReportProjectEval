import json
import re

from rctreportviewer.html.write_evaluation_html import (
    write_html_file as write_evaluation_summary_html_file,
)
from rctreportviewer.html.write_model_html import (
    write_html_file as write_model_summary_html_file,
)
from rctreportviewer.html.write_html import write_html_file
from rctreportviewer.constants import (
    path_to_bpf_data,
    model_type_disp_map,
    outcome_disp_map,
)
from rctreportviewer.analytics import perform_analytic_calculations
from rctreportviewer.converters import convert_model_data_units
from rctreportviewer.summarizers.rmd import summarize_rmd_data


class SummaryReportGenerator:
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
        self.bpf_data = None
        self.ruleset = None

        self.model_types = set()
        self.space_areas = {}
        self.baseline_space_types = {}
        self.space_lpd_allowances = {}
        self.hvac_system_types_b = {}
        self.baseline_total_lighting_power_allowance = 0
        self.baseline_lighting_power_allowance_by_space_type = {}
        self.bpfs_by_metric = {}
        self.proposed_model_summary = {}
        self.baseline_model_summary = {}

        self.rules_passed = []  # ALL outcomes are PASS or N/A
        self.rules_failed = []  # ANY outcome is FAIL
        self.full_eval_rules_undetermined = []  # ANY outcome is UNDETERMINED
        self.appl_eval_rules_undetermined = []  # ANY outcome is UNDETERMINED
        self.rules_undetermined_missing_data = (
            []
        )  # Undetermined specifically due to missing data
        self.rules_not_applicable = []  # ALL outcomes are N/A
        self.rule_evaluation_outcome_counts = {}
        self.rule_evaluation_message_counts = {}

    @staticmethod
    def load_file(file_path):
        """
        Reads a JSON file and returns the python equivalent data structure.
        """
        # Verify the file path is to a JSON file extension
        if not file_path.endswith((".json", ".rpd")):
            raise ValueError("Invalid file type. Please provide a JSON file.")

        with open(file_path, "r") as file:
            data = json.load(file)

        return data

    def load_files(self):
        """
        Loads the JSON files into memory that are needed to produce the HTML report.
        """
        self.evaluation_data = self.load_file(self.detailed_evaluation_report_file_path)
        self.rpd_data = [self.load_file(file_path) for file_path in self.rpd_file_paths]
        self.bpf_data = self.load_file(path_to_bpf_data)

    def load_evaluation_file(self):
        """
        Loads only the evaluation JSON file into memory.
        """
        self.evaluation_data = self.load_file(self.detailed_evaluation_report_file_path)

    def load_model_files(self):
        """
        Loads only the RPD JSON file(s) into memory.
        """
        self.rpd_data = [self.load_file(file_path) for file_path in self.rpd_file_paths]
        self.bpf_data = self.load_file(path_to_bpf_data)

    def extract_evaluation_data(self):
        """
        Extracts select evaluation data from the overall data structure for reformatting and easy presentation.
        """
        self.ruleset = self.evaluation_data.get("ruleset")

        for rpd_file in self.evaluation_data["rpd_files"]:
            self.model_types.add(
                model_type_disp_map.get(rpd_file["ruleset_model_type"])
            )

        for rule in self.evaluation_data["rules"]:
            rule_id = rule["rule_id"]
            eval_type = rule["evaluation_type"]
            outcomes = set()
            messages = set()
            has_undetermined_missing_data = False

            # Initialize the nested dictionaries if the rule_id is new
            if rule_id not in self.rule_evaluation_outcome_counts:
                self.rule_evaluation_outcome_counts[rule_id] = {}
            if rule_id not in self.rule_evaluation_message_counts:
                self.rule_evaluation_message_counts[rule_id] = {}

            for evaluation in rule["evaluations"]:
                outcome = outcome_disp_map.get(evaluation["outcome"])
                outcomes.add(outcome)

                # Normalize messages into a list of strings
                eval_messages = []
                if isinstance(evaluation["messages"], str):
                    eval_messages = [evaluation["messages"]]
                elif isinstance(evaluation["messages"], dict):
                    eval_messages = [
                        f"{k}: {v}" for k, v in evaluation["messages"].items()
                    ]
                elif isinstance(evaluation["messages"], list):
                    eval_messages = evaluation["messages"]

                for msg in eval_messages:
                    messages.add(msg)

                    if outcome == "Undetermined" and (
                        "missing:" in msg.lower()
                        or re.search(
                            r"at least one .* value must exist", msg, re.IGNORECASE
                        )
                    ):
                        has_undetermined_missing_data = True

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

                # if rule_id == "6-4" and "calculated_values" in evaluation:
                #     lpd_allowance_calc_value = next(
                #         (
                #             calc_value
                #             for calc_value in evaluation["calculated_values"]
                #             if calc_value["variable"] == "lpd_allowance_b"
                #         ),
                #         None,
                #     )
                #     if lpd_allowance_calc_value:
                #         self.space_lpd_allowances[evaluation["data_group_id"]] = float(
                #             lpd_allowance_calc_value["value"]
                #         )
                #
                # if (
                #     rule_id == "18-1"
                #     and "calculated_values" in evaluation
                #     and not self.hvac_system_types_b
                # ):
                #     hvac_system_types_b_value = next(
                #         (
                #             calc_value
                #             for calc_value in evaluation["calculated_values"]
                #             if calc_value["variable"] == "hvac_system_types_b"
                #         ),
                #         None,
                #     )
                #     if hvac_system_types_b_value:
                #         self.hvac_system_types_b = ast.literal_eval(
                #             hvac_system_types_b_value["value"]
                #         )

            # Determine rule status
            if outcomes == {"Failing"} and messages == {" ::TOLERANCE::"}:
                self.rules_passed.append(rule_id)
            elif "Failing" in outcomes:
                self.rules_failed.append(rule_id)
            elif outcomes == {"Passing"} or outcomes == {"Passing", "N/A"}:
                self.rules_passed.append(rule_id)
            elif outcomes == {"N/A"}:
                self.rules_not_applicable.append(rule_id)
            elif "Undetermined" in outcomes and not has_undetermined_missing_data:
                if eval_type == "FULL":
                    self.full_eval_rules_undetermined.append(rule_id)
                elif eval_type == "APPLICABILITY":
                    self.appl_eval_rules_undetermined.append(rule_id)

            # Only missing-data undetermined
            elif has_undetermined_missing_data:
                self.rules_undetermined_missing_data.append(rule_id)

    def extract_model_data(self):
        if len(self.rpd_data) == 1:
            self.rpd_data = self.rpd_data[0]

        else:
            merged = self.rpd_data[0]

            for rpd in self.rpd_data[1:]:
                # Extend the ruleset_model_descriptions list
                merged["ruleset_model_descriptions"].extend(
                    rpd["ruleset_model_descriptions"]
                )

                for key in rpd:
                    if key == "ruleset_model_descriptions":
                        continue  # already handled

                    val = rpd[key]
                    if key not in merged:
                        merged[key] = val
                        continue

                    # Merge non-list, non-dict values (e.g., strings and numbers)
                    if not isinstance(val, (list, dict)):
                        if merged[key] != val:
                            # TODO log this conflict visible to users
                            print(
                                "Conflicting value for key:",
                                key,
                                " keeping the first occurrence",
                            )
                            pass
                        continue

                    # Merge dicts like metadata and output, or calendar and weather in older versions of the schema
                    if isinstance(val, dict):
                        if not isinstance(merged[key], dict):
                            continue  # mismatched types, skip or log error
                        for subkey, subval in val.items():
                            if subkey not in merged[key]:
                                merged[key][subkey] = subval
                            elif merged[key][subkey] != subval:
                                # TODO log this conflict visible to users
                                print(
                                    "Conflicting value for key:",
                                    subkey,
                                    " keeping the first occurrence",
                                )
                                pass

            self.rpd_data = merged

        proposed_rmd = next(
            (
                rmd
                for rmd in self.rpd_data["ruleset_model_descriptions"]
                if rmd["type"] == "PROPOSED"
            ),
            None,
        )
        baseline_rmd = next(
            (
                rmd
                for rmd in self.rpd_data["ruleset_model_descriptions"]
                if rmd["type"] == "BASELINE_0"
            ),
            None,
        )
        if not proposed_rmd or not baseline_rmd:
            # TODO Handle the case where the proposed or baseline RMD is not found
            print("Proposed or Baseline RMD not found in the RPD data.")
            return

        self.proposed_model_summary = summarize_rmd_data(
            self, proposed_rmd, model_type="Proposed"
        )
        self.baseline_model_summary = summarize_rmd_data(
            self, baseline_rmd, model_type="Baseline"
        )

    def run(self):
        self.load_files()
        self.extract_evaluation_data()
        self.extract_model_data()
        perform_analytic_calculations(self)
        convert_model_data_units(
            self.baseline_model_summary, self.proposed_model_summary
        )
        write_html_file(self)

    def summarize_models(self):
        self.load_model_files()
        self.extract_model_data()
        perform_analytic_calculations(self)
        convert_model_data_units(
            self.baseline_model_summary, self.proposed_model_summary
        )
        write_model_summary_html_file(self)

    def summarize_evaluations(self):
        self.load_evaluation_file()
        self.extract_evaluation_data()
        write_evaluation_summary_html_file(self)
