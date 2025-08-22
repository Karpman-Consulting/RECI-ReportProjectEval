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


def write_evaluations_section(file, rct_detailed_report):
    rule_categories = {
        "Failing": rct_detailed_report.rules_failed,
        "Passing": rct_detailed_report.rules_passed,
        "Undetermined": rct_detailed_report.full_eval_rules_undetermined
        + rct_detailed_report.appl_eval_rules_undetermined,
        "N/A": rct_detailed_report.rules_not_applicable,
    }

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
            """
        )
        if category == "Undetermined":
            file.write(
                f"""
                <h3 class="mt-4">Rules Fully Evaluated</h3>
                """
            )
        file.write(
            f"""
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
            for rule_id in rct_detailed_report.full_eval_rules_undetermined:
                rule_data = next(
                    rule
                    for rule in rct_detailed_report.evaluation_data["rules"]
                    if rule["rule_id"] == rule_id
                )
                section = rule_id.split("-")[0]
                if section not in sections_seen:
                    sections_seen.add(section)
                    section_title = section_titles_with_colors.get(int(section))[0]
                    section_color = section_titles_with_colors.get(int(section))[1]
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
                outcome_summary = " | ".join(
                    [
                        f"{k}: {v}"
                        for k, v in rct_detailed_report.rule_evaluation_outcome_counts[
                            rule_id
                        ].items()
                    ]
                )

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
                    li_style = styles.get(evaluation["outcome"], styles["DEFAULT"])
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
                            cv.get("unit") for cv in evaluation["calculated_values"]
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
                                file.write(f"<td>{calculated_value['unit']}</td>")
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
            for rule_id in rct_detailed_report.appl_eval_rules_undetermined:
                rule_data = next(
                    rule
                    for rule in rct_detailed_report.evaluation_data["rules"]
                    if rule["rule_id"] == rule_id
                )
                section = rule_id.split("-")[0]
                if section not in sections_seen:
                    sections_seen.add(section)
                    section_title = section_titles_with_colors.get(int(section))[0]
                    section_color = section_titles_with_colors.get(int(section))[1]
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
                outcome_summary = " | ".join(
                    [
                        f"{k}: {v}"
                        for k, v in rct_detailed_report.rule_evaluation_outcome_counts[
                            rule_id
                        ].items()
                    ]
                )

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
                    li_style = styles.get(evaluation["outcome"], styles["DEFAULT"])
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
                            cv.get("unit") for cv in evaluation["calculated_values"]
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
                                file.write(f"<td>{calculated_value['unit']}</td>")
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
                    for rule in rct_detailed_report.evaluation_data["rules"]
                    if rule["rule_id"] == rule_id
                )
                section = rule_id.split("-")[0]
                if section not in sections_seen:
                    sections_seen.add(section)
                    section_title = section_titles_with_colors.get(int(section))[0]
                    section_color = section_titles_with_colors.get(int(section))[1]
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
                outcome_summary = " | ".join(
                    [
                        f"{k}: {v}"
                        for k, v in rct_detailed_report.rule_evaluation_outcome_counts[
                            rule_id
                        ].items()
                    ]
                )

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
                    li_style = styles.get(evaluation["outcome"], styles["DEFAULT"])
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
                            cv.get("unit") for cv in evaluation["calculated_values"]
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
                                file.write(f"<td>{calculated_value['unit']}</td>")
                            elif has_any_units:
                                file.write("<td></td>")
                            file.write("</tr>")
                        file.write("</tbody></table></li>")
                    file.write("</ul></li>")
                file.write("</ul></div></td></tr>")

        file.write("</tbody></table></div></div>")
