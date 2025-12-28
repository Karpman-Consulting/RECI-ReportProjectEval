import re
import ast


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


def is_missing_data_evaluation(evaluation) -> bool:
    if evaluation.get("outcome") != "UNDETERMINED":
        return False

    msgs = evaluation.get("messages", [])
    if isinstance(msgs, str):
        msgs = [msgs]
    elif isinstance(msgs, dict):
        msgs = msgs.values()

    for msg in msgs:
        msg_l = msg.lower()
        if (
            "missing:" in msg_l
            or "missing_" in msg_l
            or "is missing " in msg_l
            or re.search(r"at least one .* value must exist", msg_l)
        ):
            return True
    return False


def split_evaluations(rule_data):
    missing = []
    standard = []

    for ev in rule_data["evaluations"]:
        if is_missing_data_evaluation(ev):
            missing.append(ev)
        else:
            standard.append(ev)

    return standard, missing


def format_value_for_html(value):
    # Try to parse strings into native Python types (list/dict/etc.)
    if isinstance(value, str):
        try:
            parsed = ast.literal_eval(value)
            value = parsed
        except (ValueError, SyntaxError):
            pass

    if isinstance(value, dict):
        return "<br>".join(f"{k}: {v}" for k, v in value.items())

    if isinstance(value, list):
        return "<br>".join(str(v) for v in value)

    return str(value)


def write_evaluations_section(file, rct_detailed_report):
    rule_categories = {
        "Failing": rct_detailed_report.rules_failed,
        "Passing": rct_detailed_report.rules_passed,
        "Undetermined": rct_detailed_report.full_eval_rules_undetermined
        + rct_detailed_report.appl_eval_rules_undetermined,
        "N/A": rct_detailed_report.rules_not_applicable,
        "Missing Data": rct_detailed_report.rules_undetermined_missing_data,
    }

    outcome_order = {
        "FAILED": 0,
        "UNDETERMINED": 1,
        "PASS": 2,
        "NOT_APPLICABLE": 3,
    }

    styles = {
        "FAILED": "background-color:#ffcccc;color:black;font-weight:bold;padding-left:10px;border-radius:8px;border:2px solid #ff0000;",
        "PASS": "background-color:#ccffcc;color:black;font-weight:bold;padding-left:10px;border-radius:8px;border:2px solid #008000;",
        "UNDETERMINED": "background-color:#ffffcc;color:black;font-weight:bold;padding-left:10px;border-radius:8px;border:2px solid #ffcc00;",
        "DEFAULT": "padding-left:10px;border:2px solid #ccc;border-radius:8px;",
    }

    for category, rules in rule_categories.items():
        btn_class = (
            "btn-danger"
            if category == "Failing"
            else "btn-warning"
            if category == "Undetermined"
            else "btn-light"
            if category == "Missing Data"
            else "btn-success"
            if category == "Passing"
            else "btn-secondary"
        )

        collapse_id = category.replace(" ", "_")

        file.write(
            f"""
        <div class="mb-3">
            <button class="btn {btn_class} w-100 text-start sticky-top"
                type="button" data-bs-toggle="collapse"
                data-bs-target="#collapse_fully_{collapse_id}">
                <strong>{category} Rules ({len(rules)})</strong>
            </button>
            <div class="collapse mx-4" id="collapse_fully_{collapse_id}">
        """
        )

        if category == "Undetermined":
            file.write("<h3 class='mt-4'>Rules Fully Evaluated</h3>")

        file.write(
            """
        <table class="table table-bordered table-striped mt-2">
            <thead class="table-dark">
                <tr>
                    <th rowspan="2">Rule ID</th>
                    <th>Description</th>
                    <th>Standard Section</th>
                    <th>Outcome Counts</th>
                </tr>
                <tr><th colspan="3">Evaluations</th></tr>
            </thead>
            <tbody>
        """
        )

        sections_seen = set()

        for rule_id in rules:
            rule_data = next(
                r
                for r in rct_detailed_report.evaluation_data["rules"]
                if r["rule_id"] == rule_id
            )

            standard_evals, missing_evals = split_evaluations(rule_data)

            if category == "Missing Data":
                evaluations = missing_evals
                if not evaluations:
                    continue
            else:
                evaluations = standard_evals

            section = rule_id.split("-")[0]
            if section not in sections_seen:
                sections_seen.add(section)
                title, color = section_titles_with_colors.get(
                    int(section), ("Unknown Section", "#eeeeee")
                )
                file.write(
                    f"""
                </tbody>
                <thead class="table-group-divider">
                    <tr>
                        <th colspan="4" class="section-title sticky-top sticky-top-2"
                            style="background-color:{color}!important;">
                            {title}
                        </th>
                    </tr>
                </thead>
                <tbody>
                """
                )

            description = rule_data.get("description", "N/A")
            standard_section = rule_data.get("standard_section", "N/A")
            outcome_summary = " | ".join(
                f"{k}: {v}"
                for k, v in rct_detailed_report.rule_evaluation_outcome_counts[
                    rule_id
                ].items()
            )

            file.write(
                f"""
            <tr>
                <td class="rule-id" rowspan="2">{rule_id}</td>
                <td>{description}</td>
                <td>{standard_section}</td>
                <td class="outcome-summary">{outcome_summary}</td>
            </tr>
            <tr>
                <td colspan="3">
                    <button class="btn btn-primary" type="button"
                        data-bs-toggle="collapse"
                        data-bs-target="#eval_{rule_id}">
                        View Evaluations
                    </button>
                    <div class="collapse" id="eval_{rule_id}">
                        <ul>
            """
            )

            for ev in sorted(
                evaluations, key=lambda e: outcome_order.get(e["outcome"], 3)
            ):
                li_style = styles.get(ev["outcome"], styles["DEFAULT"])
                file.write(
                    f"""
                <li style="{li_style}" class="p-2 m-1">
                    {ev["data_group_id"]}
                    <ul>
                        <li><strong>Outcome:</strong> {ev["outcome"]}</li>
                """
                )

                if ev.get("messages"):
                    msgs = set()
                    if isinstance(ev["messages"], str):
                        msgs.add(ev["messages"])
                    elif isinstance(ev["messages"], dict):
                        msgs |= {f"{k}: {v}" for k, v in ev["messages"].items()}
                    else:
                        msgs |= set(ev["messages"])
                    file.write(f"<li><strong>Messages:</strong> {', '.join(msgs)}</li>")

                if ev.get("calculated_values"):
                    file.write(
                        """
                    <li><strong>Calculated Values:</strong>
                        <table class="mb-2 me-2 table table-sm table-bordered">
                            <thead>
                                <tr><th>Variable</th><th>Value</th>
                    """
                    )

                    has_units = any(cv.get("unit") for cv in ev["calculated_values"])
                    if has_units:
                        file.write("<th>Unit</th>")
                    file.write("</tr></thead><tbody>")

                    for cv in ev["calculated_values"]:
                        file.write(
                            f"""
                        <tr>
                            <td>{cv["variable"]}</td>
                            <td>{format_value_for_html(cv["value"])}</td>
                        """
                        )
                        if cv.get("unit"):
                            file.write(f"<td>{cv['unit']}</td>")
                        elif has_units:
                            file.write("<td></td>")
                        file.write("</tr>")

                    file.write("</tbody></table></li>")

                file.write("</ul></li>")

            file.write(
                """
                        </ul>
                    </div>
                </td>
            </tr>
            """
            )

        file.write(
            """
            </tbody>
        </table>
        </div>
        </div>
        """
        )
