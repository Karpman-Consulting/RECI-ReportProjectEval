def write_interior_lighting_summary(file, rct_detailed_report):
    """
    Write the interior lighting summary section to the HTML file.
    """

    b = rct_detailed_report.baseline_model_summary["int_ltg_summaries"]
    p = rct_detailed_report.proposed_model_summary["int_ltg_summaries"]

    file.write(
        """
<section class="mb-4">
  <div class="card shadow-sm">

    <!-- CLICKABLE HEADER -->
    <div class="card-header bg-light d-flex align-items-center"
         role="button"
         data-bs-toggle="collapse"
         data-bs-target="#collapse-interior-lighting-summary"
         aria-expanded="false"
         style="cursor: pointer;">
      <span class="fw-semibold">Interior Lighting Summary</span>
    </div>

    <div id="collapse-interior-lighting-summary" class="collapse">
            <div class="card-body">

                <div class="table-responsive">
                    <table class="table table-sm align-middle">
                        <thead class="table-light text-center border-bottom">
                            <tr>
                                <th colspan="2"></th>
                                <th colspan="10">Lighting Power (W)</th>
                                <th colspan="10">Lighting Power Density (W/ft²)</th>
                                <th></th>
                            </tr>
                            <tr>
                                <th colspan="2"></th>
                                <th colspan="5">Baseline</th>
                                <th colspan="5">Proposed</th>
                                <th colspan="5">Baseline</th>
                                <th colspan="5">Proposed</th>
                                <th></th>
                            </tr>
                            <tr class="small align-middle">
                                <th>Interior Lighting ID</th>
                                <th class="vertical-header">Floor Area<br>(ft²)</th>
                                <th class="vertical-header">General</th>
                                <th class="vertical-header">Decorative</th>
                                <th class="vertical-header">Retail</th>
                                <th class="vertical-header">Exempt</th>
                                <th class="vertical-header">Total</th>
                                <th class="vertical-header">General</th>
                                <th class="vertical-header">Decorative</th>
                                <th class="vertical-header">Retail</th>
                                <th class="vertical-header">Exempt</th>
                                <th class="vertical-header">Total</th>
                                <th class="vertical-header">General</th>
                                <th class="vertical-header">Decorative</th>
                                <th class="vertical-header">Retail</th>
                                <th class="vertical-header">Exempt</th>
                                <th class="vertical-header">Total</th>
                                <th class="vertical-header">General</th>
                                <th class="vertical-header">Decorative</th>
                                <th class="vertical-header">Retail</th>
                                <th class="vertical-header">Exempt</th>
                                <th class="vertical-header">Total</th>
                                <th class="vertical-header">% Savings</th>
                            </tr>
                        </thead>
                        <tbody class="small text-center">
"""
    )

    totals = {
        "baseline": dict.fromkeys(
            ["floor_area", "general", "retail", "decorative", "exempt", "total"], 0
        ),
        "proposed": dict.fromkeys(
            ["floor_area", "general", "retail", "decorative", "exempt", "total"], 0
        ),
    }

    all_ids = set(b) | set(p)

    for lid in sorted(all_ids):
        bl = b.get(lid)
        pl = p.get(lid)

        def v(obj, key):
            return obj[key] if obj else 0

        floor = v(bl, "floor_area") or v(pl, "floor_area")

        totals["baseline"]["floor_area"] += v(bl, "floor_area")
        totals["proposed"]["floor_area"] += v(pl, "floor_area")

        for k in ["general", "retail", "decorative", "exempt", "total"]:
            totals["baseline"][k] += v(bl, f"int_ltg_power_{k}")
            totals["proposed"][k] += v(pl, f"int_ltg_power_{k}")

        savings = (
            (v(bl, "int_ltg_power_total") - v(pl, "int_ltg_power_total"))
            / v(bl, "int_ltg_power_total") * 100
            if v(bl, "int_ltg_power_total") > 0
            else 0
        )

        file.write(
            f"""
                            <tr>
                                <td>{lid}</td>
                                <td>{round(floor):,}</td>

                                <td>{round(v(bl, "int_ltg_power_general")):,}</td>
                                <td>{round(v(bl, "int_ltg_power_decorative")):,}</td>
                                <td>{round(v(bl, "int_ltg_power_retail")):,}</td>
                                <td>{round(v(bl, "int_ltg_power_exempt")):,}</td>
                                <td>{round(v(bl, "int_ltg_power_total")):,}</td>

                                <td>{round(v(pl, "int_ltg_power_general")):,}</td>
                                <td>{round(v(pl, "int_ltg_power_decorative")):,}</td>
                                <td>{round(v(pl, "int_ltg_power_retail")):,}</td>
                                <td>{round(v(pl, "int_ltg_power_exempt")):,}</td>
                                <td>{round(v(pl, "int_ltg_power_total")):,}</td>

                                <td>{round(v(bl, "int_ltg_power_general") / floor, 2) if floor else 0}</td>
                                <td>{round(v(bl, "int_ltg_power_decorative") / floor, 2) if floor else 0}</td>
                                <td>{round(v(bl, "int_ltg_power_retail") / floor, 2) if floor else 0}</td>
                                <td>{round(v(bl, "int_ltg_power_exempt") / floor, 2) if floor else 0}</td>
                                <td>{round(v(bl, "int_ltg_power_total") / floor, 2) if floor else 0}</td>

                                <td>{round(v(pl, "int_ltg_power_general") / floor, 2) if floor else 0}</td>
                                <td>{round(v(pl, "int_ltg_power_decorative") / floor, 2) if floor else 0}</td>
                                <td>{round(v(pl, "int_ltg_power_retail") / floor, 2) if floor else 0}</td>
                                <td>{round(v(pl, "int_ltg_power_exempt") / floor, 2) if floor else 0}</td>
                                <td>{round(v(pl, "int_ltg_power_total") / floor, 2) if floor else 0}</td>

                                <td>{round(savings, 1)}%</td>
                            </tr>
"""
        )

    file.write(
        """
                        </tbody>
                    </table>
                </div>

            </div>
        </div>
    </div>
</section>
"""
    )
