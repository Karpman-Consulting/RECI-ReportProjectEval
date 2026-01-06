def write_interior_lighting_details(file, rct_detailed_report):
    """
    Write the interior lighting details as a collapsible subsection
    styled to match HVAC Details.
    """

    b = rct_detailed_report.baseline_model_summary["int_ltg_summaries"]
    p = rct_detailed_report.proposed_model_summary["int_ltg_summaries"]

    file.write(
        """
<hr>

<style>
#interior-lighting-table th.sortable {
  position: relative;
  padding-right: 1.25rem;
}
#interior-lighting-table th.sortable { cursor: pointer; white-space: nowrap; }
#interior-lighting-table .sort-icon i { font-size: 0.85em; opacity: 0.9; }
</style>

<div class="card shadow-sm mb-3">
  <div class="card-header bg-light d-flex align-items-center"
       role="button"
       data-bs-toggle="collapse"
       data-bs-target="#collapse-interior-lighting-details"
       aria-expanded="false"
       style="cursor: pointer;">
    <span class="fw-semibold">Interior Lighting Details</span>
  </div>
</div>

<div id="collapse-interior-lighting-details" class="collapse mt-3">
<div class="table-responsive">
<table id="interior-lighting-table" class="table table-sm align-middle">
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
  <th class="sortable" data-type="string">Interior Lighting ID<span class="sort-icon ms-1"></span></th>
  <th class="sortable vertical-header" data-type="number">Floor Area<br><span class="sort-icon ms-1"></span>(ft²)</th>

  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>General</th>
  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>Decorative</th>
  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>Retail</th>
  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>Exempt</th>
  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>Total</th>

  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>General</th>
  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>Decorative</th>
  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>Retail</th>
  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>Exempt</th>
  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>Total</th>

  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>General</th>
  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>Decorative</th>
  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>Retail</th>
  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>Exempt</th>
  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>Total</th>

  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>General</th>
  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>Decorative</th>
  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>Retail</th>
  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>Exempt</th>
  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>Total</th>

  <th class="sortable vertical-header" data-type="number"><span class="sort-icon ms-1"></span>% Savings</th>
</tr>
</thead>
<tbody class="small text-center">
"""
    )

    all_ids = set(b) | set(p)

    for lid in sorted(all_ids):
        bl = b.get(lid)
        pl = p.get(lid)

        def v(obj, key):
            return obj[key] if obj else 0

        floor = v(bl, "floor_area") or v(pl, "floor_area")

        savings = (
            (v(bl, "int_ltg_power_total") - v(pl, "int_ltg_power_total"))
            / v(bl, "int_ltg_power_total")
            * 100
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
"""
    )
