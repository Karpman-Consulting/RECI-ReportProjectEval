from rctreportviewer.constants import efficiency_display_map


def _has_any_positive(*vals):
    return any(v > 0 for v in vals)


def write_hvac_summary(file, rct_detailed_report):
    b = rct_detailed_report.baseline_model_summary
    p = rct_detailed_report.proposed_model_summary

    # ==========================================================
    # HVAC SUMMARY (COLLAPSIBLE)
    # ==========================================================
    file.write(
        """
<section class="mb-4">
  <div class="card shadow-sm">

    <div class="card-header bg-light d-flex align-items-center"
         role="button"
         data-bs-toggle="collapse"
         data-bs-target="#collapse-hvac-summary"
         aria-expanded="false"
         style="cursor: pointer;">
      <span class="fw-semibold">HVAC Summary</span>
    </div>

    <div id="collapse-hvac-summary" class="collapse">
      <div class="card-body">
"""
    )

    # ==========================================================
    # Cooling Plant Summary
    # ==========================================================
    if b["chiller_count"] + p["chiller_count"] > 0:
        file.write(
            """
<h5>Cooling Plant Summary</h5>
<div class="table-responsive">
<table class="table table-sm">
<thead class="table-light text-center">
<tr class="border-top">
  <th rowspan="2" class="border-start">Fuel Type</th>
  <th colspan="4" class="border-start">Baseline</th>
  <th colspan="4" class="border-start border-end">Proposed</th>
</tr>
<tr>
  <th class="border-start">Chillers</th><th class="border-start">Capacity (ton)</th><th class="border-start">CT GPM</th><th class="border-start">CT HP</th>
  <th class="border-start">Chillers</th><th class="border-start">Capacity (ton)</th><th class="border-start">CT GPM</th><th class="border-start border-end">CT HP</th>
</tr>
</thead>
<tbody class="text-center small">
"""
        )

        if _has_any_positive(
            b.get("electric_chiller_count", 0), p.get("electric_chiller_count", 0)
        ):
            file.write(
                f"""
<tr>
  <td class="border-start">Electricity</td>
  <td class="border-start">{round(b.get("electric_chiller_count", 0)):,}</td>
  <td class="border-start">{round(b.get("electric_chiller_plant_capacity", 0),1):,}</td>
  <td class="border-start">{round(b.get("cooling_tower_gpm", 0),1):,}</td>
  <td class="border-start">{round(b.get("cooling_tower_hp", 0),1):,}</td>
  <td class="border-start">{round(p.get("electric_chiller_count", 0)):,}</td>
  <td class="border-start">{round(p.get("electric_chiller_plant_capacity", 0),1):,}</td>
  <td class="border-start">{round(p.get("cooling_tower_gpm", 0),1):,}</td>
  <td class="border-start border-end">{round(p.get("cooling_tower_hp", 0),1):,}</td>
</tr>
"""
            )

        if _has_any_positive(p.get("fossil_fuel_chiller_count", 0)):
            file.write(
                f"""
<tr>
  <td class="border-start">Fossil Fuel</td>
  <td class="border-start" colspan="4">—</td>
  <td class="border-start">{round(p.get("fossil_fuel_chiller_count", 0)):,}</td>
  <td class="border-start">{round(p.get("fossil_fuel_chiller_plant_capacity", 0),1):,}</td>
  <td class="border-start border-end" colspan="2">—</td>
</tr>
"""
            )

        file.write(
            f"""
<tr class="fw-bold border-top">
  <td>Total</td>
  <td class="border-start">{round(b.get("electric_chiller_count", 0)):,}</td>
  <td class="border-start">{round(b.get("electric_chiller_plant_capacity", 0),1):,}</td>
  <td class="border-start">{round(b.get("cooling_tower_gpm", 0),1):,}</td>
  <td class="border-start">{round(b.get("cooling_tower_hp", 0),1):,}</td>
  <td class="border-start">{round(p.get("chiller_count", 0)):,}</td>
  <td class="border-start">{round(p.get("electric_chiller_plant_capacity", 0) + p.get("fossil_fuel_chiller_plant_capacity", 0),1):,}</td>
  <td class="border-start">{round(p.get("cooling_tower_gpm", 0),1):,}</td>
  <td class="border-start border-end">{round(p.get("cooling_tower_hp", 0),1):,}</td>
</tr>
</tbody>
</table>
</div>
"""
        )

    # ==========================================================
    # Heating Plant Summary
    # ==========================================================
    if b["boiler_count"] + p["boiler_count"] > 0:
        file.write(
            """
<h5>Heating Plant Summary</h5>
<div class="table-responsive">
<table class="table table-sm">
<thead class="table-light text-center">
<tr class="border-top">
  <th rowspan="2" class="border-start">Fuel Type</th>
  <th colspan="2" class="border-start">Baseline</th>
  <th colspan="2" class="border-start border-end">Proposed</th>
</tr>
<tr>
  <th class="border-start">Boilers</th><th class="border-start">Capacity (kBtu/hr)</th>
  <th class="border-start">Boilers</th><th class="border-start border-end">Capacity (kBtu/hr)</th>
</tr>
</thead>
<tbody class="text-center small">
"""
        )

        if _has_any_positive(p.get("electric_boiler_count", 0)):
            file.write(
                f"""
<tr>
  <td>Electricity</td>
  <td colspan="2">—</td>
  <td class="border-start">{round(p.get("electric_boiler_count", 0)):,}</td>
  <td class="border-start border-end">{round(p.get("electric_boiler_plant_capacity", 0),1):,}</td>
</tr>
"""
            )

        if _has_any_positive(
            b.get("fossil_fuel_boiler_count", 0), p.get("fossil_fuel_boiler_count", 0)
        ):
            file.write(
                f"""
<tr>
  <td class="border-start">Fossil Fuel</td>
  <td class="border-start">{round(b.get("fossil_fuel_boiler_count", 0)):,}</td>
  <td class="border-start">{round(b.get("fossil_fuel_boiler_plant_capacity", 0),1):,}</td>
  <td class="border-start">{round(p.get("fossil_fuel_boiler_count", 0)):,}</td>
  <td class="border-start border-end">{round(p.get("fossil_fuel_boiler_plant_capacity", 0),1):,}</td>
</tr>
"""
            )

        file.write(
            f"""
<tr class="fw-bold border-top">
  <td class="border-start">Total</td>
  <td class="border-start">{round(b.get("fossil_fuel_boiler_count", 0)):,}</td>
  <td class="border-start">{round(b.get("fossil_fuel_boiler_plant_capacity", 0),1):,}</td>
  <td class="border-start">{round(p.get("boiler_count", 0)):,}</td>
  <td class="border-start border-end">{round(p.get("electric_boiler_plant_capacity", 0) + p.get("fossil_fuel_boiler_plant_capacity", 0),1):,}</td>
</tr>
</tbody>
</table>
</div>
"""
        )

    # ==========================================================
    # Air-Side HVAC Capacity Summary
    # ==========================================================
    file.write(
        """
<h5>Air-Side HVAC Capacity Summary</h5>
<div class="table-responsive">
<table class="table table-sm">
<thead class="table-light text-center">
<tr class="border-top">
  <th rowspan="2" class="border-start">Fuel Type</th>
  <th colspan="2" class="border-start">Baseline</th>
  <th colspan="2" class="border-start border-end">Proposed</th>
</tr>
<tr>
  <th class="border-start">Heating (kBtu/hr)</th><th class="border-start">Cooling (kBtu/hr)</th>
  <th class="border-start">Heating (kBtu/hr)</th><th class="border-start border-end">Cooling (kBtu/hr)</th>
</tr>
</thead>
<tbody class="text-center small">
"""
    )

    fuel_rows = [
        "Electricity",
        "Fossil Fuel",
        "On-site Boiler Plant",
        "Purchased Heat",
        "On-site Chiller Plant",
        "Purchased CHW",
    ]

    for fuel in fuel_rows:
        bh = b["heating_capacity_by_fuel_type"].get(fuel, 0)
        bc = b["cooling_capacity_by_fuel_type"].get(fuel, 0)
        ph = p["heating_capacity_by_fuel_type"].get(fuel, 0)
        pc = p["cooling_capacity_by_fuel_type"].get(fuel, 0)

        if _has_any_positive(bh, bc, ph, pc):
            file.write(
                f"""
<tr>
  <td class="border-start">{fuel}</td>
  <td class="border-start">{round(bh):,}</td>
  <td class="border-start">{round(bc):,}</td>
  <td class="border-start">{round(ph):,}</td>
  <td class="border-start border-end">{round(pc):,}</td>
</tr>
"""
            )

    file.write(
        f"""
<tr class="fw-bold border-top">
  <td class="border-start">Total</td>
  <td class="border-start">{round(b["heating_capacity_by_fuel_type"].get("Total", 0)):,}</td>
  <td class="border-start">{round(b["cooling_capacity_by_fuel_type"].get("Total", 0)):,}</td>
  <td class="border-start">{round(p["heating_capacity_by_fuel_type"].get("Total", 0)):,}</td>
  <td class="border-start border-end">{round(p["cooling_capacity_by_fuel_type"].get("Total", 0)):,}</td>
</tr>
</tbody>
</table>
</div>
"""
    )

    # ==========================================================
    # HVAC DETAILS (Systems)
    # ==========================================================
    file.write(
        """
<hr>

<div class="card shadow-sm mb-3">
  <div class="card-header bg-light d-flex align-items-center"
       role="button"
       data-bs-toggle="collapse"
       data-bs-target="#collapse-hvac-details"
       aria-expanded="false"
       style="cursor: pointer;">
    <span class="fw-semibold">HVAC Details</span>
  </div>
</div>

<div id="collapse-hvac-details" class="collapse mt-3">
<div class="table-responsive">
<table class="table table-sm">
<thead class="table-light text-center">
<tr class="border-top">
  <th rowspan="2" class="border-start">System</th>
  <th rowspan="2" class="border-start">Type</th>
  <th rowspan="2" class="border-start">Zones</th>
  <th colspan="6" class="border-start">Heating</th>
  <th colspan="5" class="border-start border-end">Cooling</th>
</tr>
<tr>
  <th class="border-start">Equip</th><th class="border-start">Fuel</th><th class="border-start">Cap</th><th class="border-start">Units</th><th>Eff</th><th class="border-start">Eff Units</th>
  <th class="border-start">Equip</th><th class="border-start">Cap</th><th class="border-start">Units</th><th class="border-start">Eff</th><th class="border-start border-end">Eff Units</th>
</tr>
</thead>
<tbody class="small text-center">
"""
    )

    for s in b["hvac_system_summaries"]:

        def eff(vals, types):
            if not vals or not types:
                return "-", "-"
            return (
                ", ".join(str(round(v, 3)) for v in vals),
                ", ".join(efficiency_display_map.get(t, t) for t in types),
            )

        h_val, h_typ = eff(
            s.get("heating_efficiency_metric_values"),
            s.get("heating_efficiency_metric_types"),
        )
        c_val, c_typ = eff(
            s.get("cooling_efficiency_metric_values"),
            s.get("cooling_efficiency_metric_types"),
        )

        file.write(
            f"""
<tr>
  <td class="border-start">{s.get("name","-")}</td>
  <td class="border-start">{s.get("type","-")}</td>
  <td class="border-start">{s.get("zone_qty",0)}</td>
  <td class="border-start">{s.get("heating_equipment_type","-").replace("_"," ").title()}</td>
  <td class="border-start">{s.get("heating_energy_source","-").replace("_"," ").title()}</td>
  <td class="border-start">{round(s.get("heating_capacity",0)):,}</td>
  <td class="border-start">{s.get("heating_capacity_units","-")}</td>
  <td class="border-start">{h_val}</td>
  <td class="border-start">{h_typ}</td>
  <td class="border-start">{s.get("cooling_equipment_type","-").replace("_"," ").title()}</td>
  <td class="border-start">{round(s.get("cooling_capacity",0)):,}</td>
  <td class="border-start">{s.get("cooling_capacity_units","-")}</td>
  <td class="border-start">{c_val}</td>
  <td class="border-start border-end">{c_typ}</td>
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
