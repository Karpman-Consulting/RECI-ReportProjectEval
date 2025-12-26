from rctreportviewer.constants import efficiency_display_map


def _has_any_positive(*vals):
    return any(v > 0 for v in vals)


def write_hvac_summary(file, rct_detailed_report):
    b = rct_detailed_report.baseline_model_summary
    p = rct_detailed_report.proposed_model_summary

    file.write("""
<section class="mb-4">
  <div class="card shadow-sm">

    <!-- CLICKABLE HEADER -->
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
""")

    # ==========================================================
    # Cooling Plant Summary
    # ==========================================================
    if b["chiller_count"] + p["chiller_count"] > 0:
        file.write("""
<h4>Cooling Plant Summary</h4>
<div class="table-responsive">
<table class="table table-sm">
<thead class="table-light text-center">
<tr>
  <th rowspan="2">Fuel Type</th>
  <th colspan="4">Baseline</th>
  <th colspan="4">Proposed</th>
</tr>
<tr>
  <th>Chillers</th>
  <th>Capacity (ton)</th>
  <th>CT GPM</th>
  <th>CT HP</th>
  <th>Chillers</th>
  <th>Capacity (ton)</th>
  <th>CT GPM</th>
  <th>CT HP</th>
</tr>
</thead>
<tbody class="text-center small">
""")

        if _has_any_positive(
            b.get("electric_chiller_count", 0),
            p.get("electric_chiller_count", 0),
        ):
            file.write(f"""
<tr>
  <td>Electricity</td>
  <td>{round(b.get("electric_chiller_count", 0)):,}</td>
  <td>{round(b.get("electric_chiller_plant_capacity", 0),1):,}</td>
  <td>{round(b.get("cooling_tower_gpm", 0),1):,}</td>
  <td>{round(b.get("cooling_tower_hp", 0),1):,}</td>
  <td>{round(p.get("electric_chiller_count", 0)):,}</td>
  <td>{round(p.get("electric_chiller_plant_capacity", 0),1):,}</td>
  <td>{round(p.get("cooling_tower_gpm", 0),1):,}</td>
  <td>{round(p.get("cooling_tower_hp", 0),1):,}</td>
</tr>
""")

        if _has_any_positive(
            p.get("fossil_fuel_chiller_count", 0),
            p.get("fossil_fuel_chiller_plant_capacity", 0),
        ):
            file.write(f"""
<tr>
  <td>Fossil Fuel</td>
  <td colspan="4">—</td>
  <td>{round(p.get("fossil_fuel_chiller_count", 0)):,}</td>
  <td>{round(p.get("fossil_fuel_chiller_plant_capacity", 0),1):,}</td>
  <td colspan="2">—</td>
</tr>
""")

        file.write(f"""
<tr class="fw-bold border-top">
  <td>Total</td>
  <td>{round(b.get("electric_chiller_count", 0)):,}</td>
  <td>{round(b.get("electric_chiller_plant_capacity", 0),1):,}</td>
  <td>{round(b.get("cooling_tower_gpm", 0),1):,}</td>
  <td>{round(b.get("cooling_tower_hp", 0),1):,}</td>
  <td>{round(p.get("chiller_count", 0)):,}</td>
  <td>{round(p.get("electric_chiller_plant_capacity", 0) + p.get("fossil_fuel_chiller_plant_capacity", 0),1):,}</td>
  <td>{round(p.get("cooling_tower_gpm", 0),1):,}</td>
  <td>{round(p.get("cooling_tower_hp", 0),1):,}</td>
</tr>
</tbody>
</table>
</div>
""")

    # ==========================================================
    # Heating Plant Summary
    # ==========================================================
    if b["boiler_count"] + p["boiler_count"] > 0:
        file.write("""
<h4>Heating Plant Summary</h4>
<div class="table-responsive">
<table class="table table-sm">
<thead class="table-light text-center">
<tr>
  <th rowspan="2">Fuel Type</th>
  <th colspan="2">Baseline</th>
  <th colspan="2">Proposed</th>
</tr>
<tr>
  <th>Boilers</th>
  <th>Capacity (kBtu/hr)</th>
  <th>Boilers</th>
  <th>Capacity (kBtu/hr)</th>
</tr>
</thead>
<tbody class="text-center small">
""")

        if _has_any_positive(
            p.get("electric_boiler_count", 0),
            p.get("electric_boiler_plant_capacity", 0),
        ):
            file.write(f"""
<tr>
  <td>Electricity</td>
  <td colspan="2">—</td>
  <td>{round(p.get("electric_boiler_count", 0)):,}</td>
  <td>{round(p.get("electric_boiler_plant_capacity", 0),1):,}</td>
</tr>
""")

        if _has_any_positive(
            b.get("fossil_fuel_boiler_count", 0),
            p.get("fossil_fuel_boiler_count", 0),
        ):
            file.write(f"""
<tr>
  <td>Fossil Fuel</td>
  <td>{round(b.get("fossil_fuel_boiler_count", 0)):,}</td>
  <td>{round(b.get("fossil_fuel_boiler_plant_capacity", 0),1):,}</td>
  <td>{round(p.get("fossil_fuel_boiler_count", 0)):,}</td>
  <td>{round(p.get("fossil_fuel_boiler_plant_capacity", 0),1):,}</td>
</tr>
""")

        file.write(f"""
<tr class="fw-bold border-top">
  <td>Total</td>
  <td>{round(b.get("fossil_fuel_boiler_count", 0)):,}</td>
  <td>{round(b.get("fossil_fuel_boiler_plant_capacity", 0),1):,}</td>
  <td>{round(p.get("boiler_count", 0)):,}</td>
  <td>{round(p.get("electric_boiler_plant_capacity", 0) + p.get("fossil_fuel_boiler_plant_capacity", 0),1):,}</td>
</tr>
</tbody>
</table>
</div>
""")

    # ==========================================================
    # HVAC DETAILS (Systems)
    # ==========================================================
    file.write("""
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
<tr>
  <th rowspan="2">System</th>
  <th rowspan="2">Type</th>
  <th rowspan="2">Zones</th>
  <th colspan="6">Heating</th>
  <th colspan="5">Cooling</th>
</tr>
<tr>
  <th>Equip</th><th>Fuel</th><th>Cap</th><th>Units</th><th>Eff</th><th>Eff Units</th>
  <th>Equip</th><th>Cap</th><th>Units</th><th>Eff</th><th>Eff Units</th>
</tr>
</thead>
<tbody class="small text-center">
""")

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

        file.write(f"""
<tr>
  <td>{s.get("name","-")}</td>
  <td>{s.get("type","-")}</td>
  <td>{s.get("zone_qty",0)}</td>
  <td>{s.get("heating_equipment_type","-").replace("_"," ").title()}</td>
  <td>{s.get("heating_energy_source","-").replace("_"," ").title()}</td>
  <td>{round(s.get("heating_capacity",0)):,}</td>
  <td>{s.get("heating_capacity_units","-")}</td>
  <td>{h_val}</td>
  <td>{h_typ}</td>
  <td>{s.get("cooling_equipment_type","-").replace("_"," ").title()}</td>
  <td>{round(s.get("cooling_capacity",0)):,}</td>
  <td>{s.get("cooling_capacity_units","-")}</td>
  <td>{c_val}</td>
  <td>{c_typ}</td>
</tr>
""")

    file.write("""
</tbody>
</table>
</div>
</div>

</div>
</div>
</section>
""")
