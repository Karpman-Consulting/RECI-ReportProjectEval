def build_segment_cards(b, p):
    segment_ids = sorted(
        set(b.get("system_count_by_building_segment", {}).keys())
        | set(p.get("system_count_by_building_segment", {}).keys())
    )

    cards = []
    for seg in segment_ids:
        cards.append(
            f"""
            <div class="col">
              <div class="card h-100 border-light shadow-sm">

                <div class="card-header bg-light fw-semibold py-2 px-3">
                  {seg}
                </div>

                <div class="card-body p-3">

                  <!-- Baseline / Proposed headers -->
                  <div class="row small fw-semibold border-bottom pb-1 mb-2">
                    <div class="col-6"></div>
                    <div class="col-3 text-center">B</div>
                    <div class="col-3 text-center">P</div>
                  </div>

                  <div class="row small mb-1">
                    <div class="col-6">Floor Area (ft<sup>2</sup>)</div>
                    <div class="col-3 text-center">{round(b["floor_area_by_building_segment"].get(seg, 0)):,}</div>
                    <div class="col-3 text-center">{round(p["floor_area_by_building_segment"].get(seg, 0)):,}</div>
                  </div>
                  
                  <div class="row small mb-1">
                    <div class="col-6">Systems</div>
                    <div class="col-3 text-center">{b["system_count_by_building_segment"].get(seg, 0)}</div>
                    <div class="col-3 text-center">{p["system_count_by_building_segment"].get(seg, 0)}</div>
                  </div>

                  <div class="row small mb-1">
                    <div class="col-6">Zones</div>
                    <div class="col-3 text-center">{b["zone_count_by_building_segment"].get(seg, 0)}</div>
                    <div class="col-3 text-center">{p["zone_count_by_building_segment"].get(seg, 0)}</div>
                  </div>

                  <div class="row small mb-1">
                    <div class="col-6">Lighting (W)</div>
                    <div class="col-3 text-center">{b["lighting_power_by_building_segment"].get(seg, 0):,.0f}</div>
                    <div class="col-3 text-center">{p["lighting_power_by_building_segment"].get(seg, 0):,.0f}</div>
                  </div>

                  <div class="row small">
                    <div class="col-6">Equipment (W)</div>
                    <div class="col-3 text-center">{b["miscellaneous_equipment_power_by_building_segment"].get(seg, 0):,.0f}</div>
                    <div class="col-3 text-center">{p["miscellaneous_equipment_power_by_building_segment"].get(seg, 0):,.0f}</div>
                  </div>
                  
                  <div class="row small mb-1">
                    <div class="col-6">Occupants</div>
                    <div class="col-3 text-center">{b["occupants_by_building_segment"].get(seg, 0):,.0f}</div>
                    <div class="col-3 text-center">{p["occupants_by_building_segment"].get(seg, 0):,.0f}</div>
                  </div>

                </div>
              </div>
            </div>
            """
        )

    return "\n".join(cards)


def write_component_summary(file, rct_detailed_report):
    tooltip_html = ""

    b = rct_detailed_report.baseline_model_summary
    p = rct_detailed_report.proposed_model_summary

    segment_cards_html = build_segment_cards(b, p)

    file.write(
        f"""
<section class="mb-4">
  <div class="card shadow-sm">

    <!-- CLICKABLE HEADER -->
    <div class="card-header bg-light d-flex align-items-center sticky-top"
         role="button"
         data-bs-toggle="collapse"
         data-bs-target="#collapse-model-component-summary"
         aria-expanded="false"
         style="cursor: pointer;">
      <span class="fw-semibold">Model Component Summary</span>
    </div>

    <div id="collapse-model-component-summary" class="collapse">
      <div class="card-body">

        <div class="row g-4">

          <!-- SECTION 1: PROJECT SUMMARY -->
        <div class="col-12 col-lg-6">
          <div class="fw-semibold mb-2">Project Summary</div>
          <div class="card border-light shadow-sm">
        
            <div class="card-header bg-light fw-semibold py-2 px-3">
              {rct_detailed_report.project_name}
            </div>
        
            <div class="card-body p-3">
        
              <!-- Baseline / Proposed headers -->
              <div class="row small fw-semibold border-bottom pb-1 mb-2">
                <div class="col-6"></div>
                <div class="col-3 text-center">B</div>
                <div class="col-3 text-center">P</div>
              </div>
        
              <div class="row small mb-1">
                <div class="col-6">Buildings</div>
                <div class="col-3 text-center">{b["building_count"]}</div>
                <div class="col-3 text-center">{p["building_count"]}</div>
              </div>
        
              <div class="row small mb-1">
                <div class="col-6">Total Floor Area (ft<sup>2</sup>)</div>
                <div class="col-3 text-center">{round(b["total_floor_area"]):,}</div>
                <div class="col-3 text-center">{round(p["total_floor_area"]):,}</div>
              </div>
        
              <div class="row small mb-1">
                <div class="col-6">Building Areas</div>
                <div class="col-3 text-center">{b["building_segment_count"]}</div>
                <div class="col-3 text-center">{p["building_segment_count"]}</div>
              </div>
        
              <div class="row small mb-1">
                <div class="col-6">Systems</div>
                <div class="col-3 text-center">{b["system_count"]}</div>
                <div class="col-3 text-center">{p["system_count"]}</div>
              </div>
        
              <div class="row small mb-1">
                <div class="col-6">Zones</div>
                <div class="col-3 text-center">{b["zone_count"]}</div>
                <div class="col-3 text-center">{p["zone_count"]}</div>
              </div>
        
              <div class="row small">
                <div class="col-6">Spaces</div>
                <div class="col-3 text-center">{b["space_count"]}</div>
                <div class="col-3 text-center">{p["space_count"]}</div>
              </div>
        
            </div>
          </div>
        </div>

          <!-- SECTION 2: BUILDING SEGMENT SUBCARDS -->
          <div class="col-12 col-lg-6">
            <div class="fw-semibold mb-2">Building Area Summary</div>

            <div class="row row-cols-1 row-cols-md-2 g-3">
              {segment_cards_html}
            </div>
          </div>

        </div>
      </div>
    </div>
  </div>
</section>
"""
    )
