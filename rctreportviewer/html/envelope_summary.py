def write_envelope_summary(file, rct_detailed_report):
    """
    Write the envelope summary section of the RCT detailed report.
    """

    b = rct_detailed_report.baseline_model_summary
    p = rct_detailed_report.proposed_model_summary

    file.write(
        """
<section class="mb-4">
  <div class="card shadow-sm">

    <!-- CLICKABLE HEADER -->
    <div class="card-header bg-light d-flex align-items-center sticky-top"
         role="button"
         data-bs-toggle="collapse"
         data-bs-target="#collapse-envelope-summary"
         aria-expanded="false"
         style="cursor: pointer;">
      <span class="fw-semibold">Envelope Summary</span>
    </div>

    <div id="collapse-envelope-summary" class="collapse">
      <div class="card-body">

                <div class="table-responsive">
                    <table class="table table-sm align-middle">
                        <thead class="table-light text-center border-bottom">
                            <tr class="border-top">
                                <th colspan="2" class="border-start"></th>
                                <th colspan="6" class="border-start">Baseline</th>
                                <th colspan="6" class="border-start border-end">Proposed</th>
                            </tr>
                            <tr>
                                <th rowspan="2" class="border-start" style="vertical-align: top;">Building Area</th>
                                <th rowspan="2" class="border-start" style="vertical-align: top;">Surface Type</th>
                                <th colspan="3" class="border-start">Opaque Surface</th>
                                <th colspan="3" class="border-start">Fenestration</th>
                                <th colspan="3" class="border-start">Opaque Surface</th>
                                <th colspan="3" class="border-start border-end">Fenestration</th>
                            </tr>
                            <tr>
                                <th class="border-start">Area (ft²)</th>
                                <th class="border-start">%</th>
                                <th class="border-start">U-Factor</th>
                                <th class="border-start">Area (ft²)</th>
                                <th class="border-start">%</th>
                                <th class="border-start">U-Factor</th>
                                <th class="border-start">Area (ft²)</th>
                                <th class="border-start">%</th>
                                <th class="border-start">U-Factor</th>
                                <th class="border-start">Area (ft²)</th>
                                <th class="border-start">%</th>
                                <th class="border-start border-end">U-Factor</th>
                            </tr>
                        </thead>
                        <tbody class="small text-center">
"""
    )

    last_segment = None

    for segment in b["floor_area_by_building_segment"]:
        if segment != last_segment:
            file.write(
                f"""
              <tr class="table-secondary fw-semibold">
                <td colspan="14" class="border-start border-end text-start">{segment}</td>
              </tr>
"""
            )
            last_segment = segment

        # ---------- Roof ----------
        if segment in b["total_roof_area_by_building_segment"]:
            b_roof = b["total_roof_area_by_building_segment"].get(segment, 0)
            b_sky = b["total_skylight_area_by_building_segment"].get(segment, 0)
            p_roof = p["total_roof_area_by_building_segment"].get(segment, 0)
            p_sky = p["total_skylight_area_by_building_segment"].get(segment, 0)

            file.write(
                f"""
              <tr>
                <td class="border-start"></td>
                <td class="text-start">Roof</td>

                <td class="border-start">{round(b_roof - b_sky):,}</td>
                <td class="border-start text-muted">{round((b_roof - b_sky) / b_roof * 100, 1) if b_roof else 0}</td>
                <td class="border-start">{round(b["overall_roof_u_factor_by_building_segment"].get(segment, 0), 3)}</td>

                <td class="border-start">{round(b_sky):,}</td>
                <td class="border-start text-muted">{round(b_sky / b_roof * 100, 1) if b_roof else 0}</td>
                <td class="border-start">{round(b["overall_skylight_u_factor_by_building_segment"].get(segment, 0), 3)}</td>

                <td class="border-start">{round(p_roof - p_sky):,}</td>
                <td class="border-start text-muted">{round((p_roof - p_sky) / p_roof * 100, 1) if p_roof else 0}</td>
                <td class="border-start">{round(p["overall_roof_u_factor_by_building_segment"].get(segment, 0), 3)}</td>

                <td class="border-start">{round(p_sky):,}</td>
                <td class="border-start text-muted">{round(p_sky / p_roof * 100, 1) if p_roof else 0}</td>
                <td class="border-start border-end">{round(p["overall_skylight_u_factor_by_building_segment"].get(segment, 0), 3)}</td>
              </tr>
"""
            )

        # ---------- Exterior Wall ----------
        if segment in b["total_wall_area_by_building_segment"]:
            b_wall = b["total_wall_area_by_building_segment"].get(segment, 0)
            b_win = b["total_window_area_by_building_segment"].get(segment, 0)
            p_wall = p["total_wall_area_by_building_segment"].get(segment, 0)
            p_win = p["total_window_area_by_building_segment"].get(segment, 0)

            file.write(
                f"""
              <tr>
                <td class="border-start"></td>
                <td class="text-start">Ext. Wall</td>

                <td class="border-start">{round(b_wall - b_win):,}</td>
                <td class="border-start text-muted">{round((b_wall - b_win) / b_wall * 100, 1) if b_wall else 0}</td>
                <td class="border-start">{round(b["overall_wall_u_factor_by_building_segment"].get(segment, 0), 3)}</td>

                <td class="border-start">{round(b_win):,}</td>
                <td class="border-start text-muted">{round(b_win / b_wall * 100, 1) if b_wall else 0}</td>
                <td class="border-start">{round(b["overall_window_u_factor_by_building_segment"].get(segment, 0), 3)}</td>

                <td class="border-start">{round(p_wall - p_win):,}</td>
                <td class="border-start text-muted">{round((p_wall - p_win) / p_wall * 100, 1) if p_wall else 0}</td>
                <td class="border-start">{round(p["overall_wall_u_factor_by_building_segment"].get(segment, 0), 3)}</td>

                <td class="border-start">{round(p_win):,}</td>
                <td class="border-start text-muted">{round(p_win / p_wall * 100, 1) if p_wall else 0}</td>
                <td class="border-start border-end">{round(p["overall_window_u_factor_by_building_segment"].get(segment, 0), 3)}</td>
              </tr>
"""
            )

    file.write(
        """
            </tbody>
          </table>
        </div>

        <p class="small text-muted mt-2">
          * U-factors are area-weighted averages by surface type.
        </p>

      </div>
    </div>
  </div>
</section>
"""
    )
