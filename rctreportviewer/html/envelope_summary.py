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
    <div class="card-header bg-light d-flex align-items-center"
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
                            <tr>
                                <th colspan="2"></th>
                                <th colspan="6">Baseline</th>
                                <th colspan="6">Proposed</th>
                            </tr>
                            <tr>
                                <th rowspan="2">Building Area</th>
                                <th rowspan="2">Surface Type</th>
                                <th colspan="3">Opaque Surface</th>
                                <th colspan="3">Fenestration</th>
                                <th colspan="3">Opaque Surface</th>
                                <th colspan="3">Fenestration</th>
                            </tr>
                            <tr>
                                <th>Area (ft²)</th>
                                <th>%</th>
                                <th>U-Factor</th>
                                <th>Area (ft²)</th>
                                <th>%</th>
                                <th>U-Factor</th>
                                <th>Area (ft²)</th>
                                <th>%</th>
                                <th>U-Factor</th>
                                <th>Area (ft²)</th>
                                <th>%</th>
                                <th>U-Factor</th>
                            </tr>
                        </thead>
                        <tbody class="small text-center">
"""
    )

    for segment in b["total_floor_area_by_building_segment"]:

        # ---------- Roof ----------
        if segment in b["total_roof_area_by_building_segment"]:
            b_roof = b["total_roof_area_by_building_segment"].get(segment, 0)
            b_sky = b["total_skylight_area_by_building_segment"].get(segment, 0)

            p_roof = p["total_roof_area_by_building_segment"].get(segment, 0)
            p_sky = p["total_skylight_area_by_building_segment"].get(segment, 0)

            file.write(
                f"""
                            <tr>
                                <td>{segment}</td>
                                <td>Roof</td>
                                <td>{round(b_roof - b_sky):,}</td>
                                <td>{round((b_roof - b_sky) / b_roof * 100, 1) if b_roof else 0}</td>
                                <td>{round(b["overall_roof_u_factor_by_building_segment"].get(segment, 0), 3)}</td>
                                <td>{round(b_sky):,}</td>
                                <td>{round(b_sky / b_roof * 100, 1) if b_roof else 0}</td>
                                <td>{round(b["overall_skylight_u_factor_by_building_segment"].get(segment, 0), 3)}</td>

                                <td>{round(p_roof - p_sky):,}</td>
                                <td>{round((p_roof - p_sky) / p_roof * 100, 1) if p_roof else 0}</td>
                                <td>{round(p["overall_roof_u_factor_by_building_segment"].get(segment, 0), 3)}</td>
                                <td>{round(p_sky):,}</td>
                                <td>{round(p_sky / p_roof * 100, 1) if p_roof else 0}</td>
                                <td>{round(p["overall_skylight_u_factor_by_building_segment"].get(segment, 0), 3)}</td>
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
                                <td>{segment}</td>
                                <td>Ext. Wall</td>
                                <td>{round(b_wall - b_win):,}</td>
                                <td>{round((b_wall - b_win) / b_wall * 100, 1) if b_wall else 0}</td>
                                <td>{round(b["overall_wall_u_factor_by_building_segment"].get(segment, 0), 3)}</td>
                                <td>{round(b_win):,}</td>
                                <td>{round(b_win / b_wall * 100, 1) if b_wall else 0}</td>
                                <td>{round(b["overall_window_u_factor_by_building_segment"].get(segment, 0), 3)}</td>

                                <td>{round(p_wall - p_win):,}</td>
                                <td>{round((p_wall - p_win) / p_wall * 100, 1) if p_wall else 0}</td>
                                <td>{round(p["overall_wall_u_factor_by_building_segment"].get(segment, 0), 3)}</td>
                                <td>{round(p_win):,}</td>
                                <td>{round(p_win / p_wall * 100, 1) if p_wall else 0}</td>
                                <td>{round(p["overall_window_u_factor_by_building_segment"].get(segment, 0), 3)}</td>
                            </tr>
"""
            )

    file.write(
        """
                        </tbody>
                    </table>
                </div>

                <p class="small text-muted mt-2">
                    * U-factors represent area-weighted averages for the corresponding building area and surface type.
                </p>

            </div>
        </div>
    </div>
</section>
"""
    )
