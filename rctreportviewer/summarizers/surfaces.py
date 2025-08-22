def summarize_rmd_surface_data(building_segment, zone, rmd_building_summary):
    constructions = rmd_building_summary["constructions"]
    for surface in zone.get("surfaces", []):
        if (
            surface.get("classification") == "WALL"
            and surface.get("adjacent_to") == "EXTERIOR"
        ):
            if "area" in surface:
                rmd_building_summary["total_exterior_wall_area"] += surface["area"]
                rmd_building_summary["total_wall_area_by_building_segment"][
                    building_segment["id"]
                ] = (
                    rmd_building_summary["total_wall_area_by_building_segment"].get(
                        building_segment["id"], 0
                    )
                    + surface["area"]
                )
            construction = next(cons for cons in constructions if cons["id"] == surface.get("construction"))
            if construction and "u_factor" in construction:
                rmd_building_summary["overall_wall_ua_by_building_segment"][
                    building_segment["id"]
                ] = (
                    rmd_building_summary["overall_wall_ua_by_building_segment"].get(
                        building_segment["id"], 0
                    )
                    + construction["u_factor"] * surface["area"]
                )
        if (
            surface.get("classification") == "CEILING"
            and surface.get("adjacent_to") == "EXTERIOR"
        ):
            if "area" in surface:
                rmd_building_summary["total_roof_area"] += surface["area"]
                rmd_building_summary["total_roof_area_by_building_segment"][
                    building_segment["id"]
                ] = (
                    rmd_building_summary["total_roof_area_by_building_segment"].get(
                        building_segment["id"], 0
                    )
                    + surface["area"]
                )
            construction = next(cons for cons in constructions if cons["id"] == surface.get("construction"))
            if construction and "u_factor" in construction:
                rmd_building_summary["overall_roof_ua_by_building_segment"][
                    building_segment["id"]
                ] = (
                    rmd_building_summary["overall_roof_ua_by_building_segment"].get(
                        building_segment["id"], 0
                    )
                    + construction["u_factor"] * surface["area"]
                )

        for subsurface in surface.get("subsurfaces", []):
            if (
                surface.get("adjacent_to") == "EXTERIOR"
                and subsurface.get("classification") == "WINDOW"
            ):
                if "glazed_area" in subsurface:
                    rmd_building_summary["total_window_area"] += subsurface[
                        "glazed_area"
                    ]
                    rmd_building_summary["total_window_area_by_building_segment"][
                        building_segment["id"]
                    ] = (
                        rmd_building_summary[
                            "total_window_area_by_building_segment"
                        ].get(building_segment["id"], 0)
                        + subsurface["glazed_area"]
                    )
                if "u_factor" in subsurface:
                    rmd_building_summary["overall_window_ua_by_building_segment"][
                        building_segment["id"]
                    ] = (
                        rmd_building_summary[
                            "overall_window_ua_by_building_segment"
                        ].get(building_segment["id"], 0)
                        + subsurface["u_factor"] * subsurface["glazed_area"]
                    )
            elif (
                surface.get("adjacent_to") == "EXTERIOR"
                and subsurface.get("classification") == "SKYLIGHT"
            ):
                if "glazed_area" in subsurface:
                    rmd_building_summary["total_skylight_area"] += subsurface[
                        "glazed_area"
                    ]
                    rmd_building_summary["total_skylight_area_by_building_segment"][
                        building_segment["id"]
                    ] = (
                        rmd_building_summary[
                            "total_skylight_area_by_building_segment"
                        ].get(building_segment["id"], 0)
                        + subsurface["glazed_area"]
                    )
                if "u_factor" in subsurface:
                    rmd_building_summary["overall_skylight_ua_by_building_segment"][
                        building_segment["id"]
                    ] = (
                        rmd_building_summary[
                            "overall_skylight_ua_by_building_segment"
                        ].get(building_segment["id"], 0)
                        + subsurface["u_factor"] * subsurface["glazed_area"]
                    )
