def summarize_rmd_space_data(
    rct_report_viewer, building_segment, zone, rmd_building_summary
):
    baseline_space_types = {}

    def add_internal_gain_from_occupancy(spc, sch):
        """Calculate the occupant internal heat gain for a space."""
        sensible_gain = spc.get("occupant_sensible_heat_gain", 0.0)
        latent_gain = spc.get("occupant_latent_heat_gain", 0.0)
        occupancy_gain = (sensible_gain + latent_gain) * spc.get(
            "number_of_occupants", 0
        )
        rmd_building_summary["occ_peak_internal_gain_by_schedule"][
            sch
        ] += occupancy_gain

    for space in zone.get("spaces", []):
        schedule_areas_added = []
        if "floor_area" in space:
            rmd_building_summary["total_floor_area"] += space["floor_area"]
            rmd_building_summary["total_floor_area_by_building_segment"][
                building_segment["id"]
            ] = (
                rmd_building_summary["total_floor_area_by_building_segment"].get(
                    building_segment["id"], 0
                )
                + space["floor_area"]
            )
            rct_report_viewer.space_areas[space["id"]] = space["floor_area"]
        if (
            "lighting_space_type" in space
            and rmd_building_summary["rmd_type"] == "Baseline"
        ):
            baseline_space_types[space["id"]] = space["lighting_space_type"]
        if "number_of_occupants" in space:
            rmd_building_summary["total_occupants"] += space["number_of_occupants"]
            if "lighting_space_type" in space:
                rmd_building_summary["total_occupants_by_space_type"][
                    space["lighting_space_type"]
                ] = rmd_building_summary["total_occupants_by_space_type"].get(
                    space["lighting_space_type"], 0
                ) + space.get(
                    "number_of_occupants", 0
                )

        for interior_lighting in space.get(
            "interior_lighting", [{"power_per_area": 0}]
        ):
            interior_lighting_summary = {
                "space_id": space.get("id"),
                "lighting_space_type": space.get("lighting_space_type"),
                "floor_area": space.get("floor_area", 0),
                "power_per_area": interior_lighting.get("power_per_area", 0),
                "int_ltg_power_general": 0,
                "int_ltg_power_retail": 0,
                "int_ltg_power_decorative": 0,
                "int_ltg_power_exempt": 0,
                "int_ltg_power_total": 0,
            }
            int_ltg_id = interior_lighting.get("id")

            if "power_per_area" in interior_lighting and "floor_area" in space:
                int_ltg_power = (
                    interior_lighting["power_per_area"] * space["floor_area"]
                )

                # Populate interior lighting summary data
                if "purpose_type" in interior_lighting:
                    interior_lighting_summary["purpose_type"] = interior_lighting[
                        "purpose_type"
                    ]
                    if interior_lighting["purpose_type"] in ["GENERAL", "TASK"]:
                        interior_lighting_summary[
                            "int_ltg_power_general"
                        ] = int_ltg_power
                    elif interior_lighting["purpose_type"] == "RETAIL_DISPLAY":
                        interior_lighting_summary[
                            "int_ltg_power_retail"
                        ] = int_ltg_power
                    elif interior_lighting["purpose_type"] == "DECORATIVE":
                        interior_lighting_summary[
                            "int_ltg_power_decorative"
                        ] = int_ltg_power
                    elif interior_lighting["purpose_type"] == "UNREGULATED":
                        interior_lighting_summary[
                            "int_ltg_power_exempt"
                        ] = int_ltg_power
                else:
                    interior_lighting_summary["int_ltg_power_general"] = int_ltg_power
                interior_lighting_summary["int_ltg_power_total"] = int_ltg_power
                rmd_building_summary["int_ltg_summaries"][
                    int_ltg_id
                ] = interior_lighting_summary

                rmd_building_summary["total_lighting_power"] += int_ltg_power
                if "lighting_space_type" in space:
                    rmd_building_summary["total_floor_area_by_space_type"][
                        space["lighting_space_type"]
                    ] = (
                        rmd_building_summary["total_floor_area_by_space_type"].get(
                            space["lighting_space_type"], 0
                        )
                        + space["floor_area"]
                    )
                    rmd_building_summary["total_lighting_power_by_space_type"][
                        space["lighting_space_type"]
                    ] = (
                        rmd_building_summary["total_lighting_power_by_space_type"].get(
                            space["lighting_space_type"], 0
                        )
                        + interior_lighting["power_per_area"] * space["floor_area"]
                    )

                # Save lighting schedule data
                schedule = interior_lighting.get("lighting_multiplier_schedule")
                for dictionary in [
                    rmd_building_summary["int_ltg_power_by_schedule"],
                    rmd_building_summary["floor_area_by_schedule"],
                    rmd_building_summary["occ_peak_internal_gain_by_schedule"],
                ]:
                    if schedule and schedule not in dictionary:
                        dictionary[schedule] = 0.0
                rmd_building_summary["int_ltg_power_by_schedule"][
                    schedule
                ] += int_ltg_power
                if schedule not in schedule_areas_added:
                    rmd_building_summary["floor_area_by_schedule"][schedule] += space[
                        "floor_area"
                    ]
                    schedule_areas_added.append(schedule)
                add_internal_gain_from_occupancy(space, schedule)

        for miscellaneous_equipment in space.get(
            "miscellaneous_equipment", [{"power": 0}]
        ):
            if "power" in miscellaneous_equipment and "floor_area" in space:
                rmd_building_summary[
                    "total_equipment_power"
                ] += miscellaneous_equipment["power"]
                if "lighting_space_type" in space:
                    rmd_building_summary[
                        "total_miscellaneous_equipment_power_by_space_type"
                    ][space["lighting_space_type"]] = (
                        rmd_building_summary[
                            "total_miscellaneous_equipment_power_by_space_type"
                        ].get(space["lighting_space_type"], 0)
                        + miscellaneous_equipment["power"]
                    )

                # Save equipment schedule data
                schedule = miscellaneous_equipment.get("multiplier_schedule")
                for dictionary in [
                    rmd_building_summary["equip_power_by_schedule"],
                    rmd_building_summary["floor_area_by_schedule"],
                    rmd_building_summary["occ_peak_internal_gain_by_schedule"],
                ]:
                    if schedule and schedule not in dictionary:
                        dictionary[schedule] = 0.0
                rmd_building_summary["equip_power_by_schedule"][
                    schedule
                ] += miscellaneous_equipment["power"]
                if schedule not in schedule_areas_added:
                    rmd_building_summary["floor_area_by_schedule"][schedule] += space[
                        "floor_area"
                    ]
                    schedule_areas_added.append(schedule)
                add_internal_gain_from_occupancy(space, schedule)

        # Save occupancy schedule data
        if "occupant_multiplier_schedule" in space and "floor_area" in space:
            schedule = space["occupant_multiplier_schedule"]
            for dictionary in [
                rmd_building_summary["floor_area_by_schedule"],
                rmd_building_summary["occ_peak_internal_gain_by_schedule"],
            ]:
                if schedule and schedule not in dictionary:
                    dictionary[schedule] = 0.0
            if schedule not in schedule_areas_added:
                rmd_building_summary["floor_area_by_schedule"][schedule] += space[
                    "floor_area"
                ]
                schedule_areas_added.append(schedule)
            add_internal_gain_from_occupancy(space, schedule)

        for swh_use_id in space.get("service_water_heating_uses", []):
            rmd_building_summary["swh_use_id_to_area_types"].setdefault(
                swh_use_id, set()
            ).add(space.get("service_water_heating_area_type"))
