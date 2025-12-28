from rctreportviewer.summarizers.zones import summarize_rmd_zone_data
from rctreportviewer.summarizers.systems import summarize_rmd_system_data
from rctreportviewer.summarizers.systems import summarize_heating_cooling_capacity_data


def summarize_building_segment_data(rct_report_viewer, building, rmd_building_summary):
    for building_segment in building.get("building_segments", []):
        rmd_building_summary["zone_count"] += len(building_segment.get("zones", []))
        rmd_building_summary["zone_count_by_building_segment"][
            building_segment["id"]
        ] = len(building_segment.get("zones", []))
        rmd_building_summary["system_count"] += len(
            building_segment.get("heating_ventilating_air_conditioning_systems", [])
        )
        rmd_building_summary["system_count_by_building_segment"][
            building_segment["id"]
        ] = len(
            building_segment.get("heating_ventilating_air_conditioning_systems", [])
        )

        rmd_building_summary["lighting_area_type_by_building_segment"][
            building_segment["id"]
        ] = building_segment.get("lighting_building_area_type")

        for swh_use_id in building_segment.get("service_water_heating_uses", []):
            rmd_building_summary["swh_use_id_to_area_types"].setdefault(
                swh_use_id, set()
            ).add(building_segment.get("service_water_heating_area_type", "ALL_OTHERS"))

        summarize_rmd_zone_data(
            rct_report_viewer, building_segment, rmd_building_summary
        )

        summarize_rmd_system_data(
            rct_report_viewer, building_segment, rmd_building_summary
        )

        summarize_heating_cooling_capacity_data(building_segment, rmd_building_summary)
