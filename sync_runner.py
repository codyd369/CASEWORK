"""
Sync Runner — Orchestrator Script.

Runs the various sync/check operations to keep the exhibit system in sync:
1. Index the all-docs database
2. Scan for new documents in the folder → add to exhibit list
3. Auto-populate metadata for incomplete exhibit rows
4. Check for missing documents → highlight RED
5. Update document open links
6. Optionally regenerate the CaseMap timeline

Can run as a one-shot or on a scheduled interval.

Usage:
    python sync_runner.py                    # Run once
    python sync_runner.py --schedule         # Run on a loop (every 5 min)
    python sync_runner.py --skip-index       # Skip all-docs indexing (faster)
    python sync_runner.py --timeline         # Also regenerate timeline
    python sync_runner.py --import-facts precompiled.xlsx  # Import pre-compiled facts
"""

import argparse
import logging
import sys
import time
import traceback

from config.settings import SYNC_INTERVAL_SECONDS
from graph_api_client import GraphAPIClient
from all_docs_indexer import AllDocsIndexer
from exhibit_list_manager import ExhibitListManager
from fact_sheet_manager import FactSheetManager
from casemap_visualizer import CaseMapVisualizer

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("sync_runner.log"),
    ],
)
logger = logging.getLogger("sync_runner")


def run_sync(
    skip_index: bool = False,
    regenerate_timeline: bool = False,
    import_facts_file: str = "",
    local_mode: bool = False,
    local_all_docs_folder: str = "",
    local_exhibit_list: str = "",
) -> dict:
    """
    Run a full sync cycle.

    Args:
        skip_index: Skip loading the all-docs index (faster for repeated runs).
        regenerate_timeline: Also regenerate the CaseMap timeline HTML.
        import_facts_file: Path to a pre-compiled facts Excel file to import.
        local_mode: If True, work with local files instead of SharePoint.
        local_all_docs_folder: Local folder path for all-docs files.
        local_exhibit_list: Local path for the exhibit list.

    Returns: Summary dict with counts for each operation.
    """
    logger.info("=" * 60)
    logger.info("SYNC CYCLE STARTING")
    logger.info("=" * 60)

    results = {}

    # ── Step 1: Initialize API client ──
    if local_mode:
        client = None
        logger.info("Running in LOCAL mode (no SharePoint).")
    else:
        logger.info("Step 1: Connecting to Microsoft Graph API...")
        client = GraphAPIClient()
        # Verify connection by resolving site/drive
        site_id = client.get_site_id()
        drive_id = client.get_drive_id()
        logger.info(f"  Connected. Site: {site_id[:20]}..., Drive: {drive_id[:20]}...")

    # ── Step 2: Load all-docs index ──
    indexer = AllDocsIndexer(client)
    if not skip_index:
        logger.info("Step 2: Loading all-docs index...")
        if local_mode and local_all_docs_folder:
            indexer.load_from_local_files(local_all_docs_folder)
        elif not local_mode:
            indexer.load()
        results["all_docs_indexed"] = indexer.count
        logger.info(f"  Indexed {indexer.count} documents.")
    else:
        logger.info("Step 2: SKIPPED (--skip-index)")

    # ── Step 3: Load exhibit list ──
    logger.info("Step 3: Loading exhibit list...")
    exhibit_mgr = ExhibitListManager(client, indexer)
    try:
        if local_mode and local_exhibit_list:
            import openpyxl
            exhibit_mgr._wb = openpyxl.load_workbook(local_exhibit_list)
            exhibit_mgr._find_data_sheet()
        elif not local_mode:
            exhibit_mgr.load()
    except Exception as e:
        logger.error(f"  Could not load exhibit list: {e}")
        logger.error("  Create Exhibit_List.xlsx in SharePoint first (see README).")
        logger.info("  Continuing without exhibit list — steps 4-7 will be skipped.")
        results["exhibit_list_error"] = str(e)
        return results

    existing_count = len(exhibit_mgr.get_all_bates_numbers())
    results["existing_exhibits"] = existing_count
    logger.info(f"  Found {existing_count} exhibits.")

    # ── Step 4: Scan for new documents ──
    if not local_mode:
        logger.info("Step 4: Scanning document folder for new exhibits...")
        added = exhibit_mgr.add_new_exhibits_from_folder()
        results["new_exhibits_added"] = added
        logger.info(f"  Added {added} new exhibits.")
    else:
        logger.info("Step 4: SKIPPED (local mode)")

    # ── Step 5: Auto-populate metadata ──
    if not skip_index:
        logger.info("Step 5: Auto-populating metadata...")
        populated = exhibit_mgr.auto_populate_metadata()
        results["metadata_populated"] = populated
        logger.info(f"  Populated metadata for {populated} rows.")
    else:
        logger.info("Step 5: SKIPPED (no index loaded)")

    # ── Step 6: Check missing documents ──
    if not local_mode:
        logger.info("Step 6: Checking for missing documents...")
        missing = exhibit_mgr.check_missing_documents()
        results["missing_documents"] = len(missing)
        logger.info(f"  {len(missing)} documents missing.")
        if missing:
            for row, bates in missing[:10]:
                logger.warning(f"    MISSING: Row {row} — {bates}")
            if len(missing) > 10:
                logger.warning(f"    ... and {len(missing) - 10} more.")
    else:
        logger.info("Step 6: SKIPPED (local mode)")

    # ── Step 7: Update document links ──
    if not local_mode:
        logger.info("Step 7: Updating document open links...")
        links_updated = exhibit_mgr.add_document_links()
        results["links_updated"] = links_updated
        logger.info(f"  Updated {links_updated} document links.")
    else:
        logger.info("Step 7: SKIPPED (local mode)")

    # ── Step 8: Save exhibit list ──
    logger.info("Step 8: Saving exhibit list...")
    if local_mode and local_exhibit_list:
        exhibit_mgr.save_local(local_exhibit_list)
    elif not local_mode:
        exhibit_mgr.save()
    logger.info("  Saved.")

    # ── Step 9: Import pre-compiled facts (if requested) ──
    if import_facts_file:
        logger.info(f"Step 9: Importing pre-compiled facts from {import_facts_file}...")
        fact_mgr = FactSheetManager(client)
        if not local_mode:
            fact_mgr.load()
        else:
            fact_mgr._create_new()

        import_result = fact_mgr.import_with_fuzzy_matching(
            import_facts_file,
            all_bates_numbers=indexer.get_all_bates_numbers(),
        )
        results["facts_imported"] = import_result["imported"]
        results["facts_fuzzy_matched"] = import_result.get("fuzzy_matched", 0)
        results["facts_flagged"] = import_result["flagged"]

        if not local_mode:
            fact_mgr.save()
        else:
            fact_mgr.save_local("Fact_Sheet_output.xlsx")
        logger.info(f"  Import complete: {import_result}")
    else:
        logger.info("Step 9: SKIPPED (no import file specified)")

    # ── Step 10: Regenerate timeline (if requested) ──
    if regenerate_timeline:
        logger.info("Step 10: Regenerating CaseMap timeline...")
        if not local_mode:
            fact_mgr = FactSheetManager(client)
            fact_mgr.load()
            viz = CaseMapVisualizer(client=client, fact_manager=fact_mgr)
            viz.load_facts()
            viz.load_document_urls()
        else:
            viz = CaseMapVisualizer()
            if import_facts_file:
                viz.load_facts_from_local("Fact_Sheet_output.xlsx")
        output = viz.generate_html()
        results["timeline_generated"] = output
        logger.info(f"  Timeline generated: {output}")
    else:
        logger.info("Step 10: SKIPPED (use --timeline to enable)")

    # ── Summary ──
    logger.info("=" * 60)
    logger.info("SYNC CYCLE COMPLETE")
    logger.info(f"Results: {results}")
    logger.info("=" * 60)

    return results


def main():
    parser = argparse.ArgumentParser(
        description="Arbitration Exhibit System — Sync Runner"
    )
    parser.add_argument("--schedule", action="store_true",
                        help="Run on a recurring schedule")
    parser.add_argument("--interval", type=int, default=SYNC_INTERVAL_SECONDS,
                        help=f"Interval between sync cycles in seconds (default: {SYNC_INTERVAL_SECONDS})")
    parser.add_argument("--skip-index", action="store_true",
                        help="Skip loading the all-docs index")
    parser.add_argument("--timeline", action="store_true",
                        help="Regenerate the CaseMap timeline")
    parser.add_argument("--import-facts",
                        help="Path to pre-compiled facts Excel file to import")
    parser.add_argument("--local", action="store_true",
                        help="Run in local mode (no SharePoint)")
    parser.add_argument("--local-all-docs",
                        help="Local folder containing all-docs Excel files")
    parser.add_argument("--local-exhibit-list",
                        help="Local path to the exhibit list Excel file")
    args = parser.parse_args()

    if args.schedule:
        logger.info(f"Running sync on schedule every {args.interval} seconds. Press Ctrl+C to stop.")
        while True:
            try:
                run_sync(
                    skip_index=args.skip_index,
                    regenerate_timeline=args.timeline,
                    import_facts_file=args.import_facts or "",
                    local_mode=args.local,
                    local_all_docs_folder=args.local_all_docs or "",
                    local_exhibit_list=args.local_exhibit_list or "",
                )
            except Exception as e:
                logger.error(f"Sync cycle failed: {e}")
                traceback.print_exc()

            logger.info(f"Next sync in {args.interval} seconds...")
            time.sleep(args.interval)
    else:
        run_sync(
            skip_index=args.skip_index,
            regenerate_timeline=args.timeline,
            import_facts_file=args.import_facts or "",
            local_mode=args.local,
            local_all_docs_folder=args.local_all_docs or "",
            local_exhibit_list=args.local_exhibit_list or "",
        )


if __name__ == "__main__":
    main()
