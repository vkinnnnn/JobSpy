#!/usr/bin/env python
"""Routine: scrape every JobSpy target role into a single Excel workbook.

Runs one search per (location, target role) across the chosen job boards,
combines everything into one de-duplicated table (keeping every scrapable
field), and writes a multi-sheet .xlsx:

    - "All Jobs"      : every unique posting, all fields
    - "Last 24h"      : postings dated within ~24 hours
    - "Remote"        : postings flagged remote
    - "By Role"       : count of postings per target role
    - "Run Info"      : parameters, per-site counts, and any failures

Usage (defaults do the full run):

    python scrape_target_roles_to_excel.py
    python scrape_target_roles_to_excel.py --limit-roles 2 --sites indeed --results 5

Note: JobSpy only supports LinkedIn, Indeed, Glassdoor, Google, ZipRecruiter,
Bayt, Naukri and BDJobs. Startup-specific boards (Y Combinator / Work at a
Startup, Wellfound/AngelList, etc.) are NOT supported; Google Jobs partially
covers those as an aggregator.
"""
from __future__ import annotations

import argparse
import re
import sys
import warnings
from datetime import date, datetime, timedelta

import pandas as pd

from jobspy import scrape_jobs, TARGET_ROLES
from jobspy.util import create_logger, desired_order

# Concatenating per-role frames with differing all-NA columns is just noisy here.
warnings.simplefilter("ignore", FutureWarning)

log = create_logger("TargetRoles")

# Title keywords that mark a posting as genuinely on-target. Indeed/Glassdoor
# match the job *description* too, so a title gate keeps the output to the
# target roles only (see the "Raw (all)" sheet for the unfiltered set).
TITLE_SIGNALS = [
    # AI / ML engineering
    "ai engineer", "a.i. engineer", "artificial intelligence engineer",
    "machine learning", "ml engineer", "ml ops", "mlops", "ml platform",
    "ml infrastructure", "ml data", "applied ai", "applied scientist",
    "ai software", "agentic", "ai agent", "agent engineer", "agent architect",
    "llm", "llmops", "large language model", "rag engineer", "retrieval augmented",
    "prompt engineer", "generative ai", "genai", "gen ai",
    "ai infrastructure", "ai platform", "computer vision", "deep learning",
    "nlp", "natural language",
    # MLOps / platform / infra
    "platform engineer", "infrastructure engineer",
    # AI product / governance / specialized
    "ai product", "ai governance", "ai ethics", "responsible ai",
    "ai security", "red team", "ai trainer", "data annotat",
    "evals", "ai evaluation", "model evaluat", "ai solution", "ai architect",
    "forward deployed",
    # Data science / analytics / engineering
    "data scientist", "data science", "data engineer", "analytics engineer",
    "data analyst", "data analytics", "business intelligence", "reporting analyst",
    "data coordinator", "research scientist",
    # Software engineering
    "software engineer", "software developer", "full stack", "full-stack",
    "fullstack", "backend", "back end", "back-end",
]


def title_is_relevant(title) -> bool:
    if not isinstance(title, str):
        return False
    padded = f" {title.lower()} "
    return any(sig in padded for sig in TITLE_SIGNALS)

# All supported US-relevant boards. Indeed/Google are the most reliable;
# LinkedIn & Glassdoor rate-limit aggressively and may return partial data.
DEFAULT_SITES = ["indeed", "google", "linkedin", "zip_recruiter", "glassdoor"]

# Source searches. Remote is derived from the is_remote column afterwards
# because Indeed cannot combine is_remote with the hours_old recency filter.
LOCATION_PROFILES = [
    {"name": "USA (nationwide)", "location": "United States", "country_indeed": "usa"},
    {"name": "Boston, MA", "location": "Boston, MA", "country_indeed": "usa"},
]

# openpyxl/Excel cannot store these control chars or > 32,767 chars per cell.
_ILLEGAL_XLSX = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_MAX_CELL = 32000

# Final column order: our search context first, then all native JobSpy fields.
EXTRA_COLS = ["matched_roles", "matched_locations", "title_relevant", "posted_within_24h"]


def _clean_cell(value):
    if isinstance(value, str):
        value = _ILLEGAL_XLSX.sub("", value)
        if len(value) > _MAX_CELL:
            value = value[:_MAX_CELL] + " …[truncated]"
    return value


def _google_term(role: str, profile: dict) -> str:
    return f"{role} jobs near {profile['location']} since last week"


def _scrape_one(
    sites: list[str],
    role: str,
    profile: dict,
    results: int,
    hours_old: int,
    fetch_linkedin_desc: bool,
    failures: list,
) -> pd.DataFrame:
    """Scrape one (role, location) across all sites, falling back per-site."""
    kwargs = dict(
        search_term=role,
        google_search_term=_google_term(role, profile),
        location=profile["location"],
        country_indeed=profile["country_indeed"],
        results_wanted=results,
        hours_old=hours_old,
        description_format="markdown",
        linkedin_fetch_description=fetch_linkedin_desc,
        verbose=0,
    )
    try:
        return scrape_jobs(site_name=sites, **kwargs)
    except Exception as e:  # one site raised -> isolate them so others survive
        log.warning(f"combined scrape failed ({role} @ {profile['name']}): {e}")
        frames = []
        for site in sites:
            try:
                part = scrape_jobs(site_name=[site], **kwargs)
                if not part.empty:
                    frames.append(part)
            except Exception as e2:
                failures.append(
                    {
                        "location": profile["name"],
                        "role": role,
                        "site": site,
                        "error": str(e2)[:300],
                    }
                )
                log.error(f"  {site} failed ({role} @ {profile['name']}): {e2}")
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def scrape_all(
    sites: list[str],
    roles: list[str],
    locations: list[dict],
    results: int,
    hours_old: int,
    fetch_linkedin_desc: bool,
):
    frames: list[pd.DataFrame] = []
    failures: list[dict] = []
    total = len(locations) * len(roles)
    done = 0
    for profile in locations:
        for role in roles:
            done += 1
            log.info(f"[{done}/{total}] {profile['name']} :: {role}")
            df = _scrape_one(
                sites, role, profile, results, hours_old, fetch_linkedin_desc, failures
            )
            if df is not None and not df.empty:
                df = df.copy()
                df["search_role"] = role
                df["search_location"] = profile["name"]
                frames.append(df)
                log.info(f"      -> {len(df)} rows")
    return frames, failures


def combine(frames: list[pd.DataFrame]) -> pd.DataFrame:
    """Concatenate, de-duplicate by job_url, and aggregate matched context."""
    if not frames:
        return pd.DataFrame()

    raw = pd.concat(frames, ignore_index=True)
    if "job_url" not in raw.columns:
        return raw

    # Aggregate every role/location a given posting matched.
    agg = (
        raw.groupby("job_url")
        .agg(
            matched_roles=("search_role", lambda s: ", ".join(sorted(set(s.dropna())))),
            matched_locations=(
                "search_location",
                lambda s: ", ".join(sorted(set(s.dropna()))),
            ),
        )
        .reset_index()
    )

    # Keep the richest row per posting (prefer one that has a description).
    raw["_has_desc"] = raw.get("description").notna() if "description" in raw else False
    raw = raw.sort_values(
        by=["_has_desc", "date_posted"], ascending=[False, False], na_position="last"
    )
    deduped = raw.drop_duplicates(subset="job_url", keep="first").drop(
        columns=["_has_desc", "search_role", "search_location"], errors="ignore"
    )
    combined = deduped.merge(agg, on="job_url", how="left")

    # Recency flag (date granularity is per-day).
    cutoff = date.today() - timedelta(days=1)
    dp = pd.to_datetime(combined.get("date_posted"), errors="coerce")
    combined["posted_within_24h"] = dp.dt.date.apply(
        lambda d: bool(d is not None and pd.notna(d) and d >= cutoff)
    )

    # On-target flag: does the posting *title* match a target-role signal?
    combined["title_relevant"] = combined.get("title").apply(title_is_relevant)

    # Order columns: context first, then native fields, then anything leftover.
    front = EXTRA_COLS + [c for c in desired_order if c in combined.columns]
    rest = [c for c in combined.columns if c not in front]
    combined = combined[front + rest]

    return combined.sort_values(
        by=["date_posted", "matched_roles"], ascending=[False, True], na_position="last"
    ).reset_index(drop=True)


def write_excel(combined: pd.DataFrame, failures: list, meta: dict, out_path: str):
    raw_all = combined.map(_clean_cell) if not combined.empty else combined

    # Main views are limited to on-target titles ("target roles only").
    on_target = (
        raw_all[raw_all["title_relevant"] == True]  # noqa: E712
        if "title_relevant" in raw_all
        else raw_all
    )

    last_24h = (
        on_target[on_target["posted_within_24h"] == True]  # noqa: E712
        if "posted_within_24h" in on_target
        else pd.DataFrame()
    )
    remote = (
        on_target[on_target["is_remote"] == True]  # noqa: E712
        if "is_remote" in on_target
        else pd.DataFrame()
    )
    by_role = (
        on_target["matched_roles"]
        .str.split(", ")
        .explode()
        .value_counts()
        .rename_axis("target_role")
        .reset_index(name="postings")
        if "matched_roles" in on_target and not on_target.empty
        else pd.DataFrame()
    )

    by_site = (
        on_target["site"].value_counts().rename_axis("site").reset_index(name="postings")
        if "site" in on_target and not on_target.empty
        else pd.DataFrame()
    )

    run_info = pd.DataFrame(
        [{"parameter": k, "value": str(v)} for k, v in meta.items()]
    )
    failures_df = pd.DataFrame(failures) if failures else pd.DataFrame(
        [{"info": "no failures"}]
    )

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        (on_target if not on_target.empty else pd.DataFrame([{"info": "no jobs"}])).to_excel(
            writer, sheet_name="All Jobs", index=False
        )
        (last_24h if not last_24h.empty else pd.DataFrame([{"info": "none"}])).to_excel(
            writer, sheet_name="Last 24h", index=False
        )
        (remote if not remote.empty else pd.DataFrame([{"info": "none"}])).to_excel(
            writer, sheet_name="Remote", index=False
        )
        (by_role if not by_role.empty else pd.DataFrame([{"info": "none"}])).to_excel(
            writer, sheet_name="By Role", index=False
        )
        # Full unfiltered set (includes off-target description matches).
        (raw_all if not raw_all.empty else pd.DataFrame([{"info": "no jobs"}])).to_excel(
            writer, sheet_name="Raw (all)", index=False
        )
        run_info.to_excel(writer, sheet_name="Run Info", index=False, startrow=0)
        by_site.to_excel(writer, sheet_name="Run Info", index=False, startrow=len(run_info) + 2)
        failures_df.to_excel(
            writer, sheet_name="Run Info", index=False, startrow=len(run_info) + len(by_site) + 5
        )

        # Freeze header + add autofilter on the browsable sheets.
        for sheet in ("All Jobs", "Last 24h", "Remote", "Raw (all)"):
            ws = writer.sheets[sheet]
            ws.freeze_panes = "A2"
            if ws.max_row > 1:
                ws.auto_filter.ref = ws.dimensions


def parse_args(argv):
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--sites", nargs="+", default=DEFAULT_SITES)
    p.add_argument("--results", type=int, default=25, help="results per role per site")
    p.add_argument("--hours", type=int, default=168, help="hours_old recency filter")
    p.add_argument("--limit-roles", type=int, default=None, help="only first N roles (testing)")
    p.add_argument(
        "--linkedin-descriptions",
        action="store_true",
        help="fetch full LinkedIn descriptions (much slower, rate-limits fast)",
    )
    p.add_argument("--out", default=None)
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv or sys.argv[1:])
    roles = TARGET_ROLES[: args.limit_roles] if args.limit_roles else TARGET_ROLES
    out_path = args.out or f"target_roles_jobs_{date.today():%Y%m%d}.xlsx"

    started = datetime.now()
    log.info(
        f"Scraping {len(roles)} roles x {len(LOCATION_PROFILES)} locations "
        f"x {len(args.sites)} sites (results={args.results}, hours_old={args.hours})"
    )

    frames, failures = scrape_all(
        sites=args.sites,
        roles=roles,
        locations=LOCATION_PROFILES,
        results=args.results,
        hours_old=args.hours,
        fetch_linkedin_desc=args.linkedin_descriptions,
    )
    combined = combine(frames)

    meta = {
        "generated_at": started.strftime("%Y-%m-%d %H:%M:%S"),
        "finished_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "roles_searched": len(roles),
        "locations": ", ".join(p["name"] for p in LOCATION_PROFILES),
        "sites": ", ".join(args.sites),
        "results_per_role_per_site": args.results,
        "hours_old": args.hours,
        "linkedin_descriptions": args.linkedin_descriptions,
        "total_unique_jobs": 0 if combined.empty else len(combined),
        "on_target_jobs": 0
        if combined.empty or "title_relevant" not in combined
        else int(combined["title_relevant"].sum()),
        "note": "Main sheets = on-target titles only; 'Raw (all)' = unfiltered. "
        "Remote & Last 24h are derived from is_remote/date_posted.",
    }

    write_excel(combined, failures, meta, out_path)
    log.info(
        f"DONE: {0 if combined.empty else len(combined)} unique jobs -> {out_path} "
        f"({len(failures)} site failures)"
    )
    print(f"OUTPUT_PATH={out_path}")
    print(f"UNIQUE_JOBS={0 if combined.empty else len(combined)}")


if __name__ == "__main__":
    main()
