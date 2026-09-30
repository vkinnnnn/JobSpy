"""Deterministic digest enrichment for the target-roles Excel.

Adds daily_digest-style columns WITHOUT an LLM. This is a **triage** pass: it
gives every row a sane heuristic score + honest skill overlap so the list can be
sorted. The authoritative score comes from the in-chat expert pass (apply-digest
skill), which overwrites these columns and sets score_source="expert".

Fixes over the naive first version:
- word-boundary skill matching (no more spurious "r"/"java" substring hits);
- symbol-aware patterns (c++, c#, node.js, ci/cd, a/b testing);
- case-sensitive matching for ambiguous short tokens (R, Go);
- years-of-experience + seniority penalties for an early-career candidate;
- score_source column so the reader knows which rows are trustworthy.
"""
from __future__ import annotations

import re
from datetime import date

import pandas as pd

# ---- Skill vocabulary -------------------------------------------------------
# (display, regex_source, case_insensitive, owned_by_chirag)
_SPECIAL = {
    "c++": r"c\+\+",
    "c#": r"c#",
    ".net": r"\.net",
    "node.js": r"node\.?js",
    "ci/cd": r"ci/?cd",
    "a/b testing": r"a/b\s*test",
}


def _tok(token: str) -> str:
    if token in _SPECIAL:
        return _SPECIAL[token]
    # boundary that respects alphanumerics but allows internal . and spaces
    return r"(?<![A-Za-z0-9])" + re.escape(token) + r"(?![A-Za-z0-9])"


# owned = skills Chirag genuinely has
_OWNED = [
    "python", "sql", "typescript", "javascript", "java", "c++",
    "react", "next.js", "redux", "tailwindcss", "streamlit",
    "fastapi", "flask", "node.js", "rest", "grpc", "graphql", "websockets",
    "microservices", "pydantic",
    "langchain", "langgraph", "crewai", "openai", "mcp", "rag", "tool-calling",
    "fine-tuning", "lora", "prompt engineering", "guardrails", "llm",
    "generative ai", "agentic", "nlp", "computer vision", "deep learning",
    "machine learning",
    "pytorch", "tensorflow", "scikit-learn", "xgboost", "lightgbm", "hugging face",
    "pandas", "numpy", "spark", "pyspark", "postgresql", "pgvector", "pinecone",
    "chromadb", "redis", "mongodb", "bigquery", "etl", "elt",
    "airflow", "mlflow", "databricks", "ci/cd", "github actions", "docker",
    "kubernetes", "aws", "gcp", "azure", "vertex ai", "sagemaker", "textract",
    "lambda", "power bi", "tableau", "plotly", "matplotlib", "seaborn",
    "statistics", "a/b testing", "data pipelines", "data modeling", "ci",
]
# gap = commonly-demanded skills Chirag does NOT have (true "missing" signal)
_GAP = [
    "go", "rust", "scala", "kotlin", "c#", ".net", "php", "ruby",
    "django", "spring", "angular", "vue", "svelte",
    "terraform", "ansible", "jenkins", "helm", "prometheus", "grafana", "datadog",
    "kafka", "rabbitmq", "snowflake", "redshift", "dbt", "hadoop", "hive",
    "elasticsearch", "cassandra", "dynamodb", "oracle", "sas", "matlab", "looker",
    "salesforce", "sap", "kubeflow", "flink",
]
# ambiguous tokens matched case-sensitively as standalone words
_CASE_SENSITIVE = {"r", "go", "c"}
_R_PAT = re.compile(r"(?<![A-Za-z])R(?![A-Za-z])")
_GO_PAT = re.compile(r"\b(Go|Golang)\b")

_OWNED_PATS = [(s, re.compile(_tok(s), re.I)) for s in _OWNED if s not in _CASE_SENSITIVE]
_OWNED_PATS.append(("r", _R_PAT))
_GAP_PATS = [(s, re.compile(_tok(s), re.I)) for s in _GAP if s not in _CASE_SENSITIVE]
_GAP_PATS.append(("go", _GO_PAT))

_VARIANT_RULES = [
    (r"\b(ml ?ops|ml platform|ai infrastructure|ai platform)\b", "mlops"),
    (r"\b(computer vision|cv engineer|nlp engineer|llm|generative|agentic|applied ai|ai engineer|machine learning|ml engineer)\b", "ml_ai"),
    (r"\b(data engineer|etl|elt|data pipeline|analytics engineer)\b", "data_eng"),
    (r"\b(data scientist|applied scientist|research scientist|ml scientist)\b", "data_science"),
    (r"\b(data analyst|business intelligence|reporting analyst|data analytics)\b", "data_analytics"),
    (r"\b(full[ -]?stack)\b", "fullstack"),
    (r"\b(ai product|product manager|governance|responsible ai|ai ethics)\b", "ai_product"),
    (r"\b(backend|back[ -]?end|software engineer|software developer|platform engineer|forward deployed)\b", "backend_swe"),
]

# Human-readable name of WHICH of Chirag's resumes is scored against the job.
VARIANT_NAMES = {
    "ml_ai": "AI/ML Engineer resume",
    "data_eng": "Data Engineer resume",
    "data_science": "Data Scientist resume",
    "backend_swe": "Backend / Software Engineer resume",
    "fullstack": "Full Stack Engineer resume",
    "ai_product": "AI Product / Governance resume",
    "mlops": "MLOps / ML Platform resume",
    "data_analytics": "Data Analyst resume",
}

_CITIZEN = re.compile(r"\b(u\.?s\.?\s*citizen|citizenship required|must be a citizen|security clearance|ts/sci|active clearance|public trust|polygraph)\b", re.I)
_SPONSOR = re.compile(r"\b(h-?1b|visa sponsorship|will sponsor|sponsorship available)\b", re.I)
_NO_SPONSOR = re.compile(r"\b(no sponsorship|not able to sponsor|without sponsorship|stem opt|\bopt\b|\bcpt\b)\b", re.I)
# JD explicitly EXCLUDES OPT/CPT/F-1 candidates (a hard blocker for Chirag), e.g.
# "F-1 OPT ... not eligible". Distinct from a generic "no sponsorship" (which is fine for him).
_OPT_EXCLUDED = re.compile(r"(f-?1|opt|cpt)[^.\n]{0,55}(not eligible|ineligible|not be considered|excluded|do(?:es)? not qualify)", re.I)
_YEARS = re.compile(r"(\d{1,2})\+?\s*(?:\+|to \d+)?\s*years?", re.I)

_SENIOR = re.compile(r"\b(senior|sr\.?|staff|principal|lead|manager|director|head of|vp|distinguished|architect)\b", re.I)
_JUNIOR = re.compile(r"\b(junior|jr\.?|entry[- ]?level|new[- ]?grad|associate|early career|graduate|\bi{1,2}\b)\b", re.I)


def _text(row) -> str:
    return f"{row.get('title', '') or ''}\n{row.get('description', '') or ''}"


def detect_auth(text: str):
    if _CITIZEN.search(text):
        return "CITIZEN/CLEARANCE_REQUIRED", False
    if _OPT_EXCLUDED.search(text):
        return "OPT_EXCLUDED", False
    opt = bool(_NO_SPONSOR.search(text))
    if _SPONSOR.search(text):
        return "H1B_SPONSOR", opt
    if opt:
        return "OPT_FRIENDLY", True
    return "UNKNOWN", False


def years_required(text: str):
    yrs = [int(m.group(1)) for m in _YEARS.finditer(text) if int(m.group(1)) <= 20]
    return min(yrs) if yrs else None


def skill_overlap(text: str):
    matching = [s for s, p in _OWNED_PATS if p.search(text)]
    missing = [s for s, p in _GAP_PATS if p.search(text)]
    return matching, missing


def resume_variant(title: str, matched_roles: str = "") -> str:
    # Drive off the actual job TITLE first; matched_roles is a noisy aggregate.
    t = (title or "").lower()
    for pat, variant in _VARIANT_RULES:
        if re.search(pat, t):
            return variant
    hay = (matched_roles or "").lower()
    for pat, variant in _VARIANT_RULES:
        if re.search(pat, hay):
            return variant
    return "backend_swe"


def _score(matching, missing, title, title_relevant, auth_type, has_desc, yrs):
    if auth_type.startswith("CITIZEN"):
        return 1.5, 15
    if auth_type == "OPT_EXCLUDED":
        return 2.0, 20   # JD explicitly excludes OPT/CPT/F-1 -> Chirag can't apply
    s = 5.0
    s += min(2.5, 0.4 * len(matching))         # skill coverage
    s -= min(2.0, 0.6 * len(missing))          # real gaps
    s += 1.2 if title_relevant else -3.0       # on-target title
    t = title or ""
    if _SENIOR.search(t):
        s -= 1.6                                # early-career fit
    elif _JUNIOR.search(t):
        s += 1.0
    if yrs is not None:
        if yrs >= 6:
            s -= 2.0
        elif yrs >= 4:
            s -= 1.0
        elif yrs >= 3:
            s -= 0.4
    if not has_desc:
        s -= 1.2                                # can't verify without JD text
    if auth_type == "OPT_FRIENDLY":
        s += 0.5
    elif auth_type == "H1B_SPONSOR":
        s += 0.3
    composite = round(max(0.0, min(9.5, s)), 1)  # reserve >9.5 for expert pass
    return composite, int(round(composite * 10))


def enrich(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    today = date.today()
    out = {k: [] for k in ("auth_type", "opt_friendly_signal", "matching_skills",
                           "missing_skills", "resume_variant", "match_score",
                           "composite_score", "days_listed", "min_years_required")}
    dp = pd.to_datetime(df.get("date_posted"), errors="coerce")
    tr = df["title_relevant"] if "title_relevant" in df else pd.Series([True] * len(df))

    for i, (_, row) in enumerate(df.iterrows()):
        t = _text(row)
        a, o = detect_auth(t)
        m, mm = skill_overlap(t)
        has_desc = bool(str(row.get("description") or "").strip())
        yrs = years_required(t) if has_desc else None
        c, ms = _score(m, mm, row.get("title", ""), bool(tr.iloc[i]), a, has_desc, yrs)
        out["auth_type"].append(a)
        out["opt_friendly_signal"].append(o)
        out["matching_skills"].append("; ".join(m[:15]))
        out["missing_skills"].append("; ".join(mm[:10]))
        out["resume_variant"].append(resume_variant(row.get("title", ""), row.get("matched_roles", "")))
        out["match_score"].append(ms)
        out["composite_score"].append(c)
        out["min_years_required"].append(yrs)
        d = dp.iloc[i]
        out["days_listed"].append(int((today - d.date()).days) if pd.notna(d) else None)

    df = df.copy()
    for k, v in out.items():
        df[k] = v
    # Name the resume being scored for this role (no resume is built anymore).
    df["resume_matched"] = [VARIANT_NAMES.get(v, v) for v in out["resume_variant"]]
    df["score_source"] = "heuristic"
    for col in ("reason", "apply_email_subject", "apply_email_body", "linkedin_inmail"):
        df[col] = ""
    return df
