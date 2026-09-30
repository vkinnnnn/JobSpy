from __future__ import annotations

"""Curated set of target job roles to search for with JobSpy.

These are the roles a job seeker is actively targeting. They are grouped into
categories for readability and to allow scraping a single slice at a time, and
also exposed as a single flat, de-duplicated list (:data:`TARGET_ROLES`).

Each string is meant to be passed straight through to
:func:`jobspy.scrape_jobs` as the ``search_term`` (see
:func:`jobspy.scrape_target_roles`).
"""

# Roles grouped by category. Order is preserved and each role appears once.
TARGET_ROLES_BY_CATEGORY: dict[str, list[str]] = {
    "ai_ml_engineering": [
        "AI Engineer",
        "Junior AI Engineer",
        "Machine Learning Engineer",
        "Applied AI Engineer",
        "AI Software Engineer",
        "Agentic AI Engineer",
        "Agent Engineer",
        "LLM Engineer",
        "RAG Engineer",
        "Prompt Engineer",
        "Computer Vision Engineer",
    ],
    "mlops_platform_infra": [
        "LLMOps Engineer",
        "MLOps Engineer",
        "ML Platform Engineer",
        "AI Infrastructure Engineer",
    ],
    "software_engineering": [
        "Software Engineer",
        "Full Stack Engineer",
        "Backend Engineer",
        "Platform Engineer",
        "Forward Deployed Engineer",
    ],
    "data_analytics": [
        "Junior Data Analyst",
        "Reporting Analyst",
        "Business Intelligence Analyst",
        "Analytics Engineer",
        "Data Coordinator",
    ],
    "data_science": [
        "Junior Data Scientist",
        "Data Scientist",
        "Applied Scientist",
        "Research Data Scientist",
    ],
    "data_engineering": [
        "Data Engineer",
        "Junior Data Engineer",
        "ML Data Engineer",
    ],
    "ai_product_governance": [
        "AI Product Manager",
        "AI Solutions Architect",
        "AI Agent Architect",
        "AI Governance Specialist",
        "AI Data Governance Manager",
        "AI Ethics & Compliance Officer",
        "Responsible AI Specialist",
        "AI Security & Red Teaming Specialist",
        "Evals Engineer",
        "Model Evaluator",
        "Data Annotator / AI Trainer",
    ],
}

# Flat, de-duplicated list of every target role (insertion order preserved).
TARGET_ROLES: list[str] = list(
    dict.fromkeys(
        role
        for roles in TARGET_ROLES_BY_CATEGORY.values()
        for role in roles
    )
)


def get_target_roles(category: str | None = None) -> list[str]:
    """Return target roles, optionally limited to a single category.

    :param category: one of the keys of :data:`TARGET_ROLES_BY_CATEGORY`. When
        ``None`` (default) all roles are returned.
    :raises ValueError: if ``category`` is not a known category.
    """
    if category is None:
        return list(TARGET_ROLES)
    try:
        return list(TARGET_ROLES_BY_CATEGORY[category])
    except KeyError:
        valid = ", ".join(TARGET_ROLES_BY_CATEGORY)
        raise ValueError(
            f"Unknown category '{category}'. Valid categories are: {valid}"
        ) from None
