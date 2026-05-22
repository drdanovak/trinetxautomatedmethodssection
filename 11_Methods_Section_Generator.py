"""
TriNetX Methods Section Generator (Scaffolder)

A checkbox-driven tool that produces a STROBE-aligned first draft of the Methods
section for a TriNetX observational study. The output is a STARTING DRAFT that
requires author verification of every clinical and code-list detail.

Designed to be dropped into the TriNetX Publication Toolkit alongside the other
Streamlit pages.
"""

import io
import datetime as _dt
from typing import Dict, List, Optional

import streamlit as st

# python-docx is already used elsewhere in the toolkit (PSM Table Generator).
try:
    from docx import Document
    from docx.shared import Pt, Inches
    DOCX_AVAILABLE = True
except Exception:
    DOCX_AVAILABLE = False


# ============================================================
# Page setup
# ============================================================
st.set_page_config(
    page_title="Methods Section Generator",
    page_icon="📝",
    layout="wide",
)

st.title("📝 TriNetX Methods Section Scaffolder")
st.caption(
    "A checkbox-driven first draft generator for the Methods section of an "
    "observational TriNetX study. The output is a starting scaffold, not a "
    "finished methods section. Every clinical detail, code list, and date "
    "must be verified against your actual TriNetX query before submission."
)

with st.expander("⚠️ Important: how to use this tool responsibly", expanded=False):
    st.markdown("""
This tool produces **boilerplate prose** based on the design choices you check below.
It is intended to:

- Lower the activation energy for novice TriNetX users writing their first methods section.
- Ensure the resulting draft addresses every relevant STROBE reporting item.
- Standardize language about the TriNetX platform, propensity score matching, and
  the statistical outputs the platform produces.

It cannot:

- Verify that your study design is appropriate for your research question.
- Confirm that your code lists, eligibility criteria, or outcome definitions
  reflect what you actually queried.
- Substitute for senior methodological review.

A "verification gate" appears before the Word export to remind you of the items
that require your independent check.
""")


# ============================================================
# Schema: every checkbox the user can flip
# ============================================================
TRINETX_NETWORKS = [
    "TriNetX Research Network",
    "TriNetX Diamond Network",
    "TriNetX Global Collaborative Network",
    "TriNetX Dataworks – USA",
    "TriNetX US Collaborative Network",
    "TriNetX EMEA Collaborative Network",
    "TriNetX LATAM Collaborative Network",
    "TriNetX APAC Collaborative Network",
    "Other (specify in custom text)",
]

STUDY_DESIGNS = [
    "Retrospective cohort study",
    "Retrospective case-control study",
    "Cross-sectional study",
    "Self-controlled case series",
]

USER_DESIGNS = [
    "New-user (incident) design",
    "Prevalent-user design",
    "Not applicable",
]

COMPARATOR_TYPES = [
    "Active comparator",
    "Non-user / unexposed comparator",
    "Historical comparator",
    "Other (specify in custom text)",
]

EFFECT_ESTIMATES = [
    "Risk ratio (RR) with 95% CI",
    "Odds ratio (OR) with 95% CI",
    "Hazard ratio (HR) with 95% CI from Cox proportional hazards",
    "Risk difference with 95% CI",
    "Kaplan–Meier survival with log-rank test",
    "Number needed to treat / harm (NNT/NNH)",
]

# How each effect-estimate item should appear mid-sentence.
EFFECT_ESTIMATES_INLINE = {
    "Risk ratio (RR) with 95% CI": "risk ratios (RR) with 95% confidence intervals",
    "Odds ratio (OR) with 95% CI": "odds ratios (OR) with 95% confidence intervals",
    "Hazard ratio (HR) with 95% CI from Cox proportional hazards":
        "hazard ratios (HR) with 95% confidence intervals from Cox proportional hazards models",
    "Risk difference with 95% CI": "risk differences with 95% confidence intervals",
    "Kaplan–Meier survival with log-rank test":
        "Kaplan–Meier survival estimates with log-rank tests",
    "Number needed to treat / harm (NNT/NNH)":
        "number needed to treat or harm (NNT/NNH)",
}

MULTIPLE_COMPARISONS_METHODS = [
    "No correction applied (single primary outcome)",
    "Bonferroni",
    "Holm–Bonferroni",
    "Benjamini–Hochberg FDR",
    "Benjamini–Yekutieli",
]

SENSITIVITY_ANALYSES = [
    "E-value calculation for unmeasured confounding",
    "Varying the lookback window",
    "Varying the outcome window",
    "Alternative matching ratio (e.g., 1:2)",
    "Alternative caliper width",
    "Exclusion of patients with prior outcome events",
    "Restricting to patients with ≥1 year of prior healthcare contact",
    "Subgroup analysis by age",
    "Subgroup analysis by sex",
    "Subgroup analysis by race/ethnicity",
]

SENSITIVITY_ANALYSES_INLINE = {
    "E-value calculation for unmeasured confounding":
        "E-value calculation for unmeasured confounding",
    "Varying the lookback window": "varying the lookback window",
    "Varying the outcome window": "varying the outcome window",
    "Alternative matching ratio (e.g., 1:2)": "an alternative matching ratio (e.g., 1:2)",
    "Alternative caliper width": "an alternative caliper width",
    "Exclusion of patients with prior outcome events":
        "exclusion of patients with any prior occurrence of the outcome",
    "Restricting to patients with ≥1 year of prior healthcare contact":
        "restriction to patients with ≥1 year of prior healthcare contact",
    "Subgroup analysis by age": "subgroup analysis by age",
    "Subgroup analysis by sex": "subgroup analysis by sex",
    "Subgroup analysis by race/ethnicity": "subgroup analysis by race/ethnicity",
}


# ============================================================
# Helper: safe text
# ============================================================
def _t(s: Optional[str], fallback: str = "") -> str:
    if s is None:
        return fallback
    s = str(s).strip()
    return s if s else fallback


def _bracket(s: Optional[str], fallback: str) -> str:
    """Return user-supplied text, or a bracketed placeholder if blank.

    Bracketed placeholders make it obvious in the draft that the author still
    needs to fill in a detail.
    """
    val = _t(s)
    return val if val else f"[{fallback}]"


def _list_join(items: List[str]) -> str:
    items = [i for i in items if i]
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    if len(items) == 2:
        return f"{items[0]} and {items[1]}"
    return ", ".join(items[:-1]) + f", and {items[-1]}"


def _decap_first(s: str) -> str:
    """Lowercase only the first character, but only if the second character is
    lowercase. Preserves acronyms (RR, HR, CI) and hyphenated proper terms
    (E-value, K-M)."""
    if not s or len(s) < 2:
        return s
    if s[1].isupper() or s[1] == "-":
        return s
    return s[0].lower() + s[1:]


# ============================================================
# Paragraph generators
# Each function takes the consolidated `c` (choices) dict and returns
# one paragraph of methods prose plus a list of STROBE items it covers.
# ============================================================
def para_data_source(c: Dict) -> Dict:
    network = c["network"]
    if network == "Other (specify in custom text)":
        network = _bracket(c.get("network_other"), "specify TriNetX network")

    hco_count = _bracket(c.get("hco_count"), "number")
    query_date = _bracket(c.get("query_date"), "query date")
    geography = _t(c.get("geography"), "primarily in the United States")

    text = (
        f"Data for this study were obtained from the {network}, a federated "
        f"global health research network providing access to electronic health "
        f"record (EHR) data from approximately {hco_count} healthcare organizations "
        f"(HCOs) located {geography}. The network aggregates de-identified patient-level "
        f"data including demographics, diagnoses (coded in ICD-10-CM), procedures "
        f"(CPT, HCPCS, ICD-10-PCS), medications (RxNorm, VA Class), laboratory "
        f"measurements (LOINC), and genomic data where available. All analyses "
        f"were conducted on the TriNetX Analytics platform on {query_date}."
    )

    if c.get("irb_exempt"):
        text += (
            " Because the TriNetX platform returns only aggregated, de-identified "
            "counts and statistics that meet criteria for a Limited Data Set under "
            "the U.S. Health Insurance Portability and Accountability Act (HIPAA) "
            "Privacy Rule, this study was deemed exempt from Institutional Review "
            "Board oversight and informed consent was not required."
        )
    elif c.get("irb_approved"):
        irb_name = _bracket(c.get("irb_name"), "name of IRB")
        irb_number = _bracket(c.get("irb_number"), "approval number")
        text += (
            f" The study protocol was reviewed and approved by the {irb_name} "
            f"(approval number {irb_number})."
        )

    return {
        "heading": "Data source",
        "text": text,
        "strobe_items": ["4 (Study design)", "5 (Setting)"],
    }


def para_study_design(c: Dict) -> Dict:
    design = c["study_design"]
    user_design = c.get("user_design", "Not applicable")
    index_event = _bracket(c.get("index_event"), "describe index event")
    study_period_start = _bracket(c.get("study_period_start"), "start date")
    study_period_end = _bracket(c.get("study_period_end"), "end date")

    text = (
        f"We conducted a {design.lower()} of adult patients in the network. "
        f"The index event was defined as {index_event}. The study period spanned "
        f"{study_period_start} to {study_period_end}."
    )

    if user_design == "New-user (incident) design":
        washout = _bracket(c.get("washout"), "washout window, e.g., 365 days")
        text += (
            f" A new-user (incident) design was applied: patients were required "
            f"to have no record of the exposure of interest in the {washout} prior "
            f"to the index date, in order to reduce prevalent-user bias and exclude "
            f"patients with established treatment patterns."
        )
    elif user_design == "Prevalent-user design":
        text += (
            " A prevalent-user design was used; we did not exclude patients with "
            "prior exposure to the agent of interest. Limitations of this approach, "
            "including potential depletion of susceptibles, are addressed in the "
            "Discussion."
        )

    if c.get("study_design_rationale"):
        text += " " + _t(c["study_design_rationale"])

    return {
        "heading": "Study design",
        "text": text,
        "strobe_items": ["4 (Study design)", "5 (Setting, dates)"],
    }


def para_eligibility(c: Dict) -> Dict:
    age_min = _bracket(c.get("age_min"), "minimum age")
    age_max = _t(c.get("age_max"))
    age_str = f"≥{age_min} years" if not age_max else f"{age_min}–{age_max} years"

    inclusion = _bracket(c.get("inclusion_criteria"),
                         "list inclusion criteria including diagnosis codes and required encounters")
    exclusion = _bracket(c.get("exclusion_criteria"),
                         "list exclusion criteria with codes")

    text = (
        f"Eligible patients were aged {age_str} at the index date. "
        f"Inclusion criteria were: {inclusion}. "
        f"Exclusion criteria were: {exclusion}."
    )

    if c.get("require_prior_encounter"):
        prior_window = _bracket(c.get("prior_encounter_window"),
                                "lookback window, e.g., 365 days")
        text += (
            f" To ensure adequate baseline information capture, we required at least "
            f"one healthcare encounter within {prior_window} prior to the index date."
        )

    if c.get("exclude_prior_outcome"):
        text += (
            " Patients with a record of the primary outcome at any time prior to "
            "the index date were excluded."
        )

    return {
        "heading": "Study population and eligibility",
        "text": text,
        "strobe_items": ["6 (Participants)", "7 (Variables)"],
    }


def para_exposure_comparator(c: Dict) -> Dict:
    exposure_name = _bracket(c.get("exposure_name"), "exposure name")
    exposure_codes = _bracket(c.get("exposure_codes"),
                              "code list, e.g., RxNorm codes for the medication")
    comparator_name = _bracket(c.get("comparator_name"), "comparator name")
    comparator_type = c.get("comparator_type", "Active comparator")

    if comparator_type == "Active comparator":
        comp_phrase = (
            f"an active comparator cohort comprising {comparator_name}"
        )
    elif comparator_type == "Non-user / unexposed comparator":
        comp_phrase = (
            f"a comparator cohort of unexposed patients ({comparator_name})"
        )
    elif comparator_type == "Historical comparator":
        comp_phrase = f"a historical comparator cohort ({comparator_name})"
    else:
        custom = _bracket(c.get("comparator_other"), "describe comparator")
        comp_phrase = custom

    text = (
        f"The exposed cohort consisted of patients with a record of {exposure_name} "
        f"({exposure_codes}) on or after the index date. The cohort was compared with "
        f"{comp_phrase}. Time zero for each patient was defined as the date of first "
        f"qualifying exposure (exposed cohort) or the date of first eligibility "
        f"(comparator cohort)."
    )

    if c.get("require_two_codes"):
        text += (
            " To improve specificity of exposure identification, at least two "
            "records of the qualifying code on separate dates were required."
        )

    return {
        "heading": "Exposure and comparator",
        "text": text,
        "strobe_items": ["7 (Variables)", "8 (Data sources / measurement)"],
    }


def para_outcomes(c: Dict) -> Dict:
    primary_outcome = _bracket(c.get("primary_outcome"), "primary outcome with codes")
    secondary_outcomes = _t(c.get("secondary_outcomes"))
    outcome_window_start = _bracket(c.get("outcome_window_start"),
                                    "start of outcome window, e.g., 1 day after index")
    outcome_window_end = _bracket(c.get("outcome_window_end"),
                                  "end of outcome window, e.g., 5 years after index")

    text = (
        f"The primary outcome was {primary_outcome}. Outcomes were ascertained from "
        f"{outcome_window_start} to {outcome_window_end} relative to the index date. "
        f"Patients were censored at the end of the outcome window, at last recorded "
        f"healthcare activity in the network, or at the occurrence of the outcome, "
        f"whichever came first."
    )

    if secondary_outcomes:
        text += f" Secondary outcomes included: {secondary_outcomes}."

    if c.get("outcome_first_only"):
        text += " Only the first occurrence of each outcome was counted per patient."

    return {
        "heading": "Outcomes",
        "text": text,
        "strobe_items": ["7 (Variables)", "8 (Data sources / measurement)"],
    }


def para_covariates_matching(c: Dict) -> Dict:
    covariates = _bracket(
        c.get("covariates"),
        "list of covariates: demographics, comorbidities (with code categories), "
        "medications, prior healthcare utilization, etc."
    )
    smd_threshold = _t(c.get("smd_threshold"), "0.1")
    matching_ratio = _t(c.get("matching_ratio"), "1:1")
    caliper = _t(c.get("caliper"), "0.1 pooled standard deviation")

    if c.get("use_psm"):
        text = (
            f"To address potential confounding by indication and baseline differences "
            f"between cohorts, we performed propensity score matching using the "
            f"TriNetX Analytics platform. Propensity scores were estimated using "
            f"logistic regression with the following covariates: {covariates}. "
            f"Patients were matched {matching_ratio} using a greedy nearest-neighbor "
            f"algorithm with a caliper of {caliper}. Balance was assessed using "
            f"standardized mean differences (SMDs), with values <{smd_threshold} "
            f"considered indicative of adequate balance."
        )
        if c.get("report_love_plot"):
            text += (
                " A Love plot displaying SMDs before and after matching was generated "
                "using the TriNetX Publication Toolkit Love Plot Generator."
            )
    elif c.get("use_adjustment"):
        text = (
            f"To address potential confounding, we adjusted for the following covariates "
            f"in multivariable models: {covariates}."
        )
    else:
        text = (
            "Crude (unadjusted) comparisons between cohorts were performed. "
            "Limitations of this approach are addressed in the Discussion."
        )

    return {
        "heading": "Covariates and confounding control",
        "text": text,
        "strobe_items": ["7 (Variables)", "9 (Bias)", "12 (Statistical methods)"],
    }


def para_statistical_analysis(c: Dict) -> Dict:
    effects = c.get("effect_estimates", [])
    if not effects:
        effects_str = "[specify which effect estimates were reported]"
    else:
        effects_str = _list_join([EFFECT_ESTIMATES_INLINE.get(e, e) for e in effects])

    alpha = _t(c.get("alpha"), "0.05")
    sided = _t(c.get("sided"), "two-sided")

    text = (
        f"For each outcome, we report {effects_str}. "
        f"Statistical significance was defined as a {sided} p-value <{alpha}."
    )

    has_km = any("Kaplan" in e or "Hazard" in e for e in effects)
    if has_km:
        text += (
            " For time-to-event analyses, Kaplan–Meier curves were generated and "
            "compared using the log-rank test, and hazard ratios were estimated "
            "using Cox proportional hazards regression. The proportional hazards "
            "assumption was assessed by visual inspection of Kaplan–Meier curves "
            "and complementary diagnostics where appropriate."
        )

    mc_method = c.get("multiple_comparisons")
    if mc_method and mc_method != "No correction applied (single primary outcome)":
        text += (
            f" To account for multiple outcome testing, p-values were adjusted using "
            f"the {mc_method} method as implemented in the TriNetX Publication "
            f"Toolkit Multiple Comparisons Correction Tool."
        )
    elif mc_method == "No correction applied (single primary outcome)":
        text += (
            " A single primary outcome was prespecified; no multiple-comparisons "
            "correction was applied to the primary analysis. Secondary outcomes are "
            "interpreted as exploratory."
        )

    sensitivities = c.get("sensitivity_analyses", [])
    if sensitivities:
        sens_str = _list_join(
            [SENSITIVITY_ANALYSES_INLINE.get(s, s) for s in sensitivities]
        )
        text += f" Pre-specified sensitivity analyses included {sens_str}."

    return {
        "heading": "Statistical analysis",
        "text": text,
        "strobe_items": ["12 (Statistical methods)"],
    }


def para_software(c: Dict) -> Dict:
    toolkit_version = _t(c.get("toolkit_version"), "current version at time of analysis")
    text = (
        f"Analyses were conducted within the TriNetX Analytics web platform. "
        f"Manuscript-ready tables, figures, and reporting diagnostics were generated "
        f"using the TriNetX Publication Toolkit ({toolkit_version}), an open-source "
        f"Streamlit application for formatting TriNetX exports into publication "
        f"outputs. Reporting completeness was assessed against the STROBE checklist "
        f"for observational studies using the toolkit's STROBE Assessment Tool."
    )
    return {
        "heading": "Software and reporting",
        "text": text,
        "strobe_items": ["12 (Statistical methods)", "STROBE compliance"],
    }


# ============================================================
# Aggregator
# ============================================================
def build_all_paragraphs(c: Dict) -> List[Dict]:
    paragraphs = [
        para_data_source(c),
        para_study_design(c),
        para_eligibility(c),
        para_exposure_comparator(c),
        para_outcomes(c),
        para_covariates_matching(c),
        para_statistical_analysis(c),
        para_software(c),
    ]
    return paragraphs


# ============================================================
# DOCX export
# ============================================================
def export_docx(paragraphs: List[Dict], title: str = "Methods") -> bytes:
    if not DOCX_AVAILABLE:
        raise RuntimeError("python-docx is not installed.")
    doc = Document()

    # Default style
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

    # Title
    h = doc.add_heading(title, level=1)
    for run in h.runs:
        run.font.size = Pt(14)

    # Provenance line
    p_prov = doc.add_paragraph()
    r = p_prov.add_run(
        f"Draft generated by the TriNetX Publication Toolkit Methods Section "
        f"Scaffolder on {_dt.date.today().isoformat()}. Verify every clinical "
        f"detail, code list, and date before submission."
    )
    r.italic = True
    r.font.size = Pt(9)

    # Body
    for p in paragraphs:
        doc.add_heading(p["heading"], level=2)
        body = doc.add_paragraph(p["text"])
        body.paragraph_format.space_after = Pt(8)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ============================================================
# UI
# ============================================================
st.header("Step 1 · Data source and approvals")

col1, col2 = st.columns(2)
with col1:
    network = st.selectbox("TriNetX network used", TRINETX_NETWORKS, index=0)
    network_other = ""
    if network == "Other (specify in custom text)":
        network_other = st.text_input("Specify network name")
    hco_count = st.text_input(
        "Approximate number of contributing HCOs at query time",
        placeholder="e.g., 120",
    )
    geography = st.text_input(
        "Geographic coverage description",
        placeholder="e.g., primarily in the United States",
    )
with col2:
    query_date = st.text_input(
        "Date the TriNetX query was executed",
        placeholder="e.g., 15 March 2025",
    )
    irb_status = st.radio(
        "IRB status",
        [
            "IRB exempt (aggregated, de-identified data only)",
            "IRB approved",
            "Not yet determined",
        ],
        index=0,
    )
    irb_name = ""
    irb_number = ""
    if irb_status == "IRB approved":
        irb_name = st.text_input("IRB name")
        irb_number = st.text_input("IRB approval number")

st.divider()
st.header("Step 2 · Study design")

col1, col2 = st.columns(2)
with col1:
    study_design = st.selectbox("Study design", STUDY_DESIGNS, index=0)
    user_design = st.selectbox(
        "New-user vs prevalent-user design", USER_DESIGNS, index=0
    )
    washout = ""
    if user_design == "New-user (incident) design":
        washout = st.text_input(
            "Washout period before index date",
            placeholder="e.g., 365 days",
        )
with col2:
    study_period_start = st.text_input(
        "Study period start", placeholder="e.g., 1 January 2010"
    )
    study_period_end = st.text_input(
        "Study period end", placeholder="e.g., 31 December 2024"
    )
    index_event = st.text_area(
        "Index event definition (will appear verbatim in draft)",
        placeholder="e.g., first recorded prescription for drug X",
        height=70,
    )

study_design_rationale = st.text_area(
    "Optional: additional design rationale to include verbatim",
    placeholder="Leave blank to skip.",
    height=60,
)

st.divider()
st.header("Step 3 · Study population")

col1, col2, col3 = st.columns(3)
with col1:
    age_min = st.text_input("Minimum age (years)", value="18")
with col2:
    age_max = st.text_input("Maximum age (years; blank for no upper limit)", value="")
with col3:
    require_prior_encounter = st.checkbox(
        "Require ≥1 prior healthcare encounter", value=True
    )

prior_encounter_window = ""
if require_prior_encounter:
    prior_encounter_window = st.text_input(
        "Prior encounter lookback window",
        value="365 days",
    )

exclude_prior_outcome = st.checkbox(
    "Exclude patients with prior occurrence of the primary outcome", value=True
)

inclusion_criteria = st.text_area(
    "Inclusion criteria (will appear verbatim in draft)",
    placeholder="e.g., adults with ≥1 inpatient or ≥2 outpatient ICD-10 codes for ...",
    height=80,
)
exclusion_criteria = st.text_area(
    "Exclusion criteria (will appear verbatim in draft)",
    placeholder="e.g., pregnancy at index, history of malignancy, ...",
    height=80,
)

st.divider()
st.header("Step 4 · Exposure and comparator")

col1, col2 = st.columns(2)
with col1:
    exposure_name = st.text_input(
        "Exposure name", placeholder="e.g., GLP-1 receptor agonist"
    )
    exposure_codes = st.text_area(
        "Exposure code list (will appear verbatim)",
        placeholder="e.g., RxNorm codes for semaglutide, liraglutide, ...",
        height=70,
    )
    require_two_codes = st.checkbox(
        "Require ≥2 records of the qualifying code on separate dates", value=False
    )
with col2:
    comparator_type = st.selectbox("Comparator type", COMPARATOR_TYPES, index=0)
    comparator_name = st.text_input(
        "Comparator cohort label",
        placeholder="e.g., DPP-4 inhibitor users",
    )
    comparator_other = ""
    if comparator_type == "Other (specify in custom text)":
        comparator_other = st.text_input("Describe the comparator")

st.divider()
st.header("Step 5 · Outcomes")

primary_outcome = st.text_area(
    "Primary outcome (will appear verbatim, including code list)",
    placeholder="e.g., incident myocardial infarction defined by ICD-10-CM I21.x",
    height=70,
)
secondary_outcomes = st.text_area(
    "Secondary outcomes (optional, will appear verbatim)",
    placeholder="Leave blank if none.",
    height=70,
)

col1, col2, col3 = st.columns(3)
with col1:
    outcome_window_start = st.text_input(
        "Outcome window start", value="1 day after index"
    )
with col2:
    outcome_window_end = st.text_input(
        "Outcome window end", value="5 years after index"
    )
with col3:
    outcome_first_only = st.checkbox("Count first occurrence only", value=True)

st.divider()
st.header("Step 6 · Confounding control")

confounding_strategy = st.radio(
    "Confounding control strategy",
    ["Propensity score matching", "Multivariable adjustment", "Crude (no adjustment)"],
    index=0,
)
use_psm = confounding_strategy == "Propensity score matching"
use_adjustment = confounding_strategy == "Multivariable adjustment"

covariates = st.text_area(
    "Covariates (will appear verbatim — include all variables used for matching/adjustment)",
    placeholder=(
        "e.g., age, sex, race/ethnicity, BMI category, smoking status, "
        "Charlson comorbidity components (ICD-10), prior medication classes, "
        "healthcare utilization in baseline window..."
    ),
    height=120,
)

if use_psm:
    col1, col2, col3 = st.columns(3)
    with col1:
        matching_ratio = st.text_input("Matching ratio", value="1:1")
    with col2:
        caliper = st.text_input("Caliper width", value="0.1 pooled standard deviation")
    with col3:
        smd_threshold = st.text_input("SMD balance threshold", value="0.1")
    report_love_plot = st.checkbox(
        "Include a Love plot of pre/post-match SMDs", value=True
    )
else:
    matching_ratio = ""
    caliper = ""
    smd_threshold = "0.1"
    report_love_plot = False

st.divider()
st.header("Step 7 · Statistical analysis")

effect_estimates = st.multiselect(
    "Effect estimates reported",
    EFFECT_ESTIMATES,
    default=[
        "Risk ratio (RR) with 95% CI",
        "Hazard ratio (HR) with 95% CI from Cox proportional hazards",
        "Kaplan–Meier survival with log-rank test",
    ],
)

col1, col2, col3 = st.columns(3)
with col1:
    alpha = st.text_input("Significance threshold (alpha)", value="0.05")
with col2:
    sided = st.selectbox("Test sidedness", ["two-sided", "one-sided"], index=0)
with col3:
    multiple_comparisons = st.selectbox(
        "Multiple comparisons correction",
        MULTIPLE_COMPARISONS_METHODS,
        index=0,
    )

sensitivity_analyses = st.multiselect(
    "Pre-specified sensitivity / subgroup analyses",
    SENSITIVITY_ANALYSES,
    default=["E-value calculation for unmeasured confounding"],
)

st.divider()
st.header("Step 8 · Software")

toolkit_version = st.text_input(
    "TriNetX Publication Toolkit version (optional)",
    placeholder="e.g., v1.0",
)


# ============================================================
# Assemble choices and render
# ============================================================
choices = {
    "network": network,
    "network_other": network_other,
    "hco_count": hco_count,
    "geography": geography,
    "query_date": query_date,
    "irb_exempt": irb_status == "IRB exempt (aggregated, de-identified data only)",
    "irb_approved": irb_status == "IRB approved",
    "irb_name": irb_name,
    "irb_number": irb_number,
    "study_design": study_design,
    "user_design": user_design,
    "washout": washout,
    "study_period_start": study_period_start,
    "study_period_end": study_period_end,
    "index_event": index_event,
    "study_design_rationale": study_design_rationale,
    "age_min": age_min,
    "age_max": age_max,
    "require_prior_encounter": require_prior_encounter,
    "prior_encounter_window": prior_encounter_window,
    "exclude_prior_outcome": exclude_prior_outcome,
    "inclusion_criteria": inclusion_criteria,
    "exclusion_criteria": exclusion_criteria,
    "exposure_name": exposure_name,
    "exposure_codes": exposure_codes,
    "require_two_codes": require_two_codes,
    "comparator_type": comparator_type,
    "comparator_name": comparator_name,
    "comparator_other": comparator_other,
    "primary_outcome": primary_outcome,
    "secondary_outcomes": secondary_outcomes,
    "outcome_window_start": outcome_window_start,
    "outcome_window_end": outcome_window_end,
    "outcome_first_only": outcome_first_only,
    "use_psm": use_psm,
    "use_adjustment": use_adjustment,
    "covariates": covariates,
    "matching_ratio": matching_ratio,
    "caliper": caliper,
    "smd_threshold": smd_threshold,
    "report_love_plot": report_love_plot,
    "effect_estimates": effect_estimates,
    "alpha": alpha,
    "sided": sided,
    "multiple_comparisons": multiple_comparisons,
    "sensitivity_analyses": sensitivity_analyses,
    "toolkit_version": toolkit_version,
}

st.divider()
st.header("Generated draft")

paragraphs = build_all_paragraphs(choices)

# Render preview
for p in paragraphs:
    st.subheader(p["heading"])
    st.markdown(p["text"])
    st.caption("STROBE items addressed: " + ", ".join(p["strobe_items"]))

# Show plain-text version for copy/paste
with st.expander("Plain-text version (for copy-paste)", expanded=False):
    full_text = ""
    for p in paragraphs:
        full_text += f"\n## {p['heading']}\n\n{p['text']}\n"
    st.code(full_text, language="markdown")

# Verification gate
st.divider()
st.header("Verification gate (required before Word export)")

st.markdown(
    "Please confirm each of the following before downloading. These items cannot "
    "be checked automatically and require your independent verification against the "
    "TriNetX query you actually ran."
)

v1 = st.checkbox(
    "I have verified that the code lists in this draft match the cohort "
    "definitions used in TriNetX."
)
v2 = st.checkbox(
    "I have verified that the index date, study period, lookback window, and "
    "outcome window match my actual analysis."
)
v3 = st.checkbox(
    "I have verified that the listed covariates are exactly those used for "
    "propensity score matching or adjustment in TriNetX."
)
v4 = st.checkbox(
    "I have reviewed the TriNetX network description and HCO count against the "
    "current TriNetX documentation."
)
v5 = st.checkbox(
    "I understand that this draft is a starting scaffold and not a substitute "
    "for senior methodological review."
)

gate_open = all([v1, v2, v3, v4, v5])

if gate_open and DOCX_AVAILABLE:
    docx_bytes = export_docx(paragraphs)
    st.download_button(
        "Download Methods section as Word document",
        data=docx_bytes,
        file_name="trinetx_methods_draft.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
elif not DOCX_AVAILABLE:
    st.warning("python-docx is not installed; Word export is unavailable.")
else:
    st.info("Tick all five verification items above to unlock the Word export.")
