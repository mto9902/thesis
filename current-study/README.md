# Current Thesis Study

This folder is the portable working package for Min Thiha Oo's current master's thesis reconstruction.

## Current route

**Factors Influencing Thai IT Students' Intentions to Use University-Issued Digital Badges for IT Micro-Credentials in Job Applications: A Quantitative Study Using an Extended Technology Acceptance Model**

The framework, 13 hypotheses, 11 constructs, and 40 coded questionnaire items now come from one published study. The source study investigated accounting students in New Zealand. The thesis changes the population to Thai university students in IT-related programs and treats a university-issued digital badge as the verifiable representation of an IT micro-credential. It does not add, remove, or rename paths in the published model.

The earlier expectancy-value-cost (EVC) files remain in this folder as an audited historical route. They are not the current design and must not be submitted as a direct replication.

## Start here

1. Read [CURRENT_STATE.md](CURRENT_STATE.md) for the plain-language status.
2. Read [DIRECT_SOURCE_AUDIT.md](DIRECT_SOURCE_AUDIT.md) for the strict source-fidelity decision.
3. Read [DIRECT_REPLICATION_SPEC.md](DIRECT_REPLICATION_SPEC.md) for the exact constructs, hypotheses, items, and disclosed changes.
4. Open `outputs/Master_Thesis_Digital_Badges_Thai_IT_Students.docx` for the supervisor-facing thesis manuscript.
5. Open `outputs/Evidence_Pack_Direct_Framework_Hypotheses_Questionnaire.docx` for the source screenshots and item-level provenance.
6. Read [SOURCE_INDEX.md](source/SOURCE_INDEX.md) before changing the framework or questionnaire.

## Rebuild the current documents

From the repository root:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r current-study/tools/requirements.txt
python current-study/tools/build_direct_replication.py --master-only
```

Omit `--master-only` only when the evidence pack also needs to be rebuilt.

Then run the literal source audit:

```bash
python current-study/tools/audit_direct_replication.py
```

The audit compares all 13 hypotheses and all 40 coded items with the published source PDF and verifies their presence in both current Word documents.

The older `build_documents.py` script rebuilds only the archived EVC route.

## Package boundaries

- The checked-in thesis manuscript contains the proposal-stage Chapters 1-4; results and discussion will be added after data collection and analysis.
- `build_direct_replication.py` is the editable source of truth for the current documents.
- The Steenkamp et al. article is retained because it contains the exact published model, hypotheses, and questionnaire appendix used here.
- Private email, health information, AU credentials, proxy settings, temporary renders, and participant data are excluded.
- No findings are claimed before real data collection and analysis.
