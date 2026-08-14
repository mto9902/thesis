# Current Thesis Study

This folder is the portable working package for Min Thiha Oo's current master's thesis reconstruction.

## Current route

**Understanding Thai IT Students' Intentions to Use University-Issued Digital Badges to Showcase Employability Skills: A Direct-Model Replication of Steenkamp, Fisher, and Nesbit (2024)**

The framework, 13 hypotheses, 11 constructs, and 40 coded questionnaire items now come from one published study. The source study investigated accounting students in New Zealand. The thesis changes the population to Thai university students in IT-related programs and treats a university-issued digital badge as the verifiable representation of an IT micro-credential. It does not add, remove, or rename paths in the published model.

The earlier expectancy-value-cost (EVC) files remain in this folder as an audited historical route. They are not the current design and must not be submitted as a direct replication.

## Start here

1. Read [CURRENT_STATE.md](CURRENT_STATE.md) for the plain-language status.
2. Read [DIRECT_SOURCE_AUDIT.md](DIRECT_SOURCE_AUDIT.md) for the strict source-fidelity decision.
3. Read [DIRECT_REPLICATION_SPEC.md](DIRECT_REPLICATION_SPEC.md) for the exact constructs, hypotheses, items, and disclosed changes.
4. Open `outputs/Master_Thesis_Direct_Replication_Digital_Badges_Thai_IT_Students.docx` for the chaptered draft.
5. Open `outputs/Evidence_Pack_Direct_Framework_Hypotheses_Questionnaire.docx` for the source screenshots and item-level provenance.
6. Read [SOURCE_INDEX.md](source/SOURCE_INDEX.md) before changing the framework or questionnaire.

## Rebuild the current documents

From the repository root:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r current-study/tools/requirements.txt
python current-study/tools/build_direct_replication.py
```

Then run the literal source audit:

```bash
python current-study/tools/audit_direct_replication.py
```

The audit compares all 13 hypotheses and all 40 coded items with the published source PDF and verifies their presence in both current Word documents.

The older `build_documents.py` script rebuilds only the archived EVC route.

## Package boundaries

- The checked-in Word files are proposal-stage review copies.
- `build_direct_replication.py` is the editable source of truth for the current documents.
- The Steenkamp et al. article is retained because it contains the exact published model, hypotheses, and questionnaire appendix used here.
- Private email, health information, AU credentials, proxy settings, temporary renders, and participant data are excluded.
- Chapters 5 and 6 contain placeholders only. No findings are claimed before real data collection and analysis.
