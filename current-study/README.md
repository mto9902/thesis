# Current Thesis Study

This folder is the portable working package for Min Thiha Oo's master's thesis.

The user restored this supervisor-accepted study on 5 September 2026. The online-certificate TPB alternative is saved separately on branch `backup/online-it-certificates-2026-09-05`; it is not active. The main Word document below is unchanged from the saved badge-study version with the subsequent supervisor-feedback revisions.

## Current study

**Factors Influencing Thai IT Students' Intentions to Use University-Issued Digital Badges for IT Micro-Credentials in Job Applications: A Quantitative Study of Usefulness, Job Relevance, Ease of Use, and Social Influence**

The study uses an author-constructed conceptual framework made from published constructs and published relationships. It is not a copy of one complete model. The final framework contains five constructs and four hypotheses:

1. Perceived Ease of Use -> Perceived Usefulness
2. Job Application Relevance -> Perceived Usefulness
3. Perceived Usefulness -> Behavioural Intention to Use DB
4. Subjective Norm -> Behavioural Intention to Use DB

All five exact construct labels, all four selected paths, and all 18 questionnaire items appear in Steenkamp, Fisher, and Nesbit (2024). TAM3 and the micro-credential model of Miao et al. (2024) provide additional theoretical and same-field support.

The English-language field questionnaire retains the 18 source statements on a five-point agreement scale. The required 40-person pilot checks clarity, comprehension, and reliability before main data collection. The analysis plan uses Cronbach's alpha, descriptive statistics, variance inflation factors, and two multiple linear regressions: one for H1-H2 and one for H3-H4.

## Start here

1. Read [CURRENT_STATE.md](CURRENT_STATE.md) for the plain-language status and remaining work.
2. Read [FRAMEWORK_SPEC.md](FRAMEWORK_SPEC.md) for the exact constructs, hypotheses, questionnaire codes, and source mapping.
3. Read [SOURCE_AUDIT.md](SOURCE_AUDIT.md) for the source-fidelity checks.
4. Open `outputs/Master_Thesis_Digital_Badges_Thai_IT_Students.docx` for the supervisor-facing Chapters 1-4 manuscript.
5. Use `outputs/Evidence_Pack_Author_Constructed_Framework_and_Questionnaire.docx` only as internal supporting material. The planned professor submission contains the main thesis file only.
6. Read [source/SOURCE_INDEX.md](source/SOURCE_INDEX.md) before changing any construct, arrow, hypothesis, or item.

## Rebuild and audit

From the repository root:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r current-study/tools/requirements.txt
python current-study/tools/build_direct_replication.py
python current-study/tools/audit_direct_replication.py
```

Use `--master-only` with the build command when only the thesis manuscript needs rebuilding.

## Package boundaries

- The thesis manuscript currently contains proposal-stage Chapters 1-4. Results and discussion require real data.
- `tools/build_direct_replication.py` is the editable source of truth despite its historical filename.
- The old EVC route and the full 11-variable Steenkamp replication are historical drafts, not the current design.
- Private correspondence, health information, AU credentials, proxy settings, participant data, and temporary renders are excluded from the portable package.
- No approval, pilot result, main-study result, or employer outcome may be claimed before it exists.
