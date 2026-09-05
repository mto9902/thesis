# Current Thesis Study

This folder is the portable working package for Min Thiha Oo's master's thesis.

## Active study

**Factors Influencing Thai IT Students' Intentions to Pursue Short Online IT Professional Certificates: A Theory of Planned Behavior Study**

The target is an assessed, short online IT programme offered by a company or university through Coursera, edX, Udacity, or a similar platform. Examples include certificates from Google, IBM, Microsoft, Meta, and universities. The study asks why adult undergraduate IT students at universities in Thailand intend to pursue one within the next 12 months.

The model uses the established core of the Theory of Planned Behavior:

1. Attitude -> Behavioral Intention
2. Subjective Norm -> Behavioral Intention
3. Perceived Behavioral Control -> Behavioral Intention

Hunsinger and Smith (2008) are the direct framework, hypothesis, variable, questionnaire, scoring, and time-horizon source. Wang (2023) supplies an independent published framework that tested and supported the same three paths among undergraduates in online learning. The study's own framework selects only these common paths and is labelled as constructed by the author.

The English questionnaire adopts 26 measurement responses printed by Hunsinger and Smith (2008). The only contextual substitution is `IT certification` -> `short online IT professional certificate`. The source's constructs, referents, control factors, seven-point response formats, and 12-month time horizon are retained.

## Start here

1. Read [CURRENT_STATE.md](CURRENT_STATE.md) for the plain-language status.
2. Read [FRAMEWORK_SPEC.md](FRAMEWORK_SPEC.md) for the exact model, hypotheses, and questionnaire map.
3. Read [SOURCE_AUDIT.md](SOURCE_AUDIT.md) for the source-fidelity result.
4. Open `outputs/Master_Thesis_Online_IT_Professional_Certificates_Thai_Students.docx` for the active Chapters 1-4 manuscript and Appendix questionnaire.
5. Read [source/SOURCE_INDEX.md](source/SOURCE_INDEX.md) before changing any variable, arrow, hypothesis, item, scale, referent, or control factor.

No separate evidence pack is part of the current submission. The published frameworks, source mapping, and complete questionnaire are included in the thesis manuscript itself.

## Rebuild and audit

From the repository root:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r current-study/tools/requirements.txt
python current-study/tools/build_online_certificates.py
python current-study/tools/audit_online_certificates.py
```

## Status boundaries

- Chapters 1-4 and the Appendix questionnaire are written, but the supervisor-feedback comparison identified unresolved scale, definition-depth, and reliability-plan issues. Read the status correction in `CURRENT_STATE.md` before treating the draft as ready for submission or pilot testing.
- The required 40-student pilot has not yet been conducted.
- Main data collection, results, discussion, and conclusions do not yet exist.
- Do not claim approval, pilot findings, main-study findings, certificate completion, learning outcomes, employer recognition, salary effects, or employment outcomes before evidence exists.
- The prior university-issued digital-badge study is preserved on branch `backup/university-issued-digital-badges-2026-09-04` and tag `university-issued-digital-badges-backup-2026-09-04`.
- Private correspondence, health information, AU credentials, proxy settings, participant data, and temporary renders remain outside the portable package.
