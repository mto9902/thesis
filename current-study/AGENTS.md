# Codex Handoff Instructions

Treat this directory as the authoritative current study. Read `README.md`, `CURRENT_STATE.md`, `DIRECT_SOURCE_AUDIT.md`, `DIRECT_REPLICATION_SPEC.md`, and `source/SOURCE_INDEX.md` before editing.

- The current design is the direct-model replication of Steenkamp, Fisher, and Nesbit (2024) documented here. Root-level proposal files and the EVC outputs are historical context only.
- Preserve the complete published Figure 1 model. Do not add, remove, redirect, or redraw paths while calling the result direct adoption.
- Preserve H1-H13 and all 11 construct labels from the source.
- Keep the exact 40 coded Appendix items and seven-point scale as the English source master. Any proposed wording change must be item-level, explicit, and approved before the pilot.
- Do not invent SN5 from the source's unnumbered employer sentence. The source states 40 items and reports SN1-SN4.
- Preserve the BI3 source typo in the English source master and evidence column; disclose any proposed correction to "them" for approval before fielding.
- Do not silently pre-delete PEC4, CPLAY4, or RES4. They were administered and then removed in the source analysis; the Thai pilot must make its own reported measurement decision.
- Describe Thailand, IT students, the IT micro-credential badge context, translation, pilot, and sampling as contextual changes, not as source-study features.
- Do not claim that the study, framework, questionnaire, ethics application, pilot, or data collection has been approved unless a later message explicitly confirms it.
- Do not invent results. Chapters 5 and 6 remain placeholders until real pilot and main-study data exist.
- Edit `tools/build_direct_replication.py`, run it, and visually inspect both regenerated DOCX files after substantive changes.
- Run `tools/audit_direct_replication.py` after rebuilding; do not treat the package as source-complete if the audit fails.
- Keep private correspondence, health information, credentials, participant data, and temporary render artifacts out of Git.
