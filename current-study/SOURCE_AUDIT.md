# Source-Fidelity Audit

Updated: 5 September 2026

This audit tests source fidelity only. It does not establish suitability for the new population, compliance with the supervisor's five-point-scale recommendation, completion of the literature requirements, approval, or readiness for fieldwork. The comparison with supervisor feedback identified open issues recorded at the top of `CURRENT_STATE.md`.

## Decision rule

A framework element passes only when the construct name and directional relationship can be located in a published framework or hypothesis. A questionnaire response passes only when its wording structure, referent or control factor, response anchors, and scoring procedure can be located in the published instrument. Every contextual change must be visible and disclosed.

## Current audit result

| Element | Evidence | Result |
|---|---|---|
| Primary published framework | Hunsinger and Smith (2008), Figure 2, p. 251 | Pass |
| Corroborating published framework | Wang (2023), Figure 1, p. 26 | Pass |
| Published source images | Pixel comparison with retained rendered source pages | Pass: unchanged direct crops |
| Author framework | Three paths common to both published models; separately labelled as constructed by the author | Pass |
| Construct labels | Attitude, Subjective Norm, Perceived Behavioral Control, and Behavioral Intention appear in both source models | Pass |
| Relationship directions | Hunsinger source H1-H3 and Wang source H3-H5 | Pass |
| Empirical support | Hunsinger and Smith (2008), Table 7 and p. 259; Wang (2023), Table 11 | Pass: all three paths supported in both samples |
| Questionnaire | Hunsinger and Smith (2008), pp. 252-254 | Pass: 26 responses represented |
| Referents | Six groups printed on p. 253 | Pass |
| Control factors | Four factors printed on p. 253 | Pass |
| Response formats | Source bipolar, semantic-differential, motivation, and facilitation scales | Pass |
| Time horizon | Next 12 months | Pass |
| Context change | `IT certification` changed consistently to the defined short online IT professional certificate | Pass with disclosed contextual substitution |
| Outcome boundary | Intention only; no claim of completion, learning, employer recognition, salary, or employment effects | Pass |

The executable audit in `tools/audit_online_certificates.py` checks the two source PDFs, direct source-image crops, the three paths, all 26 questionnaire response codes and complete item-table wording, required thesis sections, prohibited old wording, and the images embedded in the Word manuscript. Its current result is:

> ONLINE-CERTIFICATE SOURCE-FIDELITY AUDIT: PASS

## Honest limitation

The source questionnaire was developed for predominantly traditional IT certifications. The current study changes the target phrase so respondents answer about assessed short online IT professional certificates. That is a real contextual adaptation, not a verbatim replication. It is disclosed throughout the manuscript and is one reason the required pilot is necessary.

The direct framework, variables, path directions, item structures, referents, control factors, scales, scoring rules, and time horizon are published rather than invented.
