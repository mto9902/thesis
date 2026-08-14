# Direct-Source Fidelity Audit

Updated: 15 August 2026

## Decision rule

An element passes only when it can be traced to a published study without inventing a construct, relationship, hypothesis direction, or questionnaire statement. Geographic and population changes are permitted only when they are explicitly described as contextual replication changes.

## Current direct-model route

Primary source:

Steenkamp, N., Fisher, R., and Nesbit, T. (2024). *Understanding accounting students' intentions to use digital badges to showcase employability skills*. Accounting Education, 33(6), 906-934. https://doi.org/10.1080/09639284.2023.2276200

| Element | Exact source location | Verdict | Reason |
|---|---|---|---|
| Published framework image | Figure 1, article p. 10 (PDF p. 11) | Pass | The package reproduces the source figure. It is not redrawn or synthesized. |
| Thesis framework | Same complete Figure 1 | Pass | Every source node and arrow is retained; no path is added, removed, or redirected. |
| Variable names | Figure 1, Appendix, and Table 2 | Pass with disclosed source spelling variation | The source's 11 construct labels and abbreviations are retained. The article itself alternates between British and American spelling of behavioural/behavioral. |
| H1-H13 | Article pp. 11-13 (PDF pp. 12-14) | Pass | All 13 directional hypotheses are explicitly printed in the source. |
| Construct questionnaire | Appendix, article pp. 28-29 (PDF pp. 29-30) | Pass | All 40 coded source items and the seven-point scale are printed in the article. |
| Item-to-construct mapping | Appendix and Table 2 | Pass | Item codes map directly to the 11 source constructs. |
| Analysis model | Research method and results sections | Pass as methodological replication | The source specifies reflective PLS-SEM, SmartPLS, 5,000 bootstrap samples, reliability, validity, VIF, path, and R-squared assessment. |
| Thai IT population | Current study only | Context change, not a source claim | The source population is New Zealand accounting students. The new population is explicitly disclosed. |
| IT micro-credential badge context | Current study only, supported by the source's distinction between badges and micro-credentials | Context change, not a source claim | Construct items remain about university-issued digital badges. The introduction defines the badge as representing an IT micro-credential. |

## Questionnaire count check

The Appendix prints exactly 40 coded items:

| Construct | Codes | Count |
|---|---:|---:|
| Perceived Usefulness | PU1-PU4 | 4 |
| Perceived Ease of Use | PEOU1-PEOU4 | 4 |
| Computer Self-Efficacy | CSE1-CSE3 | 3 |
| Perceptions of External Control | PEC1-PEC4 | 4 |
| Computer Playfulness | CPLAY1-CPLAY4 | 4 |
| Computer Anxiety | CANX1-CANX4 | 4 |
| Subjective Norm | SN1-SN4 | 4 |
| Image | IMG1-IMG3 | 3 |
| Job Application Relevance | REL1-REL3 | 3 |
| Result Demonstrability | RES1-RES4 | 4 |
| Behavioral Intention | BI1-BI3 | 3 |
| **Total** |  | **40** |

## Disclosed source anomalies

### Unnumbered employer sentence

The source Appendix prints "Employers think that I should use university-issued digital badges when applying for jobs" after SN4 without an item code. The methods state that 40 items were used, and Table 2 reports only SN1-SN4. Therefore the current instrument does not invent SN5 or include the unnumbered sentence as a scored construct item.

### Items removed in the source analysis

The source administered 40 coded items and then removed PEC4, CPLAY4, and RES4 because their source-sample outer loadings were below .40. The current pilot will administer all 40 source items. Any removal will be based on the Thai data and reported, preserving the distinction between direct instrument adoption and sample-specific measurement decisions.

### BI3 typographical issue

The source Appendix prints "I would plan to use the when I apply for jobs..." The English source master and evidence pack reproduce that wording verbatim. A possible field correction changes only "the" to "them" and remains pending supervisor approval. No construct, referent, direction, or timeframe would change.

## Overall verdict

The current direct-model route passes the requested source-fidelity test for the framework, hypotheses, variables, and questionnaire. It should be described as a contextual replication, not as an original framework and not as a full literal replication of the source population or procedures.

## Automated verification

`tools/audit_direct_replication.py` independently checks the source PDF, direct image crops, and both Word documents. On 15 August 2026 it confirmed the pixel-exact source images, 13/13 hypotheses, 11/11 constructs, and 40/40 exact source items. The retained source PDF has SHA-256 `b4966d48b892f4fb041b1397937a25b1c3fec6c867d7a2b6ac758a4e67530752`.

The earlier EVC route remains a transparent historical draft but fails this strict standard because it inferred quantitative hypotheses from a qualitative framework and changed every questionnaire referent.
