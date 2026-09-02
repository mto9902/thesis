# Source-Fidelity Audit

Updated: 2 September 2026

## Decision rule

A framework element passes only when its construct name and directional relationship can be located in a published framework or hypothesis. A questionnaire item passes only when its wording and construct mapping can be located in the published instrument. The author's framework may select and combine published elements, but it must identify the source of every box and arrow.

## Current audit result

| Element | Evidence | Result |
|---|---|---|
| Published frameworks | Venkatesh and Bala (2008), Figure 2; Miao et al. (2024), Figure 1; Steenkamp et al. (2024), Figure 1 | Pass |
| Author framework | Five selected constructs and four selected paths, separately labelled as constructed by the author | Pass |
| Construct labels | All five appear in Steenkamp et al. (2024), Figure 1 | Pass |
| Relationships | All four appear as explicit Steenkamp hypotheses and were supported in that sample | Pass |
| Hypothesis direction | All four preserve the published positive direction | Pass |
| Questionnaire | PU1-PU4, PEOU1-PEOU4, SN1-SN4, REL1-REL3, and BI1-BI3 appear in the Steenkamp Appendix | Pass |
| Scale | Steenkamp's seven-point source scale is documented; the field instrument uses the explicitly requested five-point adaptation while retaining all item statements | Pass with disclosed response-format adaptation |
| Context changes | Thai IT population and local translation/pilot are identified as current-study decisions | Pass |
| Outcome boundary | Student intention is measured; actual employment and employer valuation are not claimed | Pass |

The executable audit in `tools/audit_direct_replication.py` checks the source PDF, direct source-image crops, selected labels and paths, all 18 questionnaire items, the disclosed BI3 correction, and both Word documents. Its final result must be `AUTHOR-FRAMEWORK AUDIT: PASS` before submission.

## Important interpretation

"Constructed by the author" does not mean invented without evidence. It means the author selected a smaller set of published constructs and published relationships for a new context, showed the source frameworks first, and cited the construction clearly.
