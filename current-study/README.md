# Current Thesis Study

This folder is the portable working package for Min Thiha Oo's current master's thesis reconstruction. It is intended to be the starting point on a new computer or in a new Codex session.

## Current topic

**Factors Influencing Intention to Use Digital Credentials for IT Micro-Credentials among Thai University Students: An Expectancy-Value-Cost Study**

This replaces the earlier survey experiment about recognition, stackability, industry endorsement, trust, and perceived value. The earlier files at the repository root are retained as research history; they are not the current design.

## Start here

1. Read [CURRENT_STATE.md](CURRENT_STATE.md) for the plain-language research status and remaining work.
2. Read [DIRECT_SOURCE_AUDIT.md](DIRECT_SOURCE_AUDIT.md) for the strict pass/fail audit requested after Professor Kimi's feedback.
3. Open [Master_Thesis_EVC_Digital_Credentials_Thai_IT_Students.docx](outputs/Master_Thesis_EVC_Digital_Credentials_Thai_IT_Students.docx) only as an audited working draft, not a submission-ready direct replication.
4. Open [Evidence_Pack_Published_Framework_and_Questionnaire.docx](outputs/Evidence_Pack_Published_Framework_and_Questionnaire.docx) for the borrowed framework, questionnaire origins, and adaptation trail.
5. Read [SOURCE_INDEX.md](source/SOURCE_INDEX.md) before changing the framework or questionnaire.

## Rebuild the documents

From the repository root on macOS or Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r current-study/tools/requirements.txt
python current-study/tools/build_documents.py
```

On Windows, activate the environment with `.venv\\Scripts\\activate` and run the same two Python commands. The builder resolves all paths relative to this folder and recreates both Word files and both figures in `current-study/outputs/`.

## Package boundaries

- The checked-in Word files are review copies that can be opened without running code.
- The builder script is the editable source of truth for generated prose, tables, questionnaire items, and diagrams.
- The Kiiskilä et al. article is included because it is the exact open-access framework source used in the evidence pack.
- Private email exports, medical information, AU credentials, proxy settings, temporary PDF renders, and document QA screenshots are intentionally excluded.
- Do not enter participant data, names, email addresses, or consent records in this public repository.
