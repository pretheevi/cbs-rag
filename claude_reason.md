# Canonical MongoDB JSON schema for the two test-case workbooks

## What I found in the files

**ALL MODULES_TCs.xlsx** (about 2,770 test cases)
- It has 7 module sheets (FrontDesk, Billing, Nurse, Message, Provider, NewCalendar, PracticeManagement). They mostly share the same columns.
- Every test has **two tracks**: Firebird (high-level) and MySQL (full regression). Each track has its own Owner, Status and Comment.
- A row with only column A filled is a **section header**, for example "Register a New patient - Insurance". Every row below it belongs to that section.
- The `Overview` and `Split` sheets are summaries: counts per module and team assignments. They are not test cases.

**InteropsTeamFunctionalityDetails.xlsx** (19 sheets, all with different layouts)
- `IntropsFeatures` is a **master catalog**. It lists each feature, the environment IDs it runs in (for example ADT runs in 93, 8 and 10), its sub-features (ADT_A04, SIU_S12 and so on) and links to config documents.
- Many sheets start with Precondition or Configuration rows that apply to every test below them (eFax, DTI, DRP, KN02).
- In eRx, a parent "Testcase - N" block holds the patient, medication and pharmacy for the child rows under it.
- The data is dirty. Statuses appear as `Pass`, `pass` and `Pass `. Row 26 says **Pass** but its comment says "Canel - Pending". The header "Excepted Result" is a typo.
- The same entities appear across sheets. For example, patients 118 and 123 are used in both eRx and ePA.

## Gaps in the current `data-preprocess.md` schema

1. It only fits eRx. It needs one shape that works for both workbooks.
2. `execution.status` is a single value. ALL MODULES needs a list of runs, one per track.
3. The full parent context is copied into every child row. It's better to store it once and keep only a small copy in each child.
4. Relationships are only implied in text. The LLM can't follow them reliably.
5. Status sits inside the keywords and the embedded text. Status changes often, and every change would force a re-embed.

## The canonical format

Use **one collection** (`qa_knowledge`) with a `doc_type` field. There are 3 document types. Keeping them in one collection means one vector index covers everything.

### 1. `test_case`: the main document (example from eRx row 27)

```json
{
  "_id": "TC:INTEROP:eRx:R27",
  "doc_type": "test_case",
  "schema_version": "1.0",

  "source": {
    "workbook": "InteropsTeamFunctionalityDetails",
    "sheet": "eRx",
    "row": 27,
    "raw_headers": { "title": "Test Scenario", "steps": "Steps", "expected": "Excepted Result" },
    "content_hash": "sha256:9f2c...",
    "ingested_at": "2026-10-02T00:00:00Z"
  },

  "hierarchy": {
    "domain": "Interop",
    "module": "eRx",
    "module_full_name": "e-Prescribing (Surescripts)",
    "section": "CancelRx",
    "group_id": "Testcase - 1",
    "sequence_no": 26,
    "breadcrumb": "Interop > eRx > Testcase - 1 > CancelRx"
  },

  "title": "Validate - RxCancelResponse to OP",
  "preconditions": [],
  "steps": [
    "In OP > e-prescribe > Cancelled/Denied cancel tab, double click on the CancelResponse transaction.",
    "Or click on the message icon of that transaction.",
    "It opens the XML.",
    "Validate all the fields of CancelRx request are populated correctly. Fields: Denial Reason, Code & Type."
  ],
  "expected_result": "All the fields must be populated correctly in the XML of the CancelResponse.",

  "context_ref": "CTX:INTEROP:eRx:TESTCASE-1",
  "context_snapshot": {
    "patient_id": "118",
    "medication": "Augmented Betamethasone 0.05% Topical Ointment",
    "pharmacy": "Shollenberger Pharmacy",
    "substance_class": "NON_CONTROLLED"
  },

  "entities": [
    { "type": "patient",      "value": "118" },
    { "type": "medication",   "value": "Augmented Betamethasone 0.05% Topical Ointment" },
    { "type": "pharmacy",     "value": "Shollenberger Pharmacy" },
    { "type": "message_type", "value": "CancelRxResponse", "standard": "NCPDP SCRIPT" },
    { "type": "system",       "value": "OP" },
    { "type": "system",       "value": "Surescripts" },
    { "type": "ui_screen",    "value": "e-prescribe > Cancelled/Denied cancel tab" }
  ],
  "environments": ["ID1", "ID93", "ID106"],

  "runs": [
    { "track": "default", "scope": null, "owner": null,
      "status": "FAIL", "status_raw": "Fail", "comment": null, "cycle": null }
  ],

  "relations": [
    { "type": "BELONGS_TO",  "target": "FEAT:INTEROP:eRx" },
    { "type": "USES_CONTEXT", "target": "CTX:INTEROP:eRx:TESTCASE-1" },
    { "type": "DEPENDS_ON",  "target": "TC:INTEROP:eRx:R26", "why": "CancelResponse must be sent first" },
    { "type": "SAME_SCENARIO", "target": "TC:INTEROP:eRx:R57", "why": "same check, Testcase - 2 / patient 123" }
  ],

  "quality": {
    "flags": ["HEADER_TYPO_FIXED"],
    "missing_fields": []
  },

  "search": {
    "keywords": ["RxCancelResponse", "CancelRx", "XML", "Denial Reason", "Surescripts", "118"],
    "aliases": ["Cancel Response", "CanelResponse", "Cancel Rx"]
  },

  "content_md": "# eRx > Testcase - 1 > CancelRx\n## Validate - RxCancelResponse to OP\n**Context:** Patient 118, Augmented Betamethasone, Shollenberger Pharmacy (non-controlled)\n**Depends on:** Send - RxCancelResponse to OP\n### Steps\n1. ...\n### Expected\nAll the fields must be populated correctly...",

  "embedding": {
    "model": "<your-embedding-model>",
    "dims": 1536,
    "text_hash": "sha256:41ab...",
    "vector": [0.0142, -0.0418, "..."]
  }
}
```

**The same shape for an ALL MODULES row** (FrontDesk row 3). Only these parts look different:

```json
{
  "_id": "TC:CORE:FrontDesk:R3",
  "hierarchy": { "domain": "Core", "module": "FrontDesk",
                 "section": "Login: Show a popup when the user enters the wrong password" },
  "attributes": { "tc_source": "Insprint", "is_high_level": true },
  "runs": [
    { "track": "Firebird", "scope": "HIGH_LEVEL",      "owner": "Marcelo", "status": "NOT_RUN", "comment": null },
    { "track": "MySQL",    "scope": "FULL_REGRESSION", "owner": "Marcelo", "status": "NOT_RUN", "comment": null }
  ]
}
```

### 2. `feature`: one per row in `IntropsFeatures`

```json
{
  "_id": "FEAT:INTEROP:HUB_ADT",
  "doc_type": "feature",
  "name": "ADT",
  "parent": "FEAT:INTEROP:HUB",
  "test_sheet": "HUB_ADT&SIU",
  "environments": ["ID93", "ID8", "ID10"],
  "directions": ["INBOUND", "OUTBOUND"],
  "sub_features": [
    { "code": "ADT_A04", "action": "Patient Creation" },
    { "code": "ADT_A08", "action": "Patient Updation" },
    { "code": "SIU_S12", "action": "Appointment Creation" }
  ],
  "doc_links": [{ "title": "ADT_TestingDocForAllEnv", "url": "https://docs.google.com/..." }],
  "content_md": "...",
  "embedding": { "...": "..." }
}
```

### 3. `context`: preconditions, configuration and parent blocks, stored once

```json
{
  "_id": "CTX:INTEROP:eRx:TESTCASE-1",
  "doc_type": "context",
  "context_kind": "TEST_DATA",
  "applies_to": { "sheet": "eRx", "rows": [3, 27] },
  "data": {
    "patient_id": "118", "patient_name": "Zachary Delaplaine",
    "pharmacy": "Shollenberger Pharmacy", "primary_dx": "Psoriasis",
    "medication": "Augmented Betamethasone 0.05% Topical Ointment",
    "quantity": 14.555, "unit": "Gram", "refills": 3, "substitution_allowed": true,
    "substance_class": "NON_CONTROLLED"
  },
  "content_md": "...",
  "embedding": { "...": "..." }
}
```

`context_kind` can be `TEST_DATA`, `PRECONDITION` or `CONFIGURATION`. The eFax and KN02 config rows and the DTI/DRP portal-preference flags all fit here.

## Why each part is there

| Field | What it fixes |
|---|---|
| `doc_type` + one collection | Both workbooks and all their layouts fit one shape and one vector index. |
| `hierarchy.breadcrumb` | Every chunk knows where it lives, even with no parent row nearby. |
| `steps[]` as an array | The LLM can quote "step 3" exactly, and you can find tests that share steps. |
| `context_ref` + `context_snapshot` | Full details are stored once. The child keeps just enough to stand on its own. |
| `entities[]` (typed) | Patient 118 in eRx and patient 118 in ePA become the same thing, so you can join across sheets. |
| `runs[]` | Handles Firebird and MySQL tracks, owners, release cycles and run history. |
| `status` + `status_raw` | Clean values for filters, and the original text for audit. |
| `relations[]` | Explicit edges let MongoDB's `$graphLookup` answer multi-hop questions ("what breaks if R26 fails?"). |
| `quality.flags` | Bad data is visible instead of silently trusted. For example, row 26 would get a `STATUS_COMMENT_CONFLICT` flag. |
| `content_hash` / `text_hash` | Re-embed only when the test text changes, not when a status changes. |
| Status kept out of `content_md` and `embedding` | Status changes often, so it lives in structured fields only. |

Two smaller notes:
- **`Overview`:** don't store it. Compute those counts on demand with an aggregation over `runs.status`.
- **`Split`:** store it as one small `policy` document, because it holds a rule: Feature Release means Full Regression on MySQL.

**Next step for our discussion:** how the LLM uses this format. That covers filtering first and then running vector search, joining `context` and `feature` documents at answer time, following `relations` for complex questions, and the markdown-plus-vector hybrid. Tell me when you want to go into it.