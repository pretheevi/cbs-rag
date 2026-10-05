Based on an in-depth analysis of your actual test case workbook (**`InteropsTeamFunctionalityDetails`** across tabs such as `eRx`, `HUB_837`, `HUB_ADT&SIU`, `ePA`, `IIS`, `PMX+`, `eFax`, etc.), here is an improved **MongoDB document schema** and **data pre-processing strategy** tailored specifically to the real-world columns, layout quirks, and healthcare interoperability domain of your sheet.

-----

### 1\. Key Insights & Challenges in Your Sheet

1.  **Hierarchical Parent-Child Structure (Row Merges / Sparse Columns)**:
      * In sheets like `eRx`, rows 2, 28, and 58 act as **Parent Context Blocks** (containing `Patient Name`, `Patient ID`, `Pharmacy`, `Primary diagnosis`, `Medication`, `Quantity`, `Refills`, and High-Level classification `Non-Controller Substances` vs. `Controller Substances`).
      * Subsequent rows (rows 3–7, 8–11, 12–23) are individual **Sub-Scenarios / Workflow Actions** (`Select Pharmacy`, `NewRx`, `Renewal/Refill Request`, `CancelRx`, `RxChange`, `RxFill`).
      * *Problem with naïve chunking*: If processed strictly row-by-row, child rows lose the patient details, medication name, and DEA classification.
2.  **Column Variations Across Sheets**:
      * Expected results appear as `Excepted Result` (typo in `eRx`), `Excepted Results` (`HUB_ADT&SIU`), `Expected Name` (`HUB_837`), `Expected Result` (`ePA`), or split across `Verification 1`, `Verification 2`, `Verification 3` (`eFax`).
      * Steps appear as `Steps`, `Test Steps`, `Testcase Steps`.
      * Separate columns for test data exist in tabs like `ePA`, `IIS`, and `DTI`, while in `eRx` it is embedded inside the scenario description.
3.  **Execution State & Defect Tracking**:
      * Columns F & G in `eRx` track real execution results (`Pass`, `Fail`) and defect/investigation remarks (e.g., `Canel - Pending`, `cancel pending`). Storing these in structured metadata unlocks queries like: *"Show all failing CancelRx testcases in OP/Surescripts"*.
4.  **Domain-Specific Interoperability Entities**:
      * The testcases span healthcare protocols: e-Prescribing (NCPDP / Surescripts `NewRx`, `RxRenewal`, `CancelRx`, `RxChange`, `RxFill`), Controlled Substances (`EPCS`, DEA Schedule), HL7 (`ADT_A04`, `SIU_S12`, `837`), and FHIR.

-----

### 2\. Improved MongoDB Document Schema

Here is the enhanced document structure adapted specifically for your testcase columns:

``` json
{
  "_id": "eRx_TC01_STEP_19",
  "source": {
    "workbook": "InteropsTeamFunctionalityDetails",
    "sheet_name": "eRx",
    "row_index": 19,
    "parent_row_index": 2
  },
  "metadata": {
    "module": "e-Prescribing",
    "sub_module": "Renewal/Refill Request",
    "feature_category": "Non-Controller Substances",
    "test_group_id": "Testcase - 1",
    "scenario_title": "Send - Renewal Response from OP - Deny",
    
    // Domain-specific extracted entities for exact filtering
    "interop": {
      "transaction_type": "RxRenewalResponse",
      "systems_involved": ["Office Practicum (OP)", "Surescripts Admin Console"],
      "action_type": "Deny",
      "controlled_substance": false,
      "dea_schedule": null
    },
    
    // Test data context inherited from parent block
    "test_data": {
      "patient_name": "Zachary Delaplaine",
      "patient_id": "118",
      "pharmacy": "Shollenberger Pharmacy",
      "medication": "Augmented Betamethasone 0.05% Topical Ointment",
      "primary_dx": "Psoriasis",
      "quantity": "14.555",
      "refills": "3"
    },
    
    // QA execution status (from Columns F & G)
    "execution": {
      "status": "Pass",
      "notes": null
    },
    
    "tags": ["eRx", "Surescripts", "Renewal", "Deny", "OP", "Prescription"]
  },

  "content": {
    "format": "markdown",
    "text": "# [eRx] Testcase - 1: Send - Renewal Response from OP - Deny\n\n## Context & Test Data\n- **Module / High Level:** e-Prescribing / Non-Controller Substances\n- **Transaction:** Renewal/Refill Request ~ Accept/Deny/Replace\n- **Patient:** Zachary Delaplaine (Pat ID: 118)\n- **Pharmacy:** Shollenberger Pharmacy\n- **Medication:** Augmented Betamethasone 0.05% Topical Ointment (Qty: 14.555, Refills: 3)\n- **Diagnosis:** Primary: Psoriasis\n\n## Preconditions / Notes\nNote that the Rx transaction will be removed from e-prescribe window. You will need to send a new RxRenewal Request to OP.\n\n## Steps to Reproduce\n1. In OP > e-prescribe > Refill/Change request, double click on the RxRenewal transaction.\n2. It opens the eRx renewal request summary window.\n3. Validate the required fields.\n4. Select the Denial reason from the dropdown.\n5. Then click on the Deny button to Deny the renewal request.\n\n## Expected Result\nMake sure RxRenewalResponse is received in Admin console (Surescript) which is sent from OP.\nImp-note: Rx transaction will be removed from the list.\n\n## Execution Status\n- **Status:** Pass"
  },

  "embedding": {
    "model": "text-embedding-3-small", // or nomic-embed-text
    "dimensions": 1536,
    "vector": [0.0142, -0.0418, 0.0211, "..."]
  },

  "search": {
    "keywords": [
      "Testcase - 1",
      "RxRenewalResponse",
      "Renewal",
      "Deny",
      "Refill",
      "Surescript",
      "Office Practicum",
      "OP",
      "Betamethasone",
      "118",
      "Shollenberger",
      "INT-RENEWALREQ",
      "Pass"
    ]
  }
}

```

-----

### 3\. Pre-Processing Pipeline Improvements

To generate this document automatically from the spreadsheet, include these steps in your data pre-processing script:

``` text
 ┌────────────────────────┐
 │ Raw Excel / Sheet Tabs │  (eRx, ePA, IIS, HUB_837, PMX+, etc.)
 └───────────┬────────────┘
             │
             ▼
 ┌────────────────────────┐
 │ 1. Column Normalizer   │  Map 'Excepted Result', 'Expected Name', 'Verification 1' -> expected_result
 └───────────┬────────────┘
             │
             ▼
 ┌────────────────────────┐
 │ 2. Context Inheritance │  Forward-fill Parent Testcase blocks (Patient, Med, Group ID) down to child steps
 └───────────┬────────────┘
             │
             ▼
 ┌────────────────────────┐
 │ 3. Entity & Keyword    │  Regex/NER extraction of Transaction Types (NewRx, CancelRx), Systems (OP, Surescripts),
 │    Extractor           │  Medications, and Execution status (Pass/Fail)
 └───────────┬────────────┘
             │
             ▼
 ┌────────────────────────┐
 │ 4. Markdown Synthesizer│  Generate clean, structured Markdown retaining complete context for LLM grounding
 └───────────┬────────────┘
             │
             ▼
 ┌────────────────────────┐
 │ 5. Embed & Upsert      │  Vector embedding on content.text + Atlas Vector & Keyword Index in MongoDB
 └────────────────────────┘

```

#### Step Details:

1.  **Column Mapping Normalization**:
    Define a unified dictionary across tabs:
    
      * **Title**: `Test Scenario` | `Test Name` | `TestcaseName` | `Scenario`
      * **Steps**: `Steps` | `Test Steps` | `Testcase Steps`
      * **Expected**: `Excepted Result` | `Excepted Results` | `Expected Name` | `Expected Result` | concatenate(`Verification 1`, `Verification 2`, `Verification 3`)
      * **Category**: `High Level` | `Workflow` | Sheet tab name
      * **Status / Results**: `Pass`/`Fail` (column F) and remarks (column G)

2.  **Parent Context Forward-Filling**:
    
      * When iterating through rows, detect parent headers (e.g., `A` starts with ` Testcase -  ` or contains multi-line patient parameters).
      * Store `current_parent_context = { patient_id, patient_name, medication, dx, dea_class }`.
      * For subsequent child rows, inject `current_parent_context` into the child's metadata and generated markdown. This ensures isolated steps like *"Close Pharmacy dropdown"* or *"Send - CancelRx"* know they belong to **Patient 118** and **Betamethasone**.

3.  **Specialized Keyword Generation**:
    
      * Combine abbreviations (`eRx`, `EPCS`, `OP`, `Surescripts`, `PA`, `FHIR`), specific codes (e.g., `INT-RENEWALREQ-1a`, `RxFill-Naproxen-2a`, `RelatesToMessageId`, `DEA Schedule`), and patient/medication names.
      * This guarantees BM25 keyword search hits on exact error messages, template IDs, and transaction types.

4.  **MongoDB Indexing Configuration**:
    
      * **Atlas Vector Search Index**: on `embedding.vector` with `cosine` similarity, indexed alongside metadata filter fields (`metadata.module`, `metadata.interop.transaction_type`, `metadata.execution.status`).
      * **Atlas Full-Text Search / Standard Indexes**: compound index on `{ "metadata.test_group_id": 1, "metadata.interop.transaction_type": 1 }` and text index on `search.keywords`.

-----

### 4\. How This Optimizes Retrieval Queries

  * **Exact Lookup**:
      * *Query*: *"Show me steps for Testcase - 1 CancelRx denial"*
      * *Execution*: Exact filter on `metadata.test_group_id: "Testcase - 1"` + `metadata.interop.transaction_type: "CancelRx"` $\\rightarrow$ instant direct MongoDB retrieval without vector latency.
  * **Semantic Search**:
      * *Query*: *"How do we handle cases where the pharmacy runs out of medication stock?"*
      * *Execution*: Vector match retrieves Row 39 (`INT-RxFill-Naproxen-2a` with reason code *"Out of Stock"*).
  * **Hybrid + Filtered Query**:
      * *Query*: *"Which EPCS test cases failed or had pending issues?"*
      * *Execution*: Pre-filter `metadata.interop.controlled_substance: true` + keyword/vector match on `metadata.execution.status: "Fail"` / `metadata.execution.notes`.
