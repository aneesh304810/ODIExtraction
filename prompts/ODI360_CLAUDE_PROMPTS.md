# ODI 360: Claude Enterprise prompts

## Step 0: Load the knowledge

Create a Claude Enterprise **Project** named `ODI 360 - PBDW`. Upload these files
from the extractor output folder:

| Upload | Why |
|---|---|
| `claude_ingest/ODI360_KNOWLEDGE_part*.md` | One section per load plan, scenario, table and DB object (the narrative facts) |
| `ODI360_datapoints.xlsx` (or every `datapoints/*.csv`) | The raw facts, one sheet per family |
| `cp360/ODI_360.csv` | Draft CP360 sheet built by the script. Claude enriches it and does not rebuild it |
| `claude_ingest/odi360_graph.json` (optional) | Lineage graph with nodes and edges |

Paste the block below into the Project **instructions**. Use prompt 1 to build the
spreadsheet. Use prompts 2 to 5 for follow-up analysis.

---

## Project instructions (paste once)

```
You are the ODI 360 analyst for the PB Data Warehouse (PBDW). Your knowledge files were
machine-extracted from the Bitbucket repos oracle-data-integrator (Master Repository,
Work Repository, Scenarios and Load Plans, Releases), prod-db-scripts (Functions, Index,
Procedure, Sequence, Tables, Views), Autosys and the other PBDW script repos.

Rules:
1. Facts come ONLY from the knowledge files. Never invent a scenario, table, job, schedule,
   Jira key or owner. If something is unknown, write "TBD" or leave the cell empty. Do not guess.
2. Columns marked (heuristic) in the knowledge files are suggestions. You may correct them,
   but say why in Status_or_Notes.
3. Keep object names exactly as extracted (UPPER case, including version numbers).
4. When you answer a question, cite the source section, for example
   "SCENARIO SCN_X v001", "TABLE FACT_HOLDINGS" or "datapoints/table_usage".
5. A load plan runs a scenario version of "-1" or blank as the latest version.
```

---

## Prompt 1: Generate the CP360 "ODI 360" spreadsheet

```
Build the CP360 ingest workbook "ODI_360.xlsx" from the project knowledge.

START FROM cp360/ODI_360.csv. It has one row per LOAD_PLAN and per SCENARIO version,
already populated with extracted facts. Your job is to VALIDATE and ENRICH it. Do not
re-derive the facts.

SHEET 1 "ODI_360". Keep these columns in exactly this order and with these exact headers:
Env, Layer, Object_Type, Object_Name, Version, Business_Domain, Description,
Source_Object_Type, Source_Object_Name, Load_Plans, Load_Plan_Step_Path, Autosys_Jobs,
Autosys_Box, Schedule, Upstream_Scenarios, Downstream_Scenarios, Scenarios_In_Load_Plan,
Source_Tables, Target_Tables, DB_Links, DB_Procedures_Called, DB_Functions_Called,
Logical_Schemas, Technologies, Knowledge_Modules, Load_Pattern, Variables, Step_Count,
Task_Count, Jira_Refs, Last_Commit_Date, Last_Commit_Author, Last_Change_Summary,
ODI_Last_Modified, ODI_Last_User, Releases, Repo_Path, Risk_Flags, Criticality,
Status_or_Notes, Project, Zone, Component, Display_Order

Enrichment rules:
- Description: if it is empty, or is generic (for example "Loads <name>"), write ONE plain-English
  sentence that says what the object does. Base it on its source/target tables, load pattern
  and load plan position. Example: "Truncate-reloads STG2_ADDEPAR_BENCHMARK from STG1 Addepar
  benchmark data, mapped through BD_MAPPING."
- Business_Domain: confirm or correct it. Use one of: Addepar, Benchmark, Performance,
  Holdings, Positions, Transactions, Accounts, Settlement, Billpay, Pricing, Securities,
  Client, Assets, API Ingestion, Housekeeping, File Handling, Reference Data, Intraday,
  Adhoc, Test / Utility, Unclassified.
- Layer: confirm or correct it. Use one of: Orchestration, Staging, Reference / Mapping,
  Warehouse Dimension, Warehouse Fact, Reporting / Mart, API Ingestion, Operational,
  Unclassified.
- Criticality: HIGH if the object is scheduled (it has an Autosys job, or it sits in a
  scheduled load plan) AND (it has downstream scenarios OR it writes a FACT_/DIM_ table OR
  it sits in an INTRADAY load plan). MEDIUM if it is scheduled or sits in any load plan.
  Otherwise LOW. Put a short rationale in Status_or_Notes.
- Status_or_Notes: replace "Auto-extracted..." with ACTIVE, ORPHAN (Risk_Flags contains
  ORPHAN_SCENARIO), SUPERSEDED (a higher Version of the same Object_Name exists and this
  version is not referenced by a load plan or job), or TEST (the name contains TEST).
  Add the rationale after " - ".
- Never change: Env, Object_Type, Object_Name, Version, Autosys_Jobs, Source_Tables,
  Target_Tables, Jira_Refs, the date columns, Repo_Path, Project, Zone, Component or
  Display_Order.
- Multi-value cells use "; " as the separator. Never use line breaks inside a cell.

SHEET 2 "ODI_360_LINEAGE": copy datapoints/lineage_edges. Columns:
src_type, src, relation, dst_type, dst, detail.

SHEET 3 "ODI_360_TABLES": copy datapoints/table_360 and add a column "Business_Meaning".
Write one short phrase per table, or "TBD".

SHEET 4 "ODI_360_FINDINGS": copy datapoints/findings and add these columns:
- "Impact": one sentence.
- "Recommended_Action": one sentence.
- "Priority": P1 for HIGH, P2 for MEDIUM, P3 for LOW and INFO. Raise the priority one level
  when the object's Criticality is HIGH.

SHEET 5 "SUMMARY" (key/value rows):
- counts per Object_Type, Layer, Business_Domain and Criticality
- the count of load plans and of Autosys jobs
- the top 10 tables with the most writers plus readers
- the top 10 findings
- every scenario that is referenced by no load plan and no job

VALIDATE before you output. List the result of each check:
[ ] Row count of ODI_360 equals the row count of cp360/ODI_360.csv
[ ] No duplicate (Object_Type, Object_Name, Version)
[ ] Every Load_Plans value exists as a LOAD_PLAN row
[ ] Every Upstream/Downstream scenario exists as a SCENARIO row
[ ] Headers match the list above exactly (sheet 1)
[ ] Display_Order is unique and ascending

OUTPUT: create the .xlsx file with the analysis/code tool. If file creation is not
available, output each sheet as a separate ```csv block, starting with ODI_360. After the
file, give a 10-line executive summary of the ODI estate.
```

> For very large estates (more than about 400 scenarios), run prompt 1 per Business_Domain
> or per load plan. Add `Only process rows where Load_Plans contains <LP_NAME>` to the prompt,
> then concatenate the resulting sheets. Keep the Display_Order values from the CSV.

---

## Prompt 2: Scenario deep dive

```
Give me the 360 view of scenario <SCENARIO_NAME>. Cover:
- what it does (business sentence)
- how it is triggered (Autosys box/job, schedule, load plan step path)
- its source and target tables and its load pattern
- the DB procedures and functions it calls, and what those DB objects read and write
- upstream and downstream scenarios
- its versions and releases
- its Jira history and last change
- its risk flags, with the fix for each
Finish with a mermaid flowchart of Autosys -> LP -> scenario -> tables -> downstream.
```

## Prompt 3: Impact analysis

```
I plan to change table <SCHEMA.TABLE> (for example, add or drop a column, or reload it).
List every impacted object:
- scenarios that write it or read it
- DB views, procedures and functions that use it
- indexes and foreign keys on it
- the load plans and Autosys jobs that run the impacted scenarios
- downstream scenarios, 2 hops out
Rank them by Criticality. Output a table: Object_Type, Object_Name, Relation, Hop,
Criticality, Autosys_Job, Action_Needed.
```

## Prompt 4: Load plan run book

```
For load plan <LP_NAME>, produce a run book:
- the step tree in execution order (serial/parallel), with each scenario's targets
- restart guidance: which steps can be re-run safely (TRUNCATE+RELOAD or MERGE) and which
  are not idempotent (APPEND, DELETE+INSERT without a date filter)
- disabled steps
- steps that reference missing scenarios
- the Autosys trigger and schedule
- the tables to check after a failure
```

## Prompt 5: Clean-up and tech debt

```
Using ODI_360_FINDINGS and ODI_360, list:
(a) orphan scenarios that are safe to retire
(b) superseded scenario versions
(c) DB objects that nothing references
(d) hard-coded credentials and endpoints (P1)
(e) scenarios with no Jira trace
Group the results into a 3-sprint remediation plan with effort (S/M/L).
```
