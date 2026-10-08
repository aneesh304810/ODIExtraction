# ODI 360: PBDW data-point extractor

`odi360_extract.py` reads the PBDW Bitbucket repos. It links every ODI scenario and
load plan to the Autosys jobs that schedule it, the tables it reads and writes, and the
DB procedures and functions it calls. It also links each object to its git and Jira
history. The output is:

1. **Data points** (CSV and XLSX) to ingest into Claude Enterprise
2. A **CP360 ingest sheet** called `ODI_360` (one row per load plan and per scenario version)
3. A **knowledge base** in Markdown plus a lineage graph, for a Claude Project

It needs Python 3.8 or later and only the standard library. `pip install openpyxl` is
optional and enables the `.xlsx` outputs.

## 0. Profile first (optional, recommended)

`odi360_profile.py` writes one redacted Markdown report (`odi360_profile.md`) that
describes the repos' structure: the ODI XML object classes and field names, the folder
layout, file-name patterns, DDL types and script keywords. Share that report so the
parser can be tuned to your exact ODI export format.

```powershell
python odi360_profile.py --root C:\path\to\PBDW --out odi360_profile.md
python odi360_profile.py --root C:\path\to\PBDW --out odi360_profile.md --no-samples   # structure only
```

## 1. Clone the repos side by side

```
D:\bitbucket\PBDW\
  oracle-data-integrator\     Master Repository, Work Repository, Scenarios and Load Plans, Releases
  prod-db-scripts\            Database Scripts\PBDWAPP\{Functions,Index,Procedure,Sequence,Tables,Views}
  Autosys\                    JIL
  pb-dwh-scripts\  pbinboundprocess\  pbschedulerdaemonprocess\  portal-queries\  billpay\ ...
```

Use full clones, not shallow ones, so the script can read the commit history and Jira keys
(for example "DW-1299 AddeparBenchmark Load").

## 2. Run

```powershell
python odi360_extract.py --root D:\bitbucket\PBDW --out D:\odi360_out --env PROD --project PBDW
```

With `--root`, the script sorts each folder under the root by its name and contents:

- A folder with `Master Repository`, `Work Repository` or `Scenarios and Load Plans` is read as the ODI repo.
- A folder whose name contains `db-scripts`, or that holds `Database Scripts`, is read as DB code.
- A folder whose name contains `autosys` is read as Autosys JIL.
- Every other folder is scanned as scripts. The script looks for `startscen`, `startloadplan`,
  scenario names and procedure names.

To choose the repos yourself, use `--odi-repo`, `--db-scripts`, `--autosys` and `--scripts`.
Each option can be repeated. `--no-git` skips the git history.

To try it on the bundled sample data:

```
python odi360_extract.py --root sample_input --out out_sample
```

## 3. Outputs

```
out/
  datapoints/                     raw facts, one CSV per family
    scenarios.csv                 per scenario version: source object, LP, Autosys, reads/writes, procs, KMs, git, Jira
    scenario_tasks.csv            per ODI task: step, KM, logical schemas, tables read/written, code preview
    loadplans.csv                 per load plan
    loadplan_steps.csv            load plan step tree (path, serial/parallel type, scenario, enabled)
    table_usage.csv               scenario x table x operation (READ/INSERT/MERGE/TRUNCATE/...)
    table_360.csv                 per table: writers, readers, DB code using it, indexes
    scenario_dependencies.csv     A writes T and B reads T, so A is upstream of B
    db_objects.csv                every CREATE in prod-db-scripts, with its reads, writes, calls and used_by
    db_foreign_keys.csv
    autosys_jobs.csv              job, box, command, machine, condition, schedule, ODI target
    script_references.csv         other repos' files that start or reference scenarios, LPs or procs
    odi_topology.csv              contexts, data servers (JDBC), physical/logical schemas, mappings, agents
    odi_work_objects.csv          projects, folders, mappings, packages, procedures, models, datastores
    odi_variables.csv
    releases.csv                  what each Releases/<folder> contains
    lineage_edges.csv             the whole 360 graph as edges
    findings.csv                  risks (see below)
    parse_errors.csv
  ODI360_datapoints.xlsx          all of the above, one sheet each
  cp360/ODI_360.csv|.xlsx         CP360 ingest sheet (+ LINEAGE / TABLES / FINDINGS sheets)
  claude_ingest/
    ODI360_KNOWLEDGE_partNN.md    upload to the Claude Project (split into parts of about 350 KB)
    odi360_graph.json             nodes and edges
    odi360_manifest.json          inputs and counts
```

### Findings produced

| Code | Severity | Meaning |
|---|---|---|
| HARDCODED_CREDENTIAL | HIGH | Password-like literal in scenario code |
| LP_MISSING_SCENARIO | HIGH | A load plan step runs a scenario that is not in the repo |
| ORPHAN_SCENARIO | MEDIUM | No load plan, Autosys job or script runs it |
| UNSCHEDULED_LOAD_PLAN | MEDIUM | No Autosys job or script starts it |
| HARDCODED_ENDPOINT | MEDIUM | JDBC URL or host name in code |
| MULTIPLE_VERSIONS | LOW | More than one exported version of a scenario |
| LP_STEP_DISABLED | LOW | Disabled load plan step |
| SELECT_STAR | LOW | `SELECT *` (risk of breaking when the schema changes) |
| NO_JIRA_TRACE | LOW | No Jira key in any commit message for the file |
| TABLE_NOT_IN_DB_SCRIPTS | INFO | The code uses a table that is not in prod-db-scripts (a source-system table, or the DDL is missing) |
| DB_OBJECT_UNREFERENCED | INFO | Table, view or proc that nothing in scope uses |
| NO_SQL_LINEAGE | INFO | Scenario with only OS, Jython or tool steps |

## 4. Claude Enterprise

See [`prompts/ODI360_CLAUDE_PROMPTS.md`](prompts/ODI360_CLAUDE_PROMPTS.md). It contains:

- the Project instructions
- **Prompt 1**, which builds the CP360 `ODI_360.xlsx`: it validates the draft sheet,
  enriches it (description, domain, layer, criticality, status) and adds the
  lineage, tables, findings and summary sheets
- prompts for a scenario deep dive, impact analysis, a load plan run book and tech-debt clean-up

## 5. Tuning

- `DOMAIN_RULES` and `LAYER_RULES` at the top of the script are regex-to-label lists. The
  first match wins. Add your own naming conventions.
- `ODI_WORK_TABLE` sets the prefixes the script skips from lineage
  (`C$_`, `I$_`, `E$_`, `J$`, `SNP_`, ...).
- **ODI XML parsing.** The parser flattens every `<Object class=...><Field name=...>` in the
  SunopsisExport and VCS XML, then reads fields by trying several candidate names, because
  field names vary between ODI versions. SQL lineage also runs on every long text field in a
  scenario file, so table-level lineage works even when the task and text layout differs.
- **Heuristic columns.** Business_Domain, Layer and Criticality are guesses. Prompt 1 refines them.
