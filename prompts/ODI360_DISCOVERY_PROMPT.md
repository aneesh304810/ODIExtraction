# ODI 360: discovery prompt

Run this in Claude with the repos available. That means either Claude Code opened on
the folder that holds the clones, or a Claude Enterprise chat with `odi360_profile.md`
and sample files attached. It produces `ODI360_DISCOVERY.md`, the input needed to build
the v2 ODI Intelligence Model parser.

```
You are doing a read-only discovery of the PBDW ODI codebase to design a lineage parser.
Repos (sibling folders): oracle-data-integrator, prod-db-scripts, pb-dwh-scripts,
pbinboundprocess, pbschedulerdaemonprocess, portal-queries (and Autosys, if it is present).
If odi360_profile.py exists, run it first:
    python odi360_profile.py --root <this folder> --out odi360_profile.md
Use its output as a starting point, then inspect the actual files to answer every
question below.

RULES
- Read-only. Do not modify, move or delete anything.
- Never output passwords, credentials, hostnames, IPs, JDBC URLs, e-mails, wallet or
  certificate content. Replace them with <REDACTED>, <HOST> and so on.
- Excerpts: at most 15 lines each, redacted. Show structure, not whole files.
- Every answer must cite the relative file path(s) it came from. If something cannot be
  determined, write "UNKNOWN - <what would be needed>". Never guess.
- Ignore CHG* folders, rollback/backout scripts, certificates and binaries, except where
  a question asks about them.

QUESTIONS (answer each under its own heading, in this order)

Q1 ODI version and export format
- The RepositoryVersion attribute of <Admin> in 3 XML files (Scenarios, Work Repository,
  Master Repository).
- Does the Work Repository contain SnpMapping/SnpMapComp (12c) or SnpPop/SnpPopCol (11g
  interfaces)? Give counts.
- Is it the ODI VCS/Git layout (one object per file) or Smart Export files?

Q2 Class and field inventory
For each class below, list: the number of files it appears in, all field names, and
1 redacted example value per field (at most 60 chars). Classes: SnpScen, SnpScenStep,
SnpScenTask, SnpVarScen, SnpLoadPlan, SnpLpStep, SnpLpVar, SnpPackage, SnpStep,
SnpMapping, SnpMapComp, SnpMapCp, SnpMapConn, SnpMapAttr, SnpMapExpr, SnpMapExprRef,
SnpMapRef, SnpTrt, SnpLineTrt, SnpTxtHeader, SnpVar, SnpModel, SnpTable, SnpCol,
SnpLschema, SnpPschemaCont, SnpContext.
Also list any other Snp* classes with a count of more than 20.

Q3 Where the executable text lives
- In one scenario file, where is the SQL of a task stored? Give the class, field name,
  and how the text links to its task (key fields). Give a 10-line redacted excerpt.
- The same for an ODI procedure (SnpLineTrt / SnpTxtHeader) and a package step.
- How does a SnpMapExpr reference an attribute (by name, by internal ID via SnpMapRef,
  by GlobalId)? Show 1 example chain: expression -> attribute -> component -> datastore.

Q4 Identity and cross-references
- Which ID fields exist (IScen, ScenNo, GlobalId, IMapping, ITable, I*...)? Are GlobalIds
  present?
- How does a scenario link to its source object (field names such as
  OrigObjType/OrigObjName, or others)?
- How does a load plan step reference its scenario (name+version, or an ID)?

Q5 Load plan structure
For 1 real load plan: the step tree (at most 30 lines) showing the step type codes used
(serial, parallel, run-scenario, case/when/else, exception) and what each code means,
based on the evidence. Note restart and exception fields.

Q6 Package step types
List the distinct SnpStep type codes with counts and 1 example each (scenario call,
mapping, procedure, variable refresh/set/evaluate, OdiStartScen, OS command, e-mail).
Is OdiStartScen ever called with -SYNC_MODE=2 (asynchronous)? Give counts.

Q7 Active vs historical
- List the top-level folders of each repo and classify each one as ACTIVE / RELEASE /
  ROLLBACK / TEST / BACKUP / OTHER, with the evidence for the label.
- How many scenarios exist in several versions? Give 3 examples with their versions.
- Are there copied or backup objects (_BKP, _OLD, _COPY, _TEST, dated suffixes)? Count them.

Q8 Scheduler entry points
- Is Autosys JIL anywhere in these repos? Give paths and counts of insert_job,
  job_type: BOX / CMD / FW (file watcher).
- List every distinct way ODI is started: startscen.sh, startloadplan.sh, the agent
  web service, OdiStartScen, a wrapper script. For each, give the count and 1 redacted
  example with its path.
- Which wrapper scripts take the scenario name as a parameter (so the scenario name is
  not visible in the JIL)?

Q9 Database code
- The folder layout under prod-db-scripts/Database Scripts: schemas, and counts per
  object type.
- Are there packages (spec/body), triggers, synonyms, DB links, external tables?
  Give counts.
- How much dynamic SQL is there (EXECUTE IMMEDIATE, DBMS_SQL)? Give the count and
  2 redacted examples.
- Are table DDL files column-complete (CREATE TABLE with columns), or partial ALTERs?

Q10 Files in and reports out
- Inbound files: file-technology models/datastores in ODI, file paths in file watchers
  or scripts, sqlldr .ctl files, external tables. Give counts and redacted examples.
- Consumers: what portal-queries contains (file types, counts, 2 example FROM clauses
  with table names only). Is there anything else that reads the warehouse tables?

Q11 Contexts and schemas
List the contexts, and for each logical schema its physical schema per context
(schema names only, no hosts). Which context is production?

Q12 Candidate gold flows
Propose 3 real end-to-end flows for hand validation. Each flow must go from a scheduler
entry, through a scenario (and its package/mapping/procedure), to the target table.
Pick 1 simple mapping, 1 package with branches, and 1 PL/SQL-heavy flow. Give the
names and paths at each hop, and mark any hop you could not resolve.

Q13 Surprises
Anything that breaks common ODI assumptions: generated code with odiRef still
unresolved, encrypted or unreadable text fields, non-XML exports, very large files,
encodings, or naming conventions that the parser must know.

OUTPUT
Write everything to ODI360_DISCOVERY.md. Start with a 10-line summary table:
question -> status (ANSWERED / PARTIAL / UNKNOWN). Keep the total under about 60 KB.
```

## Optional: runtime evidence (to decide which objects are "active")

Run these against the ODI **work repository schema** as a read-only user. Share only the
aggregated results.

```sql
-- scenarios executed in the last 90 days
SELECT scen_name, scen_version, COUNT(*) runs, MAX(sess_beg) last_run,
       SUM(CASE WHEN sess_status = 'E' THEN 1 ELSE 0 END) errors
FROM   snp_session
WHERE  sess_beg > SYSDATE - 90 AND scen_name IS NOT NULL
GROUP  BY scen_name, scen_version ORDER BY runs DESC;

-- load plans executed in the last 90 days
SELECT i.load_plan_name, COUNT(*) runs, MAX(r.start_date) last_run
FROM   snp_lp_inst i JOIN snp_lpi_run r ON r.i_lp_inst = i.i_lp_inst
WHERE  r.start_date > SYSDATE - 90
GROUP  BY i.load_plan_name ORDER BY runs DESC;
```

Autosys, if the JIL is not in Bitbucket: `autorep -J PBDW% -q > pbdw_jobs.jil`. Replace the
`PBDW%` job-name filter with your own naming pattern, and redact machine names before
sharing.
