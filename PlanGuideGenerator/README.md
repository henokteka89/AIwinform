# PlanGuideGenerator

A WinForms tool, written in Windows PowerShell 5.1, that writes a reviewed script to pin a query's plan. It can produce either:
- a **plan guide** (`sp_create_plan_guide`), or
- the **Query Store** equivalent (`sp_query_store_set_hints` / `sp_query_store_force_plan`).

You give it a statement plus either a **query hint** or a **showplan XML**. Or you start from the **saved good and/or bad execution plans** and let it work out the rest.

**It never creates anything.** The output is text: Copy, or Save .sql. The optional server check runs SELECTs only, through a read-only guard, on a background runspace. Errors in button clicks show a message box and are logged to `%TEMP%\PlanGuideGenerator.log`.

## Start
Double-click **`Start-PlanGuideGenerator.cmd`**. The window opens with no console behind it, and no admin rights are needed. You can set `SERVER` / `DATABASE` at the top of the launcher.

Or from PowerShell:
```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File .\PlanGuideGenerator.ps1
.\PlanGuideGenerator.ps1 -Server SQLAG01 -Database Sales
. .\PlanGuideGenerator.ps1 -NoGui        # load the engine only, for scripting
```

Scripting:
```powershell
. .\PlanGuideGenerator.ps1 -NoGui
$r = New-PlanGuideScript -Options @{ Statement = 'SELECT ...'; Params = '@p int'; Hint = 'RECOMPILE'; Database = 'Sales' }
$r.Findings; $r.Script | Set-Content .\guide.sql -Encoding UTF8

# from saved plans
$d = Resolve-PlanGuideFromPlans -PlanA (Get-Content -Raw .\good.sqlplan) -LabelA Good -PlanB (Get-Content -Raw .\bad.sqlplan)
$d.Findings; $d.Report; $d.Suggestions
if ($d.Options) { (New-PlanGuideScript -Options ($d.Options + @{ Database = 'Sales' })).Script }                         # plan guide
if ($d.Options) { (New-PlanGuideScript -Options ($d.Options + @{ Database = 'Sales'; Method = 'QueryStore' })).Script }  # Query Store
```

## Requirements
- Windows PowerShell 5.1 (.NET 4.x), or PowerShell 7+ on Windows.
- Plan guides: SQL Server 2008 and later.
- Query Store output:
  - Forcing a plan: SQL Server 2016+ with Query Store on.
  - Query Store hints: SQL Server 2022+ (or Azure SQL).
- Server check (optional), least privilege:
  - `VIEW SERVER STATE`, for the plan-cache check
  - `VIEW DEFINITION` on the module, for module checks
  - read access to the Query Store catalog views

  The person who runs the generated script needs `ALTER` on the database (SQL / TEMPLATE guides), `ALTER` on the module (OBJECT guides), or the rights `sp_query_store_*` requires.

## 0. Start from saved plans (optional)
Load **Plan 1** and, optionally, **Plan 2** (Open .sqlplan or Paste). Label each one *Let the tool judge*, *Good* or *Bad*. You can type a hint, then click **Analyze and fill the form**. If nothing is loaded here, Analyze uses the plan in section 3.

- The statement (plus @params or module) is **always** loaded from the plan, even when no decision can be made yet. The label next to the button says what is still needed.
- Picking a **suggested hint** applies it at once.

**The rule**
| Situation | Plan guide output | Query Store output |
|---|---|---|
| A hint is given (good or bad plan) | Query-hint guide; statement from the plan (the good one if known) | `sp_query_store_set_hints` |
| No hint, and a confirmed good plan | USE PLAN guide with the good plan's XML (that statement only) | `sp_query_store_force_plan` on the good plan's plan_id |
| No hint, only a bad plan | Nothing to force: it stops and lists suggested hints | same |

**How it decides good vs bad**
- **Two plans**
  - The same statement is matched in both by QueryHash, falling back to the text.
  - Evidence, in order of strength:
    1. actual elapsed time (1.5x or more apart)
    2. CPU (2x or more)
    3. actual logical reads (2x or more)
    4. estimated cost, which is weak evidence because a bad plan usually looked cheap to the optimizer.
  - If your labels contradict strong evidence, it stops, unless you tick *Use my labels even if the numbers disagree*.
  - With no labels and only estimates, or numbers that are too close, it asks you to label the plans.
  - Both plans labelled the same is an error. The same QueryPlanHash on both means forcing changes nothing, and USE PLAN is refused.
- **One plan**: it must be labelled (or come with a hint). A plan labelled Good still gets warnings for red flags:
  - tempdb spills
  - a row misestimate of 100x or more (where 1,000+ rows are involved)
  - an implicit conversion that prevents a seek
  - a join with no join predicate
  - memory grant warnings
  - 10,000 or more key/RID lookups
  - missing-index suggestions

**What it extracts**
- The statement text. A `(@p ...)` prefix is split off, and the PSP/OPPO `OPTION (PLAN PER VALUE ...)` is removed (see below).
- @params, rebuilt from the plan's ParameterList (name + ParameterDataType) and flagged so you confirm it matches the app's string exactly.
- QueryHash and QueryPlanHash, which are used to find the query_id and plan_id for Query Store output.
- **OBJECT guides**:
  - When the statement sits under `StoredProc` (an actual plan of `EXEC proc`), it becomes an OBJECT guide on `[schema].[proc]`.
  - When a PSP variant's ObjectID is not 0, it switches to OBJECT too.
- **Local variables**: a variable with a runtime value but no compiled value means a test batch with DECLARE. It warns that the app's parameterized text and plan will differ.
- **Auto-parameterized text**: suggests the Template type.
- **ARITHABORT differences**: the "fast in SSMS, slow in the app" case.

**Comparison report**: verdict, statement #, actual or estimated, elapsed, CPU, logical reads, cost, DOP, CE model, memory grant/used, spills, worst misestimate, plan hash, compiled values, ARITHABORT, operator shape and warnings.

**Suggested hints**
- `OPTIMIZE FOR (@p = <good plan's compiled value>)` when the compiled values differ (parameter sniffing)
- `MAXDOP 1` when the good plan is serial and the bad one parallel; `USE HINT('ENABLE_PARALLEL_PLAN_PREFERENCE')` the other way round
- `FORCE_LEGACY_` / `FORCE_DEFAULT_CARDINALITY_ESTIMATION` when the CE versions differ
- a single join type when the good plan uses only one
- `OPTIMIZE FOR UNKNOWN` and `RECOMPILE`

## Query Store output (section 4: Generate)
A saved plan does **not** contain the Query Store query_id. The only Query Store attributes in plan XML (QueryStoreStatementHintText / Id / Source) appear after a hint has been applied. The plan does carry QueryHash and QueryPlanHash, which the script uses to find the query_id and plan_id.

- **Hint** → `sp_query_store_set_hints`. The target is found by one of:
  - **PSP query**: the parent query, through `sys.query_store_query_variant`. Variants inherit the parent's hint. A commented line targets the single variant instead.
  - **Query hash from the plan**: every query_id with that hash. For OBJECT it also filters on `object_id = OBJECT_ID(module)`.
  - **No plan**: exact `query_sql_text` match (BIN2), with or without the `(@params)` prefix.
- **Good plan** → `sp_query_store_force_plan` on the newest plan_id with the good plan's `query_plan_hash` (and query_hash, when it isn't PSP). That plan must already be in this database's Query Store.
- **What the script contains**:
  - **Pre-flight checks**: database, server, AG primary, Query Store READ_WRITE, and 2022+ for hints.
  - **Safety latch**.
  - **Preview**: a `#pg_target` list of what will change, plus any existing Query Store hint, which `set_hints` replaces. It stops if nothing matches.
  - **Apply**: a cursor applying the change to each target.
  - **Verify**: the hint failure reason, or `is_forced_plan` / `last_force_failure_reason_desc`.
  - **Rollback**: `sp_query_store_clear_hints` / `sp_query_store_unforce_plan`.
- **Template** (forced parameterization) has no Query Store equivalent here, so it is an Error.
- Query Store hints override plan guides and hard-coded hints. They aren't supported for statements eligible for simple parameterization. Query Store forcing recompiles toward the plan and usually gets it, but that isn't guaranteed.

## Parameter Sensitive Plan optimization (SQL Server 2022+, compat 160)
- **The issue:** a plan saved from a PSP query is a **query variant**. Its StatementText ends with an engine-generated `option (PLAN PER VALUE(ObjectID = .., QueryVariantID = .., predicate_range(..)))`. That hint can't be used directly and isn't parsed if added manually. The app sends the parent query without it, so a guide on the variant text would never match.
- **The fix:**
  - The clause is stripped everywhere the tool reads plan text. If other options shared the same OPTION clause, only the PLAN PER VALUE item goes.
  - Typing it into the statement box is an **Error**.
- **ObjectID** is the object_id of the module the statement belongs to, and 0 for ad hoc or sp_executesql. A non-zero ObjectID switches the form to **Inside a module** and gives `SELECT OBJECT_SCHEMA_NAME(<id>) + '.' + OBJECT_NAME(<id>)` to find the name. A SQL guide with a non-zero ObjectID is an Error.
- **Warnings:**
  - A plan guide on a PSP query applies one plan or one hint to **every** parameter value, replacing PSP's per-range plans.
  - Different good/bad variants (QueryVariantID) mean the narrower fix is per variant, through Query Store output or the commented Query Store alternative in plan guide output.
- RECOMPILE and DISABLE_PARAMETER_SNIFFING turn PSP off.
- **OPPO (SQL Server 2025, `optional_predicate`)**: a variant can be a rewritten query, so its text is flagged as untrustworthy. Use the server check, then **Use cached text**.
- **Dispatcher plans** only route to variants. The tool asks for the actual plan of the variant that ran.

## Inputs (manual)
1. **Statement.** Paste an `exec sp_executesql N'...', N'@p int', ...` call or plan-cache text like `(@p int)SELECT...`, and **Split it** moves the declaration into @params. A stats line shows the character and line counts, the line-break type (CRLF/LF/mixed), tabs, trailing spaces and a trailing `;`. **Demo query** loads an example. It is replaced automatically when a plan is loaded.
2. **Where it runs**
   - **Ad hoc / sp_executesql**: SQL guide, with @params and an optional whole batch.
   - **Inside a module**: OBJECT guide, schema.name.
   - **Template**: `sp_get_query_template` + PARAMETERIZATION FORCED/SIMPLE, plus an optional second SQL guide carrying a hint for the parameterized form.
3. **What to force**
   - **Query hint**, with or without `OPTION ( )`. A quick-hint list is provided. Hints are checked for balance, known names and variables missing from @params.
   - **Plan XML**: paste it or open a .sqlplan file, then pick the statement. Loading a plan fills the query if the box is empty or holds the demo query. The tool can keep only the chosen statement (USE PLAN needs exactly one) and strip actual runtime stats. **Use this statement's text** refills the query, @params and module from the plan.
4. **Script options**: name (auto `PG_<scope>_<hash>`), database (USE + check), expected @@SERVERNAME, line breaks (As pasted / CRLF / LF), NCHAR() line-break encoding, replace existing guide, safety latch, text check, and **Generate: Plan guide / Query Store**.
5. **Server check (optional, read-only)**: Windows or SQL login, Encrypt / Trust cert.

## The generated plan guide script
- **Pre-flight** (any failure sets NOEXEC ON): DB_NAME(), the optional @@SERVERNAME, and not an AG secondary. Then the **safety latch** (RAISERROR + SET NOEXEC ON), which you delete by hand.
- The statement is a `CAST(N'' AS nvarchar(max)) + NCHAR(13) + NCHAR(10) + ...` literal, so copy/paste and editors can't change line breaks, and nothing truncates at 4000 characters.
- A **binary text check** (`COLLATE Latin1_General_BIN2` + DATALENGTH):
  - SQL guides: against `dm_exec_sql_text`, as `(@params)` + statement.
  - OBJECT guides: against `sys.sql_modules`. It also reports module type, encryption and a CRLF/LF mismatch.
- Optional drop of a same-name guide, then `sp_create_plan_guide`, then `sys.fn_validate_plan_guide`.
- USE PLAN mode adds a commented `sp_create_plan_guide_from_handle` alternative, found by query_plan_hash. PSP queries also get a commented Query Store alternative.
- A footer covers how to confirm the guide is used (PlanGuideName in the actual plan, XE `plan_guide_successful/unsuccessful`), plus DISABLE / DROP rollback.

## Checks (Checks tab)
**Errors** block the script. They include:
- no statement, or a PLAN PER VALUE clause in the statement
- a SQL guide on a statement whose PSP ObjectID says it is in a module
- bad hint parentheses, PARAMETERIZATION outside Template, or a hint variable that isn't in the statement or @params
- a statement not found in the batch
- invalid plan XML, a multi-statement plan without trimming, or dynamic/keyset cursors
- Template with Query Store output, or a plan without a hash for Query Store forcing
- the good/bad plan errors from section 0

**Warnings**:
- an unsplit sp_executesql call
- an existing OPTION clause (the guide's hints replace it)
- undeclared variables
- unknown hint names
- mixed line breaks
- plan text or line breaks that differ from the query
- remote operators
- DML outside SELECT/INSERT/UPDATE/DELETE/MERGE
- PSP query variants
- Query Store text matching without a plan

## Server check (read-only)
SELECTs only; `LOCK_TIMEOUT 3000`, `DEADLOCK_PRIORITY LOW`, READ UNCOMMITTED, 30 s timeout, MAXDOP 1, Application Name `PlanGuideGenerator (read-only)`. Plan-cache scans are skipped while the server is busy (average runnable tasks > 2 or any blocking).
- **Always**: version, edition, AG role, database parameterization (Template guides only take effect when it is the opposite setting), ALTER and VIEW SERVER STATE rights.
- **Plan guide mode**:
  - Whether a guide already has this name, and whether another guide (enabled or disabled) already targets the same text.
  - **SQL**: is the exact text in the cache, and in which database? If not, a loose LIKE search (first 8 tokens) shows the **first differing character** with `\r \n \t` made visible, and says whether only the line breaks or only the whitespace differ. **Use cached text** replaces the query with what really runs.
  - **OBJECT**: whether the module exists and its type is supported, whether the definition is visible, and the exact position (line number) or the first difference.
- **Query Store mode**:
  - Query Store state, and matching query_id(s) by hash or text, or the PSP parent.
  - Whether the good plan is in Query Store, and whether it is already forced.
  - Existing hints that would be replaced.
  - Missing hint/variant views, which mean the server is older than 2022.

## Things to know
- **Exact matching:** plan guides match character for character, comments and trailing whitespace included, and @params must match the app's declaration string exactly.
  - That is why the server check matters after filling from a plan: plan files can lose or alter whitespace, and test batches in SSMS differ from the app's text.
  - Query Store output found by query_hash avoids this problem.
- **Where they live:** guides and Query Store data live in the database, so they follow backups and AG replicas. Create them on the primary. A guide applies only to queries executed in that database.
- **PSP and plan guides:** Microsoft's PSP documentation covers Query Store hints and forcing per variant. It doesn't describe how plan guides interact with dispatcher plans or variants, so test a PSP plan guide in a lab before production.
- **Test before production use.** Run the generated script in a lab first, with the safety latch on, and review the checks.
