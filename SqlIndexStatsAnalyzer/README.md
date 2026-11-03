# SqlIndexStatsAnalyzer

A **read-only** WinForms tool, written in Windows PowerShell 5.1, for analyzing SQL Server indexes and statistics. It is built to be safe on production databases from a few MB up to 50 TB+.

A single self-contained script: the GUI and the read-only engine are in `SqlIndexStatsAnalyzer.ps1`. All SQL runs on a background runspace, so the window never freezes.

## Start
Double-click **`Start-SqlIndexStatsAnalyzer.cmd`**. The window opens with no console behind it, and no admin rights are needed. To connect straight away, set `SERVER` (and optionally `DATABASE`) at the top of the launcher.

Or from PowerShell:
```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File .\SqlIndexStatsAnalyzer.ps1
.\SqlIndexStatsAnalyzer.ps1 -Server SQLAG01 -Database Sales     # connect straight away
. .\SqlIndexStatsAnalyzer.ps1 -NoGui                            # load only the read-only functions (scripting)
```

## Requirements
- Windows PowerShell 5.1 (.NET 4.x), or PowerShell 7+ on Windows (without the histogram chart, which needs `System.Windows.Forms.DataVisualization`; the histogram is then shown as a grid).
- SQL Server 2012 SP1 and later. Features are switched on by version and edition (see Capabilities).
- Permissions: see "Things to know" below.

## Guardrails: why it is safe to point at production
| Layer | What it does |
|---|---|
| **Read-only guard** | Every statement goes through one function (`Invoke-SafeQuery`). Before it connects, a tokenizer strips comments, strings and quoted names, then refuses anything that is not SELECT, DECLARE, harmless SET options or `DBCC SHOW_STATISTICS`. Refused: INSERT/UPDATE/DELETE/MERGE, ALTER/CREATE/DROP/TRUNCATE, EXEC, `sp_`/`xp_`, SELECT INTO, OPENROWSET/OPENQUERY, BACKUP/RESTORE, KILL, WAITFOR, transactions, UPDATE STATISTICS, REBUILD/REORGANIZE, any other DBCC, and more. The exact text sent (with the session prefix) is checked a second time. The central server-list query goes through the same guard. |
| **No execution path** | Maintenance is generated as **text** in the Script tab. There is no Execute button, and the read-only guard itself rejects every generated script. |
| **Polite sessions** | `LOCK_TIMEOUT 3000`, `DEADLOCK_PRIORITY LOW`, `READ UNCOMMITTED`, 30 s command timeout, `OPTION (MAXDOP 1)` on heavier metadata reads, connection pool capped at 2, one query at a time, Cancel / Esc kills the running command. Application Name is `SqlIndexStatsAnalyzer (read-only)`, so you can see it in `sys.dm_exec_sessions`. |
| **Metadata only** | Row counts and sizes come from `sys.dm_db_partition_stats`. No `COUNT(*)`, and no table data is read. Histograms read only the statistics object. |
| **Fragmentation is opt-in** | `sys.dm_db_index_physical_stats` runs only when you click the button. It scans one index at a time, in **LIMITED** or **SAMPLED** mode only (never DETAILED), at MAXDOP 1. It is **refused** above a size cap (20 GB on production, 200 GB elsewhere, 2 GB for SAMPLED), and **refused while the server is busy** (average runnable tasks > 2 or any blocking). On production it is also refused when the load cannot be checked. The cap matters: even LIMITED reads the level above the leaf, which is about 0.3% of a 30 TB index (around 100 GB). |
| **Production by default** | A server is Production unless the central list or a name pattern says otherwise (`dev\|tst\|test\|qa\|uat\|lab\|stg...`). A red banner shows the environment. |
| **Script safety** | Every generated script checks `@@SERVERNAME` and `DB_NAME()`, and checks that it is on the AG primary. It uses `SET NOEXEC ON` if any check fails. Production scripts start with a **safety latch** that you must delete by hand. Options the server can't do are downgraded with a NOTE (ONLINE on Standard, RESUMABLE before 2017, RESUMABLE + SORT_IN_TEMPDB, MAXDOP on old builds, REORGANIZE on a heap or with page locks off). REBUILD on a disabled index warns that it will enable the index. NORECOMPUTE is kept, because dropping it silently re-enables auto-update. Each script includes estimates of the space, log and tempdb it will need. |
| **Transparency** | The Query log tab lists every statement sent, with its duration, row count and status. |

## Using it
1. **Server**: type a name, pick one from the cache, or use **Server list...** to load from a central table. The default query is `SELECT SqlInstance AS ServerName FROM DBAPatch.dbo.PatchInstance`. Any SELECT that returns `ServerName`/`SqlInstance`/`Name` and, optionally, `Environment` works. You can tag environments in the grid. Windows or SQL auth (the password stays in memory and is never saved), plus optional Encrypt and ApplicationIntent=ReadOnly.
2. **Database**: the list shows size, AG role and accessibility.
3. **Find a table**: the list loads once from metadata (tens of thousands of tables is fine), then searching is local and instant. Matching runs exact → prefix → contains → tokens (`order hdr`) → typos (`custmer`, `ordrheader`) → abbreviations (`ordhdr`), and also supports `*`/`%` wildcards and `schema.table`. A search over 50,000 tables takes about 0.2 s. With an empty search box the largest tables are listed first.
4. **Overview & priorities**: a table, database and server summary, plus one ranked to-do list across indexes and statistics. Double-click a row to jump to it.
5. **Indexes**: size, rows, seeks/scans/lookups/writes, last read and write, compression, fill factor, page allocations (a proxy for page splits), lock and latch waits, forwarded fetches, and duplicate/overlap detection. *Check fragmentation* fills in Frag% (LIMITED). *Check page density* adds fullness (SAMPLED, small indexes only). Columnstore indexes get a row group check instead (deleted %, delta stores, small row groups). The bottom pane explains the selected index in plain English.
6. **Statistics**: last updated, rows, rows sampled, **sample %**, modifications, **% of the auto-update threshold**, steps, NORECOMPUTE, incremental, filtered, and persisted sample. Click a statistic to see the **histogram chart** (dark bars = rows equal to the bucket value, light bars = rows in between), the density vector, and an explanation for non-DBAs: how fresh it is, how big the "survey" was, how unique the values are, whether the data is skewed (a parameter-sniffing risk), and ascending-key warnings.
7. **Missing indexes**: the optimizer's requests, with a commented-out CREATE INDEX and a reminder to check overlaps first.
8. **Generate script...** (Indexes or Statistics tab): choose REBUILD or REORGANIZE, ONLINE, WAIT_AT_LOW_PRIORITY (ABORT_AFTER_WAIT = SELF), RESUMABLE + MAX_DURATION, SORT_IN_TEMPDB, MAXDOP, a single PARTITION, LOB_COMPACTION, and COMPRESS_ALL_ROW_GROUPS. For statistics: FULLSCAN / SAMPLE n PERCENT / RESAMPLE [ON PARTITIONS], MAXDOP, and PERSIST_SAMPLE_PERCENT. The preview updates live. Production defaults are ONLINE + low-priority wait + resumable where available. You can optionally add the equivalent **Ola Hallengren IndexOptimize** call (the tool detects whether IndexOptimize is installed). Copy or save the result as `.sql`.

## Procedure / query analysis (opt-in tab)
Nothing runs until you click. Everything reads plans that **already exist** (plan cache or Query Store). Nothing is executed, compiled or recompiled.

**Modules** (left tab)
- **Load modules** lists procedures, scalar functions, multi-statement TVFs and triggers (metadata only).
- **Only procedures that use the selected table** filters the list using `sys.sql_expression_dependencies`.
- **Analyze** reads the module's cached plan from `dm_exec_procedure_stats`, `dm_exec_function_stats` (2016+) or `dm_exec_trigger_stats`, via `dm_exec_text_query_plan`. Per-statement runtime comes from `dm_exec_query_stats`. If the plan isn't cached, it falls back to Query Store.

**Find query** (left tab): search by **SQL text**, **query hash**, **query plan hash**, **Query Store query_id** or **plan handle**.
- The plan cache is searched across the whole server: ad hoc, prepared and dynamic SQL included.
- Query Store is searched in the current database.
- Each source returns the top 50 matches, at MAXDOP 1, with the values passed as parameters.
- A plan-cache text search needs at least 6 characters and is skipped while the server is busy.
- For a Query Store query, every plan it has had appears in the **Plan** picker. A large duration spread between plans is flagged as a plan regression.

**What the analysis covers**
- **Every statement kind**, including the IF/WHILE condition's own query (`StmtCond`) and cursors. Statements in branches that haven't run since the plan was cached (IF/ELSE not taken, early RETURN) are listed as *not run since cached*, not as missing history. Runtime numbers are matched by query hash, falling back to the statement text.
- **Statistics the optimizer used** (recorded in plans from SQL 2017 and 2016 SP2). Each is scored on its current freshness and sample, and rises in priority when that statement's row estimate is 10x or more off. On older versions, the stats of the tables the plan reads are assessed instead. Columns with no statistics are High.
- **Index maintenance only where it helps:** only indexes the plan **scans**, and only after a LIMITED fragmentation check (button) shows fragmentation or low page density. Seek-only indexes are "no action".
- **Dynamic SQL** (`sp_executesql`, `EXEC(@sql)`) is found from the plan and from the module definition. These statements are cached on their own, so the finding suggests a text fragment. Double-click it to search for them.
- **Nested procedures, scalar UDFs and multi-statement TVFs** are listed. Double-click one to analyze it next.
- **Problems maintenance won't fix:**
  - scans that read 100x or more rows than they return
  - heavy key lookups and missing-index hints
  - implicit conversions that prevent seeks
  - tempdb spills and missing join predicates
  - parameter sensitivity, with the compiled parameter values shown
- **Scripts:** statistics updates and REBUILD/REORGANIZE go to the Script tab as text, with the safety header. **Save as .sqlplan** opens the plan in SSMS or Plan Explorer.

Plan XML is parsed with DTD processing and external resolution disabled. Every query passes the read-only guard.

## How priorities are scored (0-100, heuristics, not orders)
**Statistics**
- +50 when changes have passed the auto-update threshold (`MIN(500 + 20%, SQRT(1000 × rows))` at compat 130+, `500 + 20%` below that); +30 at 50% of it, +15 at 20%.
- +20 or +10 when 20% or 10% of the rows have changed.
- +15 for a sample under 1% on tables over 10M rows; +8 for a sample under 5% on tables over 1M rows.
- Age with changes: +10 after 30 days, +5 after 7.
- +10 for NORECOMPUTE with changes; +10 when the database has auto-update OFF.
- A statistic that was never built on a non-empty table scores 90.

Suggested method: FULLSCAN up to 10M rows. Above that, a sample of about 10M rows (at least 1%).

**Indexes** (indexes under 1,000 pages are ignored)
- Fragmentation ≥ 30% → REBUILD (+45). 5-30% → REORGANIZE (+20).
- Page density below 70% (and below fill factor − 10) → +30. This follows [Microsoft's current guidance](https://learn.microsoft.com/sql/relational-databases/indexes/reorganize-and-rebuild-indexes) that page density often matters more than fragmentation.
- Size adds +5 or +10. Scan-heavy use adds +10. Seek-only use subtracts 10, since fragmentation barely affects seeks.
- When fragmentation hasn't been measured, leaf allocations above 10% of the index size flag likely page splits.
- Heaps with forwarded fetches → REBUILD (rebuilding a heap also rebuilds all its nonclustered indexes).
- Columnstore: ≥ 20% deleted rows, open delta stores, or row groups under 100K rows.
- **Review**, in the spirit of sp_BlitzIndex: unused nonclustered indexes (0 reads, with writes), exact duplicates, same keys with different INCLUDEs, left-prefix overlaps, disabled and hypothetical indexes. The decision there is whether to drop, not how to maintain. The tool warns when uptime is under 14 days.

Every threshold is in `%APPDATA%\SqlIndexStatsAnalyzer\settings.json` (Server list... > Open settings file): `ReorgThreshold`, `RebuildThreshold`, `PageDensityThreshold`, `MinPageCount`, `MaxScanGBProduction`, `MaxScanGBNonProduction`, `MaxSampledScanGB`, `MaxAvgRunnableTasks`, `CommandTimeoutSec`, `ScanTimeoutSec`, `LockTimeoutMs`, `DefaultMaxDop`, `ProdNamePattern`, `NonProdNamePattern`, `UnknownIsProduction`, and `MaxTablesDisplayed`.

## Capabilities detected per server
| Feature | Needs |
|---|---|
| ONLINE rebuild | Enterprise / Developer / Azure (EngineEdition 3, 5, 8) |
| ONLINE single partition, WAIT_AT_LOW_PRIORITY | 2014+ |
| RESUMABLE rebuild | 2017+ Enterprise |
| ONLINE columnstore rebuild | 2019+ |
| `sys.dm_db_stats_histogram` | 2016 SP1 CU2+ (older versions fall back to `DBCC SHOW_STATISTICS WITH HISTOGRAM`) |
| UPDATE STATISTICS ... MAXDOP | 2014 SP3, 2016 SP2, 2017 CU3+ |
| PERSIST_SAMPLE_PERCENT | 2016 SP1 CU4, 2017 CU1+ |
| COMPRESS_ALL_ROW_GROUPS, row group DMV | 2016+ |

## Things to know
- Usage counters (seeks/scans/lookups, missing indexes, operational stats) reset at restart and are **per replica** on AGs. ReadOnly intent through a listener lands you on a secondary.
- Least-privilege login: `VIEW SERVER STATE` (2022+: `VIEW SERVER PERFORMANCE STATE`), `VIEW ANY DEFINITION`, `VIEW ANY DATABASE`, plus `VIEW DATABASE STATE` and `db_datareader` in each database. Without SELECT on a table, SQL Server hides its statistics details. The tool warns when you connect as sysadmin.
- In-memory OLTP tables are shown, but rowstore maintenance doesn't apply to them.
