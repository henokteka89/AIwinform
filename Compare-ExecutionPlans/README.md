# Compare-ExecutionPlans

A WinForms tool that loads two or more SQL Server execution plans, compares them side by side, flags the differences that matter, and builds an AI-ready Markdown prompt for deeper analysis.

## Start
Double-click **`Start-Compare-ExecutionPlans.cmd`**. The window opens with no console behind it, and no admin rights are needed.

Or from PowerShell:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -STA -File .\Compare-ExecutionPlans.ps1
```

## Loading plans
- **Files:** `.sqlplan` / `.xml`, using the button or drag and drop onto the list.
- **Clipboard:** plan XML copied from SSMS (right-click the plan, choose *Show Execution Plan XML*, then Ctrl+A, Ctrl+C).
- **SQL Server:**
  - Query Store, by `query_id`, `plan_id` or query text.
  - The plan cache, by query text or `query_hash`.

## What you get
| Tab | Contents |
|---|---|
| Comparison | Side-by-side metrics; rows that differ are highlighted |
| Findings | Differences between plans plus per-plan heuristics: spills, skew, key lookups, implicit conversions, memory grants… |
| Operators | Every operator from every plan in one sortable grid |
| Plan trees | Indented operator trees with estimates, actuals, predicates and warnings |
| AI prompt | Markdown package (instructions, facts, trees, optional raw XML) to paste into an AI assistant |
| JSON export | Structured data for other tooling |

## Requirements
- Windows PowerShell 5.1, or PowerShell 7+ on Windows.
- Capturing plans from a server needs **VIEW SERVER STATE** (plan cache) or **VIEW DATABASE STATE** (Query Store).

## Tips
- **Actual plans give much better analysis than estimated plans,** because estimated plans have no runtime numbers. So do Query Store plans with runtime stats.
- To investigate a regression, compare the good plan against the bad one for the same query.
