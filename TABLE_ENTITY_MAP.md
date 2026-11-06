# Table → Entity-Type Map — maintenance guide

## What this is

`src/taskpane/table-entity-map.json` classifies every real table in the Table Builder catalog
(`DOWNLOADTABLE_CATALOG_URL`, `src/taskpane/taskpane.js`) by what kind of thing a row represents —
Company, Sector, Mutual Fund Scheme, PMS Provider, etc. — plus which column holds that entity's id
and which holds its display name. Given the table a selected metric resolves to (either
`metric-registry.json`'s `annual.prefer`/`quarterly.prefer`, or the catalog's own table for a
metric outside that registry), this answers "what should the Entity dropdown let the user pick,
and from which column."

**Two copies exist, on purpose:**

- `src/taskpane/table-entity-map.json` — the **source of truth**, committed to this repo, reviewed
  and diffed like any other file. Never imported/bundled by webpack.
- `https://ir-feed.s3.ap-south-1.amazonaws.com/excel-addin-metadata/table-entity-map.json` — a
  **public, unauthenticated-GET copy** of the exact same file, fetched by the add-in at runtime
  (`getTableEntityMap()` / `getEntityTypeForTable()` in `taskpane.js`, cached 24h in
  `localStorage` under `qd_table_entity_map_cache_v1`, same pattern as the catalog itself). The app
  never reads the repo file directly — **editing the repo file alone does nothing until you
  re-upload it** (see Step 4).

`ir-feed` is a different bucket from `ir-reports` (where `index.json` itself lives) — that's just
where upload access was available when this was set up, not a meaningful distinction. The bucket
has public ACLs enabled; only this one object (`excel-addin-metadata/table-entity-map.json`) is
`public-read` — nothing else in that bucket is, and nothing else should be made public as part of
this workflow.

## File shape

```jsonc
{
  "entity_types": ["Company", "Sector", "Mutual Fund AMC", "Mutual Fund Scheme", "PMS Provider", "PMS Investment Approach", "Economic Indicator"],
  "tables": [
    {
      "table": "GIAStock_Data.dbo.basic_info",   // exact name as it appears in index.json's keys
      "domain": "COMPANY MASTER",                 // index.json's own domain grouping, for reference
      "is_entity_table": true,                    // false = reference/metadata table, no entity to pick (see below)
      "entity_type": "Company",                   // one of entity_types, or null when is_entity_table is false
      "entity_id_column": "fincode",               // real column name holding the entity's id — null if there isn't one on this table
      "entity_name_column": "CompName",            // real column name holding the entity's display name — null if it must be resolved via a join
      "related_entity_type": "...",                // OPTIONAL — only for a junction/pivoted table that mixes two entity types in one row
      "related_entity_id_column": "...",           // OPTIONAL — the related entity's id column, when related_entity_type is set
      "notes": "..."                               // anything a future reader needs that isn't obvious from the columns alone
    }
  ]
}
```

`is_entity_table: false` is for pure reference/lookup tables — a metric catalog (e.g.
`global_filter`, `thurro_metrics2`) or a JSON-key mapping (`mf_general_name_mapping`) — that have
no entity of their own to show in a dropdown; they exist only to be joined against. Leave
`entity_type`/`entity_id_column`/`entity_name_column` as `null` for these.

## When to update this

- A new table is added to the Table Builder catalog (`index.json` gains a new `tableName` under
  some domain) that isn't already in `table-entity-map.json`.
- A new metric is added to `src/taskpane/metric-registry.json` whose `annual.prefer`/
  `quarterly.prefer` points at a table not yet in the map.
- An existing table's columns change shape (rare, but would invalidate `entity_id_column`/
  `entity_name_column`).

There's no automatic drift detection today — `getEntityTypeForTable()` just returns `null` on a
miss (treated as "no entity type known," not an error), so a missing table fails quietly rather
than loudly. When adding a brand-new table/metric to this app, check this file as part of that
work rather than waiting for something to look wrong.

## Step-by-step: adding a new table

### 1. Get the new table's real columns and summary

Pull the current catalog and find the new table's entry:

```bash
curl -s https://ir-reports.s3.ap-south-1.amazonaws.com/dbcatalog/index.json -o catalog.json
node -e "
const data = require('./catalog.json');
for (const [domain, domainTables] of Object.entries(data)) {
  for (const entry of domainTables) {
    for (const [tableName, meta] of Object.entries(entry)) {
      if (tableName === 'YOUR_NEW_TABLE_NAME_HERE') {
        console.log(domain, tableName);
        console.log(Object.keys(meta.columns || {}));
        console.log(meta.summary);
      }
    }
  }
}
"
```

If the table's columns come back as a **template placeholder** (something containing `<...>`,
e.g. `<filter_name where Type='Y'>`) instead of real column names — like `sc_year_data`/
`sc_quater_data`/`technical_analysis` already do — resolve the real columns via
`INFORMATION_SCHEMA.COLUMNS` the same way `resolveTemplatedTableColumns` does in `taskpane.js`,
or just open the table's own schema doc at
`https://ir-reports.s3.ap-south-1.amazonaws.com/dbcatalog/<tableName>.md`.

**Don't trust the catalog's `summary` text on its own** — it's written for an LLM's query-planning
use case, not for entity classification, and at least one existing table's summary
(`GIS.dbo.company_holding_percent`) is a verbatim copy-paste of a *different* table's summary and
describes the wrong structure entirely. Always classify from the actual column list.

### 2. Decide the entity type from the columns

Look for an id+name column pair and match it against what's already in the map — new tables are
almost always a variant of an existing entity type, not a genuinely new one:

| Columns look like | Entity type |
|---|---|
| `fincode` (+ `CompName`/`company_name`/`Company`/`name`) | Company |
| `sector_id`/`Sector`/`sector_name` | Sector |
| `pms_id`/`pms_name` (provider-level) | PMS Provider |
| `ia_id`/`ia_name` (+ a `pms_id` it belongs to) | PMS Investment Approach |
| `FSCBI-ProviderCompanyID`/`FSCBI-ProviderCompanyName` | Mutual Fund AMC |
| `scheme_isin`/`FSCBI-ISIN`/`fundID`/`FSCBI-FundName` | Mutual Fund Scheme |
| `identifier`/`title` (macro/economic series) | Economic Indicator |

If none of those fit, it's a genuinely new entity type — add it to the top-level `entity_types`
array too, not just to that one table's row.

If the table has an id column but **no** name column (common — e.g. `ratios_ind`, `mva_technical`,
`comp_about_us` only carry `fincode`), set `entity_name_column: null` and note which table resolves
the name (almost always `basic_info` for Company, the relevant `*_details`/`*_providers` table for
everything else).

If a single row mixes two entity types — a holdings/relationship table like
`current_fund_holdings` (scheme ↔ held company), or a pivoted table like
`new_morningstar_amc_sector` (one row per Sector, one *column* per AMC name) — set `entity_type`/
`entity_id_column` to the table's own **row** identity, and add `related_entity_type` /
`related_entity_id_column` for the other side.

If the table has no id+name pair representing a real-world thing at all — it's a pure lookup/
mapping table like `global_filter` or `mf_general_name_mapping` — set `is_entity_table: false`
and leave the entity fields `null`.

### 3. Add the entry locally

Edit `src/taskpane/table-entity-map.json` and append the new object to the `tables` array (keep
the array sorted the same way `index.json` groups them — by domain — so a diff stays easy to
read). Validate it parses and the count is what you expect:

```bash
node -e "const m=require('./src/taskpane/table-entity-map.json'); console.log('tables:', m.tables.length);"
```

### 4. Re-upload to S3 (the live copy the app actually reads)

Needs AWS credentials for the `ir-feed` bucket (ask whoever holds them if you don't — do not commit
them to this repo). Overwrite the **same key**, `excel-addin-metadata/table-entity-map.json` — this
replaces the object in place, it does not create a new file or need a new URL:

```python
import boto3

s3 = boto3.client(
    "s3",
    region_name="ap-south-1",
    aws_access_key_id="...",       # from whoever holds the ir-feed credentials
    aws_secret_access_key="...",
)
with open("src/taskpane/table-entity-map.json", "rb") as f:
    s3.put_object(
        Bucket="ir-feed",
        Key="excel-addin-metadata/table-entity-map.json",
        Body=f.read(),
        ContentType="application/json",
        ACL="public-read",   # must be repeated on every overwrite — ACL doesn't "stick" across a new PutObject otherwise
    )
```

### 5. Verify

```bash
curl -s https://ir-feed.s3.ap-south-1.amazonaws.com/excel-addin-metadata/table-entity-map.json | node -e "
let d=''; process.stdin.on('data', c => d += c);
process.stdin.on('end', () => { const j = JSON.parse(d); console.log('tables:', j.tables.length); });
"
```

Confirm the count matches your local file, and spot-check the new table's entry is present with
the fields you expect.

### 6. Cache staleness

The add-in caches this file for 24h in `localStorage` (`qd_table_entity_map_cache_v1`). A session
that already loaded the old version won't see your update until that cache expires or the
workbook/task pane is closed and reopened fresh (same behavior as the catalog's own 24h cache — no
special-casing needed, this isn't urgent enough to warrant a cache-busting mechanism). If you ever
need everyone to pick up a change immediately, that cache key is where you'd look — but that's not
part of the normal add-a-table workflow.

## Where the actual lookup is used in code

- `TABLE_ENTITY_MAP_URL` — the S3 URL constant, next to `DOWNLOADTABLE_CATALOG_URL` in
  `src/taskpane/taskpane.js`.
- `getTableEntityMap()` — fetches + caches, returns a `Map` keyed by table name.
- `getEntityTypeForTable(tableName)` — convenience lookup; returns the table's row from the map, or
  `null` for an unmapped/non-entity table.

Both live in the same `downloadtable` closure as `getCatalogEntries`, right after it.
