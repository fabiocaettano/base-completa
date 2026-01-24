# AI Agent Instructions for Excel-to-JSON Data Pipeline

## Project Overview
This codebase consolidates multi-sheet Excel data (pedidos/orders, lotes/batches, objetos_postado/posted objects, mcco/cost centers) into structured JSON for a bot/RPA system. The primary output is `dados_para_bot` with unified order records containing hierarchical relationships between orders, items, and delivery tracking.

## Architecture & Data Flow

**Four Input Sheets** → **Consolidation Logic** → **JSON Output + Excel Summary**

### Sheet Processing Order
1. **base_pedidos**: Orders with item-level details (CLIENTE, NUM_PEDIDO_SISCAP, MCCO, status codes)
   - Key transformation: Filter by status pairs (520→540, 541→540 = RESERVED; 980/982/983→999 = CANCELED)
   - Creates unique key: `NUM_PEDIDO_SISCAP|CLIENTE`

2. **base_lotes**: Batch/lot grouping for fulfillment (references pedido via PEDIDO column)
   - Enriches with zone, status_impressao (print status), roteirizador (routing)

3. **base_objeto_postado**: SRO tracking data (when orders are shipped)
   - Adds REGISTRO (tracking number), DATA_EXPEDICAO, DATA_CANCELAMENTO

4. **base_mcco**: Lookup table mapping MCCO codes to DESCRICAO_MCCO (regional descriptions)

### Key Consolidation Rules
- **Uniqueness**: Each NUM_PEDIDO_SISCAP appears once in output (never duplicated)
- **Grouping**: All items for an order grouped under PEDIDO_NOVO → RESERVED/CANCELED arrays
- **Status Mapping**: Hardcoded status codes (520, 541, 980, 982, 983, 540, 999) determine item classification
- **MCCO Lookup**: Join base_pedidos.MCCO with base_mcco.MCCO to populate DESCRICAO_SE

## Code Patterns & Conventions

### File Versions & Evolution
- `processar_excel_json.py` (v1): Initial implementation, single-sheet focused
- `processar_excel_json_v02.py` (v2): **Full implementation** with all three sheet joins (RECOMMENDED)
- `processar_excel_json_v03.py` (v3): Latest iteration (check for improvements)
- `teste2.py`: Functional test/example with 5-stage pipeline commented

### Memory & Performance Optimization
```python
# Always configure pandas first
configurar_pandas()

# Use dtype specifications to reduce memory:
dtype={'CLIENTE': 'int32', 'QTD': 'int32', 'NUM_PEDIDO_SISCAP': str}

# Filter early (usecols, nrows):
usecols=['DR', 'PED', 'CLIENTE', ...]  # Only needed columns
nrows=limit  # For testing

# Garbage collection for large datasets
gc.collect()
```

### Data Cleaning Pattern
```python
# Strip whitespace from string columns
str_cols = df.select_dtypes(include=['object']).columns
for col in str_cols:
    df[col] = df[col].astype(str).str.strip()

# Numeric conversion with error handling
df['ULT_STATUS'] = pd.to_numeric(df['ULT_STATUS'], errors='coerce')
```

### Status Filtering Pattern
Always use this exact logic for reservado/cancelado:
```python
# RESERVED: 520→540 OR 541→540
mask_reserved = ((df['ULT_STATUS'] == 520) & (df['PROX_STATUS'] == 540)) | \
                ((df['ULT_STATUS'] == 541) & (df['PROX_STATUS'] == 540))

# CANCELED: 980→999 OR 982→999 OR 983→999
mask_canceled = ((df['ULT_STATUS'] == 980) & (df['PROX_STATUS'] == 999)) | \
                ((df['ULT_STATUS'] == 982) & (df['PROX_STATUS'] == 999)) | \
                ((df['ULT_STATUS'] == 983) & (df['PROX_STATUS'] == 999))
```

### JSON Output Structure
```python
{
  'NUM_PEDIDO_SISCAP': '0023476325260nu',
  'CLIENTE': '9906821',
  'NOME_CLIENTE': 'AGF MAJOR ARTHUR',
  'MCCO': '74',
  'DESCRICAO_SE': 'from base_mcco',
  'PEDIDO_NOVO': [
    {
      'DATA_PEDIDO': '2026-01-19',
      'PEDIDO': '154926',
      'TIPO_PEDIDO': 'FM',
      'RESERVED': [{CODIGO, DESCRICAO_CODIGO, UM, QTDE}],
      'CANCELED': [{CODIGO, DESCRICAO_CODIGO, UM, QTDE}]
    }
  ],
  'PEDIDO_WMS': [...],
  'PEDIDO_POSTADO': [...]
}
```

## Critical Integration Points

### Joining Multiple DataFrames
- **Pedidos ↔ Lotes**: Via `PEDIDO` + `TIPO_PEDIDO` columns
- **Lotes ↔ Objetos Postado**: Via `LOTE` column matching
- **Pedidos ↔ MCCO**: Via `MCCO` column lookup

When joining:
1. Ensure column names are normalized (rename columns in etapa 02)
2. Handle NaN values from mismatched joins (use `pd.merge(..., how='left')`)
3. Verify join coverage (compare before/after row counts)

### Output Formats
- **Individual JSONs**: Saved per order (`pedido_{NUM_PEDIDO|CLIENTE}.json`)
- **Excel Summary**: Consolidated view with JSON column for bot import
- **File naming**: Replace pipe `|` with underscore in filenames

## Development Workflow

### Testing & Validation
1. Run `verificar_instalacao.py` to confirm pandas, openpyxl, numpy installed
2. Start with `nrows=10000` or less for testing large files
3. Use `teste2.py` as reference for the 5-stage pipeline
4. Print progress markers (✓ / ✗) for each stage

### Common Debugging Points
- **Missing columns**: Check sheet_name spelling (case-sensitive)
- **Status mismatches**: Verify ULT_STATUS + PROX_STATUS combinations
- **Date parsing**: Always use `pd.to_datetime(..., errors='coerce')`
- **Encoding**: Use `encoding='utf-8'` for JSON output to handle Portuguese chars

### File Location Assumptions
- Input: `001-base_completa.xlsx` in working directory (see teste2.py)
- Output: `output/` or `results/` directory (create with `os.makedirs(..., exist_ok=True)`)

## Language & Environment
- **Python 3.7+** (check with `verificar_instalacao.py`)
- **Key packages**: pandas, openpyxl, numpy (openpyxl required for Excel I/O)
- **Locale**: Portuguese (pt-BR) - column names, descriptions, comments in Portuguese
- **Character encoding**: Always UTF-8 for JSON with `ensure_ascii=False`

