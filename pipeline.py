"""
BoroughLens — Data Pipeline
Reads raw Excel datasets, filters to London boroughs,
outputs one clean JSON file per metric to data/

Run: python3 pipeline.py
Requires: pip install openpyxl pandas
"""

import pandas as pd
import json
import os
from pathlib import Path

# ── CONFIG ────────────────────────────────────────────────────────────────────

DATA_DIR = Path('./data')

FILES = {
    'house_price_earnings': './raw/ratio-house-price-earnings-residence-based.xlsx',
    'fuel_poverty_old':     './raw/fuel_poverty.xlsx',
    'fuel_poverty_new':     './raw/fuel-poverty-sub-regional-statistics-5-december-2024-update.xlsx',
    'gdhi':                 './raw/regionalgrossdisposablehouseholdincomelocalauthorities2023.xlsx',
    'prs': './raw/privaterentalmarketstatistics231220.xls',
}

# Canonical borough list — single source of truth for names and ordering
BOROUGHS = {
    'E09000001': 'City of London',
    'E09000002': 'Barking and Dagenham',
    'E09000003': 'Barnet',
    'E09000004': 'Bexley',
    'E09000005': 'Brent',
    'E09000006': 'Bromley',
    'E09000007': 'Camden',
    'E09000008': 'Croydon',
    'E09000009': 'Ealing',
    'E09000010': 'Enfield',
    'E09000011': 'Greenwich',
    'E09000012': 'Hackney',
    'E09000013': 'Hammersmith and Fulham',
    'E09000014': 'Haringey',
    'E09000015': 'Harrow',
    'E09000016': 'Havering',
    'E09000017': 'Hillingdon',
    'E09000018': 'Hounslow',
    'E09000019': 'Islington',
    'E09000020': 'Kensington and Chelsea',
    'E09000021': 'Kingston upon Thames',
    'E09000022': 'Lambeth',
    'E09000023': 'Lewisham',
    'E09000024': 'Merton',
    'E09000025': 'Newham',
    'E09000026': 'Redbridge',
    'E09000027': 'Richmond upon Thames',
    'E09000028': 'Southwark',
    'E09000029': 'Sutton',
    'E09000030': 'Tower Hamlets',
    'E09000031': 'Waltham Forest',
    'E09000032': 'Wandsworth',
    'E09000033': 'Westminster',
}

# ── HELPERS ───────────────────────────────────────────────────────────────────

def write_json(filename, data):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    out = DATA_DIR / filename
    with open(out, 'w') as f:
        json.dump(data, f, indent=2)
    print(f'  Written: {out}')

def round2(n):
    return round(float(n), 2)

def safe_float(val):
    """Return float or None — handles NaN, '[x]', None."""
    if val is None:
        return None
    try:
        f = float(val)
        return None if pd.isna(f) else f
    except (ValueError, TypeError):
        return None

# ── METRIC 1: HOUSE PRICE TO EARNINGS ────────────────────────────────────────

def build_house_price_earnings():
    print('\nBuilding house-price-to-earnings...')
    df = pd.read_excel(
        FILES['house_price_earnings'],
        sheet_name='Median Earnings to Prices ratio',
        header=0
    )

    # Year columns start after New Code, Code, Area
    year_cols = [c for c in df.columns if isinstance(c, int)]
    london = df[df['New Code'].astype(str).str.startswith('E09')]

    boroughs = []
    for _, row in london.iterrows():
        ons_code = row['New Code']
        values = {}
        for year in year_cols:
            val = safe_float(row[year])
            if val is not None:
                values[str(year)] = round2(val)
        boroughs.append({
            'ons_code': ons_code,
            'name': BOROUGHS.get(ons_code, row['Area']),
            'values': values,
        })

    all_years = [str(y) for y in year_cols]
    write_json('house_price_earnings.json', {
        'metric': 'house_price_earnings',
        'description': 'Median house price to annual earnings ratio by borough',
        'source': 'MHCLG / London Datastore',
        'years': all_years,
        'derived': False,
        'boroughs': boroughs,
    })

# ── METRIC 2: FUEL POVERTY ────────────────────────────────────────────────────

def build_fuel_poverty():
    print('\nBuilding fuel poverty...')

    # --- Old file: 2009–2017 ---
    raw_old = pd.read_excel(FILES['fuel_poverty_old'], sheet_name='LA and Regions', header=None)
    headers = raw_old.iloc[1].tolist()  # ['Code ', 'Area  Name ', 2009, 2010 ...]
    year_cols = headers[2:]
    old_data = {}
    for _, row in raw_old.iloc[2:].iterrows():
        code = str(row.iloc[0]).strip() if row.iloc[0] else ''
        if not code.startswith('E09'):
            continue
        old_data[code] = {}
        for i, year in enumerate(year_cols):
            val = safe_float(row.iloc[2 + i])
            if val is not None:
                old_data[code][str(int(float(year)))] = round2(val)

    # --- New file: 2022 snapshot ---
    raw_new = pd.read_excel(FILES['fuel_poverty_new'], sheet_name='Table 2', header=None)
    new_data = {}
    for _, row in raw_new.iterrows():
        code = str(row.iloc[0]) if row.iloc[0] else ''
        if not code.startswith('E09'):
            continue
        val = safe_float(row.iloc[6])  # col 6 = proportion fuel poor %
        if val is not None:
            new_data[code] = {'2022': round2(val)}

    # --- Merge and build output ---
    boroughs = []
    for ons_code, name in BOROUGHS.items():
        values = {**old_data.get(ons_code, {}), **new_data.get(ons_code, {})}
        boroughs.append({'ons_code': ons_code, 'name': name, 'values': values})

    all_years = sorted({y for b in boroughs for y in b['values']})
    write_json('fuel_poverty.json', {
        'metric': 'fuel_poverty',
        'description': 'Proportion of households in fuel poverty (%) by borough',
        'source': 'DESNZ / London Datastore',
        'years': all_years,
        'derived': False,
        'boroughs': boroughs,
    })

# ── METRIC 3: GDHI PER HEAD ───────────────────────────────────────────────────

def build_gdhi():
    print('\nBuilding GDHI per head...')
    raw = pd.read_excel(FILES['gdhi'], sheet_name='Table 3', header=None)

    header_row = raw.iloc[1].tolist()
    # Cols: Region(0), LAD code(1), Region name(2), 1997(3) ... 2023
    year_cols = header_row[3:]
    years = [str(int(float(y))) for y in year_cols if safe_float(y) is not None]

    boroughs = []
    for _, row in raw.iloc[2:].iterrows():
        ons_code = str(row.iloc[1]) if row.iloc[1] else ''
        if not ons_code.startswith('E09'):
            continue
        values = {}
        for i, year in enumerate(years):
            val = safe_float(row.iloc[3 + i])
            if val is not None:
                values[year] = int(round(val))
        boroughs.append({
            'ons_code': ons_code,
            'name': BOROUGHS.get(ons_code, str(row.iloc[2])),
            'values': values,
        })

    write_json('gdhi_per_head.json', {
        'metric': 'gdhi_per_head',
        'description': 'Gross disposable household income per head (£) by borough',
        'source': 'ONS Regional GDHI',
        'years': years,
        'derived': False,
        'boroughs': boroughs,
    })

# ── METRIC 4: RENT TO INCOME (DERIVED) ───────────────────────────────────────
#
# Formula:
#   monthly_income   = gdhi_per_head_2022 / 12
#   rent_to_income % = (median_monthly_rent / monthly_income) * 100
#
# Periodicity note: PRS covers Oct 2022–Sep 2023.
# Aligned to GDHI 2022 as closest full calendar year.

def build_rent_to_income():
    print('\nBuilding rent-to-income (derived)...')

    # --- PRS median monthly rent ---
    prs_raw = pd.read_excel(FILES['prs'], sheet_name='Table2.7', header=None)
    prs_rents = {}
    for _, row in prs_raw.iterrows():
        code = str(row.iloc[2]) if row.iloc[2] else ''
        if not code.startswith('E09'):
            continue
        median = safe_float(row.iloc[7])  # col 7 = median rent
        if median is not None:
            prs_rents[code] = median

    # --- GDHI per head 2022 ---
    gdhi_raw = pd.read_excel(FILES['gdhi'], sheet_name='Table 3', header=None)
    header_row = gdhi_raw.iloc[1].tolist()
    year_cols = header_row[3:]
    years = [str(int(float(y))) for y in year_cols if safe_float(y) is not None]
    col_2022 = years.index('2022') + 3  # offset to raw column index

    gdhi_2022 = {}
    for _, row in gdhi_raw.iloc[2:].iterrows():
        code = str(row.iloc[1]) if row.iloc[1] else ''
        if not code.startswith('E09'):
            continue
        val = safe_float(row.iloc[col_2022])
        if val is not None:
            gdhi_2022[code] = val

    # --- Derive and build output ---
    boroughs = []
    warnings = []

    for ons_code, name in BOROUGHS.items():
        rent = prs_rents.get(ons_code)
        gdhi = gdhi_2022.get(ons_code)

        if not rent:
            warnings.append(f'No PRS rent data for {ons_code} ({name})')
        if not gdhi:
            warnings.append(f'No GDHI data for {ons_code} ({name})')

        values = {}
        if rent and gdhi:
            monthly_income = gdhi / 12
            values['2022-23'] = round2((rent / monthly_income) * 100)

        boroughs.append({
            'ons_code': ons_code,
            'name': name,
            'rent_monthly': rent,
            'gdhi_annual_2022': int(round(gdhi)) if gdhi else None,
            'values': values,
        })

    if warnings:
        print('  Warnings:')
        for w in warnings:
            print(f'    {w}')

    write_json('rent_to_income.json', {
        'metric': 'rent_to_income',
        'description': 'Median monthly rent as % of estimated monthly income by borough',
        'source': 'Derived: ONS PRS Statistics (rent) + ONS GDHI (income)',
        'years': ['2022-23'],
        'derived': True,
        'derivation': 'median_monthly_rent / (gdhi_per_head_2022 / 12) * 100',
        'periodicity_note': 'PRS data covers Oct 2022–Sep 2023; GDHI aligned to calendar year 2022',
        'boroughs': boroughs,
    })

# ── RUN ───────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    print('BoroughLens data pipeline starting...')
    print(f'Output directory: {DATA_DIR}')

    build_house_price_earnings()
    build_fuel_poverty()
    build_gdhi()
    build_rent_to_income()

    print('\nAll metrics written successfully.')
