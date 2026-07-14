"""
All the fixed configuration for the Ventus composition report:
column order, colors, fonts and file paths.

If the report layout ever needs to change (new column, different color),
this is the only file that should need editing.
"""

from pathlib import Path

# ---------------------------------------------------------------------------
# Output sheet layout
# ---------------------------------------------------------------------------

SHEET_NAME_OUTPUT = 'Reranged'
HEADER_ROW = 2  # row with column titles in the output sheet
FIRST_DATA_ROW = 3  # first row of actual product data

DESIRED_ORDER = [
    'Sortiment',
    'Foto 1',
    'Foto 2',
    'Nazev',
    'Skupina',
    'Prodejní cena 2024-Y24',
    'Prodejní cena 2025-Y25',
    'Prodejní cena 2026-Y26',
    'Nákupní cena',
    'Skut. cena / kus',
    'VOZ CENÍK',
    'Marže 2024-Y24',
    'Marže 2025-Y25',
    'Marže 2026-Y26',
    'Marže % 2024-Y24',
    'Marže % 2025-Y25',
    'Marže % 2026-Y26',
    'Množství 2024-Y24',
    'Množství 2025-Y25',
    'Množství 2026-Y26',
    'Příjem obd. 2024-Y24',
    'Příjem obd. 2025-Y25',
    'Příjem obd. 2026-Y26',
    'Sklade fyzicky',
    'Skladem volné',
    'Suma zbývá dodat',
    '(Skladem + zbývá dodat) / prodej 2025',
    'Order ctn',
    'Order pcs',
    'Komentař Jarmila',
    'Komentař nakup',
    'Balení 3',
    'Minimální odběr pro OV',
    'Objem',
    'Objem celkem',
    'Poznámka',
    'Slevová skupina',
    'Prim. dodavatel',
    'Dod. skupina',
    'EAN',
]

# Column (by name) whose values sort the whole report, descending.
SORT_COLUMN = 'Prodejní cena 2025-Y25'

# Column that holds the product code used to match the "top sellers" list.
CODE_COLUMN = 'Sortiment'

# Source columns that get divided by 100 (stored as e.g. 42 meaning 42%).
PERCENT_COLUMNS = {'Marže % 2024-Y24', 'Marže % 2025-Y25', 'Marže % 2026-Y26'}

# Source column name that may appear multiple times (multiple deliveries).
DELIVERY_QTY_COLUMN = 'Zbývá dodat ks'
DELIVERY_DATE_COLUMN = 'Datum dodání z.'

HIDDEN_COLUMNS = ('I', 'J', 'AG', 'AM', 'AN')

# ---------------------------------------------------------------------------
# Colors (hex, no '#') and fonts
# ---------------------------------------------------------------------------

COLOR_HEX = {
    'header_blue': 'd4ebff',
    'yellow': 'fbff1f',
    'blue': '1c86ff',
    'violet': 'bb1cff',
    'light_blue': 'c9e6ff',
    'light_green': 'd0ffb3',
    'light_yellow': 'ffe496',
    'soft_gray': 'd1d1d1',
    'soft_orange': 'ffd5bd',
    'soft_green': '94f7af',
    'bestseller': 'ffd700',  # gold highlight for top-30 items
}

FONT_NAME = 'Segoe UI'
BASE_FONT_SIZE = 12
HEADER_FONT_SIZE = 9

# ---------------------------------------------------------------------------
# App / persistence
# ---------------------------------------------------------------------------

APP_TITLE = 'Ventus Composition Processor'
CONFIG_DIR = Path.home() / '.ventus_processor'
CONFIG_FILE = CONFIG_DIR / 'config.json'
