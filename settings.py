OUTPUT_COLUMNS = {
    'report_date': 1,
    'source': 2,
    'port': 3,
    'vessel': 6,
    'previous_port': 9,
    'load_discharge': 12,
    'cargo_type': 13,
    'grade': 14,
    'quantity': 15,
    'ETA': 17,
    'ETB': 18,
    'ETD': 19,
    'buyer': 22
}
FORMAT = {
    'quantity': '0',
    'ETA': 'dd\\-mm\\-yy',
    'ETB': 'dd\\-mm\\-yy',
    'ETD': 'dd\\-mm\\-yy'
}
STATIC_ATTRS = {
    'source': 'gac.com',
    'load_discharge': 'D',
    'cargo_type': 'CRUDE OIL'
}
DYNAMIC_ATTRS = ['report_date', 'source']
INPUT_COLUMNS = {
    'vessel': 1,
    'ETA': 2,
    'ETB': 3,
    'ETD': 4,
    'grade': 5,
    'quantity': 6,
    'buyer': 7,
    'previous_port': 8
}
INPUT_TABLE_HEADER = (3, 1)
INPUT_PORT_COLUMN = 1
INPUT_FIRST_PORT = (4, 1)