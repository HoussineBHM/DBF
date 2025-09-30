import io
import os
import sys
import pandas as pd

# Ensure the project root is on sys.path so tests can import the Streamlit app module
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from streamlit_app import transform_excel, df_to_dbf_records, write_dbf_bytes, SCHEMA, records_to_preview_df


def sample_odoo_df():
    return pd.DataFrame([
        {
            'Numéro': 'INV001',
            'Compte': '401000',
            'Partenaire': 'ACME SARL',
            'Date': '15/08/2023',
            'Débit': 0.0,
            'Crédit': 123.45,
            'Libellé': 'Vente',
            'Taxe d\'origine': '21%'
        },
        {
            'Numéro': 'INV002',
            'Compte': '700000',
            'Partenaire': 'FOO SA',
            'Date': '01/09/2023',
            'Débit': 0.0,
            'Crédit': 200.0,
            'Libellé': 'Ventes 700',
            'Taxe d\'origine': '0%'
        }
    ])


def test_convert_parity_bytes_and_preview():
    df = sample_odoo_df()
    recs_t, unb = transform_excel(df, keep_other_70x=False, map21='211400', map00='211100')
    bytes_t = write_dbf_bytes(recs_t, SCHEMA, encoding='latin-1')

    recs_c = df_to_dbf_records(df)
    bytes_c = write_dbf_bytes(recs_c, SCHEMA, encoding='latin-1')

    assert bytes_t == bytes_c, "DBF bytes differ between Transform and Convert"

    pv_t = records_to_preview_df(recs_t, SCHEMA)
    pv_c = records_to_preview_df(recs_c, SCHEMA)
    # DataFrame equality in content
    assert pv_t.equals(pv_c), "Preview DataFrames differ"
