
import pandas as pd
import numpy as np
import os

def load_data(file_path):
    """Loads all required sheets from the Excel file into a dictionary of pandas DataFrames."""
    xls = pd.ExcelFile(file_path)
    sheets = ["BASE 2026", "PVTA_MAT VK13", "PVTA_FTE VK13", "CONT_VTA ZSDD4501", "CONT_COMPRA TRAOPE"]
    dataframes = {}
    for sheet in sheets:
        if sheet == "BASE 2026":
            df = pd.read_excel(xls, sheet_name=sheet, header=None) # Read without header to inspect
            # Print first 15 rows and first 10 columns for debugging
            print(f"Raw data head for sheet '{sheet}':")
            print(df.iloc[:15, :10].to_string())
        else:
            df = pd.read_excel(xls, sheet_name=sheet)
        dataframes[sheet] = df
    return dataframes

def clean_data(df):
    """Performs the specified hygiene transformations."""
    if 'Unidad' in df.columns:
        df['Unidad'] = df['Unidad'].replace('TO', 'TN')
    return df

def generate_ruta_key(df):
    """Creates the 'RUTA' key in the main DataFrame."""
    # Ensure relevant columns are string type before concatenation
    df['CEDIS'] = df['CEDIS'].astype(str)
    df['SF'] = df['SF'].astype(str)
    df['NO PLANTA / OBRA'] = df['NO PLANTA / OBRA'].astype(str)
    df['ARTICULO'] = df['ARTICULO'].astype(str)
    df['MODALIDAD'] = df['MODALIDAD'].astype(str)

    df['RUTA'] = df['CEDIS'] + '-' + \
                 df['SF'] + '-' + \
                 df['NO PLANTA / OBRA'] + '-' + \
                 df['ARTICULO'] + '-' + \
                 df['MODALIDAD'].str[-1]
    return df

def perform_joins(base_df, pvta_mat_df, pvta_fte_df, cont_compra_df):
    """Executes all the left join operations."""

    # PVTA_MAT Join
    base_df = pd.merge(base_df, pvta_mat_df[['RUTA', 'Importe (MP VTA SNOW)', 'Unidad (UM MAT)']],
                       on='RUTA', how='left', suffixes=('', '_PVTA_MAT'))

    # PVTA_FTE Join
    base_df = pd.merge(base_df, pvta_fte_df[['RUTA', 'Unidad']],
                       on='RUTA', how='left', suffixes=('', '_PVTA_FTE'))
    
    # Fill NaN in 'Unidad_PVTA_FTE' with "RECOGIDO"
    if 'Unidad_PVTA_FTE' in base_df.columns:
        base_df['Unidad_PVTA_FTE'] = base_df['Unidad_PVTA_FTE'].fillna('RECOGIDO')
    else:
        # If 'Unidad_PVTA_FTE' column was not created due to no matching 'Unidad' in PVTA_FTE,
        # it means there were no matches. We should create a new column and fill it.
        # This assumes the original 'Unidad' column from BASE 2026 is preserved
        # and 'Unidad' from PVTA_FTE is a new column.
        # If there's a specific column name for PVTA_FTE's 'Unidad', it should be specified.
        # For now, I'm assuming 'Unidad_PVTA_FTE' is the target column after the merge.
        # If the 'Unidad' column in PVTA_FTE is meant to overwrite the original 'Unidad'
        # then the logic would be different.
        pass # The instruction was to fill if the unit *remains* null, which means it should be there.
    

    # CONT_COMPRA Join
    # Ensure 'MODALIDAD' is string type and handle potential NaN before checking last digit
    base_df['MODALIDAD'] = base_df['MODALIDAD'].astype(str).replace('nan', '')

    cont_compra_modalidad_4 = base_df[base_df['MODALIDAD'].str.endswith('4')]
    cont_compra_other = base_df[~base_df['MODALIDAD'].str.endswith('4')]

    if not cont_compra_modalidad_4.empty:
        cont_compra_modalidad_4 = pd.merge(cont_compra_modalidad_4, cont_compra_df,
                                           on='RUTA', how='left', suffixes=('', '_CONT_COMPRA'))

    if not cont_compra_other.empty:
        # Ensure join keys are string type for cont_compra_other
        cont_compra_other['CEDIS'] = cont_compra_other['CEDIS'].astype(str)
        cont_compra_other['SF'] = cont_compra_other['SF'].astype(str)
        cont_compra_other['ARTICULO'] = cont_compra_other['ARTICULO'].astype(str)
        
        cont_compra_df['CEDIS'] = cont_compra_df['CEDIS'].astype(str)
        cont_compra_df['SF'] = cont_compra_df['SF'].astype(str)
        cont_compra_df['ARTICULO'] = cont_compra_df['ARTICULO'].astype(str)

        cont_compra_other = pd.merge(cont_compra_other, cont_compra_df,
                                     on=['CEDIS', 'SF', 'ARTICULO'], how='left', suffixes=('', '_CONT_COMPRA'))

    base_df = pd.concat([cont_compra_modalidad_4, cont_compra_other]).sort_index()

    return base_df


def export_data(df, output_path):
    """Saves the final DataFrame to an Excel file without formulas."""
    df.to_excel(output_path, index=False)
    print(f"Processed data exported to {output_path}")

def main(file_path, output_path):
    print(f"Starting ETL pipeline for {file_path}")
    dataframes = load_data(file_path)
    base_df = dataframes["BASE 2026"]
    pvta_mat_df = dataframes["PVTA_MAT VK13"]
    pvta_fte_df = dataframes["PVTA_FTE VK13"]
    cont_compra_df = dataframes["CONT_COMPRA TRAOPE"]

    base_df = clean_data(base_df)
    base_df = generate_ruta_key(base_df)
    base_df = perform_joins(base_df, pvta_mat_df, pvta_fte_df, cont_compra_df)

    export_data(base_df, output_path)
    print("ETL pipeline completed.")

if __name__ == "__main__":
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    input_excel_path = os.path.join(BASE_DIR, "EXTRACCIÓN_DATOS_AGREGADOS_RESPALDO_v0.xlsx")
    output_excel_path = os.path.join(BASE_DIR, "BASE_2026_PROCESADA.xlsx")

    # Create the directory if it doesn't exist
    output_dir = os.path.dirname(output_excel_path)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    main(input_excel_path, output_excel_path)
