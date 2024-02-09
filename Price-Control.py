import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.backends.backend_pdf as pdf_backend

def apply_inflation_to_prices(prices, inflation_percentage):
    """Apply inflation to prices and fill NaN values with 'No price'."""
    prices_before_inflation = prices.copy()
    prices = prices * (1 + inflation_percentage / 100)
    prices.fillna("No price", inplace=True)
    return prices, prices_before_inflation

def update_prices_with_inflation(excel_file_path, output_excel_file_path, output_pdf_file_path):
    """
    Update prices in an Excel file with inflation and save as PDF.

    Args:
    - excel_file_path (str): Path to the input Excel file.
    - output_excel_file_path (str): Path to the output Excel file.
    - output_pdf_file_path (str): Path to the output PDF file.

    Returns:
    None
    """
    # Read all sheets into a dictionary of DataFrames
    excel_sheets = pd.read_excel(excel_file_path, sheet_name=None)

    # Create a new Excel writer to save modified sheets
    with pd.ExcelWriter(output_excel_file_path, engine='xlsxwriter') as writer:
        # Iterate through each sheet and apply inflation
        for sheet_name, df in excel_sheets.items():
            if 'Productos' not in df.columns or 'Precio' not in df.columns:
                print(f"Error: 'Productos' or 'Precio' columns not found in sheet '{sheet_name}'. Skipping this sheet.")
                continue

            try:
                inflation_percentage = float(input(f"Enter inflation percentage for sheet '{sheet_name}': "))
            except ValueError:
                print("Error: Please enter a valid numeric inflation percentage.")
                continue

            df['Precio Nuevo'], prices_before_inflation = apply_inflation_to_prices(df['Precio'], inflation_percentage)

            print(f"All prices in sheet '{sheet_name}' updated with {inflation_percentage}% inflation. Empty prices filled with 'No price'.")

            # Add columns for inflation percentage and prices before inflation
            df['Porcentaje de Inflación Aplicado'] = inflation_percentage

            df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Save the updated DataFrame to a PDF using matplotlib
    with pdf_backend.PdfPages(output_pdf_file_path) as pdf:
        for sheet_name, df in pd.read_excel(output_excel_file_path, sheet_name=None).items():
            plt.figure(figsize=(10, 6))
            plt.axis('off')
            plt.table(cellText=df.values,
                      colLabels=df.columns,
                      cellLoc='center',
                      loc='center')
            plt.title(sheet_name)
            pdf.savefig()
            plt.close()

# Example usage with manual input of inflation percentages for each sheet
excel_file_path = r'C:\Users\andres.pangrazi\Desktop\Almacen.xlsx'
output_excel_file_path = r'C:\Users\andres.pangrazi\Desktop\updated_Almacen.xlsx'
output_pdf_file_path = r'C:\Users\andres.pangrazi\Desktop\updated_Almacen.pdf'

update_prices_with_inflation(excel_file_path, output_excel_file_path, output_pdf_file_path)
