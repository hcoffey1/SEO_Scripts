# Hayden Coffey
import sys
import os
import pdfplumber
import pandas as pd
from openpyxl import load_workbook

def get_pdf_table(file):
    MAGIC_PHRASE = "Well"
    all_tables = []
    with pdfplumber.open(file) as pdf:
        
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if not table:
                    continue

                df = pd.DataFrame(table)

                # Check if the MAGIC_PHRASE column exists in the header
                header = df.iloc[0]
                if MAGIC_PHRASE in header.values:
                    df.columns = header  # Set first row as header
                    df = df[1:]  # Drop the header row from data
                    all_tables.append(df)

    # Combine all relevant tables
    final_df = pd.concat(all_tables, ignore_index=True)

    return final_df

def get_input_ref_df(file):
    wb = load_workbook(file)

    source = wb.active
    ws = wb.copy_worksheet(source)

    val_list = (list(ws.values))

    i = 0
    for l in val_list:
        if all(value is None for value in l):
            break
        i+=1
    

    content_ref_df = pd.DataFrame(val_list[0:i]).drop(index=0).set_index(0)
    color_ref_df = content_ref_df.copy()

    #Last part to do is add the color information to this dataframe so we can crossreference with other table.
    row = 0
    col = 0

    for r in ws[2:i]:
        col = 0
        for c in r[1:]:
            if (color_ref_df.iloc[row,col]): #c.fill.fgColor.theme):
                color_ref_df.iloc[row,col] = c.fill.fgColor.theme
            else:
                break
            col += 1
        row += 1
    
    #Strip whitespace
    content_ref_df.index = content_ref_df.index.astype(str).str.strip()
    #content_ref_df.columns = content_ref_df.columns.astype(str).str.strip()

    color_row = (ws[i+2])
    
    color_table = []
    for c in color_row:
        if c.value == None:
            continue

        color_table.append([c.fill.fgColor.theme, c.value])
    
    color_df = (pd.DataFrame(color_table))
    color_df=color_df.set_index(0)

    color_match_dict = (color_df[1].to_dict())

    color_ref_df = color_ref_df.replace(color_match_dict)

    return content_ref_df, color_ref_df

def get_target_value(well, df_lookup):
    """Extract the lookup value from df_lookup based on a well identifier like 'A1'."""
    # Parse the well string: 
    # Row letter is the first character; column number is the rest (it could be multiple digits)
    row_label = well[0].upper()
    col_number = int(well[1:])

    # Get the corresponding lookup value from df_lookup
    return df_lookup.loc[row_label, col_number]


def main():

    if len(sys.argv) < 3:
        print("Usage: python {} input.pdf input_ref.xlsx (optional:) output.xlsx".format(sys.argv[0]))
        return
    
    input_pdf_file = sys.argv[1]
    input_ref_file = sys.argv[2]

    output_file = "out_hc.xlsx"

    if len(sys.argv) == 4:
        output_file = sys.argv[3] 

    pdf_table_df = get_pdf_table(input_pdf_file)
    input_ref_df, color_ref_df = get_input_ref_df(input_ref_file)

    pdf_table_df["Content"] = pdf_table_df["Well"].apply(lambda well: get_target_value(well, input_ref_df))
    pdf_table_df["Target"] = pdf_table_df["Well"].apply(lambda well: get_target_value(well, color_ref_df))

    pdf_table_df = pdf_table_df.drop(['Fluor'], axis=1)

    if os.path.exists(output_file):
         raise FileExistsError(f"The file '{output_file}' already exists. Will not overwrite.")

    pdf_table_df.to_excel(output_file, index=False)

if __name__ == "__main__":
    main()
