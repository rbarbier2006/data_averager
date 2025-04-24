from flask import Flask, render_template, request, send_file
import pandas as pd
from io import BytesIO

app = Flask(__name__)

def find_table_start(df, keyword):
    for row_index in range(len(df)):
        for col_index in range(len(df.columns)):
            cell_value = str(df.iat[row_index, col_index]).strip()
            if cell_value.lower() == keyword.lower():
                return row_index, col_index
    raise ValueError(f"Keyword '{keyword}' not found.")

def clean_table(df_raw, keyword, num_rows, num_cols):
    start_row, start_col = find_table_start(df_raw, keyword)
    end_row = start_row + num_rows + 1
    end_col = start_col + num_cols
    return df_raw.iloc[start_row:end_row, start_col:end_col]

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    file = request.files['file']
    keyword = request.form['keyword']
    rows = int(request.form['rows'])
    cols = int(request.form['cols'])

    xls = pd.ExcelFile(file)
    usable_sheets = xls.sheet_names[5:]  # start from 6th sheet

    dfs = []
    for sheet in usable_sheets:
        df_raw = xls.parse(sheet, header=None)
        df_cleaned = clean_table(df_raw, keyword, rows, cols)
        dfs.append(df_cleaned)

    num_rows, num_cols = dfs[0].shape
    avg_df = pd.DataFrame(index=range(1, num_rows), columns=range(num_cols))

    for i in range(1, num_rows):
        for j in range(num_cols):
            values = []
            for df in dfs:
                try:
                    val = df.iloc[i, j]
                    if pd.notna(val) and str(val).strip() != "":
                        values.append(float(val))
                except:
                    pass
            if values:
                avg_df.loc[i, j] = round(sum(values) / len(values), 2)

    avg_df.loc[0] = dfs[0].iloc[0]
    avg_df.sort_index(inplace=True)

    output = BytesIO()
    avg_df.to_excel(output, index=False, header=False)
    output.seek(0)

    return send_file(output, download_name="averaged_output.xlsx", as_attachment=True)

import os

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
