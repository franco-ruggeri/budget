import pandas as pd


def summarize(sheet, data_frame, group_by, drop_columns):
    # Prepare dataframe
    data_frame = (
        data_frame
        .drop(drop_columns, axis=1)
        .groupby(by=group_by).sum(numeric_only=True)
        .reset_index()
    )
    columns = data_frame.columns.tolist()
    columns = columns[-1:] + columns[:-1]
    data_frame = data_frame[columns]

    # Write dataframe to Excel range
    row_offset = sum([table.range.rows.count + 2 for table in sheet.tables])
    n_rows, n_cols = data_frame.shape
    range_ = sheet[row_offset:row_offset+n_rows+1, 0:n_cols]
    range_.options(index=False).value = data_frame

    # Improve style
    sheet.tables.add(source=range_, table_style_name="TableStyleMedium1")
    range_.row_height = 20
    range_.column_width = 20
    range_.number_format = "#,##0.00"


def table_to_df(table):
    return table.range.expand().options(pd.DataFrame, index=0).value
