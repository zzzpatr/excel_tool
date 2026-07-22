import hashlib

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Excel 資料查詢", page_icon="📊", layout="wide")
st.title("📊 Excel 資料查詢與加入清單")

uploaded_file = st.file_uploader("上傳 Excel 檔案", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("請先上傳 Excel 檔案。")
    st.stop()

try:
    file_bytes = uploaded_file.getvalue()
    excel_file = pd.ExcelFile(uploaded_file)
    selected_sheet = excel_file.sheet_names[0]
    check_df = pd.read_excel(excel_file, sheet_name=selected_sheet)
except Exception as error:
    st.error(f"無法讀取 Excel 檔案：{error}")
    st.stop()

if check_df.empty:
    st.warning("這個工作表沒有資料。")
    st.stop()

# 更換檔案或工作表時，清除上一份資料的搜尋及累積結果。
dataset_key = f"{hashlib.md5(file_bytes).hexdigest()}:{selected_sheet}"
if st.session_state.get("dataset_key") != dataset_key:
    st.session_state.dataset_key = dataset_key
    st.session_state.search_results = pd.DataFrame()
    st.session_state.added_rows = pd.DataFrame(
        columns=["加入數量", *check_df.columns]
    )
    st.session_state.editor_revision = 0

# 相容更新程式前已經存在的瀏覽器 session。
if "editor_revision" not in st.session_state:
    st.session_state.editor_revision = 0

look_up_columns = st.multiselect(
    "查看／比對的欄位",
    list(check_df.columns),
    placeholder="請選擇一個或多個欄位",
)

if look_up_columns:
    with st.form("search_form"):
        st.subheader("輸入查詢條件")
        query_values = {}
        for column in look_up_columns:
            # selectbox 可以直接輸入部分文字，篩選 Excel 中既有的完整值。
            available_values = sorted(
                {
                    str(value).strip()
                    for value in check_df[column].dropna().tolist()
                    if str(value).strip()
                },
                key=str.casefold,
            )
            query_values[column] = st.selectbox(
                f"{column}",
                available_values,
                index=None,
                placeholder="輸入部分文字搜尋，或展開選擇",
                key=f"query_{dataset_key}_{column}",
            )
        search_submitted = st.form_submit_button("查找相同資料", type="primary")

    if search_submitted:
        filled_queries = {
            column: value
            for column, value in query_values.items()
            if value is not None
        }

        if not filled_queries:
            st.warning("請至少輸入一個查詢值。")
            st.session_state.search_results = pd.DataFrame()
        else:
            matches = pd.Series(True, index=check_df.index)
            for column, value in filled_queries.items():
                normalized_column = check_df[column].fillna("").astype(str).str.strip()
                matches &= normalized_column.str.casefold() == value.casefold()

            st.session_state.search_results = check_df.loc[matches].copy()

    search_results = st.session_state.search_results
    if not search_results.empty:
        st.success(f"找到 {len(search_results)} 筆相同資料。")

        # 用查詢欄位內容作為選項標籤，實際選取值則使用結果中的位置。
        result_positions = list(range(len(search_results)))

        def format_result(position: int) -> str:
            return f"第 {position + 1} 筆"

        with st.form("add_row_form"):
            show_df = search_results.copy()
            show_df["第 n 筆"] = range(1,len(show_df)+1,1)
            show_df = show_df[["第 n 筆"] + list(search_results.columns)]
            st.dataframe(
                show_df,
                use_container_width=True,
                hide_index=True,
            )
            selected_position = st.selectbox(
                "選擇要加入的產品",
                result_positions,
                format_func=format_result,
            )
            selected_number = st.number_input(
                "加入數量",
                min_value=1,
                value=1,
                step=1,
            )
            add_submitted = st.form_submit_button("加入清單", type="primary")
        if add_submitted:
            selected_row = search_results.iloc[[selected_position]].copy()

            # 僅將 Excel 中真正的數值型欄位乘上加入數量。

            selected_row.insert(0, "數量", selected_number)

            st.session_state.added_rows = pd.concat(
                [st.session_state.added_rows, selected_row],
                ignore_index=True,
            )
            st.session_state.editor_revision += 1
            st.success("已加入清單。你可以繼續查詢並加入其他資料。")
    elif search_submitted:
        st.warning("找不到符合所有輸入條件的資料。")
else:
    st.info("請先選擇要比對的欄位。")

st.divider()
st.subheader(f"已加入的資料（{len(st.session_state.added_rows)} 筆）")

if st.session_state.added_rows.empty:
    st.info("目前尚未加入任何資料。")
else:
    available_display_columns = list(st.session_state.added_rows.columns)
    display_columns = ["產品名稱","產品編號","規格型號","單位","數量","建議售價"]
    display_names = ["品名","型號","規格","單位","數量","單價"]

    rename_dict = dict(zip(display_columns, display_names))

    added_rows_display = st.session_state.added_rows.loc[:, display_columns].copy()
    added_rows_display = added_rows_display.rename(columns=rename_dict)
    added_rows_display["金額"] = added_rows_display["單價"] * added_rows_display["數量"]
    added_rows_display.insert(len(added_rows_display.columns), "刪除", False)

    edited_rows = st.data_editor(
        added_rows_display,
        use_container_width=True,
        hide_index=True,
        disabled=[column for column in added_rows_display.columns if column != "刪除"],
        column_config={
            "刪除": st.column_config.CheckboxColumn(
                "刪除",
                help="勾選要從清單移除的資料",
                default=False,
            )
        },
        key=(
            f"added_rows_editor_{dataset_key}_"
        ),
    )

    selected_for_deletion = edited_rows["刪除"].fillna(False).astype(bool)
    if st.button(
        "刪除勾選資料",
        disabled=not selected_for_deletion.any(),
    ):
        st.session_state.added_rows = (
            st.session_state.added_rows.loc[~selected_for_deletion.to_numpy()]
            .reset_index(drop=True)
        )
        st.session_state.editor_revision += 1
        st.rerun()
