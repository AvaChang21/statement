import pandas as pd
import streamlit as st

@st.cache_data
def convert_df(df):
    # IMPORTANT: Cache the conversion to prevent computation on every rerun
    return df.to_csv().encode("utf-8")

uploaded_file = st.file_uploader("选择 CSV 文件：")
if uploaded_file is not None:
    dataframe = pd.read_csv(uploaded_file, encoding='gbk')
    df = dataframe.groupby('用户名称')['电量(MWh)'].sum()

    csv = convert_df(df)

    st.download_button(
        label="下载汇总后的数据",
        data=csv,
        file_name="new_data.csv",
        mime="text/csv",
    )
