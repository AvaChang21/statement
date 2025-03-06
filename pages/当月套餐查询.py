import pandas as pd
import re
import streamlit as st
import io


def read_sheets(df):
    data = {}
    for name in df.sheet_names[:6]:
        df_pre = df.parse(sheet_name=name)
        data[name] = df_pre

    return data


def repack_data(data):
    clean_data = {}

    for k, v in data.items():
        ndf = pd.DataFrame(
            columns=['零售用户名称', '合同开始日期', '合同结束日期', '固定价格比例', '固定价格', '长协比例',
                     '长协低于部分甲方', '长协低于部分乙方',
                     '月竞比例', '月竞低于部分甲方', '月竞低于部分乙方', '代购比例', '代购盈利甲方占比',
                     '代购盈利乙方占比',
                     '混合浮动长协', '混合浮动月竞', '混合浮动代购', '年度长协交易均价', '月度竞价出清价', '代购价格',
                     '固定服务费', '归属', '基准价', '预估电量', '备注'])

        aa = pd.DataFrame(
            columns=['零售用户名称', '合同开始日期', '合同结束日期', '固定价格比例', '固定价格', '长协比例',
                     '长协低于部分甲方', '长协低于部分乙方',
                     '月竞比例', '月竞低于部分甲方', '月竞低于部分乙方', '代购比例', '代购盈利甲方占比',
                     '代购盈利乙方占比',
                     '混合浮动长协', '混合浮动月竞', '混合浮动代购', '年度长协交易均价', '月度竞价出清价', '代购价格',
                     '固定服务费', '归属', '基准价', '预估电量', '备注'])

        bb = pd.DataFrame(
            columns=['零售用户名称', '合同开始日期', '合同结束日期', '固定价格比例', '固定价格', '长协比例',
                     '长协低于部分甲方', '长协低于部分乙方',
                     '月竞比例', '月竞低于部分甲方', '月竞低于部分乙方', '代购比例', '代购盈利甲方占比',
                     '代购盈利乙方占比',
                     '混合浮动长协', '混合浮动月竞', '混合浮动代购', '年度长协交易均价', '月度竞价出清价', '代购价格',
                     '固定服务费', '归属', '基准价', '预估电量', '备注'])

        pattern1 = r'\d+\.\d+'
        if k == '标准套餐':
            v['固定价格电量比例'] = v['固定价格电量比例'].apply(pd.to_numeric, errors='coerce')
            a = v[v['固定价格电量比例'] == 100]
            b = v[v['固定价格电量比例'] == 0]

            aa['零售用户名称'] = a['零售用户名称']
            aa['合同开始日期'] = a['合同开始日期']
            aa['合同结束日期'] = a['合同结束日期']
            aa['固定价格'] = a['零售固定成交价']

            bb['零售用户名称'] = b['零售用户名称']
            bb['合同开始日期'] = b['合同开始日期']
            bb['合同结束日期'] = b['合同结束日期']
            bb['长协比例'] = b['长协比例']
            bb['长协低于部分甲方'] = b['长协盈利甲方占比']
            bb['长协低于部分乙方'] = b['长协盈利乙方占比']

            bb['月竞比例'] = b['月竞比例']
            bb['月竞低于部分甲方'] = b['月度盈利甲方占比']
            bb['月竞低于部分乙方'] = b['月度盈利乙方占比']

            bb['基准价'] = b['分成基准价'].str.extract(r'(\d+\.\d+)')

            ndf = pd.concat([aa, bb], ignore_index=True)

            # ndf['基准价'] = v['分成基准价'].apply(lambda x: re.findall(pattern1, x)[0])
            # ndf['基准价'] = v['分成基准价'].str.extract(r'(\d+\.\d+)')

            clean_data[k] = ndf

        if k == '固定手续费':
            ndf['零售用户名称'] = v['零售用户名称']
            ndf['合同开始日期'] = v['合同开始日期']
            ndf['合同结束日期'] = v['合同结束日期']
            ndf['长协比例'] = v['长协比例']
            ndf['月竞比例'] = v['月竞比例']
            ndf['年度长协交易均价'] = v['年度长协交易均价']
            ndf['月度竞价出清价'] = v['月度竞价出清价']

            clean_data[k] = ndf

        if k == '分成加保底':
            ndf['零售用户名称'] = v['零售用户名称']
            ndf['合同开始日期'] = v['合同开始日期']
            ndf['合同结束日期'] = v['合同结束日期']
            ndf['固定价格'] = v['零售固定成交价/相对让利价']

            clean_data[k] = ndf

        if k == '固定价格类':
            ndf['零售用户名称'] = v['零售用户名称']
            ndf['合同开始日期'] = v['合同开始日期']
            ndf['合同结束日期'] = v['合同结束日期']
            ndf['固定价格'] = v['固定价格']

            clean_data[k] = ndf

        if k == '价格浮动类':
            ndf['零售用户名称'] = v['零售用户名称']
            ndf['合同开始日期'] = v['合同开始日期']
            ndf['合同结束日期'] = v['合同结束日期']

            ndf['长协比例'] = v['价格浮动类长协占比']
            ndf['月竞比例'] = v['价格浮动类月竞占比']
            ndf['代购比例'] = v['价格浮动类代购占比']

            ndf['年度长协交易均价'] = v['价格浮动类长协价格']
            ndf['月度竞价出清价'] = v['价格浮动类月竞价格']
            ndf['代购价格'] = v['价格浮动类代购价格']

            clean_data[k] = ndf

        if k == '比例分成类':
            ndf['零售用户名称'] = v['零售用户名称']
            ndf['合同开始日期'] = v['合同开始日期']
            ndf['合同结束日期'] = v['合同结束日期']

            # ndf['基准价'] = v['分成参考价'].apply(lambda x: re.findall(pattern1, x)[0])
            ndf['基准价'] = v['分成参考价'].str.extract(r'(\d+\.\d+)')

            ndf['长协比例'] = v['分成长协占比']
            ndf['长协低于部分甲方'] = v['分成长协盈利甲方占比']
            ndf['长协低于部分乙方'] = v['分成长协盈利乙方占比']

            ndf['月竞比例'] = v['分成月竞占比']
            ndf['月竞低于部分甲方'] = v['分成月竞盈利甲方占比']
            ndf['月竞低于部分乙方'] = v['分成月竞盈利乙方占比']

            ndf['代购比例'] = v['分成代购占比']
            ndf['代购盈利甲方占比'] = v['分成代购盈利甲方占比']
            ndf['代购盈利乙方占比'] = v['分成代购盈利乙方占比']

            clean_data[k] = ndf

        if k == '混合套餐（固定价格类，价格浮动类）':
            ndf['零售用户名称'] = v['零售用户名称']
            ndf['合同开始日期'] = v['合同开始日期']
            ndf['合同结束日期'] = v['合同结束日期']

            ndf['固定价格比例'] = v['固定价格类比例']
            ndf['固定价格'] = v['固定价格']

            ndf['长协比例'] = v['价格浮动类长协占比']
            ndf['月竞比例'] = v['价格浮动类月竞占比']
            ndf['代购比例'] = v['价格浮动类代购占比']

            ndf['年度长协交易均价'] = v['价格浮动类长协价格']
            ndf['月度竞价出清价'] = v['价格浮动类月竞价格']
            ndf['代购价格'] = v['价格浮动类代购价格']

            clean_data[k] = ndf

        if k == '混合套餐（固定价格类，比例分成类）':
            ndf['零售用户名称'] = v['零售用户名称']
            ndf['合同开始日期'] = v['合同开始日期']
            ndf['合同结束日期'] = v['合同结束日期']

            ndf['固定价格比例'] = v['固定价格类比例']
            ndf['固定价格'] = v['固定价格']

            # ndf['基准价'] = v['分成参考价'].apply(lambda x: re.findall(pattern1, x)[0])
            ndf['基准价'] = v['分成参考价'].str.extract(r'(\d+\.\d+)')

            ndf['长协比例'] = v['分成长协占比']
            ndf['长协低于部分甲方'] = v['分成长协盈利甲方占比']
            ndf['长协低于部分乙方'] = v['分成长协盈利乙方占比']

            ndf['月竞比例'] = v['分成月竞占比']
            ndf['月竞低于部分甲方'] = v['分成月竞盈利甲方占比']
            ndf['月竞低于部分乙方'] = v['分成月竞盈利乙方占比']

            ndf['代购比例'] = v['分成代购占比']
            ndf['代购盈利甲方占比'] = v['分成代购盈利甲方占比']
            ndf['代购盈利乙方占比'] = v['分成代购盈利乙方占比']

            clean_data[k] = ndf

        if k == '混合套餐（价格浮动类，比例分成类）':
            ndf['零售用户名称'] = v['零售用户名称']
            ndf['合同开始日期'] = v['合同开始日期']
            ndf['合同结束日期'] = v['合同结束日期']

            ndf['混合浮动长协'] = v['价格浮动类长协占比']
            ndf['混合浮动月竞'] = v['价格浮动类月竞占比']
            ndf['混合浮动代购'] = v['价格浮动类代购占比']

            ndf['年度长协交易均价'] = v['价格浮动类长协价格']
            ndf['月度竞价出清价'] = v['价格浮动类月竞价格']
            ndf['代购价格'] = v['价格浮动类代购价格']

            # ndf['基准价'] = v['分成参考价'].apply(lambda x: re.findall(pattern1, x)[0])
            ndf['基准价'] = v['分成参考价'].str.extract(r'(\d+\.\d+)')

            ndf['长协比例'] = v['分成长协占比']
            ndf['长协低于部分甲方'] = v['分成长协盈利甲方占比']
            ndf['长协低于部分乙方'] = v['分成长协盈利乙方占比']

            ndf['月竞比例'] = v['分成月竞占比']
            ndf['月竞低于部分甲方'] = v['分成月竞盈利甲方占比']
            ndf['月竞低于部分乙方'] = v['分成月竞盈利乙方占比']

            ndf['代购比例'] = v['分成代购占比']
            ndf['代购盈利甲方占比'] = v['分成代购盈利甲方占比']
            ndf['代购盈利乙方占比'] = v['分成代购盈利乙方占比']

            clean_data[k] = ndf

        if k == '混合套餐（固定价格类，价格浮动类，比例分成类）':
            ndf['零售用户名称'] = v['零售用户名称']
            ndf['合同开始日期'] = v['合同开始日期']
            ndf['合同结束日期'] = v['合同结束日期']
            ndf['固定价格比例'] = v['固定价格类比例']
            ndf['固定价格'] = v['固定价格']

            ndf['混合浮动长协'] = v['价格浮动类长协占比']
            ndf['混合浮动月竞'] = v['价格浮动类月竞占比']
            ndf['混合浮动代购'] = v['价格浮动类代购占比']

            ndf['年度长协交易均价'] = v['价格浮动类长协价格']
            ndf['月度竞价出清价'] = v['价格浮动类月竞价格']
            ndf['代购价格'] = v['价格浮动类代购价格']

            ndf['基准价'] = v['分成参考价'].apply(lambda x: re.findall(pattern1, x)[0])

            ndf['长协比例'] = v['分成长协占比']
            ndf['长协低于部分甲方'] = v['分成长协盈利甲方占比']
            ndf['长协低于部分乙方'] = v['分成长协盈利乙方占比']

            ndf['月竞比例'] = v['分成月竞占比']
            ndf['月竞低于部分甲方'] = v['分成月竞盈利甲方占比']
            ndf['月竞低于部分乙方'] = v['分成月竞盈利乙方占比']

            ndf['代购比例'] = v['分成代购占比']
            ndf['代购盈利甲方占比'] = v['分成代购盈利甲方占比']
            ndf['代购盈利乙方占比'] = v['分成代购盈利乙方占比']

            clean_data[k] = ndf

    return clean_data


def find_date(df, year_month):
    df['合同开始日期'] = pd.to_datetime(df['合同开始日期']).dt.date
    df['合同结束日期'] = pd.to_datetime(df['合同结束日期']).dt.date

    start_date = pd.to_datetime(year_month + '-01').date()  # '2025-03-01'
    end_date = pd.to_datetime(year_month + '-01') + pd.offsets.MonthEnd(0)  # '2025-03-31'
    end_date = end_date.date()

    filtered_df = df[(df['合同开始日期'] <= start_date) & (df['合同结束日期'] >= end_date)]

    return filtered_df


uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx", "xls"])

rename_file = st.file_uploader('如有用户更名，请上传表格', type=["xlsx"])

rename = pd.DataFrame(columns=['零售用户名称', '套餐表格原称'])
rename.loc[0] = ['新名称', '原名称']
st.write('更名表模板')
st.write(rename)

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names

    data = {}
    for sheet in sheet_names:
        data[sheet] = pd.read_excel(xls, sheet_name=sheet)

    clean_data = repack_data(data)

    a = pd.concat(clean_data)

    nan_value = float("NaN")
    a.replace("", nan_value, inplace=True)
    a.dropna(how='all', axis=1, inplace=True)

    if rename_file:
        rn = pd.read_excel(rename_file)
        mapping_dict = dict(zip(rn['套餐表格原称'], rn['零售用户名称']))
        st.write(mapping_dict)
        a['零售用户名称'] = a['零售用户名称'].map(mapping_dict).fillna(a['零售用户名称'])
        # a['零售用户名称'] = a['零售用户名称'].replace(mapping_dict)

    year_month = st.text_input('请输入当前年月, e.g. 2025-03', '')
    if year_month:
        df = find_date(a, year_month)

        # 将 DataFrame 写入 Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")
        output.seek(0)  # 重要：重置指针位置

        # 显示下载按钮
        st.download_button(
            label="📥 下载整理后的套餐文件",
            data=output,
            file_name=f"{year_month}套餐.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
