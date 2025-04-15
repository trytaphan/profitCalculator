import pandas as pd
import streamlit as st
import openpyxl


def read_excel_data(_file_path, sheet):
    data = pd.read_excel(_file_path, sheet_name=sheet)[["记账日期", "对方户名", "金额", "备注"]]
    data = data.sort_values(by="记账日期", ignore_index=True)
    return data


def match_payments(_payments, _collections):
    i = 0  # 付款指针
    j = 0  # 回款指针
    results = []
    payments_copy = _payments.copy()
    collections_copy = _collections.copy()
    payments_copy["剩余金额"] = payments_copy["金额"]
    collections_copy ["剩余金额"] = collections_copy ["金额"]

    while i < len(payments_copy) and j < len(collections_copy):
        # 计算匹配金额
        match_amount = min(payments_copy.loc[i, "剩余金额"], collections_copy.loc[j, "剩余金额"])

        # 更新剩余金额
        payments_copy.loc[i, "剩余金额"] -= match_amount
        collections_copy.loc[j, "剩余金额"] -= match_amount

        # 记录结果
        results.append({"付款索引":i,
                        "回款索引":j,
                        "拆分金额":match_amount})

        #移动指针
        if payments_copy.loc[i, "剩余金额"] < 1e-2:
            i += 1
        if collections_copy.loc[j, "剩余金额"] < 1e-2:
            j += 1
    return results

def result_to_df(_payments, _collections, _match_result):
    result_df = pd.DataFrame(_match_result)
    payments_renamed = _payments.rename(columns={"记账日期":"付款日期",
                                                 "对方户名":"付款对象",
                                                 "金额":"付款金额",
                                                 "备注": "付款备注"})
    collections_renamed = _collections.rename(columns={"记账日期":"回款日期",
                                                 "对方户名":"回款客户",
                                                 "金额":"回款金额",
                                                 "备注": "回款备注"})
    result_df = (result_df. merge(right = payments_renamed,
                                left_on="付款索引",
                                right_index=True).
                            merge(right = collections_renamed,
                                left_on="回款索引",
                                right_index=True))
    return result_df

def decorate_df(_df):
    decorated_df = _df.copy()
    decorated_df.loc[_df.duplicated(subset=["付款索引"], keep="first"), "付款金额"] = ""
    decorated_df.loc[_df.duplicated(subset=["回款索引"], keep="first"), "回款金额"] = ""
    column_list = ["付款日期", "付款对象", "付款金额", "付款备注", "拆分金额",
                   "回款日期", "回款客户", "回款金额", "回款备注"]
    decorated_df = decorated_df.loc[:, column_list]
    return decorated_df
if __name__ == '__main__':
    uploaded_file = st.file_uploader(label="请上传对账表", type = ["xlsx"])
    if uploaded_file:
        if st.button("开始计算"):
            payments = read_excel_data(uploaded_file, "付款明细表")
            collections = read_excel_data(uploaded_file, "回款明细表")
            match_result = match_payments(payments, collections)
            df = result_to_df(payments, collections, match_result)
            st.dataframe(decorate_df(df))
