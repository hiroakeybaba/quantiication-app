import streamlit as st
import pandas as pd
from io import BytesIO

st.title("LC/MS 定量アプリ（血漿・糞便対応版）")

id_col = "データファイル名"
INTERNAL_STANDARD_KEYWORD = "d4"

meta_file = st.file_uploader("メタデータExcel", type=["xlsx"])
sample_file = st.file_uploader("サンプルデータExcel", type=["xlsx"])
cal_file = st.file_uploader("検量線Excel", type=["xlsx"])

sample_type = st.radio(
    "サンプル種別",
    ["血漿", "糞便"],
    horizontal=True
)

if meta_file and sample_file and cal_file:

    # ======================
    # 読み込み
    # ======================
    meta = pd.read_excel(meta_file)
    sample = pd.read_excel(sample_file)
    cal = pd.read_excel(cal_file)

    st.success("読み込み完了")

    # ======================
    # 列整理
    # ======================
    for df_tmp in [meta, sample, cal]:
        df_tmp.columns = df_tmp.columns.str.strip()

    sample[id_col] = sample[id_col].astype(str)
    meta[id_col] = meta[id_col].astype(str)

    # ======================
    # wide → long
    # ======================
    area_cols = [c for c in sample.columns if "面積" in str(c)]

    if len(area_cols) == 0:
        st.error("面積列が見つかりません")
        st.stop()

    sample = sample.melt(
        id_vars=[id_col],
        value_vars=area_cols,
        var_name="Compound",
        value_name="Area"
    )

    sample["Compound"] = (
        sample["Compound"]
        .str.replace(" : 面積", "", regex=False)
        .str.strip()
    )

    sample = sample.dropna(subset=["Area"])

    # ======================
    # 内標
    # ======================
    is_list = [
        c for c in sample["Compound"].unique()
        if INTERNAL_STANDARD_KEYWORD in str(c)
    ]

    if len(is_list) == 0:
        st.error("内標(d4)が見つかりません")
        st.stop()

    is_name = is_list[0]

    is_area = sample.loc[
        sample["Compound"] == is_name,
        [id_col, "Area"]
    ].copy()

    is_area = (
        is_area
        .groupby(id_col, as_index=False)["Area"]
        .mean()
        .rename(columns={"Area": "IS_Area"})
    )

    sample = sample.merge(
        is_area,
        on=id_col,
        how="left"
    )

    # ======================
    # response
    # ======================
    sample["response"] = sample["Area"] / sample["IS_Area"]

    # ======================
    # blank補正
    # ======================
    blank = sample[
        sample[id_col].str.contains(
            "blank",
            case=False,
            na=False
        )
    ]

    blank_mean = (
        blank.groupby("Compound")["response"]
        .mean()
        .reset_index()
        .rename(columns={"response": "blank_response"})
    )

    sample = sample.merge(
        blank_mean,
        on="Compound",
        how="left"
    )

    sample["blank_response"] = sample["blank_response"].fillna(0)

    sample["response_corrected"] = (
        sample["response"]
        - sample["blank_response"]
    ).clip(lower=0)

    # ======================
    # merge
    # ======================
    df = sample.merge(
        cal,
        on="Compound",
        how="left"
    )

    df = df.merge(
        meta,
        on=id_col,
        how="left"
    )

    # ======================
    # 検量線
    # ======================
    df["Conc_raw"] = (
        df["response_corrected"]
        - df["Intercept"]
    ) / df["Slope"]

    # ======================
    # pg/uL算出
    # ======================
    df["Conc_pg_ul"] = (
        df["Conc_raw"]
        / df["インジェクション量"]
        * df["溶出量"]
        / df["カラム抽出持ち込み量"]
        * (df["サンプル量"] + df["溶媒添加量"])
        / df["サンプル量"]
    )

    # ======================
    # 血漿 / 糞便
    # ======================
    if sample_type == "血漿":

        df["Conc_final"] = df["Conc_pg_ul"]
        df["LLOQ_display"] = df["LLOQ"]
        final_unit = "pg/µL"

    else:

        df["Conc_final"] = (
            df["Conc_pg_ul"]
            * df["溶媒添加量"]
            / df["サンプル量"]
        )

        df["LLOQ_display"] = (
            df["LLOQ"]
            * df["溶媒添加量"]
            / df["サンプル量"]
        )

        final_unit = "pg/mg"

    # ======================
    # LOQ / ULOQ
    # ======================
    df["LOQ_flag"] = (
        df["response_corrected"]
        < df["LLOQ_response"]
    )

    df["ULOQ_flag"] = (
        df["response_corrected"]
        > df["ULOQ_response"]
    )

    df["Conc_use"] = df["Conc_final"].where(
        ~df["LOQ_flag"],
        df["LLOQ_display"]
    )

    # ======================
    # QC
    # ======================
    is_df = df[df["Compound"] == is_name]

    ref = is_df[
        is_df[id_col].str.contains(
            "抽出効率",
            na=False
        )
    ]

    ref_mean = (
        ref["Area"]
        / ref["インジェクション量"]
    ).mean()

    samp = is_df[
        ~is_df[id_col].str.contains(
            "抽出効率",
            na=False
        )
    ].copy()

    samp["ratio"] = (
        (samp["Area"] / samp["インジェクション量"])
        / ref_mean
    )

    samp["QC_flag"] = (
        (samp["ratio"] < 0.5)
        | (samp["ratio"] > 1.5)
    )

    df = df.merge(
        samp[[id_col, "QC_flag"]],
        on=id_col,
        how="left"
    )

    df["QC_flag"] = df["QC_flag"].fillna(False)

    # ======================
    # 表示用結果
    # ======================
    def format_row(r):

        if r["ULOQ_flag"]:
            return "ERROR (>ULOQ)"

        elif r["QC_flag"]:
            return "ERROR(IS QC)"

        elif r["LOQ_flag"]:
            return f"<{round(r['LLOQ_display'], 3)} {final_unit}"

        else:
            return f"{round(r['Conc_use'], 3)} {final_unit}"

    df["Result_display"] = df.apply(
        format_row,
        axis=1
    )

    # ======================
    # 画面表示
    # ======================
    st.subheader("結果（表示）")

    st.dataframe(
        df[
            [
                id_col,
                "Compound",
                "Result_display"
            ]
        ]
    )

    # ======================
    # Result①
    # ======================
    result_with_loq = df.pivot_table(
        index=id_col,
        columns="Compound",
        values="Result_display",
        aggfunc="first"
    ).reset_index()

    # ======================
    # Result②
    # ======================
    result_numeric = df.copy()

    result_numeric.loc[
        (
            result_numeric["LOQ_flag"]
            | result_numeric["ULOQ_flag"]
            | result_numeric["QC_flag"]
        ),
        "Conc_use"
    ] = pd.NA

    result_numeric_only = result_numeric.pivot_table(
        index=id_col,
        columns="Compound",
        values="Conc_use",
        aggfunc="first"
    ).reset_index()

    # ======================
    # 画面表示
    # ======================
    st.subheader("結果（<LLOQ表示あり）")
    st.dataframe(result_with_loq)

    st.subheader("結果（数値のみ）")
    st.dataframe(result_numeric_only)

    # ======================
    # Excel出力
    # ======================
    buffer = BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Long"
        )

        result_with_loq.to_excel(
            writer,
            index=False,
            sheet_name="Result_with_LOQ"
        )

        result_numeric_only.to_excel(
            writer,
            index=False,
            sheet_name="Result_numeric_only"
        )

    buffer.seek(0)

    st.download_button(
        "Excel DL",
        buffer,
        "result.xlsx"
    )
