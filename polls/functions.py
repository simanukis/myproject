import pandas as pd
import numpy as np

# from pandas.io.common import codecs

# ファイルを指定してデータフレーム化
df_jinji = pd.read_excel("/content/drive/MyDrive/data/jinji.XLS")
df_jinji = df_jinji.drop(columns=["社員名", "集計区分－１        "])
df_syain = pd.read_excel("/content/drive/MyDrive/data/syain.XLS")
df_syain = df_syain.drop(
    columns=[
        "氏 名",
        "役職ｺｰﾄﾞ",
        "役職",
        "所属分類1名",
        "所属分類2名",
        "所属分類3名",
        "所 .3",
        "所属分類4名",
        "所 .4",
        "所属分類5名",
        "備考",
        "在籍区分",
        "在籍",
        "退職年",
        "退職月",
        "退職日",
    ]
)
# 人事ファイルと社員ファイルを結合してデータフレーム化
df_m = pd.merge(df_jinji, df_syain, how="inner", on="社員ｺｰﾄﾞ")
# 役員データフレームの作成
df_yakuin_kbn = df_m[(df_m["所 "] == 0) & (df_m["所 .1"] == 0) & (df_m["所 .2"] == 0)]
df_yakuin_kbn = df_yakuin_kbn.astype("object")
df_yakuin_kbn["区分"] = "役員1"
df_yakuin_kbn = df_yakuin_kbn.rename(
    columns={"所 ": "所属1", "所 .1": "所属2", "所 .2": "所属3"}
)
# 一般間接1データフレームの作成
df_ipan_kbn103 = df_m[(df_m["集計区分－２        "] == 103)]
df_ipan_kbn103 = df_ipan_kbn103.astype("object")
df_ipan_kbn103["区分"] = "一般間接1"
# 一般間接2データフレームの作成
df_ipan_kbn104 = df_m[(df_m["集計区分－２        "] == 104)]
df_ipan_kbn104 = df_ipan_kbn104.astype("object")
df_ipan_kbn104["区分"] = "一般間接2"
# 一般間接3データフレームの作成
df_ipan_kbn105 = df_m[(df_m["集計区分－２        "] == 105)]
df_ipan_kbn105 = df_ipan_kbn105.astype("object")
df_ipan_kbn105["区分"] = "一般間接3"
# 一般間接6データフレームの作成
df_ipan_kbn108 = df_m[(df_m["集計区分－２        "] == 108)]
df_ipan_kbn108 = df_ipan_kbn108.astype("object")
df_ipan_kbn108["区分"] = "一般間接6"
# 一般販売1データフレームの作成
df_ipan_kbn109 = df_m[(df_m["集計区分－２        "] == 109)]
df_ipan_kbn109 = df_ipan_kbn109.astype("object")
df_ipan_kbn109["区分"] = "一般販売1"
# 一般販売2データフレームの作成
df_ipan_kbn110 = df_m[(df_m["集計区分－２        "] == 110)]
df_ipan_kbn110 = df_ipan_kbn110.astype("object")
df_ipan_kbn110["区分"] = "一般販売2"
# 一般管理データフレームの作成
df_ipan_kbn = pd.concat(
    [
        df_ipan_kbn103,
        df_ipan_kbn104,
        df_ipan_kbn105,
        df_ipan_kbn108,
        df_ipan_kbn109,
        df_ipan_kbn110,
    ],
    axis=0,
    ignore_index=True,
)
df_ipan_kbn = df_ipan_kbn.rename(
    columns={"所 ": "所属1", "所 .1": "所属2", "所 .2": "所属3"}
)
# 鍛造間接1のデータフレーム作成
df_tanzo_kbn211 = df_m[(df_m["集計区分－２        "] == 211)]
df_tanzo_kbn211 = df_tanzo_kbn211.astype("object")
df_tanzo_kbn211["区分"] = "間接1"
# 鍛造間接2のデータフレーム作成
df_tanzo_kbn212 = df_m[(df_m["集計区分－２        "] == 212)]
df_tanzo_kbn212 = df_tanzo_kbn212.astype("object")
df_tanzo_kbn212["区分"] = "間接2"
# 鍛造間接3のデータフレーム作成
df_tanzo_kbn213 = df_m[(df_m["集計区分－２        "] == 213)]
df_tanzo_kbn213 = df_tanzo_kbn213.astype("object")
df_tanzo_kbn213["区分"] = "間接3"
# 鍛造間接4のデータフレーム作成
df_tanzo_kbn214 = df_m[(df_m["集計区分－２        "] == 214)]
df_tanzo_kbn214 = df_tanzo_kbn214.astype("object")
df_tanzo_kbn214["区分"] = "間接4"
# 鍛造間接5のデータフレーム作成
df_tanzo_kbn215 = df_m[(df_m["集計区分－２        "] == 215)]
df_tanzo_kbn215 = df_tanzo_kbn215.astype("object")
df_tanzo_kbn215["区分"] = "間接5"
# 鍛造間接6のデータフレーム作成
df_tanzo_kbn216 = df_m[(df_m["集計区分－２        "] == 216)]
df_tanzo_kbn216 = df_tanzo_kbn216.astype("object")
df_tanzo_kbn216["区分"] = "間接6"
# 鍛造直接1のデータフレーム作成
df_tanzo_kbn218 = df_m[(df_m["集計区分－２        "] == 218)]
df_tanzo_kbn218 = df_tanzo_kbn218.astype("object")
df_tanzo_kbn218["区分"] = "直接1"
# 鍛造直接4のデータフレーム作成
df_tanzo_kbn221 = df_m[(df_m["集計区分－２        "] == 221)]
df_tanzo_kbn221 = df_tanzo_kbn221.astype("object")
df_tanzo_kbn221["区分"] = "直接4"
# 鍛造のデータフレーム作成
df_tanzo_kbn = pd.concat(
    [
        df_tanzo_kbn211,
        df_tanzo_kbn212,
        df_tanzo_kbn213,
        df_tanzo_kbn214,
        df_tanzo_kbn215,
        df_tanzo_kbn216,
        df_tanzo_kbn218,
        df_tanzo_kbn221,
    ],
    axis=0,
    ignore_index=True,
)
df_tanzo_kbn = df_tanzo_kbn.rename(
    columns={"所 ": "所属1", "所 .1": "所属2", "所 .2": "所属3"}
)
# 切削間接1のデータフレーム作成
df_sesaku_kbn311 = df_m[(df_m["集計区分－２        "] == 311)]
df_sesaku_kbn311 = df_sesaku_kbn311.astype("object")
df_sesaku_kbn311["区分"] = "間接1"
# 切削間接2のデータフレーム作成
df_sesaku_kbn312 = df_m[(df_m["集計区分－２        "] == 312)]
df_sesaku_kbn312 = df_sesaku_kbn312.astype("object")
df_sesaku_kbn312["区分"] = "間接2"
# 切削間接4のデータフレーム作成
df_sesaku_kbn314 = df_m[(df_m["集計区分－２        "] == 314)]
df_sesaku_kbn314 = df_sesaku_kbn314.astype("object")
df_sesaku_kbn314["区分"] = "間接4"
# 切削間接5のデータフレーム作成
df_sesaku_kbn315 = df_m[(df_m["集計区分－２        "] == 315)]
df_sesaku_kbn315 = df_sesaku_kbn315.astype("object")
df_sesaku_kbn315["区分"] = "間接5"
# 切削間接6のデータフレーム作成
df_sesaku_kbn316 = df_m[(df_m["集計区分－２        "] == 316)]
df_sesaku_kbn316 = df_sesaku_kbn316.astype("object")
df_sesaku_kbn316["区分"] = "間接6"
# 切削直接1のデータフレーム作成
df_sesaku_kbn318 = df_m[(df_m["集計区分－２        "] == 318)]
df_sesaku_kbn318 = df_sesaku_kbn318.astype("object")
df_sesaku_kbn318["区分"] = "直接1"
# 切削直接2のデータフレーム作成
df_sesaku_kbn319 = df_m[(df_m["集計区分－２        "] == 319)]
df_sesaku_kbn319 = df_sesaku_kbn319.astype("object")
df_sesaku_kbn319["区分"] = "直接2"
# 切削直接4のデータフレーム作成
df_sesaku_kbn321 = df_m[(df_m["集計区分－２        "] == 321)]
df_sesaku_kbn321 = df_sesaku_kbn321.astype("object")
df_sesaku_kbn321["区分"] = "直接4"
# 切削のデータフレーム作成
df_sesaku_kbn = pd.concat(
    [
        df_sesaku_kbn311,
        df_sesaku_kbn312,
        df_sesaku_kbn314,
        df_sesaku_kbn315,
        df_sesaku_kbn316,
        df_sesaku_kbn318,
        df_sesaku_kbn319,
        df_sesaku_kbn321,
    ],
    axis=0,
    ignore_index=True,
)
df_sesaku_kbn = df_sesaku_kbn.rename(
    columns={"所 ": "所属1", "所 .1": "所属2", "所 .2": "所属3"}
)
# AC間接1のデータフレーム作成
df_ac_kbn411 = df_m[(df_m["集計区分－２        "] == 411)]
df_ac_kbn411 = df_ac_kbn411.astype("object")
df_ac_kbn411["区分"] = "間接1"
# AC間接2のデータフレーム作成
df_ac_kbn412 = df_m[(df_m["集計区分－２        "] == 412)]
df_ac_kbn412 = df_ac_kbn412.astype("object")
df_ac_kbn412["区分"] = "間接2"
# AC間接4のデータフレーム作成
df_ac_kbn414 = df_m[(df_m["集計区分－２        "] == 414)]
df_ac_kbn414 = df_ac_kbn414.astype("object")
df_ac_kbn414["区分"] = "間接4"
# AC間接5のデータフレーム作成
df_ac_kbn415 = df_m[(df_m["集計区分－２        "] == 415)]
df_ac_kbn415 = df_ac_kbn415.astype("object")
df_ac_kbn415["区分"] = "間接5"
# AC直接1のデータフレーム作成
df_ac_kbn418 = df_m[(df_m["集計区分－２        "] == 418)]
df_ac_kbn418 = df_ac_kbn418.astype("object")
df_ac_kbn418["区分"] = "直接1"
# AC直接4のデータフレーム作成
df_ac_kbn421 = df_m[(df_m["集計区分－２        "] == 421)]
df_ac_kbn421 = df_ac_kbn421.astype("object")
df_ac_kbn421["区分"] = "直接4"
# ACのデータフレーム作成
df_ac_kbn = pd.concat(
    [
        df_ac_kbn411,
        df_ac_kbn412,
        df_ac_kbn414,
        df_ac_kbn415,
        df_ac_kbn418,
        df_ac_kbn421,
    ],
    axis=0,
    ignore_index=True,
)
df_ac_kbn = df_ac_kbn.rename(
    columns={"所 ": "所属1", "所 .1": "所属2", "所 .2": "所属3"}
)
# PC間接1のデータフレーム作成
df_pc_kbn511 = df_m[(df_m["集計区分－２        "] == 511)]
df_pc_kbn511 = df_pc_kbn511.astype("object")
df_pc_kbn511["区分"] = "間接1"
# PC間接2のデータフレーム作成
df_pc_kbn512 = df_m[(df_m["集計区分－２        "] == 512)]
df_pc_kbn512 = df_pc_kbn512.astype("object")
df_pc_kbn512["区分"] = "間接2"
# PC間接4のデータフレーム作成
df_pc_kbn514 = df_m[(df_m["集計区分－２        "] == 514)]
df_pc_kbn514 = df_pc_kbn514.astype("object")
df_pc_kbn514["区分"] = "間接4"
# PC間接5のデータフレーム作成
df_pc_kbn515 = df_m[(df_m["集計区分－２        "] == 515)]
df_pc_kbn515 = df_pc_kbn515.astype("object")
df_pc_kbn515["区分"] = "間接5"
# PC間接6のデータフレーム作成
df_pc_kbn516 = df_m[(df_m["集計区分－２        "] == 516)]
df_pc_kbn516 = df_pc_kbn516.astype("object")
df_pc_kbn516["区分"] = "間接6"
# PC直接1のデータフレーム作成
df_pc_kbn518 = df_m[(df_m["集計区分－２        "] == 518)]
df_pc_kbn518 = df_pc_kbn518.astype("object")
df_pc_kbn518["区分"] = "直接1"
# PC直接4のデータフレーム作成
df_pc_kbn521 = df_m[(df_m["集計区分－２        "] == 521)]
df_pc_kbn521 = df_pc_kbn521.astype("object")
df_pc_kbn521["区分"] = "直接4"
# PCのデータフレーム作成
df_pc_kbn = pd.concat(
    [
        df_pc_kbn511,
        df_pc_kbn512,
        df_pc_kbn514,
        df_pc_kbn515,
        df_pc_kbn516,
        df_pc_kbn518,
        df_pc_kbn521,
    ],
    axis=0,
    ignore_index=True,
)
df_pc_kbn = df_pc_kbn.rename(
    columns={"所 ": "所属1", "所 .1": "所属2", "所 .2": "所属3"}
)
# 宮城間接1のデータフレーム作成
df_miyagi_kbn611 = df_m[(df_m["集計区分－２        "] == 611)]
df_miyagi_kbn611 = df_miyagi_kbn611.astype("object")
df_miyagi_kbn611["区分"] = "間接1"
# 宮城間接2のデータフレーム作成
df_miyagi_kbn612 = df_m[(df_m["集計区分－２        "] == 612)]
df_miyagi_kbn612 = df_miyagi_kbn612.astype("object")
df_miyagi_kbn612["区分"] = "間接2"
# 宮城間接4のデータフレーム作成
df_miyagi_kbn614 = df_m[(df_m["集計区分－２        "] == 614)]
df_miyagi_kbn614 = df_miyagi_kbn614.astype("object")
df_miyagi_kbn614["区分"] = "間接4"
# 宮城間接6のデータフレーム作成
df_miyagi_kbn616 = df_m[(df_m["集計区分－２        "] == 616)]
df_miyagi_kbn616 = df_miyagi_kbn616.astype("object")
df_miyagi_kbn616["区分"] = "間接6"
# 宮城直接1のデータフレーム作成
df_miyagi_kbn618 = df_m[(df_m["集計区分－２        "] == 618)]
df_miyagi_kbn618 = df_miyagi_kbn618.astype("object")
df_miyagi_kbn618["区分"] = "直接1"
# 宮城のデータフレーム作成
df_miyagi_kbn = pd.concat(
    [
        df_miyagi_kbn611,
        df_miyagi_kbn612,
        df_miyagi_kbn614,
        df_miyagi_kbn616,
        df_miyagi_kbn618,
    ],
    axis=0,
    ignore_index=True,
)
df_miyagi_kbn = df_miyagi_kbn.rename(
    columns={"所 ": "所属1", "所 .1": "所属2", "所 .2": "所属3"}
)
# 住設間接2のデータフレーム作成
df_jyusetu_kbn712 = df_m[(df_m["集計区分－２        "] == 712)]
df_jyusetu_kbn712 = df_jyusetu_kbn712.astype("object")
df_jyusetu_kbn712["区分"] = "間接2"
# 住設間接4のデータフレーム作成
df_jyusetu_kbn714 = df_m[(df_m["集計区分－２        "] == 714)]
df_jyusetu_kbn714 = df_jyusetu_kbn714.astype("object")
df_jyusetu_kbn714["区分"] = "間接4"
# 住設間接6のデータフレーム作成
df_jyusetu_kbn716 = df_m[(df_m["集計区分－２        "] == 716)]
df_jyusetu_kbn716 = df_jyusetu_kbn716.astype("object")
df_jyusetu_kbn716["区分"] = "間接6"
# 住設のデータフレーム作成
df_jyusetu_kbn = pd.concat(
    [df_jyusetu_kbn712, df_jyusetu_kbn714, df_jyusetu_kbn716], axis=0, ignore_index=True
)
df_jyusetu_kbn = df_jyusetu_kbn.rename(
    columns={"所 ": "所属1", "所 .1": "所属2", "所 .2": "所属3"}
)
# 給与データファイルのデータフレーム化
df_kinsi = pd.read_excel("/content/drive/MyDrive/data/kinsi11-2023.XLS")
df_kinsi = df_kinsi.drop(columns=df_kinsi.columns[[16, 41, 61]], axis=1)
df_kinsi = df_kinsi.dropna(how="any")
df_kinsi.drop(
    labels=[
        "所属4",
        "所属5",
        "社員名",
        "支給日[年]",
        "支給日[月]",
        "支給日[日]",
        "【 勤怠 】  ",
        "深夜時間    ",
        "休業日数    ",
        "【 支給 】  ",
        "課税支給額  ",
        "【 控除 】  ",
        "健康保険    ",
        "厚生年金保険",
        "雇用保険    ",
        "社会保険計  ",
        "(内)介護保険",
        "課税対象額  ",
        "所得税      ",
        "住民税      ",
        "財形預金    ",
        "生命保険    ",
        "自動車保険  ",
        "社員積立    ",
        "労金        ",
        "その他控除  ",
        "控除合計額  ",
        "【 補助 】  ",
        "職能等級    ",
        "基本給１    ",
        "基本給２    ",
        "【 合計 】  ",
        "差引支給額  ",
        "銀行振込１  ",
        "銀行振込２  ",
        "銀行振込３  ",
        "現金支給額  ",
    ],
    axis=1,
    inplace=True,
)
df_kinsi_s = df_kinsi.set_index(["所属1", "所属2", "所属3"])
df_kinsi_s = df_kinsi_s.rename(columns={"            .1": "支給額"})

# 給与データフレームと各区分データフレームを結合する
# 役員
df_yakuin_m = pd.merge(df_yakuin_kbn, df_kinsi_s, how="inner", on="社員ｺｰﾄﾞ")
# 一般管理
df_ipan_m = pd.merge(df_ipan_kbn, df_kinsi_s, how="inner", on="社員ｺｰﾄﾞ")
# 鍛造
df_tanzo_m = pd.merge(df_tanzo_kbn, df_kinsi_s, how="inner", on="社員ｺｰﾄﾞ")
# 切削
df_sesaku_m = pd.merge(df_sesaku_kbn, df_kinsi_s, how="inner", on="社員ｺｰﾄﾞ")
# AC
df_ac_m = pd.merge(df_ac_kbn, df_kinsi_s, how="inner", on="社員ｺｰﾄﾞ")
# PC
df_pc_m = pd.merge(df_pc_kbn, df_kinsi_s, how="inner", on="社員ｺｰﾄﾞ")
# 宮城
df_miyagi_m = pd.merge(df_miyagi_kbn, df_kinsi_s, how="inner", on="社員ｺｰﾄﾞ")
# 住設
df_jyusetu_m = pd.merge(df_jyusetu_kbn, df_kinsi_s, how="inner", on="社員ｺｰﾄﾞ")

# ----集計----
# 一般管理_役員
df_yakuin_1 = df_yakuin_m.groupby("区分").get_group("役員1")
df_yakuin_1 = df_yakuin_1.drop("所属1", axis=1)
member = df_yakuin_1["所属2"] == 0
df_yakuin_1.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_yakuin_1["所属3"] > 0
df_yakuin_1.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_yakuin_1["出勤日数    "]
df_yakuin_1.insert(2, "実在籍者", real_member)
time_yukyu = df_yakuin_1["有休日数    "] * 8
df_yakuin_1.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
df_yakuin_1.insert(4, "有休時間在籍者平均", ave_yukyu.round(2))
abs_time = df_yakuin_1["欠勤日数    "] * 8
df_yakuin_1.insert(5, "欠勤時間", abs_time)
work_time = df_yakuin_1["勤務時間    "]
df_yakuin_1.insert(6, "勤務時間", work_time)
late_early_time = df_yakuin_1["遅早時間    "]
df_yakuin_1.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_yakuin_1["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
df_yakuin_1.insert(16, "出勤率", work_rate / member)
real_work_time = real_member * 8
df_yakuin_1.insert(17, "実労働時間", real_work_time)
zure_time = df_yakuin_1["ズレ時間    "]
df_yakuin_1.insert(18, "ズレ時間", zure_time)
overtime = df_yakuin_1["残業時間    "] + df_yakuin_1["深夜残業    "]
df_yakuin_1.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_yakuin_1.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_yakuin_1["法外休出    "]
df_yakuin_1.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_yakuin_1.insert(27, "法定外休出主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_yakuin_1["法定休出    "]
df_yakuin_1.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_yakuin_1.insert(29, "法定休出主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_yakuin_1["６０Ｈ超    "]
df_yakuin_1.insert(30, "時間外60時間超", overtime_60)
holiday_time = df_yakuin_1["代休時間    "] + df_yakuin_1["深夜代休    "]
df_yakuin_1.insert(31, "代休時間", holiday_time)
df_yakuin_1.insert(32, "応援時間", 0)
total_work_time = df_yakuin_1["勤務時間    "] + df_yakuin_1["残業時間    "]
df_yakuin_1.insert(33, "総労働時間", total_work_time)
basic_salary = df_yakuin_1["基 本 給    "] + df_yakuin_1["支給額"]
df_yakuin_1.insert(35, "基本給", basic_salary)
post_allowance = df_yakuin_1["役職手当    "]
df_yakuin_1.insert(36, "役職手当", post_allowance)
sales_allowance = df_yakuin_1["営業手当    "]
df_yakuin_1.insert(37, "営業手当", sales_allowance)
aria_allowance = df_yakuin_1["地域手当    "]
df_yakuin_1.insert(38, "地域手当", aria_allowance)
spe_allowance = df_yakuin_1["特殊手当    "]
df_yakuin_1.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_yakuin_1["特別技技手当"]
df_yakuin_1.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_yakuin_1["調整手当    "]
df_yakuin_1.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_yakuin_1["別居手当    "]
df_yakuin_1.insert(42, "別居手当", sep_allowance)
com_allowance = df_yakuin_1["通勤手当    "]
df_yakuin_1.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + spe_tec_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_yakuin_1.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_yakuin_1["残業手当    "]
df_yakuin_1.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_yakuin_1["休出手当    "]
df_yakuin_1.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_yakuin_1["深夜勤務手当"]
df_yakuin_1.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_yakuin_1["交替時差手当"]
df_yakuin_1.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_yakuin_1["休業手当    "]
df_yakuin_1.insert(49, "休業手当", closed_allowance)
closed_deduction = df_yakuin_1["休業控除    "]
df_yakuin_1.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_yakuin_1["代 休 他    "]
df_yakuin_1.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_yakuin_1["欠勤控除    "] + df_yakuin_1["遅早控除    "]
df_yakuin_1.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_yakuin_1["精 算 分    "]
df_yakuin_1.insert(53, "精算分", settlement)
sub_total_2 = (
    overtime_allowance
    + vacation_allowance
    + night_work_allowance
    + time_difference_allowance
    + closed_allowance
    + closed_deduction
    + compny_leave_etc
    + abs_early_deduction
    + settlement
)
df_yakuin_1.insert(54, "小計 2", sub_total_2)
total = sub_total_1 + sub_total_2
df_yakuin_1.insert(55, "総支給額", total)
df_yakuin_1.insert(56, "応援時間額", 0)
df_yakuin_1.insert(57, "役員振替", 0)
df_yakuin_1.insert(58, "部門振替", 0)
df_yakuin_1.insert(59, "合計", 0)

# 不要フィールドの削除
df_yakuin_1 = df_yakuin_1.drop("所属2", axis=1)
df_yakuin_1 = df_yakuin_1.drop("所属3", axis=1)
df_yakuin_1 = df_yakuin_1.drop("社員ｺｰﾄﾞ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("区分", axis=1)
df_yakuin_1 = df_yakuin_1.drop("出勤日数    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("有休日数    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("欠勤日数    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("残業時間    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("深夜残業    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("法外休出    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("法定休出    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("代休時間    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("深夜代休    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("６０Ｈ超    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("別居手当    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("通勤手当    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("特別技技手当", axis=1)
df_yakuin_1 = df_yakuin_1.drop("特殊手当    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("地域手当    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("営業手当    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("役職手当    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("調整手当    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("基 本 給    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("残業手当    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("休出手当    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("深夜勤務手当", axis=1)
df_yakuin_1 = df_yakuin_1.drop("交替時差手当", axis=1)
df_yakuin_1 = df_yakuin_1.drop("休業手当    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("代 休 他    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("欠勤控除    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("遅早控除    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("精 算 分    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("支給合計額  ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("休業控除    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("支給額", axis=1)
df_yakuin_1 = df_yakuin_1.drop("ズレ時間    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("雑費・食事代", axis=1)
df_yakuin_1 = df_yakuin_1.drop("雑費・衣靴代", axis=1)
df_yakuin_1 = df_yakuin_1.drop("雑費        ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("受診料・他  ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("雑費・会費等", axis=1)
df_yakuin_1 = df_yakuin_1.drop("勤務時間    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("遅早時間    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("特休日数    ", axis=1)
df_yakuin_1 = df_yakuin_1.drop("集計区分－２        ", axis=1)
df_yakuin_1 = df_yakuin_1.sum()
# 仮CSVファイルの出力（Excel出力のため)
df_yakuin_1.to_csv(
    "/content/drive/MyDrive/data/一般管理/A.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 一般管理_一般間接1
df_ipan_1 = df_ipan_m.groupby("区分").get_group("一般間接1")
df_ipan_1 = df_ipan_1.drop("所属1", axis=1)
member = df_ipan_1["所属2"] > 0
df_ipan_1.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_ipan_1["所属3"] > 0
df_ipan_1.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_ipan_1["出勤日数    "]
df_ipan_1.insert(2, "実在籍者", real_member)
time_yukyu = df_ipan_1["有休日数    "] * 8
df_ipan_1.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_ipan_1.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_ipan_1["欠勤日数    "] * 8
df_ipan_1.insert(5, "欠勤時間", abs_time)
work_time = df_ipan_1["勤務時間    "]
df_ipan_1.insert(6, "勤務時間", work_time)
late_early_time = df_ipan_1["遅早時間    "]
df_ipan_1.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_ipan_1["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_ipan_1.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_ipan_1.insert(17, "実労働時間", real_work_time)
zure_time = df_ipan_1["ズレ時間    "]
df_ipan_1.insert(18, "ズレ時間", zure_time)
overtime = df_ipan_1["残業時間    "] + df_ipan_1["深夜残業    "]
df_ipan_1.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_ipan_1.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_ipan_1["法外休出    "]
df_ipan_1.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_ipan_1.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_ipan_1["法定休出    "]
df_ipan_1.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_ipan_1.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_ipan_1["６０Ｈ超    "]
df_ipan_1.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_ipan_1["代休時間    "] + df_ipan_1["深夜代休    "]
df_ipan_1.insert(31, "代休時間", holiday_time)
df_ipan_1.insert(32, "応援時間", 0)
total_work_time = (
    df_ipan_1["勤務時間    "]
    + df_ipan_1["残業時間    "]
    + df_ipan_1["法外休出    "]
    + df_ipan_1["法定休出    "]
)
df_ipan_1.insert(33, "総労働時間", total_work_time)
basic_salary = df_ipan_1["基 本 給    "] + df_ipan_1["支給額"]
df_ipan_1.insert(35, "基本給", basic_salary)
post_allowance = df_ipan_1["役職手当    "]
df_ipan_1.insert(36, "役職手当", post_allowance)
sales_allowance = df_ipan_1["営業手当    "]
df_ipan_1.insert(37, "営業手当", sales_allowance)
aria_allowance = df_ipan_1["地域手当    "]
df_ipan_1.insert(38, "地域手当", aria_allowance)
spe_allowance = df_ipan_1["特殊手当    "]
df_ipan_1.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_ipan_1["特別技技手当"]
df_ipan_1.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_ipan_1["調整手当    "]
df_ipan_1.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_ipan_1["別居手当    "]
df_ipan_1.insert(42, "別居手当", sep_allowance)
com_allowance = df_ipan_1["通勤手当    "]
df_ipan_1.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + spe_tec_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
sub_total_1 = sub_total_1.sum()
sub_total_1 = sub_total_1 / member
sub_total_1 = sub_total_1.astype(int)
df_ipan_1.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_ipan_1["残業手当    "]
df_ipan_1.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_ipan_1["休出手当    "]
df_ipan_1.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_ipan_1["深夜勤務手当"]
df_ipan_1.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_ipan_1["交替時差手当"]
df_ipan_1.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_ipan_1["休業手当    "]
df_ipan_1.insert(49, "休業手当", closed_allowance)
teate_total = (
    overtime_allowance
    + vacation_allowance
    + night_work_allowance
    + time_difference_allowance
    + closed_allowance
)
teate_total = teate_total.sum()
closed_deduction = df_ipan_1["休業控除    "]
df_ipan_1.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_ipan_1["代 休 他    "]
df_ipan_1.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_ipan_1["欠勤控除    "] + df_ipan_1["遅早控除    "]
df_ipan_1.insert(52, "欠勤・遅早控除", abs_early_deduction)
kojyo_total = closed_deduction + compny_leave_etc + abs_early_deduction
kojyo_total = kojyo_total.sum()
settlement = df_ipan_1["精 算 分    "]
df_ipan_1.insert(53, "精算分", settlement)
sub_total_2 = (
    overtime_allowance
    + vacation_allowance
    + night_work_allowance
    + time_difference_allowance
    + closed_allowance
    + closed_deduction
    + compny_leave_etc
    + abs_early_deduction
    + settlement
)
df_ipan_1.insert(54, "小計 2", sub_total_2)
total = (sub_total_1 + sub_total_2) - kojyo_total
df_ipan_1.insert(55, "総支給額", total)
df_ipan_1.insert(56, "応援時間額", 0)
df_ipan_1.insert(57, "役員振替", 0)
df_ipan_1.insert(58, "部門振替", 0)
df_ipan_1.insert(59, "合計", 0)
# 不要フィールドの削除
df_ipan_1 = df_ipan_1.drop("所属2", axis=1)
df_ipan_1 = df_ipan_1.drop("所属3", axis=1)
df_ipan_1 = df_ipan_1.drop("社員ｺｰﾄﾞ", axis=1)
df_ipan_1 = df_ipan_1.drop("区分", axis=1)
df_ipan_1 = df_ipan_1.drop("出勤日数    ", axis=1)
df_ipan_1 = df_ipan_1.drop("有休日数    ", axis=1)
df_ipan_1 = df_ipan_1.drop("欠勤日数    ", axis=1)
df_ipan_1 = df_ipan_1.drop("残業時間    ", axis=1)
df_ipan_1 = df_ipan_1.drop("深夜残業    ", axis=1)
df_ipan_1 = df_ipan_1.drop("法外休出    ", axis=1)
df_ipan_1 = df_ipan_1.drop("法定休出    ", axis=1)
df_ipan_1 = df_ipan_1.drop("代休時間    ", axis=1)
df_ipan_1 = df_ipan_1.drop("深夜代休    ", axis=1)
df_ipan_1 = df_ipan_1.drop("６０Ｈ超    ", axis=1)
df_ipan_1 = df_ipan_1.drop("別居手当    ", axis=1)
df_ipan_1 = df_ipan_1.drop("通勤手当    ", axis=1)
df_ipan_1 = df_ipan_1.drop("特別技技手当", axis=1)
df_ipan_1 = df_ipan_1.drop("特殊手当    ", axis=1)
df_ipan_1 = df_ipan_1.drop("地域手当    ", axis=1)
df_ipan_1 = df_ipan_1.drop("営業手当    ", axis=1)
df_ipan_1 = df_ipan_1.drop("役職手当    ", axis=1)
df_ipan_1 = df_ipan_1.drop("調整手当    ", axis=1)
df_ipan_1 = df_ipan_1.drop("基 本 給    ", axis=1)
df_ipan_1 = df_ipan_1.drop("残業手当    ", axis=1)
df_ipan_1 = df_ipan_1.drop("休出手当    ", axis=1)
df_ipan_1 = df_ipan_1.drop("深夜勤務手当", axis=1)
df_ipan_1 = df_ipan_1.drop("交替時差手当", axis=1)
df_ipan_1 = df_ipan_1.drop("休業手当    ", axis=1)
df_ipan_1 = df_ipan_1.drop("代 休 他    ", axis=1)
df_ipan_1 = df_ipan_1.drop("欠勤控除    ", axis=1)
df_ipan_1 = df_ipan_1.drop("遅早控除    ", axis=1)
df_ipan_1 = df_ipan_1.drop("精 算 分    ", axis=1)
df_ipan_1 = df_ipan_1.drop("支給合計額  ", axis=1)
df_ipan_1 = df_ipan_1.drop("休業控除    ", axis=1)
df_ipan_1 = df_ipan_1.drop("支給額", axis=1)
df_ipan_1 = df_ipan_1.drop("ズレ時間    ", axis=1)
df_ipan_1 = df_ipan_1.drop("雑費・食事代", axis=1)
df_ipan_1 = df_ipan_1.drop("雑費・衣靴代", axis=1)
df_ipan_1 = df_ipan_1.drop("雑費        ", axis=1)
df_ipan_1 = df_ipan_1.drop("受診料・他  ", axis=1)
df_ipan_1 = df_ipan_1.drop("雑費・会費等", axis=1)
df_ipan_1 = df_ipan_1.drop("勤務時間    ", axis=1)
df_ipan_1 = df_ipan_1.drop("遅早時間    ", axis=1)
df_ipan_1 = df_ipan_1.drop("特休日数    ", axis=1)
df_ipan_1 = df_ipan_1.drop("集計区分－２        ", axis=1)
df_ipan_1 = df_ipan_1.sum()
df_ipan_1.to_csv(
    "/content/drive/MyDrive/data/一般管理/B.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 一般管理_一般間接2
df_ipan_2 = df_ipan_m.groupby("区分").get_group("一般間接2")
df_ipan_2 = df_ipan_2.drop("所属1", axis=1)
member = df_ipan_2["所属2"] > 0
df_ipan_2.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_ipan_2["所属3"] > 0
df_ipan_2.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_ipan_2["出勤日数    "]
df_ipan_2.insert(2, "実在籍者", real_member)
time_yukyu = df_ipan_2["有休日数    "] * 8
df_ipan_2.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_ipan_2.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_ipan_2["欠勤日数    "] * 8
df_ipan_2.insert(5, "欠勤時間", abs_time)
work_time = df_ipan_2["勤務時間    "]
df_ipan_2.insert(6, "勤務時間", work_time)
late_early_time = df_ipan_2["遅早時間    "]
df_ipan_2.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_ipan_2["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_ipan_2.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_ipan_2.insert(17, "実労働時間", real_work_time)
zure_time = df_ipan_2["ズレ時間    "]
df_ipan_2.insert(18, "ズレ時間", zure_time)
overtime = df_ipan_2["残業時間    "] + df_ipan_2["深夜残業    "]
df_ipan_2.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_ipan_2.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_ipan_2["法外休出    "]
df_ipan_2.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_ipan_2.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_ipan_2["法定休出    "]
df_ipan_2.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_ipan_2.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_ipan_2["６０Ｈ超    "]
df_ipan_2.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_ipan_2["代休時間    "] + df_ipan_2["深夜代休    "]
df_ipan_2.insert(31, "代休時間", holiday_time)
df_ipan_2.insert(32, "応援時間", 0)
total_work_time = (
    df_ipan_2["勤務時間    "]
    + df_ipan_2["残業時間    "]
    + df_ipan_2["法外休出    "]
    + df_ipan_2["法定休出    "]
)
df_ipan_2.insert(33, "総労働時間", total_work_time)
basic_salary = df_ipan_2["基 本 給    "] + df_ipan_2["支給額"]
df_ipan_2.insert(35, "基本給", basic_salary)
post_allowance = df_ipan_2["役職手当    "]
df_ipan_2.insert(36, "役職手当", post_allowance)
sales_allowance = df_ipan_2["営業手当    "]
df_ipan_2.insert(37, "営業手当", sales_allowance)
aria_allowance = df_ipan_2["地域手当    "]
df_ipan_2.insert(38, "地域手当", aria_allowance)
spe_allowance = df_ipan_2["特殊手当    "]
df_ipan_2.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_ipan_2["特別技技手当"]
df_ipan_2.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_ipan_2["調整手当    "]
df_ipan_2.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_ipan_2["別居手当    "]
df_ipan_2.insert(42, "別居手当", sep_allowance)
com_allowance = df_ipan_2["通勤手当    "]
df_ipan_2.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + spe_tec_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
sub_total_1 = sub_total_1.sum()
sub_total_1 = sub_total_1 / member
sub_total_1 = sub_total_1.astype(int)
df_ipan_2.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_ipan_2["残業手当    "]
df_ipan_2.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_ipan_2["休出手当    "]
df_ipan_2.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_ipan_2["深夜勤務手当"]
df_ipan_2.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_ipan_2["交替時差手当"]
df_ipan_2.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_ipan_2["休業手当    "]
df_ipan_2.insert(49, "休業手当", closed_allowance)
teate_total = (
    overtime_allowance
    + vacation_allowance
    + night_work_allowance
    + time_difference_allowance
    + closed_allowance
)
teate_total = teate_total.sum()
closed_deduction = df_ipan_2["休業控除    "]
df_ipan_2.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_ipan_2["代 休 他    "]
df_ipan_2.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_ipan_2["欠勤控除    "] + df_ipan_2["遅早控除    "]
df_ipan_2.insert(52, "欠勤・遅早控除", abs_early_deduction)
kojyo_total = closed_deduction + compny_leave_etc + abs_early_deduction
kojyo_total = kojyo_total.sum()
settlement = df_ipan_2["精 算 分    "]
df_ipan_2.insert(53, "精算分", settlement)
sub_total_2 = (
    overtime_allowance
    + vacation_allowance
    + night_work_allowance
    + time_difference_allowance
    + closed_allowance
    + closed_deduction
    + compny_leave_etc
    + abs_early_deduction
    + settlement
)
df_ipan_2.insert(54, "小計 2", sub_total_2)
total = (sub_total_1 + sub_total_2) - kojyo_total
df_ipan_2.insert(55, "総支給額", total)
df_ipan_2.insert(56, "応援時間額", 0)
df_ipan_2.insert(57, "役員振替", 0)
df_ipan_2.insert(58, "部門振替", 0)
df_ipan_2.insert(59, "合計", 0)
# 不要フィールドの削除
df_ipan_2 = df_ipan_2.drop("所属2", axis=1)
df_ipan_2 = df_ipan_2.drop("所属3", axis=1)
df_ipan_2 = df_ipan_2.drop("社員ｺｰﾄﾞ", axis=1)
df_ipan_2 = df_ipan_2.drop("区分", axis=1)
df_ipan_2 = df_ipan_2.drop("出勤日数    ", axis=1)
df_ipan_2 = df_ipan_2.drop("有休日数    ", axis=1)
df_ipan_2 = df_ipan_2.drop("欠勤日数    ", axis=1)
df_ipan_2 = df_ipan_2.drop("残業時間    ", axis=1)
df_ipan_2 = df_ipan_2.drop("深夜残業    ", axis=1)
df_ipan_2 = df_ipan_2.drop("法外休出    ", axis=1)
df_ipan_2 = df_ipan_2.drop("法定休出    ", axis=1)
df_ipan_2 = df_ipan_2.drop("代休時間    ", axis=1)
df_ipan_2 = df_ipan_2.drop("深夜代休    ", axis=1)
df_ipan_2 = df_ipan_2.drop("６０Ｈ超    ", axis=1)
df_ipan_2 = df_ipan_2.drop("別居手当    ", axis=1)
df_ipan_2 = df_ipan_2.drop("通勤手当    ", axis=1)
df_ipan_2 = df_ipan_2.drop("特別技技手当", axis=1)
df_ipan_2 = df_ipan_2.drop("特殊手当    ", axis=1)
df_ipan_2 = df_ipan_2.drop("地域手当    ", axis=1)
df_ipan_2 = df_ipan_2.drop("営業手当    ", axis=1)
df_ipan_2 = df_ipan_2.drop("役職手当    ", axis=1)
df_ipan_2 = df_ipan_2.drop("調整手当    ", axis=1)
df_ipan_2 = df_ipan_2.drop("基 本 給    ", axis=1)
df_ipan_2 = df_ipan_2.drop("残業手当    ", axis=1)
df_ipan_2 = df_ipan_2.drop("休出手当    ", axis=1)
df_ipan_2 = df_ipan_2.drop("深夜勤務手当", axis=1)
df_ipan_2 = df_ipan_2.drop("交替時差手当", axis=1)
df_ipan_2 = df_ipan_2.drop("休業手当    ", axis=1)
df_ipan_2 = df_ipan_2.drop("代 休 他    ", axis=1)
df_ipan_2 = df_ipan_2.drop("欠勤控除    ", axis=1)
df_ipan_2 = df_ipan_2.drop("遅早控除    ", axis=1)
df_ipan_2 = df_ipan_2.drop("精 算 分    ", axis=1)
df_ipan_2 = df_ipan_2.drop("支給合計額  ", axis=1)
df_ipan_2 = df_ipan_2.drop("休業控除    ", axis=1)
df_ipan_2 = df_ipan_2.drop("支給額", axis=1)
df_ipan_2 = df_ipan_2.drop("ズレ時間    ", axis=1)
df_ipan_2 = df_ipan_2.drop("雑費・食事代", axis=1)
df_ipan_2 = df_ipan_2.drop("雑費・衣靴代", axis=1)
df_ipan_2 = df_ipan_2.drop("雑費        ", axis=1)
df_ipan_2 = df_ipan_2.drop("受診料・他  ", axis=1)
df_ipan_2 = df_ipan_2.drop("雑費・会費等", axis=1)
df_ipan_2 = df_ipan_2.drop("勤務時間    ", axis=1)
df_ipan_2 = df_ipan_2.drop("遅早時間    ", axis=1)
df_ipan_2 = df_ipan_2.drop("特休日数    ", axis=1)
df_ipan_2 = df_ipan_2.drop("集計区分－２        ", axis=1)
df_ipan_2 = df_ipan_2.sum()
df_ipan_2.to_csv(
    "/content/drive/MyDrive/data/一般管理/C.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 一般管理_一般間接3
df_ipan_3 = df_ipan_m.groupby("区分").get_group("一般間接3")
df_ipan_3 = df_ipan_3.drop("所属1", axis=1)
member = df_ipan_3["所属2"] > 0
df_ipan_3.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_ipan_3["所属3"] > 0
df_ipan_3.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_ipan_3["出勤日数    "]
df_ipan_3.insert(2, "実在籍者", real_member)
time_yukyu = df_ipan_3["有休日数    "] * 8
df_ipan_3.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_ipan_3.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_ipan_3["欠勤日数    "] * 8
df_ipan_3.insert(5, "欠勤時間", abs_time)
work_time = df_ipan_3["勤務時間    "]
df_ipan_3.insert(6, "勤務時間", work_time)
late_early_time = df_ipan_3["遅早時間    "]
df_ipan_3.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_ipan_3["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_ipan_3.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_ipan_3.insert(17, "実労働時間", real_work_time)
zure_time = df_ipan_3["ズレ時間    "]
df_ipan_3.insert(18, "ズレ時間", zure_time)
overtime = df_ipan_3["残業時間    "] + df_ipan_3["深夜残業    "]
df_ipan_3.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_ipan_3.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_ipan_3["法外休出    "]
df_ipan_3.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_ipan_3.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_ipan_3["法定休出    "]
df_ipan_3.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_ipan_3.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_ipan_3["６０Ｈ超    "]
df_ipan_3.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_ipan_3["代休時間    "] + df_ipan_3["深夜代休    "]
df_ipan_3.insert(31, "代休時間", holiday_time)
df_ipan_3.insert(32, "応援時間", 0)
total_work_time = (
    df_ipan_3["勤務時間    "]
    + df_ipan_3["残業時間    "]
    + df_ipan_3["法外休出    "]
    + df_ipan_3["法定休出    "]
)
df_ipan_3.insert(33, "総労働時間", total_work_time)
basic_salary = df_ipan_3["基 本 給    "] + df_ipan_3["支給額"]
df_ipan_3.insert(35, "基本給", basic_salary)
post_allowance = df_ipan_3["役職手当    "]
df_ipan_3.insert(36, "役職手当", post_allowance)
sales_allowance = df_ipan_3["営業手当    "]
df_ipan_3.insert(37, "営業手当", sales_allowance)
aria_allowance = df_ipan_3["地域手当    "]
df_ipan_3.insert(38, "地域手当", aria_allowance)
spe_allowance = df_ipan_3["特殊手当    "]
df_ipan_3.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_ipan_3["特別技技手当"]
df_ipan_3.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_ipan_3["調整手当    "]
df_ipan_3.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_ipan_3["別居手当    "]
df_ipan_3.insert(42, "別居手当", sep_allowance)
com_allowance = df_ipan_3["通勤手当    "]
df_ipan_3.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + spe_tec_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
sub_total_1 = sub_total_1.sum()
sub_total_1 = sub_total_1 / member
sub_total_1 = sub_total_1.astype(int)
df_ipan_3.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_ipan_3["残業手当    "]
df_ipan_3.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_ipan_3["休出手当    "]
df_ipan_3.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_ipan_3["深夜勤務手当"]
df_ipan_3.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_ipan_3["交替時差手当"]
df_ipan_3.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_ipan_3["休業手当    "]
df_ipan_3.insert(49, "休業手当", closed_allowance)
teate_total = (
    overtime_allowance
    + vacation_allowance
    + night_work_allowance
    + time_difference_allowance
    + closed_allowance
)
teate_total = teate_total.sum()
closed_deduction = df_ipan_3["休業控除    "]
df_ipan_3.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_ipan_3["代 休 他    "]
df_ipan_3.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_ipan_3["欠勤控除    "] + df_ipan_3["遅早控除    "]
df_ipan_3.insert(52, "欠勤・遅早控除", abs_early_deduction)
kojyo_total = closed_deduction + compny_leave_etc + abs_early_deduction
kojyo_total = kojyo_total.sum()
settlement = df_ipan_3["精 算 分    "]
df_ipan_3.insert(53, "精算分", settlement)
sub_total_2 = (
    overtime_allowance
    + vacation_allowance
    + night_work_allowance
    + time_difference_allowance
    + closed_allowance
    + closed_deduction
    + compny_leave_etc
    + abs_early_deduction
    + settlement
)
df_ipan_3.insert(54, "小計 2", sub_total_2)
total = (sub_total_1 + sub_total_2) - kojyo_total
df_ipan_3.insert(55, "総支給額", total)
df_ipan_3.insert(56, "応援時間額", 0)
df_ipan_3.insert(57, "役員振替", 0)
df_ipan_3.insert(58, "部門振替", 0)
df_ipan_3.insert(59, "合計", 0)
# 不要フィールドの削除
df_ipan_3 = df_ipan_3.drop("所属2", axis=1)
df_ipan_3 = df_ipan_3.drop("所属3", axis=1)
df_ipan_3 = df_ipan_3.drop("社員ｺｰﾄﾞ", axis=1)
df_ipan_3 = df_ipan_3.drop("区分", axis=1)
df_ipan_3 = df_ipan_3.drop("出勤日数    ", axis=1)
df_ipan_3 = df_ipan_3.drop("有休日数    ", axis=1)
df_ipan_3 = df_ipan_3.drop("欠勤日数    ", axis=1)
df_ipan_3 = df_ipan_3.drop("残業時間    ", axis=1)
df_ipan_3 = df_ipan_3.drop("深夜残業    ", axis=1)
df_ipan_3 = df_ipan_3.drop("法外休出    ", axis=1)
df_ipan_3 = df_ipan_3.drop("法定休出    ", axis=1)
df_ipan_3 = df_ipan_3.drop("代休時間    ", axis=1)
df_ipan_3 = df_ipan_3.drop("深夜代休    ", axis=1)
df_ipan_3 = df_ipan_3.drop("６０Ｈ超    ", axis=1)
df_ipan_3 = df_ipan_3.drop("別居手当    ", axis=1)
df_ipan_3 = df_ipan_3.drop("通勤手当    ", axis=1)
df_ipan_3 = df_ipan_3.drop("特別技技手当", axis=1)
df_ipan_3 = df_ipan_3.drop("特殊手当    ", axis=1)
df_ipan_3 = df_ipan_3.drop("地域手当    ", axis=1)
df_ipan_3 = df_ipan_3.drop("営業手当    ", axis=1)
df_ipan_3 = df_ipan_3.drop("役職手当    ", axis=1)
df_ipan_3 = df_ipan_3.drop("調整手当    ", axis=1)
df_ipan_3 = df_ipan_3.drop("基 本 給    ", axis=1)
df_ipan_3 = df_ipan_3.drop("残業手当    ", axis=1)
df_ipan_3 = df_ipan_3.drop("休出手当    ", axis=1)
df_ipan_3 = df_ipan_3.drop("深夜勤務手当", axis=1)
df_ipan_3 = df_ipan_3.drop("交替時差手当", axis=1)
df_ipan_3 = df_ipan_3.drop("休業手当    ", axis=1)
df_ipan_3 = df_ipan_3.drop("代 休 他    ", axis=1)
df_ipan_3 = df_ipan_3.drop("欠勤控除    ", axis=1)
df_ipan_3 = df_ipan_3.drop("遅早控除    ", axis=1)
df_ipan_3 = df_ipan_3.drop("精 算 分    ", axis=1)
df_ipan_3 = df_ipan_3.drop("支給合計額  ", axis=1)
df_ipan_3 = df_ipan_3.drop("休業控除    ", axis=1)
df_ipan_3 = df_ipan_3.drop("支給額", axis=1)
df_ipan_3 = df_ipan_3.drop("ズレ時間    ", axis=1)
df_ipan_3 = df_ipan_3.drop("雑費・食事代", axis=1)
df_ipan_3 = df_ipan_3.drop("雑費・衣靴代", axis=1)
df_ipan_3 = df_ipan_3.drop("雑費        ", axis=1)
df_ipan_3 = df_ipan_3.drop("受診料・他  ", axis=1)
df_ipan_3 = df_ipan_3.drop("雑費・会費等", axis=1)
df_ipan_3 = df_ipan_3.drop("勤務時間    ", axis=1)
df_ipan_3 = df_ipan_3.drop("遅早時間    ", axis=1)
df_ipan_3 = df_ipan_3.drop("特休日数    ", axis=1)
df_ipan_3 = df_ipan_3.drop("集計区分－２        ", axis=1)
df_ipan_3 = df_ipan_3.sum()
df_ipan_3.to_csv(
    "/content/drive/MyDrive/data/一般管理/D.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 一般管理_一般間接6
df_ipan_6 = df_ipan_m.groupby("区分").get_group("一般間接6")
df_ipan_6 = df_ipan_6.drop("所属1", axis=1)
member = df_ipan_6["所属2"] > 0
df_ipan_6.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_ipan_6["所属3"] > 0
df_ipan_6.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_ipan_6["出勤日数    "]
df_ipan_6.insert(2, "実在籍者", real_member)
time_yukyu = df_ipan_6["有休日数    "] * 8
df_ipan_6.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_ipan_6.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_ipan_6["欠勤日数    "] * 8
df_ipan_6.insert(5, "欠勤時間", abs_time)
work_time = df_ipan_6["勤務時間    "]
df_ipan_6.insert(6, "勤務時間", work_time)
late_early_time = df_ipan_6["遅早時間    "]
df_ipan_6.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_ipan_6["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_ipan_6.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_ipan_6.insert(17, "実労働時間", real_work_time)
zure_time = df_ipan_6["ズレ時間    "]
df_ipan_6.insert(18, "ズレ時間", zure_time)
overtime = df_ipan_6["残業時間    "] + df_ipan_6["深夜残業    "]
df_ipan_6.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_ipan_6.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_ipan_6["法外休出    "]
df_ipan_6.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_ipan_6.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_ipan_6["法定休出    "]
df_ipan_6.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_ipan_6.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_ipan_6["６０Ｈ超    "]
df_ipan_6.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_ipan_6["代休時間    "] + df_ipan_6["深夜代休    "]
df_ipan_6.insert(31, "代休時間", holiday_time)
df_ipan_6.insert(32, "応援時間", 0)
total_work_time = (
    df_ipan_6["勤務時間    "]
    + df_ipan_6["残業時間    "]
    + df_ipan_6["法外休出    "]
    + df_ipan_6["法定休出    "]
)
df_ipan_6.insert(33, "総労働時間", total_work_time)
basic_salary = df_ipan_6["基 本 給    "] + df_ipan_6["支給額"]
df_ipan_6.insert(35, "基本給", basic_salary)
post_allowance = df_ipan_6["役職手当    "]
df_ipan_6.insert(36, "役職手当", post_allowance)
sales_allowance = df_ipan_6["営業手当    "]
df_ipan_6.insert(37, "営業手当", sales_allowance)
aria_allowance = df_ipan_6["地域手当    "]
df_ipan_6.insert(38, "地域手当", aria_allowance)
spe_allowance = df_ipan_6["特殊手当    "]
df_ipan_6.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_ipan_6["特別技技手当"]
df_ipan_6.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_ipan_6["調整手当    "]
df_ipan_6.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_ipan_6["別居手当    "]
df_ipan_6.insert(42, "別居手当", sep_allowance)
com_allowance = df_ipan_6["通勤手当    "]
df_ipan_6.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + spe_tec_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
sub_total_1 = sub_total_1.sum()
sub_total_1 = sub_total_1 / member
sub_total_1 = sub_total_1.astype(int)
df_ipan_6.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_ipan_6["残業手当    "]
df_ipan_6.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_ipan_6["休出手当    "]
df_ipan_6.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_ipan_6["深夜勤務手当"]
df_ipan_6.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_ipan_6["交替時差手当"]
df_ipan_6.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_ipan_6["休業手当    "]
df_ipan_6.insert(49, "休業手当", closed_allowance)
teate_total = (
    overtime_allowance
    + vacation_allowance
    + night_work_allowance
    + time_difference_allowance
    + closed_allowance
)
teate_total = teate_total.sum()
closed_deduction = df_ipan_6["休業控除    "]
df_ipan_6.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_ipan_6["代 休 他    "]
df_ipan_6.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_ipan_6["欠勤控除    "] + df_ipan_6["遅早控除    "]
df_ipan_6.insert(52, "欠勤・遅早控除", abs_early_deduction)
kojyo_total = closed_deduction + compny_leave_etc + abs_early_deduction
kojyo_total = kojyo_total.sum()
settlement = df_ipan_6["精 算 分    "]
df_ipan_6.insert(53, "精算分", settlement)
sub_total_2 = (
    overtime_allowance
    + vacation_allowance
    + night_work_allowance
    + time_difference_allowance
    + closed_allowance
    + closed_deduction
    + compny_leave_etc
    + abs_early_deduction
    + settlement
)
df_ipan_6.insert(54, "小計 2", sub_total_2)
total = (sub_total_1 + sub_total_2) - kojyo_total
df_ipan_6.insert(55, "総支給額", total)
df_ipan_6.insert(56, "応援時間額", 0)
df_ipan_6.insert(57, "役員振替", 0)
df_ipan_6.insert(58, "部門振替", 0)
df_ipan_6.insert(59, "合計", 0)
# 不要フィールドの削除
df_ipan_6 = df_ipan_6.drop("所属2", axis=1)
df_ipan_6 = df_ipan_6.drop("所属3", axis=1)
df_ipan_6 = df_ipan_6.drop("社員ｺｰﾄﾞ", axis=1)
df_ipan_6 = df_ipan_6.drop("区分", axis=1)
df_ipan_6 = df_ipan_6.drop("出勤日数    ", axis=1)
df_ipan_6 = df_ipan_6.drop("有休日数    ", axis=1)
df_ipan_6 = df_ipan_6.drop("欠勤日数    ", axis=1)
df_ipan_6 = df_ipan_6.drop("残業時間    ", axis=1)
df_ipan_6 = df_ipan_6.drop("深夜残業    ", axis=1)
df_ipan_6 = df_ipan_6.drop("法外休出    ", axis=1)
df_ipan_6 = df_ipan_6.drop("法定休出    ", axis=1)
df_ipan_6 = df_ipan_6.drop("代休時間    ", axis=1)
df_ipan_6 = df_ipan_6.drop("深夜代休    ", axis=1)
df_ipan_6 = df_ipan_6.drop("６０Ｈ超    ", axis=1)
df_ipan_6 = df_ipan_6.drop("別居手当    ", axis=1)
df_ipan_6 = df_ipan_6.drop("通勤手当    ", axis=1)
df_ipan_6 = df_ipan_6.drop("特別技技手当", axis=1)
df_ipan_6 = df_ipan_6.drop("特殊手当    ", axis=1)
df_ipan_6 = df_ipan_6.drop("地域手当    ", axis=1)
df_ipan_6 = df_ipan_6.drop("営業手当    ", axis=1)
df_ipan_6 = df_ipan_6.drop("役職手当    ", axis=1)
df_ipan_6 = df_ipan_6.drop("調整手当    ", axis=1)
df_ipan_6 = df_ipan_6.drop("基 本 給    ", axis=1)
df_ipan_6 = df_ipan_6.drop("残業手当    ", axis=1)
df_ipan_6 = df_ipan_6.drop("休出手当    ", axis=1)
df_ipan_6 = df_ipan_6.drop("深夜勤務手当", axis=1)
df_ipan_6 = df_ipan_6.drop("交替時差手当", axis=1)
df_ipan_6 = df_ipan_6.drop("休業手当    ", axis=1)
df_ipan_6 = df_ipan_6.drop("代 休 他    ", axis=1)
df_ipan_6 = df_ipan_6.drop("欠勤控除    ", axis=1)
df_ipan_6 = df_ipan_6.drop("遅早控除    ", axis=1)
df_ipan_6 = df_ipan_6.drop("精 算 分    ", axis=1)
df_ipan_6 = df_ipan_6.drop("支給合計額  ", axis=1)
df_ipan_6 = df_ipan_6.drop("休業控除    ", axis=1)
df_ipan_6 = df_ipan_6.drop("支給額", axis=1)
df_ipan_6 = df_ipan_6.drop("ズレ時間    ", axis=1)
df_ipan_6 = df_ipan_6.drop("雑費・食事代", axis=1)
df_ipan_6 = df_ipan_6.drop("雑費・衣靴代", axis=1)
df_ipan_6 = df_ipan_6.drop("雑費        ", axis=1)
df_ipan_6 = df_ipan_6.drop("受診料・他  ", axis=1)
df_ipan_6 = df_ipan_6.drop("雑費・会費等", axis=1)
df_ipan_6 = df_ipan_6.drop("勤務時間    ", axis=1)
df_ipan_6 = df_ipan_6.drop("遅早時間    ", axis=1)
df_ipan_6 = df_ipan_6.drop("特休日数    ", axis=1)
df_ipan_6 = df_ipan_6.drop("集計区分－２        ", axis=1)
df_ipan_6 = df_ipan_6.sum()
df_ipan_6.to_csv(
    "/content/drive/MyDrive/data/一般管理/E.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 一般管理_一般販売1
df_ipan_h_1 = df_ipan_m.groupby("区分").get_group("一般販売1")
df_ipan_h_1 = df_ipan_h_1.drop("所属1", axis=1)
member = df_ipan_h_1["所属2"] > 0
df_ipan_h_1.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_ipan_h_1["所属3"] > 0
df_ipan_h_1.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_ipan_h_1["出勤日数    "]
# real_member = real_member.round().astype(int)
df_ipan_h_1.insert(2, "実在籍者", real_member)
time_yukyu = df_ipan_h_1["有休日数    "] * 8
df_ipan_h_1.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_ipan_h_1.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_ipan_h_1["欠勤日数    "] * 8
df_ipan_h_1.insert(5, "欠勤時間", abs_time)
work_time = df_ipan_h_1["勤務時間    "]
df_ipan_h_1.insert(6, "勤務時間", work_time)
late_early_time = df_ipan_h_1["遅早時間    "]
df_ipan_h_1.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_ipan_h_1["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_ipan_h_1.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_ipan_h_1.insert(17, "実労働時間", real_work_time)
zure_time = df_ipan_h_1["ズレ時間    "]
df_ipan_h_1.insert(18, "ズレ時間", zure_time)
overtime = df_ipan_h_1["残業時間    "] + df_ipan_h_1["深夜残業    "]
df_ipan_h_1.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_ipan_h_1.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_ipan_h_1["法外休出    "]
df_ipan_h_1.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_ipan_h_1.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_ipan_h_1["法定休出    "]
df_ipan_h_1.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_ipan_h_1.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_ipan_h_1["６０Ｈ超    "]
df_ipan_h_1.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_ipan_h_1["代休時間    "] + df_ipan_h_1["深夜代休    "]
df_ipan_h_1.insert(31, "代休時間", holiday_time)
df_ipan_h_1.insert(32, "応援時間", 0)
total_work_time = (
    df_ipan_h_1["勤務時間    "]
    + df_ipan_h_1["残業時間    "]
    + df_ipan_h_1["法外休出    "]
    + df_ipan_h_1["法定休出    "]
)
df_ipan_h_1.insert(33, "総労働時間", total_work_time)
basic_salary = df_ipan_h_1["基 本 給    "] + df_ipan_h_1["支給額"]
df_ipan_h_1.insert(35, "基本給", basic_salary)
post_allowance = df_ipan_h_1["役職手当    "]
df_ipan_h_1.insert(36, "役職手当", post_allowance)
sales_allowance = df_ipan_h_1["営業手当    "]
df_ipan_h_1.insert(37, "営業手当", sales_allowance)
aria_allowance = df_ipan_h_1["地域手当    "]
df_ipan_h_1.insert(38, "地域手当", aria_allowance)
spe_allowance = df_ipan_h_1["特殊手当    "]
df_ipan_h_1.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_ipan_h_1["特別技技手当"]
df_ipan_h_1.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_ipan_h_1["調整手当    "]
df_ipan_h_1.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_ipan_h_1["別居手当    "]
df_ipan_h_1.insert(42, "別居手当", sep_allowance)
com_allowance = df_ipan_h_1["通勤手当    "]
df_ipan_h_1.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + spe_tec_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
sub_total_1 = sub_total_1.sum()
sub_total_1 = sub_total_1 / member
sub_total_1 = sub_total_1.astype(int)
df_ipan_h_1.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_ipan_h_1["残業手当    "]
df_ipan_h_1.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_ipan_h_1["休出手当    "]
df_ipan_h_1.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_ipan_h_1["深夜勤務手当"]
df_ipan_h_1.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_ipan_h_1["交替時差手当"]
df_ipan_h_1.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_ipan_h_1["休業手当    "]
df_ipan_h_1.insert(49, "休業手当", closed_allowance)
teate_total = (
    overtime_allowance
    + vacation_allowance
    + night_work_allowance
    + time_difference_allowance
    + closed_allowance
)
teate_total = teate_total.sum()
closed_deduction = df_ipan_h_1["休業控除    "]
df_ipan_h_1.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_ipan_h_1["代 休 他    "]
df_ipan_h_1.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_ipan_h_1["欠勤控除    "] + df_ipan_h_1["遅早控除    "]
df_ipan_h_1.insert(52, "欠勤・遅早控除", abs_early_deduction)
kojyo_total = closed_deduction + compny_leave_etc + abs_early_deduction
kojyo_total = kojyo_total.sum()
settlement = df_ipan_h_1["精 算 分    "]
df_ipan_h_1.insert(53, "精算分", settlement)
sub_total_2 = (
    overtime_allowance
    + vacation_allowance
    + night_work_allowance
    + time_difference_allowance
    + closed_allowance
    + closed_deduction
    + compny_leave_etc
    + abs_early_deduction
    + settlement
)
df_ipan_h_1.insert(54, "小計 2", sub_total_2)
total = (sub_total_1 + sub_total_2) - kojyo_total
df_ipan_h_1.insert(55, "総支給額", total)
df_ipan_h_1.insert(56, "応援時間額", 0)
df_ipan_h_1.insert(57, "役員振替", 0)
df_ipan_h_1.insert(58, "部門振替", 0)
df_ipan_h_1.insert(59, "合計", 0)
# 不要フィールドの削除
df_ipan_h_1 = df_ipan_h_1.drop("所属2", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("所属3", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("社員ｺｰﾄﾞ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("区分", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("出勤日数    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("有休日数    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("欠勤日数    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("残業時間    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("深夜残業    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("法外休出    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("法定休出    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("代休時間    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("深夜代休    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("６０Ｈ超    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("別居手当    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("通勤手当    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("特別技技手当", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("特殊手当    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("地域手当    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("営業手当    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("役職手当    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("調整手当    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("基 本 給    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("残業手当    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("休出手当    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("深夜勤務手当", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("交替時差手当", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("休業手当    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("代 休 他    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("欠勤控除    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("遅早控除    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("精 算 分    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("支給合計額  ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("休業控除    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("支給額", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("ズレ時間    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("雑費・食事代", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("雑費・衣靴代", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("雑費        ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("受診料・他  ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("雑費・会費等", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("勤務時間    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("遅早時間    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("特休日数    ", axis=1)
df_ipan_h_1 = df_ipan_h_1.drop("集計区分－２        ", axis=1)
df_ipan_h_1 = df_ipan_h_1.sum()
df_ipan_h_1.to_csv(
    "/content/drive/MyDrive/data/一般管理/F.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 一般管理_一般販売2
df_ipan_h_2 = df_ipan_m.groupby("区分").get_group("一般販売2")
df_ipan_h_2 = df_ipan_h_2.drop("所属1", axis=1)
member = df_ipan_h_2["所属2"] > 0
df_ipan_h_2.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_ipan_h_2["所属3"] > 0
df_ipan_h_2.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_ipan_h_2["出勤日数    "]
df_ipan_h_2.insert(2, "実在籍者", real_member)
time_yukyu = df_ipan_h_2["有休日数    "] * 8
df_ipan_h_2.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_ipan_h_2.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_ipan_h_2["欠勤日数    "] * 8
df_ipan_h_2.insert(5, "欠勤時間", abs_time)
work_time = df_ipan_h_2["勤務時間    "]
df_ipan_h_2.insert(6, "勤務時間", work_time)
late_early_time = df_ipan_h_2["遅早時間    "]
df_ipan_h_2.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_ipan_h_2["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_ipan_h_2.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_ipan_h_2.insert(17, "実労働時間", real_work_time)
zure_time = df_ipan_h_2["ズレ時間    "]
df_ipan_h_2.insert(18, "ズレ時間", zure_time)
overtime = df_ipan_h_2["残業時間    "] + df_ipan_h_2["深夜残業    "]
df_ipan_h_2.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_ipan_h_2.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_ipan_h_2["法外休出    "]
df_ipan_h_2.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_ipan_h_2.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_ipan_h_2["法定休出    "]
df_ipan_h_2.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_ipan_h_2.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_ipan_h_2["６０Ｈ超    "]
df_ipan_h_2.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_ipan_h_2["代休時間    "] + df_ipan_h_2["深夜代休    "]
df_ipan_h_2.insert(31, "代休時間", holiday_time)
df_ipan_h_2.insert(32, "応援時間", 0)
total_work_time = (
    df_ipan_h_2["勤務時間    "]
    + df_ipan_h_2["残業時間    "]
    + df_ipan_h_2["法外休出    "]
    + df_ipan_h_2["法定休出    "]
)
df_ipan_h_2.insert(33, "総労働時間", total_work_time)
basic_salary = df_ipan_h_2["基 本 給    "] + df_ipan_h_2["支給額"]
df_ipan_h_2.insert(35, "基本給", basic_salary)
post_allowance = df_ipan_h_2["役職手当    "]
df_ipan_h_2.insert(36, "役職手当", post_allowance)
sales_allowance = df_ipan_h_2["営業手当    "]
df_ipan_h_2.insert(37, "営業手当", sales_allowance)
aria_allowance = df_ipan_h_2["地域手当    "]
df_ipan_h_2.insert(38, "地域手当", aria_allowance)
spe_allowance = df_ipan_h_2["特殊手当    "]
df_ipan_h_2.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_ipan_h_2["特別技技手当"]
df_ipan_h_2.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_ipan_h_2["調整手当    "]
df_ipan_h_2.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_ipan_h_2["別居手当    "]
df_ipan_h_2.insert(42, "別居手当", sep_allowance)
com_allowance = df_ipan_h_2["通勤手当    "]
df_ipan_h_2.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + spe_tec_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
sub_total_1 = sub_total_1.sum()
sub_total_1 = sub_total_1 / member
sub_total_1 = sub_total_1.astype(int)
df_ipan_h_2.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_ipan_h_2["残業手当    "]
df_ipan_h_2.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_ipan_h_2["休出手当    "]
df_ipan_h_2.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_ipan_h_2["深夜勤務手当"]
df_ipan_h_2.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_ipan_h_2["交替時差手当"]
df_ipan_h_2.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_ipan_h_2["休業手当    "]
df_ipan_h_2.insert(49, "休業手当", closed_allowance)
teate_total = (
    overtime_allowance
    + vacation_allowance
    + night_work_allowance
    + time_difference_allowance
    + closed_allowance
)
teate_total = teate_total.sum()
closed_deduction = df_ipan_h_2["休業控除    "]
df_ipan_h_2.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_ipan_h_2["代 休 他    "]
df_ipan_h_2.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_ipan_h_2["欠勤控除    "] + df_ipan_h_2["遅早控除    "]
df_ipan_h_2.insert(52, "欠勤・遅早控除", abs_early_deduction)
kojyo_total = closed_deduction + compny_leave_etc + abs_early_deduction
kojyo_total = kojyo_total.sum()
settlement = df_ipan_h_2["精 算 分    "]
df_ipan_h_2.insert(53, "精算分", settlement)
sub_total_2 = (
    overtime_allowance
    + vacation_allowance
    + night_work_allowance
    + time_difference_allowance
    + closed_allowance
    + closed_deduction
    + compny_leave_etc
    + abs_early_deduction
    + settlement
)
df_ipan_h_2.insert(54, "小計 2", sub_total_2)
total = (sub_total_1 + sub_total_2) - kojyo_total
df_ipan_h_2.insert(55, "総支給額", total)
df_ipan_h_2.insert(56, "応援時間額", 0)
df_ipan_h_2.insert(57, "役員振替", 0)
df_ipan_h_2.insert(58, "部門振替", 0)
df_ipan_h_2.insert(59, "合計", 0)
# 不要フィールドの削除
df_ipan_h_2 = df_ipan_h_2.drop("所属2", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("所属3", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("社員ｺｰﾄﾞ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("区分", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("出勤日数    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("有休日数    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("欠勤日数    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("残業時間    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("深夜残業    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("法外休出    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("法定休出    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("代休時間    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("深夜代休    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("６０Ｈ超    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("別居手当    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("通勤手当    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("特別技技手当", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("特殊手当    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("地域手当    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("営業手当    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("役職手当    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("調整手当    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("基 本 給    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("残業手当    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("休出手当    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("深夜勤務手当", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("交替時差手当", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("休業手当    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("代 休 他    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("欠勤控除    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("遅早控除    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("精 算 分    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("支給合計額  ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("休業控除    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("支給額", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("ズレ時間    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("雑費・食事代", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("雑費・衣靴代", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("雑費        ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("受診料・他  ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("雑費・会費等", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("勤務時間    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("遅早時間    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("特休日数    ", axis=1)
df_ipan_h_2 = df_ipan_h_2.drop("集計区分－２        ", axis=1)
df_ipan_h_2 = df_ipan_h_2.sum()
df_ipan_h_2.to_csv(
    "/content/drive/MyDrive/data/一般管理/G.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 鍛造_間接1
df_tanzo_k_1 = df_tanzo_m.groupby("区分").get_group("間接1")
df_tanzo_k_1 = df_tanzo_k_1.drop("所属1", axis=1)
member = df_tanzo_k_1["所属2"] > 0
df_tanzo_k_1.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_tanzo_k_1["所属3"] > 0
df_tanzo_k_1.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_tanzo_k_1["出勤日数    "]
df_tanzo_k_1.insert(2, "実在籍者", real_member)
time_yukyu = df_tanzo_k_1["有休日数    "] * 8
df_tanzo_k_1.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_tanzo_k_1.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_tanzo_k_1["欠勤日数    "] * 8
df_tanzo_k_1.insert(5, "欠勤時間", abs_time)
work_time = df_tanzo_k_1["勤務時間    "]
df_tanzo_k_1.insert(6, "勤務時間", work_time)
late_early_time = df_tanzo_k_1["遅早時間    "]
df_tanzo_k_1.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_tanzo_k_1["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_tanzo_k_1.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_tanzo_k_1.insert(17, "実労働時間", real_work_time)
zure_time = df_tanzo_k_1["ズレ時間    "]
df_tanzo_k_1.insert(18, "ズレ時間", zure_time)
overtime = df_tanzo_k_1["残業時間    "] + df_tanzo_k_1["深夜残業    "]
df_tanzo_k_1.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_tanzo_k_1.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_tanzo_k_1["法外休出    "]
df_tanzo_k_1.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_tanzo_k_1.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_tanzo_k_1["法定休出    "]
df_tanzo_k_1.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_tanzo_k_1.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_tanzo_k_1["６０Ｈ超    "]
df_tanzo_k_1.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_tanzo_k_1["代休時間    "] + df_tanzo_k_1["深夜代休    "]
df_tanzo_k_1.insert(31, "代休時間", holiday_time)
df_tanzo_k_1.insert(32, "応援時間", 0)
total_work_time = (
    df_tanzo_k_1["勤務時間    "]
    + df_tanzo_k_1["残業時間    "]
    + df_tanzo_k_1["法外休出    "]
    + df_tanzo_k_1["法定休出    "]
)
df_tanzo_k_1.insert(33, "総労働時間", total_work_time)
basic_salary = df_tanzo_k_1["基 本 給    "] + df_tanzo_k_1["支給額"]
df_tanzo_k_1.insert(35, "基本給", basic_salary)
post_allowance = df_tanzo_k_1["役職手当    "]
df_tanzo_k_1.insert(36, "役職手当", post_allowance)
sales_allowance = df_tanzo_k_1["営業手当    "]
df_tanzo_k_1.insert(37, "営業手当", sales_allowance)
aria_allowance = df_tanzo_k_1["地域手当    "]
df_tanzo_k_1.insert(38, "地域手当", aria_allowance)
spe_allowance = df_tanzo_k_1["特殊手当    "]
df_tanzo_k_1.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_tanzo_k_1["特別技技手当"]
df_tanzo_k_1.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_tanzo_k_1["調整手当    "]
df_tanzo_k_1.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_tanzo_k_1["別居手当    "]
df_tanzo_k_1.insert(42, "別居手当", sep_allowance)
com_allowance = df_tanzo_k_1["通勤手当    "]
df_tanzo_k_1.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_tanzo_k_1.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_tanzo_k_1["残業手当    "]
df_tanzo_k_1.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_tanzo_k_1["休出手当    "]
df_tanzo_k_1.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_tanzo_k_1["深夜勤務手当"]
df_tanzo_k_1.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_tanzo_k_1["交替時差手当"]
df_tanzo_k_1.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_tanzo_k_1["休業手当    "]
df_tanzo_k_1.insert(49, "休業手当", closed_allowance)
closed_deduction = df_tanzo_k_1["休業控除    "]
df_tanzo_k_1.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_tanzo_k_1["代 休 他    "]
df_tanzo_k_1.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_tanzo_k_1["欠勤控除    "] + df_tanzo_k_1["遅早控除    "]
df_tanzo_k_1.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_tanzo_k_1["精 算 分    "]
df_tanzo_k_1.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_tanzo_k_1.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_tanzo_k_1.insert(55, "総支給額", total)
df_tanzo_k_1.insert(56, "応援時間額", 0)
df_tanzo_k_1.insert(57, "役員振替", 0)
df_tanzo_k_1.insert(58, "部門振替", 0)
df_tanzo_k_1.insert(59, "合計", 0)
# フィールドの削除
df_tanzo_k_1 = df_tanzo_k_1.drop("所属2", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("所属3", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("社員ｺｰﾄﾞ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("区分", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("出勤日数    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("有休日数    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("欠勤日数    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("残業時間    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("深夜残業    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("法外休出    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("法定休出    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("代休時間    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("深夜代休    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("６０Ｈ超    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("別居手当    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("通勤手当    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("特別技技手当", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("特殊手当    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("地域手当    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("営業手当    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("役職手当    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("調整手当    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("基 本 給    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("残業手当    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("休出手当    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("深夜勤務手当", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("交替時差手当", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("休業手当    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("代 休 他    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("欠勤控除    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("遅早控除    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("精 算 分    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("支給合計額  ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("休業控除    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("支給額", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("ズレ時間    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("雑費・食事代", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("雑費・衣靴代", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("雑費        ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("受診料・他  ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("雑費・会費等", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("勤務時間    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("遅早時間    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("特休日数    ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.drop("集計区分－２        ", axis=1)
df_tanzo_k_1 = df_tanzo_k_1.sum()
df_tanzo_k_1.to_csv(
    "/content/drive/MyDrive/data/鍛造/A.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 鍛造_間接2
df_tanzo_k_2 = df_tanzo_m.groupby("区分").get_group("間接2")
df_tanzo_k_2 = df_tanzo_k_2.drop("所属1", axis=1)
member = df_tanzo_k_2["所属2"] > 0
df_tanzo_k_2.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_tanzo_k_2["所属3"] > 0
df_tanzo_k_2.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_tanzo_k_2["出勤日数    "]
df_tanzo_k_2.insert(2, "実在籍者", real_member)
time_yukyu = df_tanzo_k_2["有休日数    "] * 8
df_tanzo_k_2.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_tanzo_k_2.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_tanzo_k_2["欠勤日数    "] * 8
df_tanzo_k_2.insert(5, "欠勤時間", abs_time)
work_time = df_tanzo_k_2["勤務時間    "]
df_tanzo_k_2.insert(6, "勤務時間", work_time)
late_early_time = df_tanzo_k_2["遅早時間    "]
df_tanzo_k_2.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_tanzo_k_2["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_tanzo_k_2.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_tanzo_k_2.insert(17, "実労働時間", real_work_time)
zure_time = df_tanzo_k_2["ズレ時間    "]
df_tanzo_k_2.insert(18, "ズレ時間", zure_time)
overtime = df_tanzo_k_2["残業時間    "] + df_tanzo_k_2["深夜残業    "]
df_tanzo_k_2.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_tanzo_k_2.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_tanzo_k_2["法外休出    "]
df_tanzo_k_2.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_tanzo_k_2.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_tanzo_k_2["法定休出    "]
df_tanzo_k_2.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_tanzo_k_2.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_tanzo_k_2["６０Ｈ超    "]
df_tanzo_k_2.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_tanzo_k_2["代休時間    "] + df_tanzo_k_2["深夜代休    "]
df_tanzo_k_2.insert(31, "代休時間", holiday_time)
df_tanzo_k_2.insert(32, "応援時間", 0)
total_work_time = (
    df_tanzo_k_2["勤務時間    "]
    + df_tanzo_k_2["残業時間    "]
    + df_tanzo_k_2["法外休出    "]
    + df_tanzo_k_2["法定休出    "]
)
df_tanzo_k_2.insert(33, "総労働時間", total_work_time)
basic_salary = df_tanzo_k_2["基 本 給    "] + df_tanzo_k_2["支給額"]
df_tanzo_k_2.insert(35, "基本給", basic_salary)
post_allowance = df_tanzo_k_2["役職手当    "]
df_tanzo_k_2.insert(36, "役職手当", post_allowance)
sales_allowance = df_tanzo_k_2["営業手当    "]
df_tanzo_k_2.insert(37, "営業手当", sales_allowance)
aria_allowance = df_tanzo_k_2["地域手当    "]
df_tanzo_k_2.insert(38, "地域手当", aria_allowance)
spe_allowance = df_tanzo_k_2["特殊手当    "]
df_tanzo_k_2.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_tanzo_k_2["特別技技手当"]
df_tanzo_k_2.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_tanzo_k_2["調整手当    "]
df_tanzo_k_2.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_tanzo_k_2["別居手当    "]
df_tanzo_k_2.insert(42, "別居手当", sep_allowance)
com_allowance = df_tanzo_k_2["通勤手当    "]
df_tanzo_k_2.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_tanzo_k_2.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_tanzo_k_2["残業手当    "]
df_tanzo_k_2.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_tanzo_k_2["休出手当    "]
df_tanzo_k_2.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_tanzo_k_2["深夜勤務手当"]
df_tanzo_k_2.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_tanzo_k_2["交替時差手当"]
df_tanzo_k_2.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_tanzo_k_2["休業手当    "]
df_tanzo_k_2.insert(49, "休業手当", closed_allowance)
closed_deduction = df_tanzo_k_2["休業控除    "]
df_tanzo_k_2.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_tanzo_k_2["代 休 他    "]
df_tanzo_k_2.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_tanzo_k_2["欠勤控除    "] + df_tanzo_k_2["遅早控除    "]
df_tanzo_k_2.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_tanzo_k_2["精 算 分    "]
df_tanzo_k_2.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_tanzo_k_2.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_tanzo_k_2.insert(55, "総支給額", total)
df_tanzo_k_2.insert(56, "応援時間額", 0)
df_tanzo_k_2.insert(57, "役員振替", 0)
df_tanzo_k_2.insert(58, "部門振替", 0)
df_tanzo_k_2.insert(59, "合計", 0)
# 不要フィールドの削除
df_tanzo_k_2 = df_tanzo_k_2.drop("所属2", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("所属3", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("社員ｺｰﾄﾞ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("区分", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("出勤日数    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("有休日数    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("欠勤日数    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("残業時間    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("深夜残業    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("法外休出    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("法定休出    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("代休時間    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("深夜代休    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("６０Ｈ超    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("別居手当    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("通勤手当    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("特別技技手当", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("特殊手当    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("地域手当    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("営業手当    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("役職手当    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("調整手当    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("基 本 給    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("残業手当    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("休出手当    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("深夜勤務手当", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("交替時差手当", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("休業手当    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("代 休 他    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("欠勤控除    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("遅早控除    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("精 算 分    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("支給合計額  ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("休業控除    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("支給額", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("ズレ時間    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("雑費・食事代", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("雑費・衣靴代", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("雑費        ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("受診料・他  ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("雑費・会費等", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("勤務時間    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("遅早時間    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("特休日数    ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.drop("集計区分－２        ", axis=1)
df_tanzo_k_2 = df_tanzo_k_2.sum()
df_tanzo_k_2.to_csv(
    "/content/drive/MyDrive/data/鍛造/B.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 鍛造_間接3
df_tanzo_k_3 = df_tanzo_m.groupby("区分").get_group("間接3")
df_tanzo_k_3 = df_tanzo_k_3.drop("所属1", axis=1)
member = df_tanzo_k_3["所属2"] > 0
df_tanzo_k_3.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_tanzo_k_3["所属3"] > 0
df_tanzo_k_3.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_tanzo_k_3["出勤日数    "]
df_tanzo_k_3.insert(2, "実在籍者", real_member)
time_yukyu = df_tanzo_k_3["有休日数    "] * 8
df_tanzo_k_3.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_tanzo_k_3.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_tanzo_k_3["欠勤日数    "] * 8
df_tanzo_k_3.insert(5, "欠勤時間", abs_time)
work_time = df_tanzo_k_3["勤務時間    "]
df_tanzo_k_3.insert(6, "勤務時間", work_time)
late_early_time = df_tanzo_k_3["遅早時間    "]
df_tanzo_k_3.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_tanzo_k_3["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_tanzo_k_3.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_tanzo_k_3.insert(17, "実労働時間", real_work_time)
zure_time = df_tanzo_k_3["ズレ時間    "]
df_tanzo_k_3.insert(18, "ズレ時間", zure_time)
overtime = df_tanzo_k_3["残業時間    "] + df_tanzo_k_3["深夜残業    "]
df_tanzo_k_3.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_tanzo_k_3.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_tanzo_k_3["法外休出    "]
df_tanzo_k_3.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_tanzo_k_3.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_tanzo_k_3["法定休出    "]
df_tanzo_k_3.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_tanzo_k_3.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_tanzo_k_3["６０Ｈ超    "]
df_tanzo_k_3.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_tanzo_k_3["代休時間    "] + df_tanzo_k_3["深夜代休    "]
df_tanzo_k_3.insert(31, "代休時間", holiday_time)
df_tanzo_k_3.insert(32, "応援時間", 0)
total_work_time = (
    df_tanzo_k_3["勤務時間    "]
    + df_tanzo_k_3["残業時間    "]
    + df_tanzo_k_3["法外休出    "]
    + df_tanzo_k_3["法定休出    "]
)
df_tanzo_k_3.insert(33, "総労働時間", total_work_time)
basic_salary = df_tanzo_k_3["基 本 給    "] + df_tanzo_k_3["支給額"]
df_tanzo_k_3.insert(35, "基本給", basic_salary)
post_allowance = df_tanzo_k_3["役職手当    "]
df_tanzo_k_3.insert(36, "役職手当", post_allowance)
sales_allowance = df_tanzo_k_3["営業手当    "]
df_tanzo_k_3.insert(37, "営業手当", sales_allowance)
aria_allowance = df_tanzo_k_3["地域手当    "]
df_tanzo_k_3.insert(38, "地域手当", aria_allowance)
spe_allowance = df_tanzo_k_3["特殊手当    "]
df_tanzo_k_3.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_tanzo_k_3["特別技技手当"]
df_tanzo_k_3.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_tanzo_k_3["調整手当    "]
df_tanzo_k_3.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_tanzo_k_3["別居手当    "]
df_tanzo_k_3.insert(42, "別居手当", sep_allowance)
com_allowance = df_tanzo_k_3["通勤手当    "]
df_tanzo_k_3.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_tanzo_k_3.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_tanzo_k_3["残業手当    "]
df_tanzo_k_3.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_tanzo_k_3["休出手当    "]
df_tanzo_k_3.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_tanzo_k_3["深夜勤務手当"]
df_tanzo_k_3.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_tanzo_k_3["交替時差手当"]
df_tanzo_k_3.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_tanzo_k_3["休業手当    "]
df_tanzo_k_3.insert(49, "休業手当", closed_allowance)
closed_deduction = df_tanzo_k_3["休業控除    "]
df_tanzo_k_3.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_tanzo_k_3["代 休 他    "]
df_tanzo_k_3.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_tanzo_k_3["欠勤控除    "] + df_tanzo_k_3["遅早控除    "]
df_tanzo_k_3.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_tanzo_k_3["精 算 分    "]
df_tanzo_k_3.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_tanzo_k_3.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_tanzo_k_3.insert(55, "総支給額", total)
df_tanzo_k_3.insert(56, "応援時間額", 0)
df_tanzo_k_3.insert(57, "役員振替", 0)
df_tanzo_k_3.insert(58, "部門振替", 0)
df_tanzo_k_3.insert(59, "合計", 0)
# 不要フィールドの削除
df_tanzo_k_3 = df_tanzo_k_3.drop("所属2", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("所属3", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("社員ｺｰﾄﾞ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("区分", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("出勤日数    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("有休日数    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("欠勤日数    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("残業時間    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("深夜残業    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("法外休出    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("法定休出    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("代休時間    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("深夜代休    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("６０Ｈ超    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("別居手当    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("通勤手当    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("特別技技手当", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("特殊手当    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("地域手当    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("営業手当    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("役職手当    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("調整手当    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("基 本 給    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("残業手当    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("休出手当    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("深夜勤務手当", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("交替時差手当", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("休業手当    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("代 休 他    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("欠勤控除    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("遅早控除    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("精 算 分    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("支給合計額  ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("休業控除    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("支給額", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("ズレ時間    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("雑費・食事代", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("雑費・衣靴代", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("雑費        ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("受診料・他  ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("雑費・会費等", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("勤務時間    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("遅早時間    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("特休日数    ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.drop("集計区分－２        ", axis=1)
df_tanzo_k_3 = df_tanzo_k_3.sum()
df_tanzo_k_3.to_csv(
    "/content/drive/MyDrive/data/鍛造/C.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 鍛造_間接4
df_tanzo_k_4 = df_tanzo_m.groupby("区分").get_group("間接4")
df_tanzo_k_4 = df_tanzo_k_4.drop("所属1", axis=1)
member = df_tanzo_k_4["所属2"] > 0
df_tanzo_k_4.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_tanzo_k_4["所属3"] > 0
df_tanzo_k_4.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_tanzo_k_4["出勤日数    "]
df_tanzo_k_4.insert(2, "実在籍者", real_member)
time_yukyu = df_tanzo_k_4["有休日数    "] * 8
df_tanzo_k_4.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_tanzo_k_4.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_tanzo_k_4["欠勤日数    "] * 8
df_tanzo_k_4.insert(5, "欠勤時間", abs_time)
work_time = df_tanzo_k_4["勤務時間    "]
df_tanzo_k_4.insert(6, "勤務時間", work_time)
late_early_time = df_tanzo_k_4["遅早時間    "]
df_tanzo_k_4.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_tanzo_k_4["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_tanzo_k_4.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_tanzo_k_4.insert(17, "実労働時間", real_work_time)
zure_time = df_tanzo_k_4["ズレ時間    "]
df_tanzo_k_4.insert(18, "ズレ時間", zure_time)
overtime = df_tanzo_k_4["残業時間    "] + df_tanzo_k_4["深夜残業    "]
df_tanzo_k_4.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_tanzo_k_4.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_tanzo_k_4["法外休出    "]
df_tanzo_k_4.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_tanzo_k_4.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_tanzo_k_4["法定休出    "]
df_tanzo_k_4.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_tanzo_k_4.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_tanzo_k_4["６０Ｈ超    "]
df_tanzo_k_4.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_tanzo_k_4["代休時間    "] + df_tanzo_k_4["深夜代休    "]
df_tanzo_k_4.insert(31, "代休時間", holiday_time)
df_tanzo_k_4.insert(32, "応援時間", 0)
total_work_time = (
    df_tanzo_k_4["勤務時間    "]
    + df_tanzo_k_4["残業時間    "]
    + df_tanzo_k_4["法外休出    "]
    + df_tanzo_k_4["法定休出    "]
)
df_tanzo_k_4.insert(33, "総労働時間", total_work_time)
basic_salary = df_tanzo_k_4["基 本 給    "] + df_tanzo_k_4["支給額"]
df_tanzo_k_4.insert(35, "基本給", basic_salary)
post_allowance = df_tanzo_k_4["役職手当    "]
df_tanzo_k_4.insert(36, "役職手当", post_allowance)
sales_allowance = df_tanzo_k_4["営業手当    "]
df_tanzo_k_4.insert(37, "営業手当", sales_allowance)
aria_allowance = df_tanzo_k_4["地域手当    "]
df_tanzo_k_4.insert(38, "地域手当", aria_allowance)
spe_allowance = df_tanzo_k_4["特殊手当    "]
df_tanzo_k_4.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_tanzo_k_4["特別技技手当"]
df_tanzo_k_4.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_tanzo_k_4["調整手当    "]
df_tanzo_k_4.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_tanzo_k_4["別居手当    "]
df_tanzo_k_4.insert(42, "別居手当", sep_allowance)
com_allowance = df_tanzo_k_4["通勤手当    "]
df_tanzo_k_4.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_tanzo_k_4.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_tanzo_k_4["残業手当    "]
df_tanzo_k_4.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_tanzo_k_4["休出手当    "]
df_tanzo_k_4.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_tanzo_k_4["深夜勤務手当"]
df_tanzo_k_4.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_tanzo_k_4["交替時差手当"]
df_tanzo_k_4.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_tanzo_k_4["休業手当    "]
df_tanzo_k_4.insert(49, "休業手当", closed_allowance)
closed_deduction = df_tanzo_k_4["休業控除    "]
df_tanzo_k_4.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_tanzo_k_4["代 休 他    "]
df_tanzo_k_4.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_tanzo_k_4["欠勤控除    "] + df_tanzo_k_4["遅早控除    "]
df_tanzo_k_4.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_tanzo_k_4["精 算 分    "]
df_tanzo_k_4.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_tanzo_k_4.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_tanzo_k_4.insert(55, "総支給額", total)
df_tanzo_k_4.insert(56, "応援時間額", 0)
df_tanzo_k_4.insert(57, "役員振替", 0)
df_tanzo_k_4.insert(58, "部門振替", 0)
df_tanzo_k_4.insert(59, "合計", 0)
df_tanzo_k_4 = df_tanzo_k_4.drop("所属2", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("所属3", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("社員ｺｰﾄﾞ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("区分", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("出勤日数    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("有休日数    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("欠勤日数    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("残業時間    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("深夜残業    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("法外休出    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("法定休出    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("代休時間    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("深夜代休    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("６０Ｈ超    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("別居手当    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("通勤手当    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("特別技技手当", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("特殊手当    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("地域手当    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("営業手当    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("役職手当    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("調整手当    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("基 本 給    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("残業手当    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("休出手当    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("深夜勤務手当", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("交替時差手当", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("休業手当    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("代 休 他    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("欠勤控除    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("遅早控除    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("精 算 分    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("支給合計額  ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("休業控除    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("支給額", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("ズレ時間    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("雑費・食事代", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("雑費・衣靴代", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("雑費        ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("受診料・他  ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("雑費・会費等", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("勤務時間    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("遅早時間    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("特休日数    ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.drop("集計区分－２        ", axis=1)
df_tanzo_k_4 = df_tanzo_k_4.sum()
df_tanzo_k_4.to_csv(
    "/content/drive/MyDrive/data/鍛造/D.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 鍛造_間接5
df_tanzo_k_5 = df_tanzo_m.groupby("区分").get_group("間接5")
df_tanzo_k_5 = df_tanzo_k_5.drop("所属1", axis=1)
member = df_tanzo_k_5["所属2"] > 0
df_tanzo_k_5.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_tanzo_k_5["所属3"] > 0
df_tanzo_k_5.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_tanzo_k_5["出勤日数    "]
df_tanzo_k_5.insert(2, "実在籍者", real_member)
time_yukyu = df_tanzo_k_5["有休日数    "] * 8
df_tanzo_k_5.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_tanzo_k_5.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_tanzo_k_5["欠勤日数    "] * 8
df_tanzo_k_5.insert(5, "欠勤時間", abs_time)
work_time = df_tanzo_k_5["勤務時間    "]
df_tanzo_k_5.insert(6, "勤務時間", work_time)
late_early_time = df_tanzo_k_5["遅早時間    "]
df_tanzo_k_5.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_tanzo_k_5["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_tanzo_k_5.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_tanzo_k_5.insert(17, "実労働時間", real_work_time)
zure_time = df_tanzo_k_5["ズレ時間    "]
df_tanzo_k_5.insert(18, "ズレ時間", zure_time)
overtime = df_tanzo_k_5["残業時間    "] + df_tanzo_k_5["深夜残業    "]
df_tanzo_k_5.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_tanzo_k_5.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_tanzo_k_5["法外休出    "]
df_tanzo_k_5.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_tanzo_k_5.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_tanzo_k_5["法定休出    "]
df_tanzo_k_5.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_tanzo_k_5.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_tanzo_k_5["６０Ｈ超    "]
df_tanzo_k_5.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_tanzo_k_5["代休時間    "] + df_tanzo_k_5["深夜代休    "]
df_tanzo_k_5.insert(31, "代休時間", holiday_time)
df_tanzo_k_5.insert(32, "応援時間", 0)
total_work_time = (
    df_tanzo_k_5["勤務時間    "]
    + df_tanzo_k_5["残業時間    "]
    + df_tanzo_k_5["法外休出    "]
    + df_tanzo_k_5["法定休出    "]
)
df_tanzo_k_5.insert(33, "総労働時間", total_work_time)
basic_salary = df_tanzo_k_5["基 本 給    "] + df_tanzo_k_5["支給額"]
df_tanzo_k_5.insert(35, "基本給", basic_salary)
post_allowance = df_tanzo_k_5["役職手当    "]
df_tanzo_k_5.insert(36, "役職手当", post_allowance)
sales_allowance = df_tanzo_k_5["営業手当    "]
df_tanzo_k_5.insert(37, "営業手当", sales_allowance)
aria_allowance = df_tanzo_k_5["地域手当    "]
df_tanzo_k_5.insert(38, "地域手当", aria_allowance)
spe_allowance = df_tanzo_k_5["特殊手当    "]
df_tanzo_k_5.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_tanzo_k_5["特別技技手当"]
df_tanzo_k_5.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_tanzo_k_5["調整手当    "]
df_tanzo_k_5.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_tanzo_k_5["別居手当    "]
df_tanzo_k_5.insert(42, "別居手当", sep_allowance)
com_allowance = df_tanzo_k_5["通勤手当    "]
df_tanzo_k_5.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_tanzo_k_5.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_tanzo_k_5["残業手当    "]
df_tanzo_k_5.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_tanzo_k_5["休出手当    "]
df_tanzo_k_5.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_tanzo_k_5["深夜勤務手当"]
df_tanzo_k_5.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_tanzo_k_5["交替時差手当"]
df_tanzo_k_5.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_tanzo_k_5["休業手当    "]
df_tanzo_k_5.insert(49, "休業手当", closed_allowance)
closed_deduction = df_tanzo_k_5["休業控除    "]
df_tanzo_k_5.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_tanzo_k_5["代 休 他    "]
df_tanzo_k_5.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_tanzo_k_5["欠勤控除    "] + df_tanzo_k_5["遅早控除    "]
df_tanzo_k_5.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_tanzo_k_5["精 算 分    "]
df_tanzo_k_5.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_tanzo_k_5.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_tanzo_k_5.insert(55, "総支給額", total)
df_tanzo_k_5.insert(56, "応援時間額", 0)
df_tanzo_k_5.insert(57, "役員振替", 0)
df_tanzo_k_5.insert(58, "部門振替", 0)
df_tanzo_k_5.insert(59, "合計", 0)
# 不要フィールドの削除
df_tanzo_k_5 = df_tanzo_k_5.drop("所属2", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("所属3", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("社員ｺｰﾄﾞ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("区分", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("出勤日数    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("有休日数    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("欠勤日数    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("残業時間    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("深夜残業    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("法外休出    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("法定休出    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("代休時間    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("深夜代休    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("６０Ｈ超    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("別居手当    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("通勤手当    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("特別技技手当", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("特殊手当    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("地域手当    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("営業手当    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("役職手当    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("調整手当    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("基 本 給    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("残業手当    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("休出手当    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("深夜勤務手当", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("交替時差手当", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("休業手当    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("代 休 他    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("欠勤控除    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("遅早控除    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("精 算 分    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("支給合計額  ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("休業控除    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("支給額", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("ズレ時間    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("雑費・食事代", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("雑費・衣靴代", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("雑費        ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("受診料・他  ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("雑費・会費等", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("勤務時間    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("遅早時間    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("特休日数    ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.drop("集計区分－２        ", axis=1)
df_tanzo_k_5 = df_tanzo_k_5.sum()
df_tanzo_k_5.to_csv(
    "/content/drive/MyDrive/data/鍛造/E.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 鍛造_間接6
df_tanzo_k_6 = df_tanzo_m.groupby("区分").get_group("間接6")
df_tanzo_k_6 = df_tanzo_k_6.drop("所属1", axis=1)
member = df_tanzo_k_6["所属2"] > 0
df_tanzo_k_6.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_tanzo_k_6["所属3"] > 0
df_tanzo_k_6.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_tanzo_k_6["出勤日数    "]
df_tanzo_k_6.insert(2, "実在籍者", real_member)
time_yukyu = df_tanzo_k_6["有休日数    "] * 8
df_tanzo_k_6.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_tanzo_k_6.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_tanzo_k_6["欠勤日数    "] * 8
df_tanzo_k_6.insert(5, "欠勤時間", abs_time)
work_time = df_tanzo_k_6["勤務時間    "]
df_tanzo_k_6.insert(6, "勤務時間", work_time)
late_early_time = df_tanzo_k_6["遅早時間    "]
df_tanzo_k_6.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_tanzo_k_6["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_tanzo_k_6.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_tanzo_k_6.insert(17, "実労働時間", real_work_time)
zure_time = df_tanzo_k_6["ズレ時間    "]
df_tanzo_k_6.insert(18, "ズレ時間", zure_time)
overtime = df_tanzo_k_6["残業時間    "] + df_tanzo_k_6["深夜残業    "]
df_tanzo_k_6.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_tanzo_k_6.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_tanzo_k_6["法外休出    "]
df_tanzo_k_6.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_tanzo_k_6.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_tanzo_k_6["法定休出    "]
df_tanzo_k_6.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_tanzo_k_6.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_tanzo_k_6["６０Ｈ超    "]
df_tanzo_k_6.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_tanzo_k_6["代休時間    "] + df_tanzo_k_6["深夜代休    "]
df_tanzo_k_6.insert(31, "代休時間", holiday_time)
df_tanzo_k_6.insert(32, "応援時間", 0)
total_work_time = (
    df_tanzo_k_6["勤務時間    "]
    + df_tanzo_k_6["残業時間    "]
    + df_tanzo_k_6["法外休出    "]
    + df_tanzo_k_6["法定休出    "]
)
df_tanzo_k_6.insert(33, "総労働時間", total_work_time)
basic_salary = df_tanzo_k_6["基 本 給    "] + df_tanzo_k_6["支給額"]
df_tanzo_k_6.insert(35, "基本給", basic_salary)
post_allowance = df_tanzo_k_6["役職手当    "]
df_tanzo_k_6.insert(36, "役職手当", post_allowance)
sales_allowance = df_tanzo_k_6["営業手当    "]
df_tanzo_k_6.insert(37, "営業手当", sales_allowance)
aria_allowance = df_tanzo_k_6["地域手当    "]
df_tanzo_k_6.insert(38, "地域手当", aria_allowance)
spe_allowance = df_tanzo_k_6["特殊手当    "]
df_tanzo_k_6.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_tanzo_k_6["特別技技手当"]
df_tanzo_k_6.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_tanzo_k_6["調整手当    "]
df_tanzo_k_6.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_tanzo_k_6["別居手当    "]
df_tanzo_k_6.insert(42, "別居手当", sep_allowance)
com_allowance = df_tanzo_k_6["通勤手当    "]
df_tanzo_k_6.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_tanzo_k_6.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_tanzo_k_6["残業手当    "]
df_tanzo_k_6.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_tanzo_k_6["休出手当    "]
df_tanzo_k_6.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_tanzo_k_6["深夜勤務手当"]
df_tanzo_k_6.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_tanzo_k_6["交替時差手当"]
df_tanzo_k_6.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_tanzo_k_6["休業手当    "]
df_tanzo_k_6.insert(49, "休業手当", closed_allowance)
closed_deduction = df_tanzo_k_6["休業控除    "]
df_tanzo_k_6.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_tanzo_k_6["代 休 他    "]
df_tanzo_k_6.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_tanzo_k_6["欠勤控除    "] + df_tanzo_k_6["遅早控除    "]
df_tanzo_k_6.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_tanzo_k_6["精 算 分    "]
df_tanzo_k_6.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_tanzo_k_6.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_tanzo_k_6.insert(55, "総支給額", total)
df_tanzo_k_6.insert(56, "応援時間額", 0)
df_tanzo_k_6.insert(57, "役員振替", 0)
df_tanzo_k_6.insert(58, "部門振替", 0)
df_tanzo_k_6.insert(59, "合計", 0)
# 不要フィールドの削除
df_tanzo_k_6 = df_tanzo_k_6.drop("所属2", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("所属3", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("社員ｺｰﾄﾞ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("区分", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("出勤日数    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("有休日数    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("欠勤日数    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("残業時間    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("深夜残業    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("法外休出    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("法定休出    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("代休時間    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("深夜代休    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("６０Ｈ超    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("別居手当    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("通勤手当    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("特別技技手当", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("特殊手当    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("地域手当    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("営業手当    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("役職手当    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("調整手当    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("基 本 給    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("残業手当    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("休出手当    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("深夜勤務手当", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("交替時差手当", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("休業手当    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("代 休 他    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("欠勤控除    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("遅早控除    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("精 算 分    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("支給合計額  ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("休業控除    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("支給額", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("ズレ時間    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("雑費・食事代", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("雑費・衣靴代", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("雑費        ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("受診料・他  ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("雑費・会費等", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("勤務時間    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("遅早時間    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("特休日数    ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.drop("集計区分－２        ", axis=1)
df_tanzo_k_6 = df_tanzo_k_6.sum()
df_tanzo_k_6.to_csv(
    "/content/drive/MyDrive/data/鍛造/F.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 鍛造_直接1
df_tanzo_t_1 = df_tanzo_m.groupby("区分").get_group("直接1")
df_tanzo_t_1 = df_tanzo_t_1.drop("所属1", axis=1)
member = df_tanzo_t_1["所属2"] > 0
df_tanzo_t_1.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_tanzo_t_1["所属3"] > 0
df_tanzo_t_1.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_tanzo_t_1["出勤日数    "]
df_tanzo_t_1.insert(2, "実在籍者", real_member)
time_yukyu = df_tanzo_t_1["有休日数    "] * 8
df_tanzo_t_1.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_tanzo_t_1.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_tanzo_t_1["欠勤日数    "] * 8
df_tanzo_t_1.insert(5, "欠勤時間", abs_time)
work_time = df_tanzo_t_1["勤務時間    "]
df_tanzo_t_1.insert(6, "勤務時間", work_time)
late_early_time = df_tanzo_t_1["遅早時間    "]
df_tanzo_t_1.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_tanzo_t_1["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_tanzo_t_1.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_tanzo_t_1.insert(17, "実労働時間", real_work_time)
zure_time = df_tanzo_t_1["ズレ時間    "]
df_tanzo_t_1.insert(18, "ズレ時間", zure_time)
overtime = df_tanzo_t_1["残業時間    "] + df_tanzo_t_1["深夜残業    "]
df_tanzo_t_1.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_tanzo_t_1.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_tanzo_t_1["法外休出    "]
df_tanzo_t_1.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_tanzo_t_1.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_tanzo_t_1["法定休出    "]
df_tanzo_t_1.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_tanzo_t_1.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_tanzo_t_1["６０Ｈ超    "]
df_tanzo_t_1.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_tanzo_t_1["代休時間    "] + df_tanzo_t_1["深夜代休    "]
df_tanzo_t_1.insert(31, "代休時間", holiday_time)
df_tanzo_t_1.insert(32, "応援時間", 0)
total_work_time = (
    df_tanzo_t_1["勤務時間    "]
    + df_tanzo_t_1["残業時間    "]
    + df_tanzo_t_1["法外休出    "]
    + df_tanzo_t_1["法定休出    "]
)
df_tanzo_t_1.insert(33, "総労働時間", total_work_time)
basic_salary = df_tanzo_t_1["基 本 給    "] + df_tanzo_t_1["支給額"]
df_tanzo_t_1.insert(35, "基本給", basic_salary)
post_allowance = df_tanzo_t_1["役職手当    "]
df_tanzo_t_1.insert(36, "役職手当", post_allowance)
sales_allowance = df_tanzo_t_1["営業手当    "]
df_tanzo_t_1.insert(37, "営業手当", sales_allowance)
aria_allowance = df_tanzo_t_1["地域手当    "]
df_tanzo_t_1.insert(38, "地域手当", aria_allowance)
spe_allowance = df_tanzo_t_1["特殊手当    "]
df_tanzo_t_1.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_tanzo_t_1["特別技技手当"]
df_tanzo_t_1.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_tanzo_t_1["調整手当    "]
df_tanzo_t_1.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_tanzo_t_1["別居手当    "]
df_tanzo_t_1.insert(42, "別居手当", sep_allowance)
com_allowance = df_tanzo_t_1["通勤手当    "]
df_tanzo_t_1.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_tanzo_t_1.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_tanzo_t_1["残業手当    "]
df_tanzo_t_1.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_tanzo_t_1["休出手当    "]
df_tanzo_t_1.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_tanzo_t_1["深夜勤務手当"]
df_tanzo_t_1.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_tanzo_t_1["交替時差手当"]
df_tanzo_t_1.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_tanzo_t_1["休業手当    "]
df_tanzo_t_1.insert(49, "休業手当", closed_allowance)
closed_deduction = df_tanzo_t_1["休業控除    "]
df_tanzo_t_1.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_tanzo_t_1["代 休 他    "]
df_tanzo_t_1.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_tanzo_t_1["欠勤控除    "] + df_tanzo_t_1["遅早控除    "]
df_tanzo_t_1.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_tanzo_t_1["精 算 分    "]
df_tanzo_t_1.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_tanzo_t_1.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_tanzo_t_1.insert(55, "総支給額", total)
df_tanzo_t_1.insert(56, "応援時間額", 0)
df_tanzo_t_1.insert(57, "役員振替", 0)
df_tanzo_t_1.insert(58, "部門振替", 0)
df_tanzo_t_1.insert(59, "合計", 0)
df_tanzo_t_1 = df_tanzo_t_1.drop("所属2", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("所属3", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("社員ｺｰﾄﾞ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("区分", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("出勤日数    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("有休日数    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("欠勤日数    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("残業時間    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("深夜残業    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("法外休出    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("法定休出    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("代休時間    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("深夜代休    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("６０Ｈ超    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("別居手当    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("通勤手当    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("特別技技手当", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("特殊手当    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("地域手当    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("営業手当    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("役職手当    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("調整手当    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("基 本 給    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("残業手当    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("休出手当    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("深夜勤務手当", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("交替時差手当", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("休業手当    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("代 休 他    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("欠勤控除    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("遅早控除    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("精 算 分    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("支給合計額  ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("休業控除    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("支給額", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("ズレ時間    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("雑費・食事代", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("雑費・衣靴代", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("雑費        ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("受診料・他  ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("雑費・会費等", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("勤務時間    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("遅早時間    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("特休日数    ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.drop("集計区分－２        ", axis=1)
df_tanzo_t_1 = df_tanzo_t_1.sum()
df_tanzo_t_1.to_csv(
    "/content/drive/MyDrive/data/鍛造/G.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 鍛造_直接4
df_tanzo_t_4 = df_tanzo_m.groupby("区分").get_group("直接4")
df_tanzo_t_4 = df_tanzo_t_4.drop("所属1", axis=1)
member = df_tanzo_t_4["所属2"] > 0
df_tanzo_t_4.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_tanzo_t_4["所属3"] > 0
df_tanzo_t_4.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_tanzo_t_4["出勤日数    "]
df_tanzo_t_4.insert(2, "実在籍者", real_member)
time_yukyu = df_tanzo_t_4["有休日数    "] * 8
df_tanzo_t_4.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_tanzo_t_4.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_tanzo_t_4["欠勤日数    "] * 8
df_tanzo_t_4.insert(5, "欠勤時間", abs_time)
work_time = df_tanzo_t_4["勤務時間    "]
df_tanzo_t_4.insert(6, "勤務時間", work_time)
late_early_time = df_tanzo_t_4["遅早時間    "]
df_tanzo_t_4.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_tanzo_t_4["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_tanzo_t_4.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_tanzo_t_4.insert(17, "実労働時間", real_work_time)
zure_time = df_tanzo_t_4["ズレ時間    "]
df_tanzo_t_4.insert(18, "ズレ時間", zure_time)
overtime = df_tanzo_t_4["残業時間    "] + df_tanzo_t_4["深夜残業    "]
df_tanzo_t_4.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_tanzo_t_4.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_tanzo_t_4["法外休出    "]
df_tanzo_t_4.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_tanzo_t_4.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_tanzo_t_4["法定休出    "]
df_tanzo_t_4.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_tanzo_t_4.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_tanzo_t_4["６０Ｈ超    "]
df_tanzo_t_4.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_tanzo_t_4["代休時間    "] + df_tanzo_t_4["深夜代休    "]
df_tanzo_t_4.insert(31, "代休時間", holiday_time)
df_tanzo_t_4.insert(32, "応援時間", 0)
total_work_time = (
    df_tanzo_t_4["勤務時間    "]
    + df_tanzo_t_4["残業時間    "]
    + df_tanzo_t_4["法外休出    "]
    + df_tanzo_t_4["法定休出    "]
)
df_tanzo_t_4.insert(33, "総労働時間", total_work_time)
basic_salary = df_tanzo_t_4["基 本 給    "] + df_tanzo_t_4["支給額"]
df_tanzo_t_4.insert(35, "基本給", basic_salary)
post_allowance = df_tanzo_t_4["役職手当    "]
df_tanzo_t_4.insert(36, "役職手当", post_allowance)
sales_allowance = df_tanzo_t_4["営業手当    "]
df_tanzo_t_4.insert(37, "営業手当", sales_allowance)
aria_allowance = df_tanzo_t_4["地域手当    "]
df_tanzo_t_4.insert(38, "地域手当", aria_allowance)
spe_allowance = df_tanzo_t_4["特殊手当    "]
df_tanzo_t_4.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_tanzo_t_4["特別技技手当"]
df_tanzo_t_4.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_tanzo_t_4["調整手当    "]
df_tanzo_t_4.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_tanzo_t_4["別居手当    "]
df_tanzo_t_4.insert(42, "別居手当", sep_allowance)
com_allowance = df_tanzo_t_4["通勤手当    "]
df_tanzo_t_4.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_tanzo_t_4.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_tanzo_t_4["残業手当    "]
df_tanzo_t_4.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_tanzo_t_4["休出手当    "]
df_tanzo_t_4.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_tanzo_t_4["深夜勤務手当"]
df_tanzo_t_4.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_tanzo_t_4["交替時差手当"]
df_tanzo_t_4.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_tanzo_t_4["休業手当    "]
df_tanzo_t_4.insert(49, "休業手当", closed_allowance)
closed_deduction = df_tanzo_t_4["休業控除    "]
df_tanzo_t_4.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_tanzo_t_4["代 休 他    "]
df_tanzo_t_4.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_tanzo_t_4["欠勤控除    "] + df_tanzo_t_4["遅早控除    "]
df_tanzo_t_4.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_tanzo_t_4["精 算 分    "]
df_tanzo_t_4.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_tanzo_t_4.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_tanzo_t_4.insert(55, "総支給額", total)
df_tanzo_t_4.insert(56, "応援時間額", 0)
df_tanzo_t_4.insert(57, "役員振替", 0)
df_tanzo_t_4.insert(58, "部門振替", 0)
df_tanzo_t_4.insert(59, "合計", 0)
# 不要フィールドの削除
df_tanzo_t_4 = df_tanzo_t_4.drop("所属2", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("所属3", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("社員ｺｰﾄﾞ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("区分", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("出勤日数    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("有休日数    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("欠勤日数    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("残業時間    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("深夜残業    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("法外休出    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("法定休出    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("代休時間    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("深夜代休    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("６０Ｈ超    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("別居手当    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("通勤手当    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("特別技技手当", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("特殊手当    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("地域手当    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("営業手当    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("役職手当    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("調整手当    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("基 本 給    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("残業手当    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("休出手当    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("深夜勤務手当", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("交替時差手当", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("休業手当    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("代 休 他    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("欠勤控除    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("遅早控除    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("精 算 分    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("支給合計額  ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("休業控除    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("支給額", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("ズレ時間    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("雑費・食事代", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("雑費・衣靴代", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("雑費        ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("受診料・他  ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("雑費・会費等", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("勤務時間    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("遅早時間    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("特休日数    ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.drop("集計区分－２        ", axis=1)
df_tanzo_t_4 = df_tanzo_t_4.sum()
df_tanzo_t_4.to_csv(
    "/content/drive/MyDrive/data/鍛造/H.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 切削_間接1
df_sesaku_k_1 = df_sesaku_m.groupby("区分").get_group("間接1")
df_sesaku_k_1 = df_sesaku_k_1.drop("所属1", axis=1)
member = df_sesaku_k_1["所属2"] > 0
df_sesaku_k_1.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_sesaku_k_1["所属3"] > 0
df_sesaku_k_1.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_sesaku_k_1["出勤日数    "]
df_sesaku_k_1.insert(2, "実在籍者", real_member)
time_yukyu = df_sesaku_k_1["有休日数    "] * 8
df_sesaku_k_1.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_sesaku_k_1.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_sesaku_k_1["欠勤日数    "] * 8
df_sesaku_k_1.insert(5, "欠勤時間", abs_time)
work_time = df_sesaku_k_1["勤務時間    "]
df_sesaku_k_1.insert(6, "勤務時間", work_time)
late_early_time = df_sesaku_k_1["遅早時間    "]
df_sesaku_k_1.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_sesaku_k_1["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_sesaku_k_1.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_sesaku_k_1.insert(17, "実労働時間", real_work_time)
zure_time = df_sesaku_k_1["ズレ時間    "]
df_sesaku_k_1.insert(18, "ズレ時間", zure_time)
overtime = df_sesaku_k_1["残業時間    "] + df_sesaku_k_1["深夜残業    "]
df_sesaku_k_1.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_sesaku_k_1.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_sesaku_k_1["法外休出    "]
df_sesaku_k_1.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_sesaku_k_1.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_sesaku_k_1["法定休出    "]
df_sesaku_k_1.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_sesaku_k_1.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_sesaku_k_1["６０Ｈ超    "]
df_sesaku_k_1.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_sesaku_k_1["代休時間    "] + df_sesaku_k_1["深夜代休    "]
df_sesaku_k_1.insert(31, "代休時間", holiday_time)
df_sesaku_k_1.insert(32, "応援時間", 0)
total_work_time = (
    df_sesaku_k_1["勤務時間    "]
    + df_sesaku_k_1["残業時間    "]
    + df_sesaku_k_1["法外休出    "]
    + df_sesaku_k_1["法定休出    "]
)
df_sesaku_k_1.insert(33, "総労働時間", total_work_time)
basic_salary = df_sesaku_k_1["基 本 給    "] + df_sesaku_k_1["支給額"]
df_sesaku_k_1.insert(35, "基本給", basic_salary)
post_allowance = df_sesaku_k_1["役職手当    "]
df_sesaku_k_1.insert(36, "役職手当", post_allowance)
sales_allowance = df_sesaku_k_1["営業手当    "]
df_sesaku_k_1.insert(37, "営業手当", sales_allowance)
aria_allowance = df_sesaku_k_1["地域手当    "]
df_sesaku_k_1.insert(38, "地域手当", aria_allowance)
spe_allowance = df_sesaku_k_1["特殊手当    "]
df_sesaku_k_1.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_sesaku_k_1["特別技技手当"]
df_sesaku_k_1.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_sesaku_k_1["調整手当    "]
df_sesaku_k_1.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_sesaku_k_1["別居手当    "]
df_sesaku_k_1.insert(42, "別居手当", sep_allowance)
com_allowance = df_sesaku_k_1["通勤手当    "]
df_sesaku_k_1.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_sesaku_k_1.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_sesaku_k_1["残業手当    "]
df_sesaku_k_1.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_sesaku_k_1["休出手当    "]
df_sesaku_k_1.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_sesaku_k_1["深夜勤務手当"]
df_sesaku_k_1.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_sesaku_k_1["交替時差手当"]
df_sesaku_k_1.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_sesaku_k_1["休業手当    "]
df_sesaku_k_1.insert(49, "休業手当", closed_allowance)
closed_deduction = df_sesaku_k_1["休業控除    "]
df_sesaku_k_1.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_sesaku_k_1["代 休 他    "]
df_sesaku_k_1.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_sesaku_k_1["欠勤控除    "] + df_sesaku_k_1["遅早控除    "]
df_sesaku_k_1.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_sesaku_k_1["精 算 分    "]
df_sesaku_k_1.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_sesaku_k_1.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_sesaku_k_1.insert(55, "総支給額", total)
df_sesaku_k_1.insert(56, "応援時間額", 0)
df_sesaku_k_1.insert(57, "役員振替", 0)
df_sesaku_k_1.insert(58, "部門振替", 0)
df_sesaku_k_1.insert(59, "合計", 0)
# 不要フィールドの削除
df_sesaku_k_1 = df_sesaku_k_1.drop("所属2", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("所属3", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("社員ｺｰﾄﾞ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("区分", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("出勤日数    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("有休日数    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("欠勤日数    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("残業時間    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("深夜残業    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("法外休出    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("法定休出    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("代休時間    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("深夜代休    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("６０Ｈ超    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("別居手当    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("通勤手当    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("特別技技手当", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("特殊手当    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("地域手当    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("営業手当    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("役職手当    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("調整手当    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("基 本 給    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("残業手当    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("休出手当    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("深夜勤務手当", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("交替時差手当", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("休業手当    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("代 休 他    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("欠勤控除    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("遅早控除    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("精 算 分    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("支給合計額  ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("休業控除    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("支給額", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("ズレ時間    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("雑費・食事代", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("雑費・衣靴代", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("雑費        ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("受診料・他  ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("雑費・会費等", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("勤務時間    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("遅早時間    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("特休日数    ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.drop("集計区分－２        ", axis=1)
df_sesaku_k_1 = df_sesaku_k_1.sum()
df_sesaku_k_1.to_csv(
    "/content/drive/MyDrive/data/切削/A.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
df_sesaku_k_2 = df_sesaku_m.groupby("区分").get_group("間接2")
df_sesaku_k_2 = df_sesaku_k_2.drop("所属1", axis=1)
member = df_sesaku_k_2["所属2"] > 0
df_sesaku_k_2.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_sesaku_k_2["所属3"] > 0
df_sesaku_k_2.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_sesaku_k_2["出勤日数    "]
df_sesaku_k_2.insert(2, "実在籍者", real_member)
time_yukyu = df_sesaku_k_2["有休日数    "] * 8
df_sesaku_k_2.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_sesaku_k_2.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_sesaku_k_2["欠勤日数    "] * 8
df_sesaku_k_2.insert(5, "欠勤時間", abs_time)
work_time = df_sesaku_k_2["勤務時間    "]
df_sesaku_k_2.insert(6, "勤務時間", work_time)
late_early_time = df_sesaku_k_2["遅早時間    "]
df_sesaku_k_2.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_sesaku_k_2["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_sesaku_k_2.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_sesaku_k_2.insert(17, "実労働時間", real_work_time)
zure_time = df_sesaku_k_2["ズレ時間    "]
df_sesaku_k_2.insert(18, "ズレ時間", zure_time)
overtime = df_sesaku_k_2["残業時間    "] + df_sesaku_k_2["深夜残業    "]
df_sesaku_k_2.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_sesaku_k_2.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_sesaku_k_2["法外休出    "]
df_sesaku_k_2.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_sesaku_k_2.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_sesaku_k_2["法定休出    "]
df_sesaku_k_2.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_sesaku_k_2.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_sesaku_k_2["６０Ｈ超    "]
df_sesaku_k_2.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_sesaku_k_2["代休時間    "] + df_sesaku_k_2["深夜代休    "]
df_sesaku_k_2.insert(31, "代休時間", holiday_time)
df_sesaku_k_2.insert(32, "応援時間", 0)
total_work_time = (
    df_sesaku_k_2["勤務時間    "]
    + df_sesaku_k_2["残業時間    "]
    + df_sesaku_k_2["法外休出    "]
    + df_sesaku_k_2["法定休出    "]
)
df_sesaku_k_2.insert(33, "総労働時間", total_work_time)
basic_salary = df_sesaku_k_2["基 本 給    "] + df_sesaku_k_2["支給額"]
df_sesaku_k_2.insert(35, "基本給", basic_salary)
post_allowance = df_sesaku_k_2["役職手当    "]
df_sesaku_k_2.insert(36, "役職手当", post_allowance)
sales_allowance = df_sesaku_k_2["営業手当    "]
df_sesaku_k_2.insert(37, "営業手当", sales_allowance)
aria_allowance = df_sesaku_k_2["地域手当    "]
df_sesaku_k_2.insert(38, "地域手当", aria_allowance)
spe_allowance = df_sesaku_k_2["特殊手当    "]
df_sesaku_k_2.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_sesaku_k_2["特別技技手当"]
df_sesaku_k_2.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_sesaku_k_2["調整手当    "]
df_sesaku_k_2.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_sesaku_k_2["別居手当    "]
df_sesaku_k_2.insert(42, "別居手当", sep_allowance)
com_allowance = df_sesaku_k_2["通勤手当    "]
df_sesaku_k_2.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_sesaku_k_2.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_sesaku_k_2["残業手当    "]
df_sesaku_k_2.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_sesaku_k_2["休出手当    "]
df_sesaku_k_2.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_sesaku_k_2["深夜勤務手当"]
df_sesaku_k_2.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_sesaku_k_2["交替時差手当"]
df_sesaku_k_2.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_sesaku_k_2["休業手当    "]
df_sesaku_k_2.insert(49, "休業手当", closed_allowance)
closed_deduction = df_sesaku_k_2["休業控除    "]
df_sesaku_k_2.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_sesaku_k_2["代 休 他    "]
df_sesaku_k_2.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_sesaku_k_2["欠勤控除    "] + df_sesaku_k_2["遅早控除    "]
df_sesaku_k_2.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_sesaku_k_2["精 算 分    "]
df_sesaku_k_2.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_sesaku_k_2.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_sesaku_k_2.insert(55, "総支給額", total)
df_sesaku_k_2.insert(56, "応援時間額", 0)
df_sesaku_k_2.insert(57, "役員振替", 0)
df_sesaku_k_2.insert(58, "部門振替", 0)
df_sesaku_k_2.insert(59, "合計", 0)
# 不要フィールドの削除
df_sesaku_k_2 = df_sesaku_k_2.drop("所属2", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("所属3", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("社員ｺｰﾄﾞ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("区分", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("出勤日数    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("有休日数    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("欠勤日数    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("残業時間    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("深夜残業    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("法外休出    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("法定休出    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("代休時間    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("深夜代休    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("６０Ｈ超    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("別居手当    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("通勤手当    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("特別技技手当", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("特殊手当    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("地域手当    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("営業手当    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("役職手当    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("調整手当    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("基 本 給    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("残業手当    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("休出手当    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("深夜勤務手当", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("交替時差手当", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("休業手当    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("代 休 他    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("欠勤控除    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("遅早控除    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("精 算 分    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("支給合計額  ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("休業控除    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("支給額", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("ズレ時間    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("雑費・食事代", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("雑費・衣靴代", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("雑費        ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("受診料・他  ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("雑費・会費等", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("勤務時間    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("遅早時間    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("特休日数    ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.drop("集計区分－２        ", axis=1)
df_sesaku_k_2 = df_sesaku_k_2.sum()
df_sesaku_k_2.to_csv(
    "/content/drive/MyDrive/data/切削/B.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 切削_間接4
df_sesaku_k_4 = df_sesaku_m.groupby("区分").get_group("間接4")
df_sesaku_k_4 = df_sesaku_k_4.drop("所属1", axis=1)
member = df_sesaku_k_4["所属2"] > 0
df_sesaku_k_4.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_sesaku_k_4["所属3"] > 0
df_sesaku_k_4.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_sesaku_k_4["出勤日数    "]
df_sesaku_k_4.insert(2, "実在籍者", real_member)
time_yukyu = df_sesaku_k_4["有休日数    "] * 8
df_sesaku_k_4.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_sesaku_k_4.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_sesaku_k_4["欠勤日数    "] * 8
df_sesaku_k_4.insert(5, "欠勤時間", abs_time)
work_time = df_sesaku_k_4["勤務時間    "]
df_sesaku_k_4.insert(6, "勤務時間", work_time)
late_early_time = df_sesaku_k_4["遅早時間    "]
df_sesaku_k_4.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_sesaku_k_4["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_sesaku_k_4.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_sesaku_k_4.insert(17, "実労働時間", real_work_time)
zure_time = df_sesaku_k_4["ズレ時間    "]
df_sesaku_k_4.insert(18, "ズレ時間", zure_time)
overtime = df_sesaku_k_4["残業時間    "] + df_sesaku_k_4["深夜残業    "]
df_sesaku_k_4.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_sesaku_k_4.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_sesaku_k_4["法外休出    "]
df_sesaku_k_4.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_sesaku_k_4.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_sesaku_k_4["法定休出    "]
df_sesaku_k_4.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_sesaku_k_4.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_sesaku_k_4["６０Ｈ超    "]
df_sesaku_k_4.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_sesaku_k_4["代休時間    "] + df_sesaku_k_4["深夜代休    "]
df_sesaku_k_4.insert(31, "代休時間", holiday_time)
df_sesaku_k_4.insert(32, "応援時間", 0)
total_work_time = (
    df_sesaku_k_4["勤務時間    "]
    + df_sesaku_k_4["残業時間    "]
    + df_sesaku_k_4["法外休出    "]
    + df_sesaku_k_4["法定休出    "]
)
df_sesaku_k_4.insert(33, "総労働時間", total_work_time)
basic_salary = df_sesaku_k_4["基 本 給    "] + df_sesaku_k_4["支給額"]
df_sesaku_k_4.insert(35, "基本給", basic_salary)
post_allowance = df_sesaku_k_4["役職手当    "]
df_sesaku_k_4.insert(36, "役職手当", post_allowance)
sales_allowance = df_sesaku_k_4["営業手当    "]
df_sesaku_k_4.insert(37, "営業手当", sales_allowance)
aria_allowance = df_sesaku_k_4["地域手当    "]
df_sesaku_k_4.insert(38, "地域手当", aria_allowance)
spe_allowance = df_sesaku_k_4["特殊手当    "]
df_sesaku_k_4.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_sesaku_k_4["特別技技手当"]
df_sesaku_k_4.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_sesaku_k_4["調整手当    "]
df_sesaku_k_4.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_sesaku_k_4["別居手当    "]
df_sesaku_k_4.insert(42, "別居手当", sep_allowance)
com_allowance = df_sesaku_k_4["通勤手当    "]
df_sesaku_k_4.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_sesaku_k_4.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_sesaku_k_4["残業手当    "]
df_sesaku_k_4.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_sesaku_k_4["休出手当    "]
df_sesaku_k_4.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_sesaku_k_4["深夜勤務手当"]
df_sesaku_k_4.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_sesaku_k_4["交替時差手当"]
df_sesaku_k_4.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_sesaku_k_4["休業手当    "]
df_sesaku_k_4.insert(49, "休業手当", closed_allowance)
closed_deduction = df_sesaku_k_4["休業控除    "]
df_sesaku_k_4.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_sesaku_k_4["代 休 他    "]
df_sesaku_k_4.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_sesaku_k_4["欠勤控除    "] + df_sesaku_k_4["遅早控除    "]
df_sesaku_k_4.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_sesaku_k_4["精 算 分    "]
df_sesaku_k_4.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_sesaku_k_4.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_sesaku_k_4.insert(55, "総支給額", total)
df_sesaku_k_4.insert(56, "応援時間額", 0)
df_sesaku_k_4.insert(57, "役員振替", 0)
df_sesaku_k_4.insert(58, "部門振替", 0)
df_sesaku_k_4.insert(59, "合計", 0)
df_sesaku_k_4 = df_sesaku_k_4.drop("所属2", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("所属3", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("社員ｺｰﾄﾞ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("区分", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("出勤日数    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("有休日数    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("欠勤日数    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("残業時間    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("深夜残業    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("法外休出    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("法定休出    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("代休時間    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("深夜代休    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("６０Ｈ超    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("別居手当    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("通勤手当    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("特別技技手当", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("特殊手当    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("地域手当    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("営業手当    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("役職手当    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("調整手当    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("基 本 給    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("残業手当    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("休出手当    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("深夜勤務手当", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("交替時差手当", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("休業手当    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("代 休 他    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("欠勤控除    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("遅早控除    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("精 算 分    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("支給合計額  ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("休業控除    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("支給額", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("ズレ時間    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("雑費・食事代", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("雑費・衣靴代", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("雑費        ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("受診料・他  ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("雑費・会費等", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("勤務時間    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("遅早時間    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("特休日数    ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.drop("集計区分－２        ", axis=1)
df_sesaku_k_4 = df_sesaku_k_4.sum()
df_sesaku_k_4.to_csv(
    "/content/drive/MyDrive/data/切削/C.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 切削_間接5
df_sesaku_k_5 = df_sesaku_m.groupby("区分").get_group("間接5")
df_sesaku_k_5 = df_sesaku_k_5.drop("所属1", axis=1)
member = df_sesaku_k_5["所属2"] > 0
df_sesaku_k_5.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_sesaku_k_5["所属3"] > 0
df_sesaku_k_5.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_sesaku_k_5["出勤日数    "]
df_sesaku_k_5.insert(2, "実在籍者", real_member)
time_yukyu = df_sesaku_k_5["有休日数    "] * 8
df_sesaku_k_5.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_sesaku_k_5.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_sesaku_k_5["欠勤日数    "] * 8
df_sesaku_k_5.insert(5, "欠勤時間", abs_time)
work_time = df_sesaku_k_5["勤務時間    "]
df_sesaku_k_5.insert(6, "勤務時間", work_time)
late_early_time = df_sesaku_k_5["遅早時間    "]
df_sesaku_k_5.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_sesaku_k_5["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_sesaku_k_5.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_sesaku_k_5.insert(17, "実労働時間", real_work_time)
zure_time = df_sesaku_k_5["ズレ時間    "]
df_sesaku_k_5.insert(18, "ズレ時間", zure_time)
overtime = df_sesaku_k_5["残業時間    "] + df_sesaku_k_5["深夜残業    "]
df_sesaku_k_5.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_sesaku_k_5.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_sesaku_k_5["法外休出    "]
df_sesaku_k_5.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_sesaku_k_5.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_sesaku_k_5["法定休出    "]
df_sesaku_k_5.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_sesaku_k_5.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_sesaku_k_5["６０Ｈ超    "]
df_sesaku_k_5.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_sesaku_k_5["代休時間    "] + df_sesaku_k_5["深夜代休    "]
df_sesaku_k_5.insert(31, "代休時間", holiday_time)
df_sesaku_k_5.insert(32, "応援時間", 0)
total_work_time = (
    df_sesaku_k_5["勤務時間    "]
    + df_sesaku_k_5["残業時間    "]
    + df_sesaku_k_5["法外休出    "]
    + df_sesaku_k_5["法定休出    "]
)
df_sesaku_k_5.insert(33, "総労働時間", total_work_time)
basic_salary = df_sesaku_k_5["基 本 給    "] + df_sesaku_k_5["支給額"]
df_sesaku_k_5.insert(35, "基本給", basic_salary)
post_allowance = df_sesaku_k_5["役職手当    "]
df_sesaku_k_5.insert(36, "役職手当", post_allowance)
sales_allowance = df_sesaku_k_5["営業手当    "]
df_sesaku_k_5.insert(37, "営業手当", sales_allowance)
aria_allowance = df_sesaku_k_5["地域手当    "]
df_sesaku_k_5.insert(38, "地域手当", aria_allowance)
spe_allowance = df_sesaku_k_5["特殊手当    "]
df_sesaku_k_5.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_sesaku_k_5["特別技技手当"]
df_sesaku_k_5.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_sesaku_k_5["調整手当    "]
df_sesaku_k_5.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_sesaku_k_5["別居手当    "]
df_sesaku_k_5.insert(42, "別居手当", sep_allowance)
com_allowance = df_sesaku_k_5["通勤手当    "]
df_sesaku_k_5.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_sesaku_k_5.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_sesaku_k_5["残業手当    "]
df_sesaku_k_5.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_sesaku_k_5["休出手当    "]
df_sesaku_k_5.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_sesaku_k_5["深夜勤務手当"]
df_sesaku_k_5.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_sesaku_k_5["交替時差手当"]
df_sesaku_k_5.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_sesaku_k_5["休業手当    "]
df_sesaku_k_5.insert(49, "休業手当", closed_allowance)
closed_deduction = df_sesaku_k_5["休業控除    "]
df_sesaku_k_5.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_sesaku_k_5["代 休 他    "]
df_sesaku_k_5.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_sesaku_k_5["欠勤控除    "] + df_sesaku_k_5["遅早控除    "]
df_sesaku_k_5.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_sesaku_k_5["精 算 分    "]
df_sesaku_k_5.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_sesaku_k_5.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_sesaku_k_5.insert(55, "総支給額", total)
df_sesaku_k_5.insert(56, "応援時間額", 0)
df_sesaku_k_5.insert(57, "役員振替", 0)
df_sesaku_k_5.insert(58, "部門振替", 0)
df_sesaku_k_5.insert(59, "合計", 0)
# 不要フィールドの削除
df_sesaku_k_5 = df_sesaku_k_5.drop("所属2", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("所属3", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("社員ｺｰﾄﾞ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("区分", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("出勤日数    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("有休日数    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("欠勤日数    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("残業時間    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("深夜残業    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("法外休出    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("法定休出    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("代休時間    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("深夜代休    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("６０Ｈ超    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("別居手当    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("通勤手当    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("特別技技手当", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("特殊手当    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("地域手当    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("営業手当    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("役職手当    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("調整手当    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("基 本 給    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("残業手当    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("休出手当    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("深夜勤務手当", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("交替時差手当", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("休業手当    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("代 休 他    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("欠勤控除    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("遅早控除    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("精 算 分    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("支給合計額  ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("休業控除    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("支給額", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("ズレ時間    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("雑費・食事代", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("雑費・衣靴代", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("雑費        ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("受診料・他  ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("雑費・会費等", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("勤務時間    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("遅早時間    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("特休日数    ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.drop("集計区分－２        ", axis=1)
df_sesaku_k_5 = df_sesaku_k_5.sum()
df_sesaku_k_5.to_csv(
    "/content/drive/MyDrive/data/切削/D.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 切削_間接6
df_sesaku_k_6 = df_sesaku_m.groupby("区分").get_group("間接6")
df_sesaku_k_6 = df_sesaku_k_6.drop("所属1", axis=1)
member = df_sesaku_k_6["所属2"] > 0
df_sesaku_k_6.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_sesaku_k_6["所属3"] > 0
df_sesaku_k_6.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_sesaku_k_6["出勤日数    "]
df_sesaku_k_6.insert(2, "実在籍者", real_member)
time_yukyu = df_sesaku_k_6["有休日数    "] * 8
df_sesaku_k_6.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_sesaku_k_6.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_sesaku_k_6["欠勤日数    "] * 8
df_sesaku_k_6.insert(5, "欠勤時間", abs_time)
work_time = df_sesaku_k_6["勤務時間    "]
df_sesaku_k_6.insert(6, "勤務時間", work_time)
late_early_time = df_sesaku_k_6["遅早時間    "]
df_sesaku_k_6.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_sesaku_k_6["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_sesaku_k_6.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_sesaku_k_6.insert(17, "実労働時間", real_work_time)
zure_time = df_sesaku_k_6["ズレ時間    "]
df_sesaku_k_6.insert(18, "ズレ時間", zure_time)
overtime = df_sesaku_k_6["残業時間    "] + df_sesaku_k_6["深夜残業    "]
df_sesaku_k_6.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_sesaku_k_6.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_sesaku_k_6["法外休出    "]
df_sesaku_k_6.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_sesaku_k_6.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_sesaku_k_6["法定休出    "]
df_sesaku_k_6.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_sesaku_k_6.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_sesaku_k_6["６０Ｈ超    "]
df_sesaku_k_6.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_sesaku_k_6["代休時間    "] + df_sesaku_k_6["深夜代休    "]
df_sesaku_k_6.insert(31, "代休時間", holiday_time)
df_sesaku_k_6.insert(32, "応援時間", 0)
total_work_time = (
    df_sesaku_k_6["勤務時間    "]
    + df_sesaku_k_6["残業時間    "]
    + df_sesaku_k_6["法外休出    "]
    + df_sesaku_k_6["法定休出    "]
)
df_sesaku_k_6.insert(33, "総労働時間", total_work_time)
basic_salary = df_sesaku_k_6["基 本 給    "] + df_sesaku_k_6["支給額"]
df_sesaku_k_6.insert(35, "基本給", basic_salary)
post_allowance = df_sesaku_k_6["役職手当    "]
df_sesaku_k_6.insert(36, "役職手当", post_allowance)
sales_allowance = df_sesaku_k_6["営業手当    "]
df_sesaku_k_6.insert(37, "営業手当", sales_allowance)
aria_allowance = df_sesaku_k_6["地域手当    "]
df_sesaku_k_6.insert(38, "地域手当", aria_allowance)
spe_allowance = df_sesaku_k_6["特殊手当    "]
df_sesaku_k_6.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_sesaku_k_6["特別技技手当"]
df_sesaku_k_6.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_sesaku_k_6["調整手当    "]
df_sesaku_k_6.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_sesaku_k_6["別居手当    "]
df_sesaku_k_6.insert(42, "別居手当", sep_allowance)
com_allowance = df_sesaku_k_6["通勤手当    "]
df_sesaku_k_6.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_sesaku_k_6.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_sesaku_k_6["残業手当    "]
df_sesaku_k_6.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_sesaku_k_6["休出手当    "]
df_sesaku_k_6.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_sesaku_k_6["深夜勤務手当"]
df_sesaku_k_6.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_sesaku_k_6["交替時差手当"]
df_sesaku_k_6.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_sesaku_k_6["休業手当    "]
df_sesaku_k_6.insert(49, "休業手当", closed_allowance)
closed_deduction = df_sesaku_k_6["休業控除    "]
df_sesaku_k_6.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_sesaku_k_6["代 休 他    "]
df_sesaku_k_6.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_sesaku_k_6["欠勤控除    "] + df_sesaku_k_6["遅早控除    "]
df_sesaku_k_6.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_sesaku_k_6["精 算 分    "]
df_sesaku_k_6.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_sesaku_k_6.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_sesaku_k_6.insert(55, "総支給額", total)
df_sesaku_k_6.insert(56, "応援時間額", 0)
df_sesaku_k_6.insert(57, "役員振替", 0)
df_sesaku_k_6.insert(58, "部門振替", 0)
df_sesaku_k_6.insert(59, "合計", 0)
# 不要フィールドの削除
df_sesaku_k_6 = df_sesaku_k_6.drop("所属2", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("所属3", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("社員ｺｰﾄﾞ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("区分", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("出勤日数    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("有休日数    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("欠勤日数    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("残業時間    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("深夜残業    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("法外休出    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("法定休出    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("代休時間    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("深夜代休    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("６０Ｈ超    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("別居手当    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("通勤手当    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("特別技技手当", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("特殊手当    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("地域手当    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("営業手当    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("役職手当    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("調整手当    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("基 本 給    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("残業手当    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("休出手当    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("深夜勤務手当", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("交替時差手当", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("休業手当    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("代 休 他    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("欠勤控除    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("遅早控除    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("精 算 分    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("支給合計額  ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("休業控除    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("支給額", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("ズレ時間    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("雑費・食事代", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("雑費・衣靴代", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("雑費        ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("受診料・他  ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("雑費・会費等", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("勤務時間    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("遅早時間    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("特休日数    ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.drop("集計区分－２        ", axis=1)
df_sesaku_k_6 = df_sesaku_k_6.sum()
df_sesaku_k_6.to_csv(
    "/content/drive/MyDrive/data/切削/E.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 切削_直接1
df_sesaku_t_1 = df_sesaku_m.groupby("区分").get_group("直接1")
df_sesaku_t_1 = df_sesaku_t_1.drop("所属1", axis=1)
member = df_sesaku_t_1["所属2"] > 0
df_sesaku_t_1.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_sesaku_t_1["所属3"] > 0
df_sesaku_t_1.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_sesaku_t_1["出勤日数    "]
df_sesaku_t_1.insert(2, "実在籍者", real_member)
time_yukyu = df_sesaku_t_1["有休日数    "] * 8
df_sesaku_t_1.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_sesaku_t_1.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_sesaku_t_1["欠勤日数    "] * 8
df_sesaku_t_1.insert(5, "欠勤時間", abs_time)
work_time = df_sesaku_t_1["勤務時間    "]
df_sesaku_t_1.insert(6, "勤務時間", work_time)
late_early_time = df_sesaku_t_1["遅早時間    "]
df_sesaku_t_1.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_sesaku_t_1["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_sesaku_t_1.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_sesaku_t_1.insert(17, "実労働時間", real_work_time)
zure_time = df_sesaku_t_1["ズレ時間    "]
df_sesaku_t_1.insert(18, "ズレ時間", zure_time)
overtime = df_sesaku_t_1["残業時間    "] + df_sesaku_t_1["深夜残業    "]
df_sesaku_t_1.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_sesaku_t_1.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_sesaku_t_1["法外休出    "]
df_sesaku_t_1.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_sesaku_t_1.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_sesaku_t_1["法定休出    "]
df_sesaku_t_1.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_sesaku_t_1.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_sesaku_t_1["６０Ｈ超    "]
df_sesaku_t_1.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_sesaku_t_1["代休時間    "] + df_sesaku_t_1["深夜代休    "]
df_sesaku_t_1.insert(31, "代休時間", holiday_time)
df_sesaku_t_1.insert(32, "応援時間", 0)
total_work_time = (
    df_sesaku_t_1["勤務時間    "]
    + df_sesaku_t_1["残業時間    "]
    + df_sesaku_t_1["法外休出    "]
    + df_sesaku_t_1["法定休出    "]
)
df_sesaku_t_1.insert(33, "総労働時間", total_work_time)
basic_salary = df_sesaku_t_1["基 本 給    "] + df_sesaku_t_1["支給額"]
df_sesaku_t_1.insert(35, "基本給", basic_salary)
post_allowance = df_sesaku_t_1["役職手当    "]
df_sesaku_t_1.insert(36, "役職手当", post_allowance)
sales_allowance = df_sesaku_t_1["営業手当    "]
df_sesaku_t_1.insert(37, "営業手当", sales_allowance)
aria_allowance = df_sesaku_t_1["地域手当    "]
df_sesaku_t_1.insert(38, "地域手当", aria_allowance)
spe_allowance = df_sesaku_t_1["特殊手当    "]
df_sesaku_t_1.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_sesaku_t_1["特別技技手当"]
df_sesaku_t_1.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_sesaku_t_1["調整手当    "]
df_sesaku_t_1.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_sesaku_t_1["別居手当    "]
df_sesaku_t_1.insert(42, "別居手当", sep_allowance)
com_allowance = df_sesaku_t_1["通勤手当    "]
df_sesaku_t_1.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_sesaku_t_1.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_sesaku_t_1["残業手当    "]
df_sesaku_t_1.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_sesaku_t_1["休出手当    "]
df_sesaku_t_1.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_sesaku_t_1["深夜勤務手当"]
df_sesaku_t_1.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_sesaku_t_1["交替時差手当"]
df_sesaku_t_1.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_sesaku_t_1["休業手当    "]
df_sesaku_t_1.insert(49, "休業手当", closed_allowance)
closed_deduction = df_sesaku_t_1["休業控除    "]
df_sesaku_t_1.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_sesaku_t_1["代 休 他    "]
df_sesaku_t_1.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_sesaku_t_1["欠勤控除    "] + df_sesaku_t_1["遅早控除    "]
df_sesaku_t_1.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_sesaku_t_1["精 算 分    "]
df_sesaku_t_1.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_sesaku_t_1.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_sesaku_t_1.insert(55, "総支給額", total)
df_sesaku_t_1.insert(56, "応援時間額", 0)
df_sesaku_t_1.insert(57, "役員振替", 0)
df_sesaku_t_1.insert(58, "部門振替", 0)
df_sesaku_t_1.insert(59, "合計", 0)
# 不要フィールドの削除
df_sesaku_t_1 = df_sesaku_t_1.drop("所属2", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("所属3", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("社員ｺｰﾄﾞ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("区分", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("出勤日数    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("有休日数    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("欠勤日数    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("残業時間    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("深夜残業    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("法外休出    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("法定休出    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("代休時間    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("深夜代休    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("６０Ｈ超    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("別居手当    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("通勤手当    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("特別技技手当", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("特殊手当    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("地域手当    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("営業手当    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("役職手当    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("調整手当    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("基 本 給    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("残業手当    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("休出手当    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("深夜勤務手当", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("交替時差手当", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("休業手当    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("代 休 他    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("欠勤控除    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("遅早控除    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("精 算 分    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("支給合計額  ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("休業控除    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("支給額", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("ズレ時間    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("雑費・食事代", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("雑費・衣靴代", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("雑費        ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("受診料・他  ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("雑費・会費等", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("勤務時間    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("遅早時間    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("特休日数    ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.drop("集計区分－２        ", axis=1)
df_sesaku_t_1 = df_sesaku_t_1.sum()
df_sesaku_t_1.to_csv(
    "/content/drive/MyDrive/data/切削/F.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 切削_直接2
df_sesaku_t_2 = df_sesaku_m.groupby("区分").get_group("直接2")
df_sesaku_t_2 = df_sesaku_t_2.drop("所属1", axis=1)
member = df_sesaku_t_2["所属2"] > 0
df_sesaku_t_2.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_sesaku_t_2["所属3"] > 0
df_sesaku_t_2.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_sesaku_t_2["出勤日数    "]
df_sesaku_t_2.insert(2, "実在籍者", real_member)
time_yukyu = df_sesaku_t_2["有休日数    "] * 8
df_sesaku_t_2.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_sesaku_t_2.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_sesaku_t_2["欠勤日数    "] * 8
df_sesaku_t_2.insert(5, "欠勤時間", abs_time)
work_time = df_sesaku_t_2["勤務時間    "]
df_sesaku_t_2.insert(6, "勤務時間", work_time)
late_early_time = df_sesaku_t_2["遅早時間    "]
df_sesaku_t_2.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_sesaku_t_2["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_sesaku_t_2.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_sesaku_t_2.insert(17, "実労働時間", real_work_time)
zure_time = df_sesaku_t_2["ズレ時間    "]
df_sesaku_t_2.insert(18, "ズレ時間", zure_time)
overtime = df_sesaku_t_2["残業時間    "] + df_sesaku_t_2["深夜残業    "]
df_sesaku_t_2.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_sesaku_t_2.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_sesaku_t_2["法外休出    "]
df_sesaku_t_2.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_sesaku_t_2.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_sesaku_t_2["法定休出    "]
df_sesaku_t_2.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_sesaku_t_2.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_sesaku_t_2["６０Ｈ超    "]
df_sesaku_t_2.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_sesaku_t_2["代休時間    "] + df_sesaku_t_2["深夜代休    "]
df_sesaku_t_2.insert(31, "代休時間", holiday_time)
df_sesaku_t_2.insert(32, "応援時間", 0)
total_work_time = (
    df_sesaku_t_2["勤務時間    "]
    + df_sesaku_t_2["残業時間    "]
    + df_sesaku_t_2["法外休出    "]
    + df_sesaku_t_2["法定休出    "]
)
df_sesaku_t_2.insert(33, "総労働時間", total_work_time)
basic_salary = df_sesaku_t_2["基 本 給    "] + df_sesaku_t_2["支給額"]
df_sesaku_t_2.insert(35, "基本給", basic_salary)
post_allowance = df_sesaku_t_2["役職手当    "]
df_sesaku_t_2.insert(36, "役職手当", post_allowance)
sales_allowance = df_sesaku_t_2["営業手当    "]
df_sesaku_t_2.insert(37, "営業手当", sales_allowance)
aria_allowance = df_sesaku_t_2["地域手当    "]
df_sesaku_t_2.insert(38, "地域手当", aria_allowance)
spe_allowance = df_sesaku_t_2["特殊手当    "]
df_sesaku_t_2.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_sesaku_t_2["特別技技手当"]
df_sesaku_t_2.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_sesaku_t_2["調整手当    "]
df_sesaku_t_2.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_sesaku_t_2["別居手当    "]
df_sesaku_t_2.insert(42, "別居手当", sep_allowance)
com_allowance = df_sesaku_t_2["通勤手当    "]
df_sesaku_t_2.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_sesaku_t_2.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_sesaku_t_2["残業手当    "]
df_sesaku_t_2.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_sesaku_t_2["休出手当    "]
df_sesaku_t_2.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_sesaku_t_2["深夜勤務手当"]
df_sesaku_t_2.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_sesaku_t_2["交替時差手当"]
df_sesaku_t_2.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_sesaku_t_2["休業手当    "]
df_sesaku_t_2.insert(49, "休業手当", closed_allowance)
closed_deduction = df_sesaku_t_2["休業控除    "]
df_sesaku_t_2.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_sesaku_t_2["代 休 他    "]
df_sesaku_t_2.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_sesaku_t_2["欠勤控除    "] + df_sesaku_t_2["遅早控除    "]
df_sesaku_t_2.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_sesaku_t_2["精 算 分    "]
df_sesaku_t_2.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_sesaku_t_2.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_sesaku_t_2.insert(55, "総支給額", total)
df_sesaku_t_2.insert(56, "応援時間額", 0)
df_sesaku_t_2.insert(57, "役員振替", 0)
df_sesaku_t_2.insert(58, "部門振替", 0)
df_sesaku_t_2.insert(59, "合計", 0)
# 不要フィールドの削除
df_sesaku_t_2 = df_sesaku_t_2.drop("所属2", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("所属3", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("社員ｺｰﾄﾞ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("区分", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("出勤日数    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("有休日数    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("欠勤日数    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("残業時間    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("深夜残業    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("法外休出    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("法定休出    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("代休時間    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("深夜代休    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("６０Ｈ超    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("別居手当    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("通勤手当    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("特別技技手当", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("特殊手当    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("地域手当    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("営業手当    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("役職手当    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("調整手当    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("基 本 給    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("残業手当    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("休出手当    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("深夜勤務手当", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("交替時差手当", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("休業手当    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("代 休 他    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("欠勤控除    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("遅早控除    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("精 算 分    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("支給合計額  ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("休業控除    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("支給額", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("ズレ時間    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("雑費・食事代", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("雑費・衣靴代", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("雑費        ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("受診料・他  ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("雑費・会費等", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("勤務時間    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("遅早時間    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("特休日数    ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.drop("集計区分－２        ", axis=1)
df_sesaku_t_2 = df_sesaku_t_2.sum()
df_sesaku_t_2.to_csv(
    "/content/drive/MyDrive/data/切削/G.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 切削_直接4
df_sesaku_t_4 = df_sesaku_m.groupby("区分").get_group("直接4")
df_sesaku_t_4 = df_sesaku_t_4.drop("所属1", axis=1)
member = df_sesaku_t_4["所属2"] > 0
df_sesaku_t_4.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_sesaku_t_4["所属3"] > 0
df_sesaku_t_4.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_sesaku_t_4["出勤日数    "]
df_sesaku_t_4.insert(2, "実在籍者", real_member)
time_yukyu = df_sesaku_t_4["有休日数    "] * 8
df_sesaku_t_4.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_sesaku_t_4.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_sesaku_t_4["欠勤日数    "] * 8
df_sesaku_t_4.insert(5, "欠勤時間", abs_time)
work_time = df_sesaku_t_4["勤務時間    "]
df_sesaku_t_4.insert(6, "勤務時間", work_time)
late_early_time = df_sesaku_t_4["遅早時間    "]
df_sesaku_t_4.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_sesaku_t_4["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_sesaku_t_4.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_sesaku_t_4.insert(17, "実労働時間", real_work_time)
zure_time = df_sesaku_t_4["ズレ時間    "]
df_sesaku_t_4.insert(18, "ズレ時間", zure_time)
overtime = df_sesaku_t_4["残業時間    "] + df_sesaku_t_4["深夜残業    "]
df_sesaku_t_4.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_sesaku_t_4.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_sesaku_t_4["法外休出    "]
df_sesaku_t_4.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_sesaku_t_4.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_sesaku_t_4["法定休出    "]
df_sesaku_t_4.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_sesaku_t_4.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_sesaku_t_4["６０Ｈ超    "]
df_sesaku_t_4.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_sesaku_t_4["代休時間    "] + df_sesaku_t_4["深夜代休    "]
df_sesaku_t_4.insert(31, "代休時間", holiday_time)
df_sesaku_t_4.insert(32, "応援時間", 0)
total_work_time = (
    df_sesaku_t_4["勤務時間    "]
    + df_sesaku_t_4["残業時間    "]
    + df_sesaku_t_4["法外休出    "]
    + df_sesaku_t_4["法定休出    "]
)
df_sesaku_t_4.insert(33, "総労働時間", total_work_time)
basic_salary = df_sesaku_t_4["基 本 給    "] + df_sesaku_t_4["支給額"]
df_sesaku_t_4.insert(35, "基本給", basic_salary)
post_allowance = df_sesaku_t_4["役職手当    "]
df_sesaku_t_4.insert(36, "役職手当", post_allowance)
sales_allowance = df_sesaku_t_4["営業手当    "]
df_sesaku_t_4.insert(37, "営業手当", sales_allowance)
aria_allowance = df_sesaku_t_4["地域手当    "]
df_sesaku_t_4.insert(38, "地域手当", aria_allowance)
spe_allowance = df_sesaku_t_4["特殊手当    "]
df_sesaku_t_4.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_sesaku_t_4["特別技技手当"]
df_sesaku_t_4.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_sesaku_t_4["調整手当    "]
df_sesaku_t_4.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_sesaku_t_4["別居手当    "]
df_sesaku_t_4.insert(42, "別居手当", sep_allowance)
com_allowance = df_sesaku_t_4["通勤手当    "]
df_sesaku_t_4.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_sesaku_t_4.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_sesaku_t_4["残業手当    "]
df_sesaku_t_4.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_sesaku_t_4["休出手当    "]
df_sesaku_t_4.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_sesaku_t_4["深夜勤務手当"]
df_sesaku_t_4.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_sesaku_t_4["交替時差手当"]
df_sesaku_t_4.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_sesaku_t_4["休業手当    "]
df_sesaku_t_4.insert(49, "休業手当", closed_allowance)
closed_deduction = df_sesaku_t_4["休業控除    "]
df_sesaku_t_4.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_sesaku_t_4["代 休 他    "]
df_sesaku_t_4.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_sesaku_t_4["欠勤控除    "] + df_sesaku_t_4["遅早控除    "]
df_sesaku_t_4.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_sesaku_t_4["精 算 分    "]
df_sesaku_t_4.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_sesaku_t_4.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_sesaku_t_4.insert(55, "総支給額", total)
df_sesaku_t_4.insert(56, "応援時間額", 0)
df_sesaku_t_4.insert(57, "役員振替", 0)
df_sesaku_t_4.insert(58, "部門振替", 0)
df_sesaku_t_4.insert(59, "合計", 0)
# 不要フィールドの削除
df_sesaku_t_4 = df_sesaku_t_4.drop("所属2", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("所属3", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("社員ｺｰﾄﾞ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("区分", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("出勤日数    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("有休日数    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("欠勤日数    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("残業時間    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("深夜残業    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("法外休出    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("法定休出    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("代休時間    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("深夜代休    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("６０Ｈ超    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("別居手当    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("通勤手当    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("特別技技手当", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("特殊手当    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("地域手当    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("営業手当    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("役職手当    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("調整手当    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("基 本 給    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("残業手当    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("休出手当    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("深夜勤務手当", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("交替時差手当", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("休業手当    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("代 休 他    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("欠勤控除    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("遅早控除    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("精 算 分    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("支給合計額  ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("休業控除    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("支給額", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("ズレ時間    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("雑費・食事代", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("雑費・衣靴代", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("雑費        ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("受診料・他  ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("雑費・会費等", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("勤務時間    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("遅早時間    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("特休日数    ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.drop("集計区分－２        ", axis=1)
df_sesaku_t_4 = df_sesaku_t_4.sum()
df_sesaku_t_4.to_csv(
    "/content/drive/MyDrive/data/切削/H.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# AC_間接1
df_ac_k_1 = df_ac_m.groupby("区分").get_group("間接1")
df_ac_k_1 = df_ac_k_1.drop("所属1", axis=1)
member = df_ac_k_1["所属2"] > 0
df_ac_k_1.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_ac_k_1["所属3"] > 0
df_ac_k_1.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_ac_k_1["出勤日数    "]
df_ac_k_1.insert(2, "実在籍者", real_member)
time_yukyu = df_ac_k_1["有休日数    "] * 8
df_ac_k_1.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_ac_k_1.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_ac_k_1["欠勤日数    "] * 8
df_ac_k_1.insert(5, "欠勤時間", abs_time)
work_time = df_ac_k_1["勤務時間    "]
df_ac_k_1.insert(6, "勤務時間", work_time)
late_early_time = df_ac_k_1["遅早時間    "]
df_ac_k_1.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_ac_k_1["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_ac_k_1.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_ac_k_1.insert(17, "実労働時間", real_work_time)
zure_time = df_ac_k_1["ズレ時間    "]
df_ac_k_1.insert(18, "ズレ時間", zure_time)
overtime = df_ac_k_1["残業時間    "] + df_ac_k_1["深夜残業    "]
df_ac_k_1.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_ac_k_1.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_ac_k_1["法外休出    "]
df_ac_k_1.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_ac_k_1.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_ac_k_1["法定休出    "]
df_ac_k_1.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_ac_k_1.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_ac_k_1["６０Ｈ超    "]
df_ac_k_1.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_ac_k_1["代休時間    "] + df_ac_k_1["深夜代休    "]
df_ac_k_1.insert(31, "代休時間", holiday_time)
df_ac_k_1.insert(32, "応援時間", 0)
total_work_time = (
    df_ac_k_1["勤務時間    "]
    + df_ac_k_1["残業時間    "]
    + df_ac_k_1["法外休出    "]
    + df_ac_k_1["法定休出    "]
)
df_ac_k_1.insert(33, "総労働時間", total_work_time)
basic_salary = df_ac_k_1["基 本 給    "] + df_ac_k_1["支給額"]
df_ac_k_1.insert(35, "基本給", basic_salary)
post_allowance = df_ac_k_1["役職手当    "]
df_ac_k_1.insert(36, "役職手当", post_allowance)
sales_allowance = df_ac_k_1["営業手当    "]
df_ac_k_1.insert(37, "営業手当", sales_allowance)
aria_allowance = df_ac_k_1["地域手当    "]
df_ac_k_1.insert(38, "地域手当", aria_allowance)
spe_allowance = df_ac_k_1["特殊手当    "]
df_ac_k_1.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_ac_k_1["特別技技手当"]
df_ac_k_1.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_ac_k_1["調整手当    "]
df_ac_k_1.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_ac_k_1["別居手当    "]
df_ac_k_1.insert(42, "別居手当", sep_allowance)
com_allowance = df_ac_k_1["通勤手当    "]
df_ac_k_1.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_ac_k_1.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_ac_k_1["残業手当    "]
df_ac_k_1.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_ac_k_1["休出手当    "]
df_ac_k_1.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_ac_k_1["深夜勤務手当"]
df_ac_k_1.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_ac_k_1["交替時差手当"]
df_ac_k_1.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_ac_k_1["休業手当    "]
df_ac_k_1.insert(49, "休業手当", closed_allowance)
closed_deduction = df_ac_k_1["休業控除    "]
df_ac_k_1.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_ac_k_1["代 休 他    "]
df_ac_k_1.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_ac_k_1["欠勤控除    "] + df_ac_k_1["遅早控除    "]
df_ac_k_1.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_ac_k_1["精 算 分    "]
df_ac_k_1.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_ac_k_1.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_ac_k_1.insert(55, "総支給額", total)
df_ac_k_1.insert(56, "応援時間額", 0)
df_ac_k_1.insert(57, "役員振替", 0)
df_ac_k_1.insert(58, "部門振替", 0)
df_ac_k_1.insert(59, "合計", 0)
# 不要フィールドの削除
df_ac_k_1 = df_ac_k_1.drop("所属2", axis=1)
df_ac_k_1 = df_ac_k_1.drop("所属3", axis=1)
df_ac_k_1 = df_ac_k_1.drop("社員ｺｰﾄﾞ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("区分", axis=1)
df_ac_k_1 = df_ac_k_1.drop("出勤日数    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("有休日数    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("欠勤日数    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("残業時間    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("深夜残業    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("法外休出    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("法定休出    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("代休時間    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("深夜代休    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("６０Ｈ超    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("別居手当    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("通勤手当    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("特別技技手当", axis=1)
df_ac_k_1 = df_ac_k_1.drop("特殊手当    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("地域手当    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("営業手当    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("役職手当    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("調整手当    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("基 本 給    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("残業手当    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("休出手当    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("深夜勤務手当", axis=1)
df_ac_k_1 = df_ac_k_1.drop("交替時差手当", axis=1)
df_ac_k_1 = df_ac_k_1.drop("休業手当    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("代 休 他    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("欠勤控除    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("遅早控除    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("精 算 分    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("支給合計額  ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("休業控除    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("支給額", axis=1)
df_ac_k_1 = df_ac_k_1.drop("ズレ時間    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("雑費・食事代", axis=1)
df_ac_k_1 = df_ac_k_1.drop("雑費・衣靴代", axis=1)
df_ac_k_1 = df_ac_k_1.drop("雑費        ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("受診料・他  ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("雑費・会費等", axis=1)
df_ac_k_1 = df_ac_k_1.drop("勤務時間    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("遅早時間    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("特休日数    ", axis=1)
df_ac_k_1 = df_ac_k_1.drop("集計区分－２        ", axis=1)
df_ac_k_1 = df_ac_k_1.sum()
df_ac_k_1.to_csv(
    "/content/drive/MyDrive/data/AC/A.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# AC_間接2
df_ac_k_2 = df_ac_m.groupby("区分").get_group("間接2")
df_ac_k_2 = df_ac_k_2.drop("所属1", axis=1)
member = df_ac_k_2["所属2"] > 0
df_ac_k_2.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_ac_k_2["所属3"] > 0
df_ac_k_2.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_ac_k_2["出勤日数    "]
df_ac_k_2.insert(2, "実在籍者", real_member)
time_yukyu = df_ac_k_2["有休日数    "] * 8
df_ac_k_2.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_ac_k_2.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_ac_k_2["欠勤日数    "] * 8
df_ac_k_2.insert(5, "欠勤時間", abs_time)
work_time = df_ac_k_2["勤務時間    "]
df_ac_k_2.insert(6, "勤務時間", work_time)
late_early_time = df_ac_k_2["遅早時間    "]
df_ac_k_2.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_ac_k_2["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_ac_k_2.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_ac_k_2.insert(17, "実労働時間", real_work_time)
zure_time = df_ac_k_2["ズレ時間    "]
df_ac_k_2.insert(18, "ズレ時間", zure_time)
overtime = df_ac_k_2["残業時間    "] + df_ac_k_2["深夜残業    "]
df_ac_k_2.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_ac_k_2.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_ac_k_2["法外休出    "]
df_ac_k_2.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_ac_k_2.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_ac_k_2["法定休出    "]
df_ac_k_2.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_ac_k_2.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_ac_k_2["６０Ｈ超    "]
df_ac_k_2.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_ac_k_2["代休時間    "] + df_ac_k_2["深夜代休    "]
df_ac_k_2.insert(31, "代休時間", holiday_time)
df_ac_k_2.insert(32, "応援時間", 0)
total_work_time = (
    df_ac_k_2["勤務時間    "]
    + df_ac_k_2["残業時間    "]
    + df_ac_k_2["法外休出    "]
    + df_ac_k_2["法定休出    "]
)
df_ac_k_2.insert(33, "総労働時間", total_work_time)
basic_salary = df_ac_k_2["基 本 給    "] + df_ac_k_2["支給額"]
df_ac_k_2.insert(35, "基本給", basic_salary)
post_allowance = df_ac_k_2["役職手当    "]
df_ac_k_2.insert(36, "役職手当", post_allowance)
sales_allowance = df_ac_k_2["営業手当    "]
df_ac_k_2.insert(37, "営業手当", sales_allowance)
aria_allowance = df_ac_k_2["地域手当    "]
df_ac_k_2.insert(38, "地域手当", aria_allowance)
spe_allowance = df_ac_k_2["特殊手当    "]
df_ac_k_2.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_ac_k_2["特別技技手当"]
df_ac_k_2.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_ac_k_2["調整手当    "]
df_ac_k_2.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_ac_k_2["別居手当    "]
df_ac_k_2.insert(42, "別居手当", sep_allowance)
com_allowance = df_ac_k_2["通勤手当    "]
df_ac_k_2.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_ac_k_2.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_ac_k_2["残業手当    "]
df_ac_k_2.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_ac_k_2["休出手当    "]
df_ac_k_2.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_ac_k_2["深夜勤務手当"]
df_ac_k_2.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_ac_k_2["交替時差手当"]
df_ac_k_2.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_ac_k_2["休業手当    "]
df_ac_k_2.insert(49, "休業手当", closed_allowance)
closed_deduction = df_ac_k_2["休業控除    "]
df_ac_k_2.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_ac_k_2["代 休 他    "]
df_ac_k_2.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_ac_k_2["欠勤控除    "] + df_ac_k_2["遅早控除    "]
df_ac_k_2.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_ac_k_2["精 算 分    "]
df_ac_k_2.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_ac_k_2.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_ac_k_2.insert(55, "総支給額", total)
df_ac_k_2.insert(56, "応援時間額", 0)
df_ac_k_2.insert(57, "役員振替", 0)
df_ac_k_2.insert(58, "部門振替", 0)
df_ac_k_2.insert(59, "合計", 0)
# 不要フィールドの削除
df_ac_k_2 = df_ac_k_2.drop("所属2", axis=1)
df_ac_k_2 = df_ac_k_2.drop("所属3", axis=1)
df_ac_k_2 = df_ac_k_2.drop("社員ｺｰﾄﾞ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("区分", axis=1)
df_ac_k_2 = df_ac_k_2.drop("出勤日数    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("有休日数    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("欠勤日数    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("残業時間    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("深夜残業    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("法外休出    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("法定休出    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("代休時間    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("深夜代休    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("６０Ｈ超    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("別居手当    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("通勤手当    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("特別技技手当", axis=1)
df_ac_k_2 = df_ac_k_2.drop("特殊手当    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("地域手当    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("営業手当    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("役職手当    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("調整手当    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("基 本 給    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("残業手当    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("休出手当    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("深夜勤務手当", axis=1)
df_ac_k_2 = df_ac_k_2.drop("交替時差手当", axis=1)
df_ac_k_2 = df_ac_k_2.drop("休業手当    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("代 休 他    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("欠勤控除    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("遅早控除    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("精 算 分    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("支給合計額  ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("休業控除    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("支給額", axis=1)
df_ac_k_2 = df_ac_k_2.drop("ズレ時間    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("雑費・食事代", axis=1)
df_ac_k_2 = df_ac_k_2.drop("雑費・衣靴代", axis=1)
df_ac_k_2 = df_ac_k_2.drop("雑費        ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("受診料・他  ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("雑費・会費等", axis=1)
df_ac_k_2 = df_ac_k_2.drop("勤務時間    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("遅早時間    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("特休日数    ", axis=1)
df_ac_k_2 = df_ac_k_2.drop("集計区分－２        ", axis=1)
df_ac_k_2 = df_ac_k_2.sum()
df_ac_k_2.to_csv(
    "/content/drive/MyDrive/data/AC/B.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# AC_間接4
df_ac_k_4 = df_ac_m.groupby("区分").get_group("間接4")
df_ac_k_4 = df_ac_k_4.drop("所属1", axis=1)
member = df_ac_k_4["所属2"] > 0
df_ac_k_4.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_ac_k_4["所属3"] > 0
df_ac_k_4.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_ac_k_4["出勤日数    "]
df_ac_k_4.insert(2, "実在籍者", real_member)
time_yukyu = df_ac_k_4["有休日数    "] * 8
df_ac_k_4.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_ac_k_4.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_ac_k_4["欠勤日数    "] * 8
df_ac_k_4.insert(5, "欠勤時間", abs_time)
work_time = df_ac_k_4["勤務時間    "]
df_ac_k_4.insert(6, "勤務時間", work_time)
late_early_time = df_ac_k_4["遅早時間    "]
df_ac_k_4.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_ac_k_4["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_ac_k_4.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_ac_k_4.insert(17, "実労働時間", real_work_time)
zure_time = df_ac_k_4["ズレ時間    "]
df_ac_k_4.insert(18, "ズレ時間", zure_time)
overtime = df_ac_k_4["残業時間    "] + df_ac_k_4["深夜残業    "]
df_ac_k_4.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_ac_k_4.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_ac_k_4["法外休出    "]
df_ac_k_4.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_ac_k_4.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_ac_k_4["法定休出    "]
df_ac_k_4.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_ac_k_4.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_ac_k_4["６０Ｈ超    "]
df_ac_k_4.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_ac_k_4["代休時間    "] + df_ac_k_4["深夜代休    "]
df_ac_k_4.insert(31, "代休時間", holiday_time)
df_ac_k_4.insert(32, "応援時間", 0)
total_work_time = (
    df_ac_k_4["勤務時間    "]
    + df_ac_k_4["残業時間    "]
    + df_ac_k_4["法外休出    "]
    + df_ac_k_4["法定休出    "]
)
df_ac_k_4.insert(33, "総労働時間", total_work_time)
basic_salary = df_ac_k_4["基 本 給    "] + df_ac_k_4["支給額"]
df_ac_k_4.insert(35, "基本給", basic_salary)
post_allowance = df_ac_k_4["役職手当    "]
df_ac_k_4.insert(36, "役職手当", post_allowance)
sales_allowance = df_ac_k_4["営業手当    "]
df_ac_k_4.insert(37, "営業手当", sales_allowance)
aria_allowance = df_ac_k_4["地域手当    "]
df_ac_k_4.insert(38, "地域手当", aria_allowance)
spe_allowance = df_ac_k_4["特殊手当    "]
df_ac_k_4.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_ac_k_4["特別技技手当"]
df_ac_k_4.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_ac_k_4["調整手当    "]
df_ac_k_4.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_ac_k_4["別居手当    "]
df_ac_k_4.insert(42, "別居手当", sep_allowance)
com_allowance = df_ac_k_4["通勤手当    "]
df_ac_k_4.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_ac_k_4.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_ac_k_4["残業手当    "]
df_ac_k_4.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_ac_k_4["休出手当    "]
df_ac_k_4.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_ac_k_4["深夜勤務手当"]
df_ac_k_4.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_ac_k_4["交替時差手当"]
df_ac_k_4.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_ac_k_4["休業手当    "]
df_ac_k_4.insert(49, "休業手当", closed_allowance)
closed_deduction = df_ac_k_4["休業控除    "]
df_ac_k_4.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_ac_k_4["代 休 他    "]
df_ac_k_4.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_ac_k_4["欠勤控除    "] + df_ac_k_4["遅早控除    "]
df_ac_k_4.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_ac_k_4["精 算 分    "]
df_ac_k_4.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_ac_k_4.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_ac_k_4.insert(55, "総支給額", total)
df_ac_k_4.insert(56, "応援時間額", 0)
df_ac_k_4.insert(57, "役員振替", 0)
df_ac_k_4.insert(58, "部門振替", 0)
df_ac_k_4.insert(59, "合計", 0)
# 不要フィールドの削除
df_ac_k_4 = df_ac_k_4.drop("所属2", axis=1)
df_ac_k_4 = df_ac_k_4.drop("所属3", axis=1)
df_ac_k_4 = df_ac_k_4.drop("社員ｺｰﾄﾞ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("区分", axis=1)
df_ac_k_4 = df_ac_k_4.drop("出勤日数    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("有休日数    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("欠勤日数    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("残業時間    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("深夜残業    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("法外休出    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("法定休出    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("代休時間    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("深夜代休    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("６０Ｈ超    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("別居手当    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("通勤手当    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("特別技技手当", axis=1)
df_ac_k_4 = df_ac_k_4.drop("特殊手当    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("地域手当    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("営業手当    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("役職手当    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("調整手当    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("基 本 給    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("残業手当    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("休出手当    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("深夜勤務手当", axis=1)
df_ac_k_4 = df_ac_k_4.drop("交替時差手当", axis=1)
df_ac_k_4 = df_ac_k_4.drop("休業手当    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("代 休 他    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("欠勤控除    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("遅早控除    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("精 算 分    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("支給合計額  ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("休業控除    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("支給額", axis=1)
df_ac_k_4 = df_ac_k_4.drop("ズレ時間    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("雑費・食事代", axis=1)
df_ac_k_4 = df_ac_k_4.drop("雑費・衣靴代", axis=1)
df_ac_k_4 = df_ac_k_4.drop("雑費        ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("受診料・他  ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("雑費・会費等", axis=1)
df_ac_k_4 = df_ac_k_4.drop("勤務時間    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("遅早時間    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("特休日数    ", axis=1)
df_ac_k_4 = df_ac_k_4.drop("集計区分－２        ", axis=1)
df_ac_k_4 = df_ac_k_4.sum()
df_ac_k_4.to_csv(
    "/content/drive/MyDrive/data/AC/C.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# AC_間接5
df_ac_k_5 = df_ac_m.groupby("区分").get_group("間接5")
df_ac_k_5 = df_ac_k_5.drop("所属1", axis=1)
member = df_ac_k_5["所属2"] > 0
df_ac_k_5.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_ac_k_5["所属3"] > 0
df_ac_k_5.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_ac_k_5["出勤日数    "]
df_ac_k_5.insert(2, "実在籍者", real_member)
time_yukyu = df_ac_k_5["有休日数    "] * 8
df_ac_k_5.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_ac_k_5.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_ac_k_5["欠勤日数    "] * 8
df_ac_k_5.insert(5, "欠勤時間", abs_time)
work_time = df_ac_k_5["勤務時間    "]
df_ac_k_5.insert(6, "勤務時間", work_time)
late_early_time = df_ac_k_5["遅早時間    "]
df_ac_k_5.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_ac_k_5["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_ac_k_5.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_ac_k_5.insert(17, "実労働時間", real_work_time)
zure_time = df_ac_k_5["ズレ時間    "]
df_ac_k_5.insert(18, "ズレ時間", zure_time)
overtime = df_ac_k_5["残業時間    "] + df_ac_k_5["深夜残業    "]
df_ac_k_5.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_ac_k_5.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_ac_k_5["法外休出    "]
df_ac_k_5.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_ac_k_5.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_ac_k_5["法定休出    "]
df_ac_k_5.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_ac_k_5.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_ac_k_5["６０Ｈ超    "]
df_ac_k_5.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_ac_k_5["代休時間    "] + df_ac_k_5["深夜代休    "]
df_ac_k_5.insert(31, "代休時間", holiday_time)
df_ac_k_5.insert(32, "応援時間", 0)
total_work_time = (
    df_ac_k_5["勤務時間    "]
    + df_ac_k_5["残業時間    "]
    + df_ac_k_5["法外休出    "]
    + df_ac_k_5["法定休出    "]
)
df_ac_k_5.insert(33, "総労働時間", total_work_time)
basic_salary = df_ac_k_5["基 本 給    "] + df_ac_k_5["支給額"]
df_ac_k_5.insert(35, "基本給", basic_salary)
post_allowance = df_ac_k_5["役職手当    "]
df_ac_k_5.insert(36, "役職手当", post_allowance)
sales_allowance = df_ac_k_5["営業手当    "]
df_ac_k_5.insert(37, "営業手当", sales_allowance)
aria_allowance = df_ac_k_5["地域手当    "]
df_ac_k_5.insert(38, "地域手当", aria_allowance)
spe_allowance = df_ac_k_5["特殊手当    "]
df_ac_k_5.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_ac_k_5["特別技技手当"]
df_ac_k_5.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_ac_k_5["調整手当    "]
df_ac_k_5.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_ac_k_5["別居手当    "]
df_ac_k_5.insert(42, "別居手当", sep_allowance)
com_allowance = df_ac_k_5["通勤手当    "]
df_ac_k_5.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_ac_k_5.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_ac_k_5["残業手当    "]
df_ac_k_5.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_ac_k_5["休出手当    "]
df_ac_k_5.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_ac_k_5["深夜勤務手当"]
df_ac_k_5.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_ac_k_5["交替時差手当"]
df_ac_k_5.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_ac_k_5["休業手当    "]
df_ac_k_5.insert(49, "休業手当", closed_allowance)
closed_deduction = df_ac_k_5["休業控除    "]
df_ac_k_5.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_ac_k_5["代 休 他    "]
df_ac_k_5.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_ac_k_5["欠勤控除    "] + df_ac_k_5["遅早控除    "]
df_ac_k_5.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_ac_k_5["精 算 分    "]
df_ac_k_5.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_ac_k_5.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_ac_k_5.insert(55, "総支給額", total)
df_ac_k_5.insert(56, "応援時間額", 0)
df_ac_k_5.insert(57, "役員振替", 0)
df_ac_k_5.insert(58, "部門振替", 0)
df_ac_k_5.insert(59, "合計", 0)
# 不要フィールドの削除
df_ac_k_5 = df_ac_k_5.drop("所属2", axis=1)
df_ac_k_5 = df_ac_k_5.drop("所属3", axis=1)
df_ac_k_5 = df_ac_k_5.drop("社員ｺｰﾄﾞ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("区分", axis=1)
df_ac_k_5 = df_ac_k_5.drop("出勤日数    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("有休日数    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("欠勤日数    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("残業時間    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("深夜残業    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("法外休出    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("法定休出    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("代休時間    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("深夜代休    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("６０Ｈ超    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("別居手当    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("通勤手当    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("特別技技手当", axis=1)
df_ac_k_5 = df_ac_k_5.drop("特殊手当    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("地域手当    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("営業手当    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("役職手当    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("調整手当    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("基 本 給    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("残業手当    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("休出手当    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("深夜勤務手当", axis=1)
df_ac_k_5 = df_ac_k_5.drop("交替時差手当", axis=1)
df_ac_k_5 = df_ac_k_5.drop("休業手当    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("代 休 他    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("欠勤控除    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("遅早控除    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("精 算 分    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("支給合計額  ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("休業控除    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("支給額", axis=1)
df_ac_k_5 = df_ac_k_5.drop("ズレ時間    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("雑費・食事代", axis=1)
df_ac_k_5 = df_ac_k_5.drop("雑費・衣靴代", axis=1)
df_ac_k_5 = df_ac_k_5.drop("雑費        ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("受診料・他  ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("雑費・会費等", axis=1)
df_ac_k_5 = df_ac_k_5.drop("勤務時間    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("遅早時間    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("特休日数    ", axis=1)
df_ac_k_5 = df_ac_k_5.drop("集計区分－２        ", axis=1)
df_ac_k_5 = df_ac_k_5.sum()
df_ac_k_5.to_csv(
    "/content/drive/MyDrive/data/AC/D.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# AC_直接1
df_ac_t_1 = df_ac_m.groupby("区分").get_group("直接1")
df_ac_t_1 = df_ac_t_1.drop("所属1", axis=1)
member = df_ac_t_1["所属2"] > 0
df_ac_t_1.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_ac_t_1["所属3"] > 0
df_ac_t_1.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_ac_t_1["出勤日数    "]
df_ac_t_1.insert(2, "実在籍者", real_member)
time_yukyu = df_ac_t_1["有休日数    "] * 8
df_ac_t_1.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_ac_t_1.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_ac_t_1["欠勤日数    "] * 8
df_ac_t_1.insert(5, "欠勤時間", abs_time)
work_time = df_ac_t_1["勤務時間    "]
df_ac_t_1.insert(6, "勤務時間", work_time)
late_early_time = df_ac_t_1["遅早時間    "]
df_ac_t_1.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_ac_t_1["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_ac_t_1.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_ac_t_1.insert(17, "実労働時間", real_work_time)
zure_time = df_ac_t_1["ズレ時間    "]
df_ac_t_1.insert(18, "ズレ時間", zure_time)
overtime = df_ac_t_1["残業時間    "] + df_ac_t_1["深夜残業    "]
df_ac_t_1.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_ac_t_1.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_ac_t_1["法外休出    "]
df_ac_t_1.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_ac_t_1.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_ac_t_1["法定休出    "]
df_ac_t_1.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_ac_t_1.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_ac_t_1["６０Ｈ超    "]
df_ac_t_1.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_ac_t_1["代休時間    "] + df_ac_t_1["深夜代休    "]
df_ac_t_1.insert(31, "代休時間", holiday_time)
df_ac_t_1.insert(32, "応援時間", 0)
total_work_time = (
    df_ac_t_1["勤務時間    "]
    + df_ac_t_1["残業時間    "]
    + df_ac_t_1["法外休出    "]
    + df_ac_t_1["法定休出    "]
)
df_ac_t_1.insert(33, "総労働時間", total_work_time)
basic_salary = df_ac_t_1["基 本 給    "] + df_ac_t_1["支給額"]
df_ac_t_1.insert(35, "基本給", basic_salary)
post_allowance = df_ac_t_1["役職手当    "]
df_ac_t_1.insert(36, "役職手当", post_allowance)
sales_allowance = df_ac_t_1["営業手当    "]
df_ac_t_1.insert(37, "営業手当", sales_allowance)
aria_allowance = df_ac_t_1["地域手当    "]
df_ac_t_1.insert(38, "地域手当", aria_allowance)
spe_allowance = df_ac_t_1["特殊手当    "]
df_ac_t_1.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_ac_t_1["特別技技手当"]
df_ac_t_1.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_ac_t_1["調整手当    "]
df_ac_t_1.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_ac_t_1["別居手当    "]
df_ac_t_1.insert(42, "別居手当", sep_allowance)
com_allowance = df_ac_t_1["通勤手当    "]
df_ac_t_1.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_ac_t_1.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_ac_t_1["残業手当    "]
df_ac_t_1.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_ac_t_1["休出手当    "]
df_ac_t_1.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_ac_t_1["深夜勤務手当"]
df_ac_t_1.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_ac_t_1["交替時差手当"]
df_ac_t_1.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_ac_t_1["休業手当    "]
df_ac_t_1.insert(49, "休業手当", closed_allowance)
closed_deduction = df_ac_t_1["休業控除    "]
df_ac_t_1.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_ac_t_1["代 休 他    "]
df_ac_t_1.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_ac_t_1["欠勤控除    "] + df_ac_t_1["遅早控除    "]
df_ac_t_1.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_ac_t_1["精 算 分    "]
df_ac_t_1.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_ac_t_1.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_ac_t_1.insert(55, "総支給額", total)
df_ac_t_1.insert(56, "応援時間額", 0)
df_ac_t_1.insert(57, "役員振替", 0)
df_ac_t_1.insert(58, "部門振替", 0)
df_ac_t_1.insert(59, "合計", 0)
# 不要フィールドの削除
df_ac_t_1 = df_ac_t_1.drop("所属2", axis=1)
df_ac_t_1 = df_ac_t_1.drop("所属3", axis=1)
df_ac_t_1 = df_ac_t_1.drop("社員ｺｰﾄﾞ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("区分", axis=1)
df_ac_t_1 = df_ac_t_1.drop("出勤日数    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("有休日数    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("欠勤日数    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("残業時間    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("深夜残業    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("法外休出    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("法定休出    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("代休時間    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("深夜代休    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("６０Ｈ超    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("別居手当    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("通勤手当    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("特別技技手当", axis=1)
df_ac_t_1 = df_ac_t_1.drop("特殊手当    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("地域手当    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("営業手当    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("役職手当    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("調整手当    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("基 本 給    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("残業手当    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("休出手当    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("深夜勤務手当", axis=1)
df_ac_t_1 = df_ac_t_1.drop("交替時差手当", axis=1)
df_ac_t_1 = df_ac_t_1.drop("休業手当    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("代 休 他    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("欠勤控除    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("遅早控除    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("精 算 分    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("支給合計額  ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("休業控除    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("支給額", axis=1)
df_ac_t_1 = df_ac_t_1.drop("ズレ時間    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("雑費・食事代", axis=1)
df_ac_t_1 = df_ac_t_1.drop("雑費・衣靴代", axis=1)
df_ac_t_1 = df_ac_t_1.drop("雑費        ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("受診料・他  ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("雑費・会費等", axis=1)
df_ac_t_1 = df_ac_t_1.drop("勤務時間    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("遅早時間    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("特休日数    ", axis=1)
df_ac_t_1 = df_ac_t_1.drop("集計区分－２        ", axis=1)
df_ac_t_1 = df_ac_t_1.sum()
df_ac_t_1.to_csv(
    "/content/drive/MyDrive/data/AC/E.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# AC_直接4
df_ac_t_4 = df_ac_m.groupby("区分").get_group("直接4")
df_ac_t_4 = df_ac_t_4.drop("所属1", axis=1)
member = df_ac_t_4["所属2"] > 0
df_ac_t_4.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_ac_t_4["所属3"] > 0
df_ac_t_4.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_ac_t_4["出勤日数    "]
df_ac_t_4.insert(2, "実在籍者", real_member)
time_yukyu = df_ac_t_4["有休日数    "] * 8
df_ac_t_4.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_ac_t_4.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_ac_t_4["欠勤日数    "] * 8
df_ac_t_4.insert(5, "欠勤時間", abs_time)
work_time = df_ac_t_4["勤務時間    "]
df_ac_t_4.insert(6, "勤務時間", work_time)
late_early_time = df_ac_t_4["遅早時間    "]
df_ac_t_4.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_ac_t_4["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_ac_t_4.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_ac_t_4.insert(17, "実労働時間", real_work_time)
zure_time = df_ac_t_4["ズレ時間    "]
df_ac_t_4.insert(18, "ズレ時間", zure_time)
overtime = df_ac_t_4["残業時間    "] + df_ac_t_4["深夜残業    "]
df_ac_t_4.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_ac_t_4.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_ac_t_4["法外休出    "]
df_ac_t_4.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_ac_t_4.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_ac_t_4["法定休出    "]
df_ac_t_4.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_ac_t_4.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_ac_t_4["６０Ｈ超    "]
df_ac_t_4.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_ac_t_4["代休時間    "] + df_ac_t_4["深夜代休    "]
df_ac_t_4.insert(31, "代休時間", holiday_time)
df_ac_t_4.insert(32, "応援時間", 0)
total_work_time = (
    df_ac_t_4["勤務時間    "]
    + df_ac_t_4["残業時間    "]
    + df_ac_t_4["法外休出    "]
    + df_ac_t_4["法定休出    "]
)
df_ac_t_4.insert(33, "総労働時間", total_work_time)
basic_salary = df_ac_t_4["基 本 給    "] + df_ac_t_4["支給額"]
df_ac_t_4.insert(35, "基本給", basic_salary)
post_allowance = df_ac_t_4["役職手当    "]
df_ac_t_4.insert(36, "役職手当", post_allowance)
sales_allowance = df_ac_t_4["営業手当    "]
df_ac_t_4.insert(37, "営業手当", sales_allowance)
aria_allowance = df_ac_t_4["地域手当    "]
df_ac_t_4.insert(38, "地域手当", aria_allowance)
spe_allowance = df_ac_t_4["特殊手当    "]
df_ac_t_4.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_ac_t_4["特別技技手当"]
df_ac_t_4.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_ac_t_4["調整手当    "]
df_ac_t_4.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_ac_t_4["別居手当    "]
df_ac_t_4.insert(42, "別居手当", sep_allowance)
com_allowance = df_ac_t_4["通勤手当    "]
df_ac_t_4.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_ac_t_4.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_ac_t_4["残業手当    "]
df_ac_t_4.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_ac_t_4["休出手当    "]
df_ac_t_4.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_ac_t_4["深夜勤務手当"]
df_ac_t_4.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_ac_t_4["交替時差手当"]
df_ac_t_4.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_ac_t_4["休業手当    "]
df_ac_t_4.insert(49, "休業手当", closed_allowance)
closed_deduction = df_ac_t_4["休業控除    "]
df_ac_t_4.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_ac_t_4["代 休 他    "]
df_ac_t_4.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_ac_t_4["欠勤控除    "] + df_ac_t_4["遅早控除    "]
df_ac_t_4.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_ac_t_4["精 算 分    "]
df_ac_t_4.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_ac_t_4.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_ac_t_4.insert(55, "総支給額", total)
df_ac_t_4.insert(56, "応援時間額", 0)
df_ac_t_4.insert(57, "役員振替", 0)
df_ac_t_4.insert(58, "部門振替", 0)
df_ac_t_4.insert(59, "合計", 0)
# 不要フィールドの削除
df_ac_t_4 = df_ac_t_4.drop("所属2", axis=1)
df_ac_t_4 = df_ac_t_4.drop("所属3", axis=1)
df_ac_t_4 = df_ac_t_4.drop("社員ｺｰﾄﾞ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("区分", axis=1)
df_ac_t_4 = df_ac_t_4.drop("出勤日数    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("有休日数    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("欠勤日数    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("残業時間    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("深夜残業    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("法外休出    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("法定休出    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("代休時間    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("深夜代休    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("６０Ｈ超    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("別居手当    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("通勤手当    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("特別技技手当", axis=1)
df_ac_t_4 = df_ac_t_4.drop("特殊手当    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("地域手当    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("営業手当    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("役職手当    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("調整手当    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("基 本 給    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("残業手当    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("休出手当    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("深夜勤務手当", axis=1)
df_ac_t_4 = df_ac_t_4.drop("交替時差手当", axis=1)
df_ac_t_4 = df_ac_t_4.drop("休業手当    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("代 休 他    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("欠勤控除    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("遅早控除    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("精 算 分    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("支給合計額  ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("休業控除    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("支給額", axis=1)
df_ac_t_4 = df_ac_t_4.drop("ズレ時間    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("雑費・食事代", axis=1)
df_ac_t_4 = df_ac_t_4.drop("雑費・衣靴代", axis=1)
df_ac_t_4 = df_ac_t_4.drop("雑費        ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("受診料・他  ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("雑費・会費等", axis=1)
df_ac_t_4 = df_ac_t_4.drop("勤務時間    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("遅早時間    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("特休日数    ", axis=1)
df_ac_t_4 = df_ac_t_4.drop("集計区分－２        ", axis=1)
df_ac_t_4 = df_ac_t_4.sum()
df_ac_t_4.to_csv(
    "/content/drive/MyDrive/data/AC/F.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# PC_間接1
df_pc_k_1 = df_pc_m.groupby("区分").get_group("間接1")
df_pc_k_1 = df_pc_k_1.drop("所属1", axis=1)
member = df_pc_k_1["所属2"] > 0
df_pc_k_1.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_pc_k_1["所属3"] > 0
df_pc_k_1.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_pc_k_1["出勤日数    "]
df_pc_k_1.insert(2, "実在籍者", real_member)
time_yukyu = df_pc_k_1["有休日数    "] * 8
df_pc_k_1.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_pc_k_1.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_pc_k_1["欠勤日数    "] * 8
df_pc_k_1.insert(5, "欠勤時間", abs_time)
work_time = df_pc_k_1["勤務時間    "]
df_pc_k_1.insert(6, "勤務時間", work_time)
late_early_time = df_pc_k_1["遅早時間    "]
df_pc_k_1.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_pc_k_1["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_pc_k_1.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_pc_k_1.insert(17, "実労働時間", real_work_time)
zure_time = df_pc_k_1["ズレ時間    "]
df_pc_k_1.insert(18, "ズレ時間", zure_time)
overtime = df_pc_k_1["残業時間    "] + df_pc_k_1["深夜残業    "]
df_pc_k_1.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_pc_k_1.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_pc_k_1["法外休出    "]
df_pc_k_1.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_pc_k_1.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_pc_k_1["法定休出    "]
df_pc_k_1.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_pc_k_1.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_pc_k_1["６０Ｈ超    "]
df_pc_k_1.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_pc_k_1["代休時間    "] + df_pc_k_1["深夜代休    "]
df_pc_k_1.insert(31, "代休時間", holiday_time)
df_pc_k_1.insert(32, "応援時間", 0)
total_work_time = (
    df_pc_k_1["勤務時間    "]
    + df_pc_k_1["残業時間    "]
    + df_pc_k_1["法外休出    "]
    + df_pc_k_1["法定休出    "]
)
df_pc_k_1.insert(33, "総労働時間", total_work_time)
basic_salary = df_pc_k_1["基 本 給    "] + df_pc_k_1["支給額"]
df_pc_k_1.insert(35, "基本給", basic_salary)
post_allowance = df_pc_k_1["役職手当    "]
df_pc_k_1.insert(36, "役職手当", post_allowance)
sales_allowance = df_pc_k_1["営業手当    "]
df_pc_k_1.insert(37, "営業手当", sales_allowance)
aria_allowance = df_pc_k_1["地域手当    "]
df_pc_k_1.insert(38, "地域手当", aria_allowance)
spe_allowance = df_pc_k_1["特殊手当    "]
df_pc_k_1.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_pc_k_1["特別技技手当"]
df_pc_k_1.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_pc_k_1["調整手当    "]
df_pc_k_1.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_pc_k_1["別居手当    "]
df_pc_k_1.insert(42, "別居手当", sep_allowance)
com_allowance = df_pc_k_1["通勤手当    "]
df_pc_k_1.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_pc_k_1.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_pc_k_1["残業手当    "]
df_pc_k_1.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_pc_k_1["休出手当    "]
df_pc_k_1.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_pc_k_1["深夜勤務手当"]
df_pc_k_1.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_pc_k_1["交替時差手当"]
df_pc_k_1.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_pc_k_1["休業手当    "]
df_pc_k_1.insert(49, "休業手当", closed_allowance)
closed_deduction = df_pc_k_1["休業控除    "]
df_pc_k_1.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_pc_k_1["代 休 他    "]
df_pc_k_1.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_pc_k_1["欠勤控除    "] + df_pc_k_1["遅早控除    "]
df_pc_k_1.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_pc_k_1["精 算 分    "]
df_pc_k_1.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_pc_k_1.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_pc_k_1.insert(55, "総支給額", total)
df_pc_k_1.insert(56, "応援時間額", 0)
df_pc_k_1.insert(57, "役員振替", 0)
df_pc_k_1.insert(58, "部門振替", 0)
df_pc_k_1.insert(59, "合計", 0)
# 不要フィールドの削除
df_pc_k_1 = df_pc_k_1.drop("所属2", axis=1)
df_pc_k_1 = df_pc_k_1.drop("所属3", axis=1)
df_pc_k_1 = df_pc_k_1.drop("社員ｺｰﾄﾞ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("区分", axis=1)
df_pc_k_1 = df_pc_k_1.drop("出勤日数    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("有休日数    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("欠勤日数    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("残業時間    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("深夜残業    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("法外休出    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("法定休出    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("代休時間    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("深夜代休    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("６０Ｈ超    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("別居手当    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("通勤手当    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("特別技技手当", axis=1)
df_pc_k_1 = df_pc_k_1.drop("特殊手当    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("地域手当    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("営業手当    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("役職手当    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("調整手当    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("基 本 給    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("残業手当    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("休出手当    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("深夜勤務手当", axis=1)
df_pc_k_1 = df_pc_k_1.drop("交替時差手当", axis=1)
df_pc_k_1 = df_pc_k_1.drop("休業手当    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("代 休 他    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("欠勤控除    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("遅早控除    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("精 算 分    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("支給合計額  ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("休業控除    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("支給額", axis=1)
df_pc_k_1 = df_pc_k_1.drop("ズレ時間    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("雑費・食事代", axis=1)
df_pc_k_1 = df_pc_k_1.drop("雑費・衣靴代", axis=1)
df_pc_k_1 = df_pc_k_1.drop("雑費        ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("受診料・他  ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("雑費・会費等", axis=1)
df_pc_k_1 = df_pc_k_1.drop("勤務時間    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("遅早時間    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("特休日数    ", axis=1)
df_pc_k_1 = df_pc_k_1.drop("集計区分－２        ", axis=1)
df_pc_k_1 = df_pc_k_1.sum()
df_pc_k_1.to_csv(
    "/content/drive/MyDrive/data/PC/A.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# PC_間接2
df_pc_k_2 = df_pc_m.groupby("区分").get_group("間接2")
df_pc_k_2 = df_pc_k_2.drop("所属1", axis=1)
member = df_pc_k_2["所属2"] > 0
df_pc_k_2.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_pc_k_2["所属3"] > 0
df_pc_k_2.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_pc_k_2["出勤日数    "]
df_pc_k_2.insert(2, "実在籍者", real_member)
time_yukyu = df_pc_k_2["有休日数    "] * 8
df_pc_k_2.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_pc_k_2.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_pc_k_2["欠勤日数    "] * 8
df_pc_k_2.insert(5, "欠勤時間", abs_time)
work_time = df_pc_k_2["勤務時間    "]
df_pc_k_2.insert(6, "勤務時間", work_time)
late_early_time = df_pc_k_2["遅早時間    "]
df_pc_k_2.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_pc_k_2["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_pc_k_2.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_pc_k_2.insert(17, "実労働時間", real_work_time)
zure_time = df_pc_k_2["ズレ時間    "]
df_pc_k_2.insert(18, "ズレ時間", zure_time)
overtime = df_pc_k_2["残業時間    "] + df_pc_k_2["深夜残業    "]
df_pc_k_2.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_pc_k_2.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_pc_k_2["法外休出    "]
df_pc_k_2.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_pc_k_2.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_pc_k_2["法定休出    "]
df_pc_k_2.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_pc_k_2.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_pc_k_2["６０Ｈ超    "]
df_pc_k_2.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_pc_k_2["代休時間    "] + df_pc_k_2["深夜代休    "]
df_pc_k_2.insert(31, "代休時間", holiday_time)
df_pc_k_2.insert(32, "応援時間", 0)
total_work_time = (
    df_pc_k_2["勤務時間    "]
    + df_pc_k_2["残業時間    "]
    + df_pc_k_2["法外休出    "]
    + df_pc_k_2["法定休出    "]
)
df_pc_k_2.insert(33, "総労働時間", total_work_time)
basic_salary = df_pc_k_2["基 本 給    "] + df_pc_k_2["支給額"]
df_pc_k_2.insert(35, "基本給", basic_salary)
post_allowance = df_pc_k_2["役職手当    "]
df_pc_k_2.insert(36, "役職手当", post_allowance)
sales_allowance = df_pc_k_2["営業手当    "]
df_pc_k_2.insert(37, "営業手当", sales_allowance)
aria_allowance = df_pc_k_2["地域手当    "]
df_pc_k_2.insert(38, "地域手当", aria_allowance)
spe_allowance = df_pc_k_2["特殊手当    "]
df_pc_k_2.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_pc_k_2["特別技技手当"]
df_pc_k_2.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_pc_k_2["調整手当    "]
df_pc_k_2.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_pc_k_2["別居手当    "]
df_pc_k_2.insert(42, "別居手当", sep_allowance)
com_allowance = df_pc_k_2["通勤手当    "]
df_pc_k_2.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_pc_k_2.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_pc_k_2["残業手当    "]
df_pc_k_2.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_pc_k_2["休出手当    "]
df_pc_k_2.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_pc_k_2["深夜勤務手当"]
df_pc_k_2.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_pc_k_2["交替時差手当"]
df_pc_k_2.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_pc_k_2["休業手当    "]
df_pc_k_2.insert(49, "休業手当", closed_allowance)
closed_deduction = df_pc_k_2["休業控除    "]
df_pc_k_2.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_pc_k_2["代 休 他    "]
df_pc_k_2.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_pc_k_2["欠勤控除    "] + df_pc_k_2["遅早控除    "]
df_pc_k_2.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_pc_k_2["精 算 分    "]
df_pc_k_2.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_pc_k_2.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_pc_k_2.insert(55, "総支給額", total)
df_pc_k_2.insert(56, "応援時間額", 0)
df_pc_k_2.insert(57, "役員振替", 0)
df_pc_k_2.insert(58, "部門振替", 0)
df_pc_k_2.insert(59, "合計", 0)
# 不要フィールドの削除
df_pc_k_2 = df_pc_k_2.drop("所属2", axis=1)
df_pc_k_2 = df_pc_k_2.drop("所属3", axis=1)
df_pc_k_2 = df_pc_k_2.drop("社員ｺｰﾄﾞ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("区分", axis=1)
df_pc_k_2 = df_pc_k_2.drop("出勤日数    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("有休日数    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("欠勤日数    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("残業時間    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("深夜残業    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("法外休出    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("法定休出    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("代休時間    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("深夜代休    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("６０Ｈ超    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("別居手当    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("通勤手当    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("特別技技手当", axis=1)
df_pc_k_2 = df_pc_k_2.drop("特殊手当    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("地域手当    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("営業手当    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("役職手当    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("調整手当    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("基 本 給    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("残業手当    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("休出手当    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("深夜勤務手当", axis=1)
df_pc_k_2 = df_pc_k_2.drop("交替時差手当", axis=1)
df_pc_k_2 = df_pc_k_2.drop("休業手当    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("代 休 他    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("欠勤控除    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("遅早控除    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("精 算 分    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("支給合計額  ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("休業控除    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("支給額", axis=1)
df_pc_k_2 = df_pc_k_2.drop("ズレ時間    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("雑費・食事代", axis=1)
df_pc_k_2 = df_pc_k_2.drop("雑費・衣靴代", axis=1)
df_pc_k_2 = df_pc_k_2.drop("雑費        ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("受診料・他  ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("雑費・会費等", axis=1)
df_pc_k_2 = df_pc_k_2.drop("勤務時間    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("遅早時間    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("特休日数    ", axis=1)
df_pc_k_2 = df_pc_k_2.drop("集計区分－２        ", axis=1)
df_pc_k_2 = df_pc_k_2.sum()
df_pc_k_2.to_csv(
    "/content/drive/MyDrive/data/PC/B.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# PC_間接4
df_pc_k_4 = df_pc_m.groupby("区分").get_group("間接4")
df_pc_k_4 = df_pc_k_4.drop("所属1", axis=1)
member = df_pc_k_4["所属2"] > 0
df_pc_k_4.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_pc_k_4["所属3"] > 0
df_pc_k_4.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_pc_k_4["出勤日数    "]
df_pc_k_4.insert(2, "実在籍者", real_member)
time_yukyu = df_pc_k_4["有休日数    "] * 8
df_pc_k_4.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_pc_k_4.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_pc_k_4["欠勤日数    "] * 8
df_pc_k_4.insert(5, "欠勤時間", abs_time)
work_time = df_pc_k_4["勤務時間    "]
df_pc_k_4.insert(6, "勤務時間", work_time)
late_early_time = df_pc_k_4["遅早時間    "]
df_pc_k_4.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_pc_k_4["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_pc_k_4.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_pc_k_4.insert(17, "実労働時間", real_work_time)
zure_time = df_pc_k_4["ズレ時間    "]
df_pc_k_4.insert(18, "ズレ時間", zure_time)
overtime = df_pc_k_4["残業時間    "] + df_pc_k_4["深夜残業    "]
df_pc_k_4.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_pc_k_4.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_pc_k_4["法外休出    "]
df_pc_k_4.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_pc_k_4.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_pc_k_4["法定休出    "]
df_pc_k_4.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_pc_k_4.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_pc_k_4["６０Ｈ超    "]
df_pc_k_4.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_pc_k_4["代休時間    "] + df_pc_k_4["深夜代休    "]
df_pc_k_4.insert(31, "代休時間", holiday_time)
df_pc_k_4.insert(32, "応援時間", 0)
total_work_time = (
    df_pc_k_4["勤務時間    "]
    + df_pc_k_4["残業時間    "]
    + df_pc_k_4["法外休出    "]
    + df_pc_k_4["法定休出    "]
)
df_pc_k_4.insert(33, "総労働時間", total_work_time)
basic_salary = df_pc_k_4["基 本 給    "] + df_pc_k_4["支給額"]
df_pc_k_4.insert(35, "基本給", basic_salary)
post_allowance = df_pc_k_4["役職手当    "]
df_pc_k_4.insert(36, "役職手当", post_allowance)
sales_allowance = df_pc_k_4["営業手当    "]
df_pc_k_4.insert(37, "営業手当", sales_allowance)
aria_allowance = df_pc_k_4["地域手当    "]
df_pc_k_4.insert(38, "地域手当", aria_allowance)
spe_allowance = df_pc_k_4["特殊手当    "]
df_pc_k_4.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_pc_k_4["特別技技手当"]
df_pc_k_4.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_pc_k_4["調整手当    "]
df_pc_k_4.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_pc_k_4["別居手当    "]
df_pc_k_4.insert(42, "別居手当", sep_allowance)
com_allowance = df_pc_k_4["通勤手当    "]
df_pc_k_4.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_pc_k_4.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_pc_k_4["残業手当    "]
df_pc_k_4.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_pc_k_4["休出手当    "]
df_pc_k_4.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_pc_k_4["深夜勤務手当"]
df_pc_k_4.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_pc_k_4["交替時差手当"]
df_pc_k_4.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_pc_k_4["休業手当    "]
df_pc_k_4.insert(49, "休業手当", closed_allowance)
closed_deduction = df_pc_k_4["休業控除    "]
df_pc_k_4.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_pc_k_4["代 休 他    "]
df_pc_k_4.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_pc_k_4["欠勤控除    "] + df_pc_k_4["遅早控除    "]
df_pc_k_4.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_pc_k_4["精 算 分    "]
df_pc_k_4.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_pc_k_4.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_pc_k_4.insert(55, "総支給額", total)
df_pc_k_4.insert(56, "応援時間額", 0)
df_pc_k_4.insert(57, "役員振替", 0)
df_pc_k_4.insert(58, "部門振替", 0)
df_pc_k_4.insert(59, "合計", 0)
# 不要フィールドの削除
df_pc_k_4 = df_pc_k_4.drop("所属2", axis=1)
df_pc_k_4 = df_pc_k_4.drop("所属3", axis=1)
df_pc_k_4 = df_pc_k_4.drop("社員ｺｰﾄﾞ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("区分", axis=1)
df_pc_k_4 = df_pc_k_4.drop("出勤日数    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("有休日数    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("欠勤日数    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("残業時間    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("深夜残業    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("法外休出    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("法定休出    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("代休時間    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("深夜代休    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("６０Ｈ超    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("別居手当    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("通勤手当    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("特別技技手当", axis=1)
df_pc_k_4 = df_pc_k_4.drop("特殊手当    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("地域手当    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("営業手当    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("役職手当    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("調整手当    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("基 本 給    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("残業手当    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("休出手当    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("深夜勤務手当", axis=1)
df_pc_k_4 = df_pc_k_4.drop("交替時差手当", axis=1)
df_pc_k_4 = df_pc_k_4.drop("休業手当    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("代 休 他    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("欠勤控除    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("遅早控除    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("精 算 分    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("支給合計額  ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("休業控除    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("支給額", axis=1)
df_pc_k_4 = df_pc_k_4.drop("ズレ時間    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("雑費・食事代", axis=1)
df_pc_k_4 = df_pc_k_4.drop("雑費・衣靴代", axis=1)
df_pc_k_4 = df_pc_k_4.drop("雑費        ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("受診料・他  ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("雑費・会費等", axis=1)
df_pc_k_4 = df_pc_k_4.drop("勤務時間    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("遅早時間    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("特休日数    ", axis=1)
df_pc_k_4 = df_pc_k_4.drop("集計区分－２        ", axis=1)
df_pc_k_4 = df_pc_k_4.sum()
df_pc_k_4.to_csv(
    "/content/drive/MyDrive/data/PC/C.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# PC_間接5
df_pc_k_5 = df_pc_m.groupby("区分").get_group("間接5")
df_pc_k_5 = df_pc_k_5.drop("所属1", axis=1)
member = df_pc_k_5["所属2"] > 0
df_pc_k_5.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_pc_k_5["所属3"] > 0
df_pc_k_5.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_pc_k_5["出勤日数    "]
df_pc_k_5.insert(2, "実在籍者", real_member)
time_yukyu = df_pc_k_5["有休日数    "] * 8
df_pc_k_5.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_pc_k_5.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_pc_k_5["欠勤日数    "] * 8
df_pc_k_5.insert(5, "欠勤時間", abs_time)
work_time = df_pc_k_5["勤務時間    "]
df_pc_k_5.insert(6, "勤務時間", work_time)
late_early_time = df_pc_k_5["遅早時間    "]
df_pc_k_5.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_pc_k_5["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_pc_k_5.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_pc_k_5.insert(17, "実労働時間", real_work_time)
zure_time = df_pc_k_5["ズレ時間    "]
df_pc_k_5.insert(18, "ズレ時間", zure_time)
overtime = df_pc_k_5["残業時間    "] + df_pc_k_5["深夜残業    "]
df_pc_k_5.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_pc_k_5.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_pc_k_5["法外休出    "]
df_pc_k_5.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_pc_k_5.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_pc_k_5["法定休出    "]
df_pc_k_5.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_pc_k_5.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_pc_k_5["６０Ｈ超    "]
df_pc_k_5.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_pc_k_5["代休時間    "] + df_pc_k_5["深夜代休    "]
df_pc_k_5.insert(31, "代休時間", holiday_time)
df_pc_k_5.insert(32, "応援時間", 0)
total_work_time = (
    df_pc_k_5["勤務時間    "]
    + df_pc_k_5["残業時間    "]
    + df_pc_k_5["法外休出    "]
    + df_pc_k_5["法定休出    "]
)
df_pc_k_5.insert(33, "総労働時間", total_work_time)
basic_salary = df_pc_k_5["基 本 給    "] + df_pc_k_5["支給額"]
df_pc_k_5.insert(35, "基本給", basic_salary)
post_allowance = df_pc_k_5["役職手当    "]
df_pc_k_5.insert(36, "役職手当", post_allowance)
sales_allowance = df_pc_k_5["営業手当    "]
df_pc_k_5.insert(37, "営業手当", sales_allowance)
aria_allowance = df_pc_k_5["地域手当    "]
df_pc_k_5.insert(38, "地域手当", aria_allowance)
spe_allowance = df_pc_k_5["特殊手当    "]
df_pc_k_5.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_pc_k_5["特別技技手当"]
df_pc_k_5.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_pc_k_5["調整手当    "]
df_pc_k_5.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_pc_k_5["別居手当    "]
df_pc_k_5.insert(42, "別居手当", sep_allowance)
com_allowance = df_pc_k_5["通勤手当    "]
df_pc_k_5.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_pc_k_5.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_pc_k_5["残業手当    "]
df_pc_k_5.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_pc_k_5["休出手当    "]
df_pc_k_5.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_pc_k_5["深夜勤務手当"]
df_pc_k_5.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_pc_k_5["交替時差手当"]
df_pc_k_5.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_pc_k_5["休業手当    "]
df_pc_k_5.insert(49, "休業手当", closed_allowance)
closed_deduction = df_pc_k_5["休業控除    "]
df_pc_k_5.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_pc_k_5["代 休 他    "]
df_pc_k_5.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_pc_k_5["欠勤控除    "] + df_pc_k_5["遅早控除    "]
df_pc_k_5.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_pc_k_5["精 算 分    "]
df_pc_k_5.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_pc_k_5.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_pc_k_5.insert(55, "総支給額", total)
df_pc_k_5.insert(56, "応援時間額", 0)
df_pc_k_5.insert(57, "役員振替", 0)
df_pc_k_5.insert(58, "部門振替", 0)
df_pc_k_5.insert(59, "合計", 0)
# 不要フィールドの削除
df_pc_k_5 = df_pc_k_5.drop("所属2", axis=1)
df_pc_k_5 = df_pc_k_5.drop("所属3", axis=1)
df_pc_k_5 = df_pc_k_5.drop("社員ｺｰﾄﾞ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("区分", axis=1)
df_pc_k_5 = df_pc_k_5.drop("出勤日数    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("有休日数    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("欠勤日数    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("残業時間    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("深夜残業    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("法外休出    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("法定休出    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("代休時間    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("深夜代休    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("６０Ｈ超    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("別居手当    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("通勤手当    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("特別技技手当", axis=1)
df_pc_k_5 = df_pc_k_5.drop("特殊手当    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("地域手当    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("営業手当    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("役職手当    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("調整手当    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("基 本 給    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("残業手当    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("休出手当    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("深夜勤務手当", axis=1)
df_pc_k_5 = df_pc_k_5.drop("交替時差手当", axis=1)
df_pc_k_5 = df_pc_k_5.drop("休業手当    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("代 休 他    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("欠勤控除    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("遅早控除    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("精 算 分    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("支給合計額  ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("休業控除    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("支給額", axis=1)
df_pc_k_5 = df_pc_k_5.drop("ズレ時間    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("雑費・食事代", axis=1)
df_pc_k_5 = df_pc_k_5.drop("雑費・衣靴代", axis=1)
df_pc_k_5 = df_pc_k_5.drop("雑費        ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("受診料・他  ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("雑費・会費等", axis=1)
df_pc_k_5 = df_pc_k_5.drop("勤務時間    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("遅早時間    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("特休日数    ", axis=1)
df_pc_k_5 = df_pc_k_5.drop("集計区分－２        ", axis=1)
df_pc_k_5 = df_pc_k_5.sum()
df_pc_k_5.to_csv(
    "/content/drive/MyDrive/data/PC/D.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# PC_間接6
df_pc_k_6 = df_pc_m.groupby("区分").get_group("間接6")
df_pc_k_6 = df_pc_k_6.drop("所属1", axis=1)
member = df_pc_k_6["所属2"] > 0
df_pc_k_6.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_pc_k_6["所属3"] > 0
df_pc_k_6.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_pc_k_6["出勤日数    "]
df_pc_k_6.insert(2, "実在籍者", real_member)
time_yukyu = df_pc_k_6["有休日数    "] * 8
df_pc_k_6.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_pc_k_6.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_pc_k_6["欠勤日数    "] * 8
df_pc_k_6.insert(5, "欠勤時間", abs_time)
work_time = df_pc_k_6["勤務時間    "]
df_pc_k_6.insert(6, "勤務時間", work_time)
late_early_time = df_pc_k_6["遅早時間    "]
df_pc_k_6.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_pc_k_6["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_pc_k_6.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_pc_k_6.insert(17, "実労働時間", real_work_time)
zure_time = df_pc_k_6["ズレ時間    "]
df_pc_k_6.insert(18, "ズレ時間", zure_time)
overtime = df_pc_k_6["残業時間    "] + df_pc_k_6["深夜残業    "]
df_pc_k_6.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_pc_k_6.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_pc_k_6["法外休出    "]
df_pc_k_6.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_pc_k_6.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_pc_k_6["法定休出    "]
df_pc_k_6.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_pc_k_6.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_pc_k_6["６０Ｈ超    "]
df_pc_k_6.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_pc_k_6["代休時間    "] + df_pc_k_6["深夜代休    "]
df_pc_k_6.insert(31, "代休時間", holiday_time)
df_pc_k_6.insert(32, "応援時間", 0)
total_work_time = (
    df_pc_k_6["勤務時間    "]
    + df_pc_k_6["残業時間    "]
    + df_pc_k_6["法外休出    "]
    + df_pc_k_6["法定休出    "]
)
df_pc_k_6.insert(33, "総労働時間", total_work_time)
basic_salary = df_pc_k_6["基 本 給    "] + df_pc_k_6["支給額"]
df_pc_k_6.insert(35, "基本給", basic_salary)
post_allowance = df_pc_k_6["役職手当    "]
df_pc_k_6.insert(36, "役職手当", post_allowance)
sales_allowance = df_pc_k_6["営業手当    "]
df_pc_k_6.insert(37, "営業手当", sales_allowance)
aria_allowance = df_pc_k_6["地域手当    "]
df_pc_k_6.insert(38, "地域手当", aria_allowance)
spe_allowance = df_pc_k_6["特殊手当    "]
df_pc_k_6.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_pc_k_6["特別技技手当"]
df_pc_k_6.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_pc_k_6["調整手当    "]
df_pc_k_6.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_pc_k_6["別居手当    "]
df_pc_k_6.insert(42, "別居手当", sep_allowance)
com_allowance = df_pc_k_6["通勤手当    "]
df_pc_k_6.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_pc_k_6.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_pc_k_6["残業手当    "]
df_pc_k_6.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_pc_k_6["休出手当    "]
df_pc_k_6.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_pc_k_6["深夜勤務手当"]
df_pc_k_6.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_pc_k_6["交替時差手当"]
df_pc_k_6.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_pc_k_6["休業手当    "]
df_pc_k_6.insert(49, "休業手当", closed_allowance)
closed_deduction = df_pc_k_6["休業控除    "]
df_pc_k_6.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_pc_k_6["代 休 他    "]
df_pc_k_6.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_pc_k_6["欠勤控除    "] + df_pc_k_6["遅早控除    "]
df_pc_k_6.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_pc_k_6["精 算 分    "]
df_pc_k_6.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_pc_k_6.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_pc_k_6.insert(55, "総支給額", total)
df_pc_k_6.insert(56, "応援時間額", 0)
df_pc_k_6.insert(57, "役員振替", 0)
df_pc_k_6.insert(58, "部門振替", 0)
df_pc_k_6.insert(59, "合計", 0)
# 不要フィールドの削除
df_pc_k_6 = df_pc_k_6.drop("所属2", axis=1)
df_pc_k_6 = df_pc_k_6.drop("所属3", axis=1)
df_pc_k_6 = df_pc_k_6.drop("社員ｺｰﾄﾞ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("区分", axis=1)
df_pc_k_6 = df_pc_k_6.drop("出勤日数    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("有休日数    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("欠勤日数    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("残業時間    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("深夜残業    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("法外休出    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("法定休出    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("代休時間    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("深夜代休    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("６０Ｈ超    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("別居手当    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("通勤手当    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("特別技技手当", axis=1)
df_pc_k_6 = df_pc_k_6.drop("特殊手当    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("地域手当    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("営業手当    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("役職手当    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("調整手当    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("基 本 給    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("残業手当    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("休出手当    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("深夜勤務手当", axis=1)
df_pc_k_6 = df_pc_k_6.drop("交替時差手当", axis=1)
df_pc_k_6 = df_pc_k_6.drop("休業手当    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("代 休 他    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("欠勤控除    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("遅早控除    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("精 算 分    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("支給合計額  ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("休業控除    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("支給額", axis=1)
df_pc_k_6 = df_pc_k_6.drop("ズレ時間    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("雑費・食事代", axis=1)
df_pc_k_6 = df_pc_k_6.drop("雑費・衣靴代", axis=1)
df_pc_k_6 = df_pc_k_6.drop("雑費        ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("受診料・他  ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("雑費・会費等", axis=1)
df_pc_k_6 = df_pc_k_6.drop("勤務時間    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("遅早時間    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("特休日数    ", axis=1)
df_pc_k_6 = df_pc_k_6.drop("集計区分－２        ", axis=1)
df_pc_k_6 = df_pc_k_6.sum()
df_pc_k_6.to_csv(
    "/content/drive/MyDrive/data/PC/E.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# PC_直接1
df_pc_t_1 = df_pc_m.groupby("区分").get_group("直接1")
df_pc_t_1 = df_pc_t_1.drop("所属1", axis=1)
member = df_pc_t_1["所属2"] > 0
df_pc_t_1.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_pc_t_1["所属3"] > 0
df_pc_t_1.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_pc_t_1["出勤日数    "]
df_pc_t_1.insert(2, "実在籍者", real_member)
time_yukyu = df_pc_t_1["有休日数    "] * 8
df_pc_t_1.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_pc_t_1.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_pc_t_1["欠勤日数    "] * 8
df_pc_t_1.insert(5, "欠勤時間", abs_time)
work_time = df_pc_t_1["勤務時間    "]
df_pc_t_1.insert(6, "勤務時間", work_time)
late_early_time = df_pc_t_1["遅早時間    "]
df_pc_t_1.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_pc_t_1["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_pc_t_1.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_pc_t_1.insert(17, "実労働時間", real_work_time)
zure_time = df_pc_t_1["ズレ時間    "]
df_pc_t_1.insert(18, "ズレ時間", zure_time)
overtime = df_pc_t_1["残業時間    "] + df_pc_t_1["深夜残業    "]
df_pc_t_1.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_pc_t_1.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_pc_t_1["法外休出    "]
df_pc_t_1.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_pc_t_1.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_pc_t_1["法定休出    "]
df_pc_t_1.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_pc_t_1.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_pc_t_1["６０Ｈ超    "]
df_pc_t_1.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_pc_t_1["代休時間    "] + df_pc_t_1["深夜代休    "]
df_pc_t_1.insert(31, "代休時間", holiday_time)
df_pc_t_1.insert(32, "応援時間", 0)
total_work_time = (
    df_pc_t_1["勤務時間    "]
    + df_pc_t_1["残業時間    "]
    + df_pc_t_1["法外休出    "]
    + df_pc_t_1["法定休出    "]
)
df_pc_t_1.insert(33, "総労働時間", total_work_time)
basic_salary = df_pc_t_1["基 本 給    "] + df_pc_t_1["支給額"]
df_pc_t_1.insert(35, "基本給", basic_salary)
post_allowance = df_pc_t_1["役職手当    "]
df_pc_t_1.insert(36, "役職手当", post_allowance)
sales_allowance = df_pc_t_1["営業手当    "]
df_pc_t_1.insert(37, "営業手当", sales_allowance)
aria_allowance = df_pc_t_1["地域手当    "]
df_pc_t_1.insert(38, "地域手当", aria_allowance)
spe_allowance = df_pc_t_1["特殊手当    "]
df_pc_t_1.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_pc_t_1["特別技技手当"]
df_pc_t_1.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_pc_t_1["調整手当    "]
df_pc_t_1.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_pc_t_1["別居手当    "]
df_pc_t_1.insert(42, "別居手当", sep_allowance)
com_allowance = df_pc_t_1["通勤手当    "]
df_pc_t_1.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_pc_t_1.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_pc_t_1["残業手当    "]
df_pc_t_1.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_pc_t_1["休出手当    "]
df_pc_t_1.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_pc_t_1["深夜勤務手当"]
df_pc_t_1.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_pc_t_1["交替時差手当"]
df_pc_t_1.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_pc_t_1["休業手当    "]
df_pc_t_1.insert(49, "休業手当", closed_allowance)
closed_deduction = df_pc_t_1["休業控除    "]
df_pc_t_1.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_pc_t_1["代 休 他    "]
df_pc_t_1.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_pc_t_1["欠勤控除    "] + df_pc_t_1["遅早控除    "]
df_pc_t_1.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_pc_t_1["精 算 分    "]
df_pc_t_1.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_pc_t_1.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_pc_t_1.insert(55, "総支給額", total)
df_pc_t_1.insert(56, "応援時間額", 0)
df_pc_t_1.insert(57, "役員振替", 0)
df_pc_t_1.insert(58, "部門振替", 0)
df_pc_t_1.insert(59, "合計", 0)
# 不要フィールドの削除
df_pc_t_1 = df_pc_t_1.drop("所属2", axis=1)
df_pc_t_1 = df_pc_t_1.drop("所属3", axis=1)
df_pc_t_1 = df_pc_t_1.drop("社員ｺｰﾄﾞ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("区分", axis=1)
df_pc_t_1 = df_pc_t_1.drop("出勤日数    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("有休日数    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("欠勤日数    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("残業時間    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("深夜残業    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("法外休出    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("法定休出    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("代休時間    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("深夜代休    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("６０Ｈ超    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("別居手当    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("通勤手当    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("特別技技手当", axis=1)
df_pc_t_1 = df_pc_t_1.drop("特殊手当    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("地域手当    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("営業手当    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("役職手当    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("調整手当    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("基 本 給    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("残業手当    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("休出手当    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("深夜勤務手当", axis=1)
df_pc_t_1 = df_pc_t_1.drop("交替時差手当", axis=1)
df_pc_t_1 = df_pc_t_1.drop("休業手当    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("代 休 他    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("欠勤控除    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("遅早控除    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("精 算 分    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("支給合計額  ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("休業控除    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("支給額", axis=1)
df_pc_t_1 = df_pc_t_1.drop("ズレ時間    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("雑費・食事代", axis=1)
df_pc_t_1 = df_pc_t_1.drop("雑費・衣靴代", axis=1)
df_pc_t_1 = df_pc_t_1.drop("雑費        ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("受診料・他  ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("雑費・会費等", axis=1)
df_pc_t_1 = df_pc_t_1.drop("勤務時間    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("遅早時間    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("特休日数    ", axis=1)
df_pc_t_1 = df_pc_t_1.drop("集計区分－２        ", axis=1)
df_pc_t_1 = df_pc_t_1.sum()
df_pc_t_1.to_csv(
    "/content/drive/MyDrive/data/PC/F.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# PC_直接4
df_pc_t_4 = df_pc_m.groupby("区分").get_group("直接4")
df_pc_t_4 = df_pc_t_4.drop("所属1", axis=1)
member = df_pc_t_4["所属2"] > 0
df_pc_t_4.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_pc_t_4["所属3"] > 0
df_pc_t_4.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_pc_t_4["出勤日数    "]
df_pc_t_4.insert(2, "実在籍者", real_member)
time_yukyu = df_pc_t_4["有休日数    "] * 8
df_pc_t_4.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_pc_t_4.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_pc_t_4["欠勤日数    "] * 8
df_pc_t_4.insert(5, "欠勤時間", abs_time)
work_time = df_pc_t_4["勤務時間    "]
df_pc_t_4.insert(6, "勤務時間", work_time)
late_early_time = df_pc_t_4["遅早時間    "]
df_pc_t_4.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_pc_t_4["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_pc_t_4.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_pc_t_4.insert(17, "実労働時間", real_work_time)
zure_time = df_pc_t_4["ズレ時間    "]
df_pc_t_4.insert(18, "ズレ時間", zure_time)
overtime = df_pc_t_4["残業時間    "] + df_pc_t_4["深夜残業    "]
df_pc_t_4.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_pc_t_4.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_pc_t_4["法外休出    "]
df_pc_t_4.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_pc_t_4.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_pc_t_4["法定休出    "]
df_pc_t_4.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_pc_t_4.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_pc_t_4["６０Ｈ超    "]
df_pc_t_4.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_pc_t_4["代休時間    "] + df_pc_t_4["深夜代休    "]
df_pc_t_4.insert(31, "代休時間", holiday_time)
df_pc_t_4.insert(32, "応援時間", 0)
total_work_time = (
    df_pc_t_4["勤務時間    "]
    + df_pc_t_4["残業時間    "]
    + df_pc_t_4["法外休出    "]
    + df_pc_t_4["法定休出    "]
)
df_pc_t_4.insert(33, "総労働時間", total_work_time)
basic_salary = df_pc_t_4["基 本 給    "] + df_pc_t_4["支給額"]
df_pc_t_4.insert(35, "基本給", basic_salary)
post_allowance = df_pc_t_4["役職手当    "]
df_pc_t_4.insert(36, "役職手当", post_allowance)
sales_allowance = df_pc_t_4["営業手当    "]
df_pc_t_4.insert(37, "営業手当", sales_allowance)
aria_allowance = df_pc_t_4["地域手当    "]
df_pc_t_4.insert(38, "地域手当", aria_allowance)
spe_allowance = df_pc_t_4["特殊手当    "]
df_pc_t_4.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_pc_t_4["特別技技手当"]
df_pc_t_4.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_pc_t_4["調整手当    "]
df_pc_t_4.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_pc_t_4["別居手当    "]
df_pc_t_4.insert(42, "別居手当", sep_allowance)
com_allowance = df_pc_t_4["通勤手当    "]
df_pc_t_4.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_pc_t_4.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_pc_t_4["残業手当    "]
df_pc_t_4.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_pc_t_4["休出手当    "]
df_pc_t_4.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_pc_t_4["深夜勤務手当"]
df_pc_t_4.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_pc_t_4["交替時差手当"]
df_pc_t_4.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_pc_t_4["休業手当    "]
df_pc_t_4.insert(49, "休業手当", closed_allowance)
closed_deduction = df_pc_t_4["休業控除    "]
df_pc_t_4.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_pc_t_4["代 休 他    "]
df_pc_t_4.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_pc_t_4["欠勤控除    "] + df_pc_t_4["遅早控除    "]
df_pc_t_4.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_pc_t_4["精 算 分    "]
df_pc_t_4.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_pc_t_4.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_pc_t_4.insert(55, "総支給額", total)
df_pc_t_4.insert(56, "応援時間額", 0)
df_pc_t_4.insert(57, "役員振替", 0)
df_pc_t_4.insert(58, "部門振替", 0)
df_pc_t_4.insert(59, "合計", 0)
# 不要フィールドの削除
df_pc_t_4 = df_pc_t_4.drop("所属2", axis=1)
df_pc_t_4 = df_pc_t_4.drop("所属3", axis=1)
df_pc_t_4 = df_pc_t_4.drop("社員ｺｰﾄﾞ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("区分", axis=1)
df_pc_t_4 = df_pc_t_4.drop("出勤日数    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("有休日数    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("欠勤日数    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("残業時間    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("深夜残業    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("法外休出    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("法定休出    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("代休時間    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("深夜代休    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("６０Ｈ超    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("別居手当    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("通勤手当    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("特別技技手当", axis=1)
df_pc_t_4 = df_pc_t_4.drop("特殊手当    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("地域手当    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("営業手当    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("役職手当    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("調整手当    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("基 本 給    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("残業手当    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("休出手当    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("深夜勤務手当", axis=1)
df_pc_t_4 = df_pc_t_4.drop("交替時差手当", axis=1)
df_pc_t_4 = df_pc_t_4.drop("休業手当    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("代 休 他    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("欠勤控除    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("遅早控除    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("精 算 分    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("支給合計額  ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("休業控除    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("支給額", axis=1)
df_pc_t_4 = df_pc_t_4.drop("ズレ時間    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("雑費・食事代", axis=1)
df_pc_t_4 = df_pc_t_4.drop("雑費・衣靴代", axis=1)
df_pc_t_4 = df_pc_t_4.drop("雑費        ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("受診料・他  ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("雑費・会費等", axis=1)
df_pc_t_4 = df_pc_t_4.drop("勤務時間    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("遅早時間    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("特休日数    ", axis=1)
df_pc_t_4 = df_pc_t_4.drop("集計区分－２        ", axis=1)
df_pc_t_4 = df_pc_t_4.sum()
df_pc_t_4.to_csv(
    "/content/drive/MyDrive/data/PC/G.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 宮城_間接1
df_miyagi_k_1 = df_miyagi_m.groupby("区分").get_group("間接1")
df_miyagi_k_1 = df_miyagi_k_1.drop("所属1", axis=1)
member = df_miyagi_k_1["所属2"] > 0
df_miyagi_k_1.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_miyagi_k_1["所属3"] > 0
df_miyagi_k_1.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_miyagi_k_1["出勤日数    "]
df_miyagi_k_1.insert(2, "実在籍者", real_member)
time_yukyu = df_miyagi_k_1["有休日数    "] * 8
df_miyagi_k_1.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_miyagi_k_1.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_miyagi_k_1["欠勤日数    "] * 8
df_miyagi_k_1.insert(5, "欠勤時間", abs_time)
work_time = df_miyagi_k_1["勤務時間    "]
df_miyagi_k_1.insert(6, "勤務時間", work_time)
late_early_time = df_miyagi_k_1["遅早時間    "]
df_miyagi_k_1.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_miyagi_k_1["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_miyagi_k_1.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_miyagi_k_1.insert(17, "実労働時間", real_work_time)
zure_time = df_miyagi_k_1["ズレ時間    "]
df_miyagi_k_1.insert(18, "ズレ時間", zure_time)
overtime = df_miyagi_k_1["残業時間    "] + df_miyagi_k_1["深夜残業    "]
df_miyagi_k_1.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_miyagi_k_1.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_miyagi_k_1["法外休出    "]
df_miyagi_k_1.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_miyagi_k_1.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_miyagi_k_1["法定休出    "]
df_miyagi_k_1.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_miyagi_k_1.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_miyagi_k_1["６０Ｈ超    "]
df_miyagi_k_1.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_miyagi_k_1["代休時間    "] + df_miyagi_k_1["深夜代休    "]
df_miyagi_k_1.insert(31, "代休時間", holiday_time)
df_miyagi_k_1.insert(32, "応援時間", 0)
total_work_time = (
    df_miyagi_k_1["勤務時間    "]
    + df_miyagi_k_1["残業時間    "]
    + df_miyagi_k_1["法外休出    "]
    + df_miyagi_k_1["法定休出    "]
)
df_miyagi_k_1.insert(33, "総労働時間", total_work_time)
basic_salary = df_miyagi_k_1["基 本 給    "] + df_miyagi_k_1["支給額"]
df_miyagi_k_1.insert(35, "基本給", basic_salary)
post_allowance = df_miyagi_k_1["役職手当    "]
df_miyagi_k_1.insert(36, "役職手当", post_allowance)
sales_allowance = df_miyagi_k_1["営業手当    "]
df_miyagi_k_1.insert(37, "営業手当", sales_allowance)
aria_allowance = df_miyagi_k_1["地域手当    "]
df_miyagi_k_1.insert(38, "地域手当", aria_allowance)
spe_allowance = df_miyagi_k_1["特殊手当    "]
df_miyagi_k_1.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_miyagi_k_1["特別技技手当"]
df_miyagi_k_1.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_miyagi_k_1["調整手当    "]
df_miyagi_k_1.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_miyagi_k_1["別居手当    "]
df_miyagi_k_1.insert(42, "別居手当", sep_allowance)
com_allowance = df_miyagi_k_1["通勤手当    "]
df_miyagi_k_1.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_miyagi_k_1.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_miyagi_k_1["残業手当    "]
df_miyagi_k_1.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_miyagi_k_1["休出手当    "]
df_miyagi_k_1.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_miyagi_k_1["深夜勤務手当"]
df_miyagi_k_1.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_miyagi_k_1["交替時差手当"]
df_miyagi_k_1.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_miyagi_k_1["休業手当    "]
df_miyagi_k_1.insert(49, "休業手当", closed_allowance)
closed_deduction = df_miyagi_k_1["休業控除    "]
df_miyagi_k_1.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_miyagi_k_1["代 休 他    "]
df_miyagi_k_1.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_miyagi_k_1["欠勤控除    "] + df_miyagi_k_1["遅早控除    "]
df_miyagi_k_1.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_miyagi_k_1["精 算 分    "]
df_miyagi_k_1.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_miyagi_k_1.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_miyagi_k_1.insert(55, "総支給額", total)
df_miyagi_k_1.insert(56, "応援時間額", 0)
df_miyagi_k_1.insert(57, "役員振替", 0)
df_miyagi_k_1.insert(58, "部門振替", 0)
df_miyagi_k_1.insert(59, "合計", 0)
# 不要フィールドの削除
df_miyagi_k_1 = df_miyagi_k_1.drop("所属2", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("所属3", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("社員ｺｰﾄﾞ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("区分", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("出勤日数    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("有休日数    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("欠勤日数    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("残業時間    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("深夜残業    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("法外休出    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("法定休出    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("代休時間    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("深夜代休    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("６０Ｈ超    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("別居手当    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("通勤手当    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("特別技技手当", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("特殊手当    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("地域手当    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("営業手当    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("役職手当    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("調整手当    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("基 本 給    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("残業手当    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("休出手当    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("深夜勤務手当", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("交替時差手当", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("休業手当    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("代 休 他    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("欠勤控除    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("遅早控除    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("精 算 分    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("支給合計額  ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("休業控除    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("支給額", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("ズレ時間    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("雑費・食事代", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("雑費・衣靴代", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("雑費        ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("受診料・他  ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("雑費・会費等", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("勤務時間    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("遅早時間    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("特休日数    ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.drop("集計区分－２        ", axis=1)
df_miyagi_k_1 = df_miyagi_k_1.sum()
df_miyagi_k_1.to_csv(
    "/content/drive/MyDrive/data/宮城/A.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 宮城_間接2
df_miyagi_k_2 = df_miyagi_m.groupby("区分").get_group("間接2")
df_miyagi_k_2 = df_miyagi_k_2.drop("所属1", axis=1)
member = df_miyagi_k_2["所属2"] > 0
df_miyagi_k_2.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_miyagi_k_2["所属3"] > 0
df_miyagi_k_2.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_miyagi_k_2["出勤日数    "]
df_miyagi_k_2.insert(2, "実在籍者", real_member)
time_yukyu = df_miyagi_k_2["有休日数    "] * 8
df_miyagi_k_2.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_miyagi_k_2.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_miyagi_k_2["欠勤日数    "] * 8
df_miyagi_k_2.insert(5, "欠勤時間", abs_time)
work_time = df_miyagi_k_2["勤務時間    "]
df_miyagi_k_2.insert(6, "勤務時間", work_time)
late_early_time = df_miyagi_k_2["遅早時間    "]
df_miyagi_k_2.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_miyagi_k_2["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_miyagi_k_2.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_miyagi_k_2.insert(17, "実労働時間", real_work_time)
zure_time = df_miyagi_k_2["ズレ時間    "]
df_miyagi_k_2.insert(18, "ズレ時間", zure_time)
overtime = df_miyagi_k_2["残業時間    "] + df_miyagi_k_2["深夜残業    "]
df_miyagi_k_2.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_miyagi_k_2.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_miyagi_k_2["法外休出    "]
df_miyagi_k_2.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_miyagi_k_2.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_miyagi_k_2["法定休出    "]
df_miyagi_k_2.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_miyagi_k_2.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_miyagi_k_2["６０Ｈ超    "]
df_miyagi_k_2.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_miyagi_k_2["代休時間    "] + df_miyagi_k_2["深夜代休    "]
df_miyagi_k_2.insert(31, "代休時間", holiday_time)
df_miyagi_k_2.insert(32, "応援時間", 0)
total_work_time = (
    df_miyagi_k_2["勤務時間    "]
    + df_miyagi_k_2["残業時間    "]
    + df_miyagi_k_2["法外休出    "]
    + df_miyagi_k_2["法定休出    "]
)
df_miyagi_k_2.insert(33, "総労働時間", total_work_time)
basic_salary = df_miyagi_k_2["基 本 給    "] + df_miyagi_k_2["支給額"]
df_miyagi_k_2.insert(35, "基本給", basic_salary)
post_allowance = df_miyagi_k_2["役職手当    "]
df_miyagi_k_2.insert(36, "役職手当", post_allowance)
sales_allowance = df_miyagi_k_2["営業手当    "]
df_miyagi_k_2.insert(37, "営業手当", sales_allowance)
aria_allowance = df_miyagi_k_2["地域手当    "]
df_miyagi_k_2.insert(38, "地域手当", aria_allowance)
spe_allowance = df_miyagi_k_2["特殊手当    "]
df_miyagi_k_2.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_miyagi_k_2["特別技技手当"]
df_miyagi_k_2.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_miyagi_k_2["調整手当    "]
df_miyagi_k_2.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_miyagi_k_2["別居手当    "]
df_miyagi_k_2.insert(42, "別居手当", sep_allowance)
com_allowance = df_miyagi_k_2["通勤手当    "]
df_miyagi_k_2.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_miyagi_k_2.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_miyagi_k_2["残業手当    "]
df_miyagi_k_2.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_miyagi_k_2["休出手当    "]
df_miyagi_k_2.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_miyagi_k_2["深夜勤務手当"]
df_miyagi_k_2.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_miyagi_k_2["交替時差手当"]
df_miyagi_k_2.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_miyagi_k_2["休業手当    "]
df_miyagi_k_2.insert(49, "休業手当", closed_allowance)
closed_deduction = df_miyagi_k_2["休業控除    "]
df_miyagi_k_2.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_miyagi_k_2["代 休 他    "]
df_miyagi_k_2.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_miyagi_k_2["欠勤控除    "] + df_miyagi_k_2["遅早控除    "]
df_miyagi_k_2.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_miyagi_k_2["精 算 分    "]
df_miyagi_k_2.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_miyagi_k_2.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_miyagi_k_2.insert(55, "総支給額", total)
df_miyagi_k_2.insert(56, "応援時間額", 0)
df_miyagi_k_2.insert(57, "役員振替", 0)
df_miyagi_k_2.insert(58, "部門振替", 0)
df_miyagi_k_2.insert(59, "合計", 0)
# 不要フィールドの削除
df_miyagi_k_2 = df_miyagi_k_2.drop("所属2", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("所属3", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("社員ｺｰﾄﾞ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("区分", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("出勤日数    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("有休日数    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("欠勤日数    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("残業時間    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("深夜残業    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("法外休出    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("法定休出    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("代休時間    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("深夜代休    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("６０Ｈ超    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("別居手当    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("通勤手当    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("特別技技手当", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("特殊手当    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("地域手当    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("営業手当    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("役職手当    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("調整手当    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("基 本 給    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("残業手当    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("休出手当    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("深夜勤務手当", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("交替時差手当", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("休業手当    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("代 休 他    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("欠勤控除    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("遅早控除    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("精 算 分    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("支給合計額  ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("休業控除    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("支給額", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("ズレ時間    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("雑費・食事代", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("雑費・衣靴代", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("雑費        ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("受診料・他  ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("雑費・会費等", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("勤務時間    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("遅早時間    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("特休日数    ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.drop("集計区分－２        ", axis=1)
df_miyagi_k_2 = df_miyagi_k_2.sum()
df_miyagi_k_2.to_csv(
    "/content/drive/MyDrive/data/宮城/B.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 宮城_間接4
df_miyagi_k_4 = df_miyagi_m.groupby("区分").get_group("間接4")
df_miyagi_k_4 = df_miyagi_k_4.drop("所属1", axis=1)
member = df_miyagi_k_4["所属2"] > 0
df_miyagi_k_4.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_miyagi_k_4["所属3"] > 0
df_miyagi_k_4.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_miyagi_k_4["出勤日数    "]
df_miyagi_k_4.insert(2, "実在籍者", real_member)
time_yukyu = df_miyagi_k_4["有休日数    "] * 8
df_miyagi_k_4.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_miyagi_k_4.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_miyagi_k_4["欠勤日数    "] * 8
df_miyagi_k_4.insert(5, "欠勤時間", abs_time)
work_time = df_miyagi_k_4["勤務時間    "]
df_miyagi_k_4.insert(6, "勤務時間", work_time)
late_early_time = df_miyagi_k_4["遅早時間    "]
df_miyagi_k_4.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_miyagi_k_4["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_miyagi_k_4.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_miyagi_k_4.insert(17, "実労働時間", real_work_time)
zure_time = df_miyagi_k_4["ズレ時間    "]
df_miyagi_k_4.insert(18, "ズレ時間", zure_time)
overtime = df_miyagi_k_4["残業時間    "] + df_miyagi_k_4["深夜残業    "]
df_miyagi_k_4.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_miyagi_k_4.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_miyagi_k_4["法外休出    "]
df_miyagi_k_4.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_miyagi_k_4.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_miyagi_k_4["法定休出    "]
df_miyagi_k_4.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_miyagi_k_4.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_miyagi_k_4["６０Ｈ超    "]
df_miyagi_k_4.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_miyagi_k_4["代休時間    "] + df_miyagi_k_4["深夜代休    "]
df_miyagi_k_4.insert(31, "代休時間", holiday_time)
df_miyagi_k_4.insert(32, "応援時間", 0)
total_work_time = (
    df_miyagi_k_4["勤務時間    "]
    + df_miyagi_k_4["残業時間    "]
    + df_miyagi_k_4["法外休出    "]
    + df_miyagi_k_4["法定休出    "]
)
df_miyagi_k_4.insert(33, "総労働時間", total_work_time)
basic_salary = df_miyagi_k_4["基 本 給    "] + df_miyagi_k_4["支給額"]
df_miyagi_k_4.insert(35, "基本給", basic_salary)
post_allowance = df_miyagi_k_4["役職手当    "]
df_miyagi_k_4.insert(36, "役職手当", post_allowance)
sales_allowance = df_miyagi_k_4["営業手当    "]
df_miyagi_k_4.insert(37, "営業手当", sales_allowance)
aria_allowance = df_miyagi_k_4["地域手当    "]
df_miyagi_k_4.insert(38, "地域手当", aria_allowance)
spe_allowance = df_miyagi_k_4["特殊手当    "]
df_miyagi_k_4.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_miyagi_k_4["特別技技手当"]
df_miyagi_k_4.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_miyagi_k_4["調整手当    "]
df_miyagi_k_4.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_miyagi_k_4["別居手当    "]
df_miyagi_k_4.insert(42, "別居手当", sep_allowance)
com_allowance = df_miyagi_k_4["通勤手当    "]
df_miyagi_k_4.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_miyagi_k_4.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_miyagi_k_4["残業手当    "]
df_miyagi_k_4.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_miyagi_k_4["休出手当    "]
df_miyagi_k_4.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_miyagi_k_4["深夜勤務手当"]
df_miyagi_k_4.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_miyagi_k_4["交替時差手当"]
df_miyagi_k_4.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_miyagi_k_4["休業手当    "]
df_miyagi_k_4.insert(49, "休業手当", closed_allowance)
closed_deduction = df_miyagi_k_4["休業控除    "]
df_miyagi_k_4.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_miyagi_k_4["代 休 他    "]
df_miyagi_k_4.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_miyagi_k_4["欠勤控除    "] + df_miyagi_k_4["遅早控除    "]
df_miyagi_k_4.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_miyagi_k_4["精 算 分    "]
df_miyagi_k_4.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_miyagi_k_4.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_miyagi_k_4.insert(55, "総支給額", total)
df_miyagi_k_4.insert(56, "応援時間額", 0)
df_miyagi_k_4.insert(57, "役員振替", 0)
df_miyagi_k_4.insert(58, "部門振替", 0)
df_miyagi_k_4.insert(59, "合計", 0)
# 不要フィールドの削除
df_miyagi_k_4 = df_miyagi_k_4.drop("所属2", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("所属3", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("社員ｺｰﾄﾞ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("区分", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("出勤日数    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("有休日数    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("欠勤日数    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("残業時間    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("深夜残業    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("法外休出    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("法定休出    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("代休時間    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("深夜代休    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("６０Ｈ超    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("別居手当    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("通勤手当    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("特別技技手当", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("特殊手当    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("地域手当    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("営業手当    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("役職手当    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("調整手当    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("基 本 給    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("残業手当    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("休出手当    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("深夜勤務手当", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("交替時差手当", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("休業手当    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("代 休 他    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("欠勤控除    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("遅早控除    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("精 算 分    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("支給合計額  ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("休業控除    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("支給額", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("ズレ時間    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("雑費・食事代", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("雑費・衣靴代", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("雑費        ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("受診料・他  ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("雑費・会費等", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("勤務時間    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("遅早時間    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("特休日数    ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.drop("集計区分－２        ", axis=1)
df_miyagi_k_4 = df_miyagi_k_4.sum()
df_miyagi_k_4.to_csv(
    "/content/drive/MyDrive/data/宮城/C.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 宮城_間接6
df_miyagi_k_6 = df_miyagi_m.groupby("区分").get_group("間接6")
df_miyagi_k_6 = df_miyagi_k_6.drop("所属1", axis=1)
member = df_miyagi_k_6["所属2"] > 0
df_miyagi_k_6.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_miyagi_k_6["所属3"] > 0
df_miyagi_k_6.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_miyagi_k_6["出勤日数    "]
df_miyagi_k_6.insert(2, "実在籍者", real_member)
time_yukyu = df_miyagi_k_6["有休日数    "] * 8
df_miyagi_k_6.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_miyagi_k_6.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_miyagi_k_6["欠勤日数    "] * 8
df_miyagi_k_6.insert(5, "欠勤時間", abs_time)
work_time = df_miyagi_k_6["勤務時間    "]
df_miyagi_k_6.insert(6, "勤務時間", work_time)
late_early_time = df_miyagi_k_6["遅早時間    "]
df_miyagi_k_6.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_miyagi_k_6["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_miyagi_k_6.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_miyagi_k_6.insert(17, "実労働時間", real_work_time)
zure_time = df_miyagi_k_6["ズレ時間    "]
df_miyagi_k_6.insert(18, "ズレ時間", zure_time)
overtime = df_miyagi_k_6["残業時間    "] + df_miyagi_k_6["深夜残業    "]
df_miyagi_k_6.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_miyagi_k_6.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_miyagi_k_6["法外休出    "]
df_miyagi_k_6.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_miyagi_k_6.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_miyagi_k_6["法定休出    "]
df_miyagi_k_6.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_miyagi_k_6.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_miyagi_k_6["６０Ｈ超    "]
df_miyagi_k_6.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_miyagi_k_6["代休時間    "] + df_miyagi_k_6["深夜代休    "]
df_miyagi_k_6.insert(31, "代休時間", holiday_time)
df_miyagi_k_6.insert(32, "応援時間", 0)
total_work_time = (
    df_miyagi_k_6["勤務時間    "]
    + df_miyagi_k_6["残業時間    "]
    + df_miyagi_k_6["法外休出    "]
    + df_miyagi_k_6["法定休出    "]
)
df_miyagi_k_6.insert(33, "総労働時間", total_work_time)
basic_salary = df_miyagi_k_6["基 本 給    "] + df_miyagi_k_6["支給額"]
df_miyagi_k_6.insert(35, "基本給", basic_salary)
post_allowance = df_miyagi_k_6["役職手当    "]
df_miyagi_k_6.insert(36, "役職手当", post_allowance)
sales_allowance = df_miyagi_k_6["営業手当    "]
df_miyagi_k_6.insert(37, "営業手当", sales_allowance)
aria_allowance = df_miyagi_k_6["地域手当    "]
df_miyagi_k_6.insert(38, "地域手当", aria_allowance)
spe_allowance = df_miyagi_k_6["特殊手当    "]
df_miyagi_k_6.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_miyagi_k_6["特別技技手当"]
df_miyagi_k_6.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_miyagi_k_6["調整手当    "]
df_miyagi_k_6.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_miyagi_k_6["別居手当    "]
df_miyagi_k_6.insert(42, "別居手当", sep_allowance)
com_allowance = df_miyagi_k_6["通勤手当    "]
df_miyagi_k_6.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_miyagi_k_6.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_miyagi_k_6["残業手当    "]
df_miyagi_k_6.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_miyagi_k_6["休出手当    "]
df_miyagi_k_6.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_miyagi_k_6["深夜勤務手当"]
df_miyagi_k_6.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_miyagi_k_6["交替時差手当"]
df_miyagi_k_6.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_miyagi_k_6["休業手当    "]
df_miyagi_k_6.insert(49, "休業手当", closed_allowance)
closed_deduction = df_miyagi_k_6["休業控除    "]
df_miyagi_k_6.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_miyagi_k_6["代 休 他    "]
df_miyagi_k_6.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_miyagi_k_6["欠勤控除    "] + df_miyagi_k_6["遅早控除    "]
df_miyagi_k_6.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_miyagi_k_6["精 算 分    "]
df_miyagi_k_6.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_miyagi_k_6.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_miyagi_k_6.insert(55, "総支給額", total)
df_miyagi_k_6.insert(56, "応援時間額", 0)
df_miyagi_k_6.insert(57, "役員振替", 0)
df_miyagi_k_6.insert(58, "部門振替", 0)
df_miyagi_k_6.insert(59, "合計", 0)
# 不要フィールドの削除
df_miyagi_k_6 = df_miyagi_k_6.drop("所属2", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("所属3", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("社員ｺｰﾄﾞ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("区分", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("出勤日数    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("有休日数    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("欠勤日数    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("残業時間    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("深夜残業    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("法外休出    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("法定休出    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("代休時間    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("深夜代休    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("６０Ｈ超    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("別居手当    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("通勤手当    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("特別技技手当", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("特殊手当    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("地域手当    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("営業手当    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("役職手当    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("調整手当    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("基 本 給    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("残業手当    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("休出手当    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("深夜勤務手当", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("交替時差手当", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("休業手当    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("代 休 他    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("欠勤控除    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("遅早控除    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("精 算 分    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("支給合計額  ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("休業控除    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("支給額", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("ズレ時間    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("雑費・食事代", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("雑費・衣靴代", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("雑費        ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("受診料・他  ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("雑費・会費等", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("勤務時間    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("遅早時間    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("特休日数    ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.drop("集計区分－２        ", axis=1)
df_miyagi_k_6 = df_miyagi_k_6.sum()
df_miyagi_k_6.to_csv(
    "/content/drive/MyDrive/data/宮城/D.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 宮城_直接1
df_miyagi_t_1 = df_miyagi_m.groupby("区分").get_group("直接1")
df_miyagi_t_1 = df_miyagi_t_1.drop("所属1", axis=1)
member = df_miyagi_t_1["所属2"] > 0
df_miyagi_t_1.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_miyagi_t_1["所属3"] > 0
df_miyagi_t_1.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_miyagi_t_1["出勤日数    "]
df_miyagi_t_1.insert(2, "実在籍者", real_member)
time_yukyu = df_miyagi_t_1["有休日数    "] * 8
df_miyagi_t_1.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_miyagi_t_1.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_miyagi_t_1["欠勤日数    "] * 8
df_miyagi_t_1.insert(5, "欠勤時間", abs_time)
work_time = df_miyagi_t_1["勤務時間    "] * 8
df_miyagi_t_1.insert(6, "勤務時間", work_time)
late_early_time = df_miyagi_t_1["遅早時間    "]
df_miyagi_t_1.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_miyagi_t_1["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_miyagi_t_1.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_miyagi_t_1.insert(17, "実労働時間", real_work_time)
zure_time = df_miyagi_t_1["ズレ時間    "]
df_miyagi_t_1.insert(18, "ズレ時間", zure_time)
overtime = df_miyagi_t_1["残業時間    "] + df_miyagi_t_1["深夜残業    "]
df_miyagi_t_1.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_miyagi_t_1.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_miyagi_t_1["法外休出    "]
df_miyagi_t_1.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_miyagi_t_1.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_miyagi_t_1["法定休出    "]
df_miyagi_t_1.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_miyagi_t_1.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_miyagi_t_1["６０Ｈ超    "]
df_miyagi_t_1.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_miyagi_t_1["代休時間    "] + df_miyagi_t_1["深夜代休    "]
df_miyagi_t_1.insert(31, "代休時間", holiday_time)
df_miyagi_t_1.insert(32, "応援時間", 0)
total_work_time = (
    df_miyagi_t_1["勤務時間    "]
    + df_miyagi_t_1["残業時間    "]
    + df_miyagi_t_1["法外休出    "]
    + df_miyagi_t_1["法定休出    "]
)
df_miyagi_t_1.insert(33, "総労働時間", total_work_time)
basic_salary = df_miyagi_t_1["基 本 給    "] + df_miyagi_t_1["支給額"]
df_miyagi_t_1.insert(35, "基本給", basic_salary)
post_allowance = df_miyagi_t_1["役職手当    "]
df_miyagi_t_1.insert(36, "役職手当", post_allowance)
sales_allowance = df_miyagi_t_1["営業手当    "]
df_miyagi_t_1.insert(37, "営業手当", sales_allowance)
aria_allowance = df_miyagi_t_1["地域手当    "]
df_miyagi_t_1.insert(38, "地域手当", aria_allowance)
spe_allowance = df_miyagi_t_1["特殊手当    "]
df_miyagi_t_1.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_miyagi_t_1["特別技技手当"]
df_miyagi_t_1.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_miyagi_t_1["調整手当    "]
df_miyagi_t_1.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_miyagi_t_1["別居手当    "]
df_miyagi_t_1.insert(42, "別居手当", sep_allowance)
com_allowance = df_miyagi_t_1["通勤手当    "]
df_miyagi_t_1.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_miyagi_t_1.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_miyagi_t_1["残業手当    "]
df_miyagi_t_1.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_miyagi_t_1["休出手当    "]
df_miyagi_t_1.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_miyagi_t_1["深夜勤務手当"]
df_miyagi_t_1.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_miyagi_t_1["交替時差手当"]
df_miyagi_t_1.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_miyagi_t_1["休業手当    "]
df_miyagi_t_1.insert(49, "休業手当", closed_allowance)
closed_deduction = df_miyagi_t_1["休業控除    "]
df_miyagi_t_1.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_miyagi_t_1["代 休 他    "]
df_miyagi_t_1.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_miyagi_t_1["欠勤控除    "] + df_miyagi_t_1["遅早控除    "]
df_miyagi_t_1.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_miyagi_t_1["精 算 分    "]
df_miyagi_t_1.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_miyagi_t_1.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_miyagi_t_1.insert(55, "総支給額", total)
df_miyagi_t_1.insert(56, "応援時間額", 0)
df_miyagi_t_1.insert(57, "役員振替", 0)
df_miyagi_t_1.insert(58, "部門振替", 0)
df_miyagi_t_1.insert(59, "合計", 0)
# 不要フィールドの削除
df_miyagi_t_1 = df_miyagi_t_1.drop("所属2", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("所属3", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("社員ｺｰﾄﾞ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("区分", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("出勤日数    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("有休日数    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("欠勤日数    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("残業時間    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("深夜残業    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("法外休出    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("法定休出    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("代休時間    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("深夜代休    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("６０Ｈ超    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("別居手当    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("通勤手当    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("特別技技手当", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("特殊手当    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("地域手当    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("営業手当    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("役職手当    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("調整手当    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("基 本 給    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("残業手当    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("休出手当    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("深夜勤務手当", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("交替時差手当", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("休業手当    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("代 休 他    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("欠勤控除    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("遅早控除    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("精 算 分    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("支給合計額  ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("休業控除    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("支給額", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("ズレ時間    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("雑費・食事代", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("雑費・衣靴代", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("雑費        ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("受診料・他  ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("雑費・会費等", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("勤務時間    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("遅早時間    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("特休日数    ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.drop("集計区分－２        ", axis=1)
df_miyagi_t_1 = df_miyagi_t_1.sum()
df_miyagi_t_1.to_csv(
    "/content/drive/MyDrive/data/宮城/E.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 住設_間接2
df_jyusetu_k_2 = df_jyusetu_m.groupby("区分").get_group("間接2")
df_jyusetu_k_2 = df_jyusetu_k_2.drop("所属1", axis=1)
member = df_jyusetu_k_2["所属2"] > 0
df_jyusetu_k_2.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_jyusetu_k_2["所属3"] > 0
df_jyusetu_k_2.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_jyusetu_k_2["出勤日数    "]
df_jyusetu_k_2.insert(2, "実在籍者", real_member)
time_yukyu = df_jyusetu_k_2["有休日数    "] * 8
df_jyusetu_k_2.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_jyusetu_k_2.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_jyusetu_k_2["欠勤日数    "] * 8
df_jyusetu_k_2.insert(5, "欠勤時間", abs_time)
work_time = df_jyusetu_k_2["勤務時間    "]
df_jyusetu_k_2.insert(6, "勤務時間", work_time)
late_early_time = df_jyusetu_k_2["遅早時間    "]
df_jyusetu_k_2.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_jyusetu_k_2["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_jyusetu_k_2.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_jyusetu_k_2.insert(17, "実労働時間", real_work_time)
zure_time = df_jyusetu_k_2["ズレ時間    "]
df_jyusetu_k_2.insert(18, "ズレ時間", zure_time)
overtime = df_jyusetu_k_2["残業時間    "] + df_jyusetu_k_2["深夜残業    "]
df_jyusetu_k_2.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_jyusetu_k_2.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_jyusetu_k_2["法外休出    "]
df_jyusetu_k_2.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_jyusetu_k_2.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_jyusetu_k_2["法定休出    "]
df_jyusetu_k_2.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_jyusetu_k_2.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_jyusetu_k_2["６０Ｈ超    "]
df_jyusetu_k_2.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_jyusetu_k_2["代休時間    "] + df_jyusetu_k_2["深夜代休    "]
df_jyusetu_k_2.insert(31, "代休時間", holiday_time)
df_jyusetu_k_2.insert(32, "応援時間", 0)
total_work_time = (
    df_jyusetu_k_2["勤務時間    "]
    + df_jyusetu_k_2["残業時間    "]
    + df_jyusetu_k_2["法外休出    "]
    + df_jyusetu_k_2["法定休出    "]
)
df_jyusetu_k_2.insert(33, "総労働時間", total_work_time)
basic_salary = df_jyusetu_k_2["基 本 給    "] + df_jyusetu_k_2["支給額"]
df_jyusetu_k_2.insert(35, "基本給", basic_salary)
post_allowance = df_jyusetu_k_2["役職手当    "]
df_jyusetu_k_2.insert(36, "役職手当", post_allowance)
sales_allowance = df_jyusetu_k_2["営業手当    "]
df_jyusetu_k_2.insert(37, "営業手当", sales_allowance)
aria_allowance = df_jyusetu_k_2["地域手当    "]
df_jyusetu_k_2.insert(38, "地域手当", aria_allowance)
spe_allowance = df_jyusetu_k_2["特殊手当    "]
df_jyusetu_k_2.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_jyusetu_k_2["特別技技手当"]
df_jyusetu_k_2.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_jyusetu_k_2["調整手当    "]
df_jyusetu_k_2.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_jyusetu_k_2["別居手当    "]
df_jyusetu_k_2.insert(42, "別居手当", sep_allowance)
com_allowance = df_jyusetu_k_2["通勤手当    "]
df_jyusetu_k_2.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_jyusetu_k_2.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_jyusetu_k_2["残業手当    "]
df_jyusetu_k_2.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_jyusetu_k_2["休出手当    "]
df_jyusetu_k_2.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_jyusetu_k_2["深夜勤務手当"]
df_jyusetu_k_2.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_jyusetu_k_2["交替時差手当"]
df_jyusetu_k_2.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_jyusetu_k_2["休業手当    "]
df_jyusetu_k_2.insert(49, "休業手当", closed_allowance)
closed_deduction = df_jyusetu_k_2["休業控除    "]
df_jyusetu_k_2.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_jyusetu_k_2["代 休 他    "]
df_jyusetu_k_2.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_jyusetu_k_2["欠勤控除    "] + df_jyusetu_k_2["遅早控除    "]
df_jyusetu_k_2.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_jyusetu_k_2["精 算 分    "]
df_jyusetu_k_2.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_jyusetu_k_2.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_jyusetu_k_2.insert(55, "総支給額", total)
df_jyusetu_k_2.insert(56, "応援時間額", 0)
df_jyusetu_k_2.insert(57, "役員振替", 0)
df_jyusetu_k_2.insert(58, "部門振替", 0)
df_jyusetu_k_2.insert(59, "合計", 0)
# 不要フィールドの削除
df_jyusetu_k_2 = df_jyusetu_k_2.drop("所属2", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("所属3", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("社員ｺｰﾄﾞ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("区分", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("出勤日数    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("有休日数    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("欠勤日数    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("残業時間    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("深夜残業    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("法外休出    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("法定休出    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("代休時間    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("深夜代休    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("６０Ｈ超    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("別居手当    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("通勤手当    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("特別技技手当", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("特殊手当    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("地域手当    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("営業手当    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("役職手当    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("調整手当    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("基 本 給    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("残業手当    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("休出手当    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("深夜勤務手当", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("交替時差手当", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("休業手当    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("代 休 他    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("欠勤控除    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("遅早控除    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("精 算 分    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("支給合計額  ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("休業控除    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("支給額", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("ズレ時間    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("雑費・食事代", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("雑費・衣靴代", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("雑費        ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("受診料・他  ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("雑費・会費等", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("勤務時間    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("遅早時間    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("特休日数    ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.drop("集計区分－２        ", axis=1)
df_jyusetu_k_2 = df_jyusetu_k_2.sum()
df_jyusetu_k_2.to_csv(
    "/content/drive/MyDrive/data/住設/A.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 住設_間接4
df_jyusetu_k_4 = df_jyusetu_m.groupby("区分").get_group("間接4")
df_jyusetu_k_4 = df_jyusetu_k_4.drop("所属1", axis=1)
member = df_jyusetu_k_4["所属2"] > 0
df_jyusetu_k_4.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_jyusetu_k_4["所属3"] > 0
df_jyusetu_k_4.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_jyusetu_k_4["出勤日数    "]
df_jyusetu_k_4.insert(2, "実在籍者", real_member)
time_yukyu = df_jyusetu_k_4["有休日数    "] * 8
df_jyusetu_k_4.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_jyusetu_k_4.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_jyusetu_k_4["欠勤日数    "] * 8
df_jyusetu_k_4.insert(5, "欠勤時間", abs_time)
work_time = df_jyusetu_k_4["勤務時間    "]
df_jyusetu_k_4.insert(6, "勤務時間", work_time)
late_early_time = df_jyusetu_k_4["遅早時間    "]
df_jyusetu_k_4.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_jyusetu_k_4["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_jyusetu_k_4.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_jyusetu_k_4.insert(17, "実労働時間", real_work_time)
zure_time = df_jyusetu_k_4["ズレ時間    "]
df_jyusetu_k_4.insert(18, "ズレ時間", zure_time)
overtime = df_jyusetu_k_4["残業時間    "] + df_jyusetu_k_4["深夜残業    "]
df_jyusetu_k_4.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_jyusetu_k_4.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_jyusetu_k_4["法外休出    "]
df_jyusetu_k_4.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_jyusetu_k_4.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_jyusetu_k_4["法定休出    "]
df_jyusetu_k_4.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_jyusetu_k_4.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_jyusetu_k_4["６０Ｈ超    "]
df_jyusetu_k_4.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_jyusetu_k_4["代休時間    "] + df_jyusetu_k_4["深夜代休    "]
df_jyusetu_k_4.insert(31, "代休時間", holiday_time)
df_jyusetu_k_4.insert(32, "応援時間", 0)
total_work_time = (
    df_jyusetu_k_4["勤務時間    "]
    + df_jyusetu_k_4["残業時間    "]
    + df_jyusetu_k_4["法外休出    "]
    + df_jyusetu_k_4["法定休出    "]
)
df_jyusetu_k_4.insert(33, "総労働時間", total_work_time)
basic_salary = df_jyusetu_k_4["基 本 給    "] + df_jyusetu_k_4["支給額"]
df_jyusetu_k_4.insert(35, "基本給", basic_salary)
post_allowance = df_jyusetu_k_4["役職手当    "]
df_jyusetu_k_4.insert(36, "役職手当", post_allowance)
sales_allowance = df_jyusetu_k_4["営業手当    "]
df_jyusetu_k_4.insert(37, "営業手当", sales_allowance)
aria_allowance = df_jyusetu_k_4["地域手当    "]
df_jyusetu_k_4.insert(38, "地域手当", aria_allowance)
spe_allowance = df_jyusetu_k_4["特殊手当    "]
df_jyusetu_k_4.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_jyusetu_k_4["特別技技手当"]
df_jyusetu_k_4.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_jyusetu_k_4["調整手当    "]
df_jyusetu_k_4.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_jyusetu_k_4["別居手当    "]
df_jyusetu_k_4.insert(42, "別居手当", sep_allowance)
com_allowance = df_jyusetu_k_4["通勤手当    "]
df_jyusetu_k_4.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_jyusetu_k_4.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_jyusetu_k_4["残業手当    "]
df_jyusetu_k_4.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_jyusetu_k_4["休出手当    "]
df_jyusetu_k_4.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_jyusetu_k_4["深夜勤務手当"]
df_jyusetu_k_4.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_jyusetu_k_4["交替時差手当"]
df_jyusetu_k_4.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_jyusetu_k_4["休業手当    "]
df_jyusetu_k_4.insert(49, "休業手当", closed_allowance)
closed_deduction = df_jyusetu_k_4["休業控除    "]
df_jyusetu_k_4.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_jyusetu_k_4["代 休 他    "]
df_jyusetu_k_4.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_jyusetu_k_4["欠勤控除    "] + df_jyusetu_k_4["遅早控除    "]
df_jyusetu_k_4.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_jyusetu_k_4["精 算 分    "]
df_jyusetu_k_4.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_jyusetu_k_4.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_jyusetu_k_4.insert(55, "総支給額", total)
df_jyusetu_k_4.insert(56, "応援時間額", 0)
df_jyusetu_k_4.insert(57, "役員振替", 0)
df_jyusetu_k_4.insert(58, "部門振替", 0)
df_jyusetu_k_4.insert(59, "合計", 0)
# 不要フィールドの削除
df_jyusetu_k_4 = df_jyusetu_k_4.drop("所属2", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("所属3", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("社員ｺｰﾄﾞ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("区分", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("出勤日数    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("有休日数    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("欠勤日数    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("残業時間    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("深夜残業    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("法外休出    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("法定休出    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("代休時間    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("深夜代休    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("６０Ｈ超    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("別居手当    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("通勤手当    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("特別技技手当", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("特殊手当    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("地域手当    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("営業手当    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("役職手当    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("調整手当    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("基 本 給    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("残業手当    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("休出手当    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("深夜勤務手当", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("交替時差手当", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("休業手当    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("代 休 他    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("欠勤控除    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("遅早控除    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("精 算 分    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("支給合計額  ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("休業控除    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("支給額", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("ズレ時間    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("雑費・食事代", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("雑費・衣靴代", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("雑費        ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("受診料・他  ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("雑費・会費等", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("勤務時間    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("遅早時間    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("特休日数    ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.drop("集計区分－２        ", axis=1)
df_jyusetu_k_4 = df_jyusetu_k_4.sum()
df_jyusetu_k_4.to_csv(
    "/content/drive/MyDrive/data/住設/B.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
# 住設_間接6
df_jyusetu_k_6 = df_jyusetu_m.groupby("区分").get_group("間接6")
df_jyusetu_k_6 = df_jyusetu_k_6.drop("所属1", axis=1)
member = df_jyusetu_k_6["所属2"] > 0
df_jyusetu_k_6.insert(0, "在籍者", member)
member = len(member)
under_mastar = df_jyusetu_k_6["所属3"] > 0
df_jyusetu_k_6.insert(1, "在籍者主幹以下人数", under_mastar)
under_mastar = (under_mastar == True).sum()
real_member = df_jyusetu_k_6["出勤日数    "]
df_jyusetu_k_6.insert(2, "実在籍者", real_member)
time_yukyu = df_jyusetu_k_6["有休日数    "] * 8
df_jyusetu_k_6.insert(3, "有休時間", time_yukyu)
ave_yukyu = time_yukyu / member
ave_yukyu = ave_yukyu.round(2)
df_jyusetu_k_6.insert(4, "有休時間在籍者平均", ave_yukyu)
abs_time = df_jyusetu_k_6["欠勤日数    "] * 8
df_jyusetu_k_6.insert(5, "欠勤時間", abs_time)
work_time = df_jyusetu_k_6["勤務時間    "]
df_jyusetu_k_6.insert(6, "勤務時間", work_time)
late_early_time = df_jyusetu_k_6["遅早時間    "]
df_jyusetu_k_6.insert(7, "遅早時間", late_early_time)
work_rate1 = (real_member * 8) - df_jyusetu_k_6["遅早時間    "]
work_rate1 = work_rate1.sum()
work_rate2 = (real_member * 8) * 100
work_rate2 = work_rate2.sum()
work_rate = (work_rate1 / work_rate2) * 100
work_rate = work_rate / member
df_jyusetu_k_6.insert(16, "出勤率", work_rate)
real_work_time = real_member * 8
df_jyusetu_k_6.insert(17, "実労働時間", real_work_time)
zure_time = df_jyusetu_k_6["ズレ時間    "]
df_jyusetu_k_6.insert(18, "ズレ時間", zure_time)
overtime = df_jyusetu_k_6["残業時間    "] + df_jyusetu_k_6["深夜残業    "]
df_jyusetu_k_6.insert(24, "残業時間", overtime.round(2))
under_mater_overtime = overtime / under_mastar
df_jyusetu_k_6.insert(25, "残業時間主幹以下平均", under_mater_overtime.round(2))
non_leave_time = df_jyusetu_k_6["法外休出    "]
df_jyusetu_k_6.insert(26, "法定外休出時間", non_leave_time)
ave_non_leave_time = non_leave_time / under_mastar
df_jyusetu_k_6.insert(27, "法定外主幹以下平均", ave_non_leave_time.round(2))
legal_leave_time = df_jyusetu_k_6["法定休出    "]
df_jyusetu_k_6.insert(28, "法定休出時間", legal_leave_time)
ave_legal_leave_time = legal_leave_time / under_mastar
df_jyusetu_k_6.insert(29, "法定主幹以下平均", ave_legal_leave_time.round(2))
overtime_60 = df_jyusetu_k_6["６０Ｈ超    "]
df_jyusetu_k_6.insert(30, "時間外60時間超え", overtime_60)
holiday_time = df_jyusetu_k_6["代休時間    "] + df_jyusetu_k_6["深夜代休    "]
df_jyusetu_k_6.insert(31, "代休時間", holiday_time)
df_jyusetu_k_6.insert(32, "応援時間", 0)
total_work_time = (
    df_jyusetu_k_6["勤務時間    "]
    + df_jyusetu_k_6["残業時間    "]
    + df_jyusetu_k_6["法外休出    "]
    + df_jyusetu_k_6["法定休出    "]
)
df_jyusetu_k_6.insert(33, "総労働時間", total_work_time)
basic_salary = df_jyusetu_k_6["基 本 給    "] + df_jyusetu_k_6["支給額"]
df_jyusetu_k_6.insert(35, "基本給", basic_salary)
post_allowance = df_jyusetu_k_6["役職手当    "]
df_jyusetu_k_6.insert(36, "役職手当", post_allowance)
sales_allowance = df_jyusetu_k_6["営業手当    "]
df_jyusetu_k_6.insert(37, "営業手当", sales_allowance)
aria_allowance = df_jyusetu_k_6["地域手当    "]
df_jyusetu_k_6.insert(38, "地域手当", aria_allowance)
spe_allowance = df_jyusetu_k_6["特殊手当    "]
df_jyusetu_k_6.insert(39, "特別手当", spe_allowance)
spe_tec_allowance = df_jyusetu_k_6["特別技技手当"]
df_jyusetu_k_6.insert(40, "特別技技手当 ", spe_tec_allowance)
adjust_allowance = df_jyusetu_k_6["調整手当    "]
df_jyusetu_k_6.insert(41, "調整手当", adjust_allowance)
sep_allowance = df_jyusetu_k_6["別居手当    "]
df_jyusetu_k_6.insert(42, "別居手当", sep_allowance)
com_allowance = df_jyusetu_k_6["通勤手当    "]
df_jyusetu_k_6.insert(43, "通勤手当", com_allowance)
sub_total_1 = (
    basic_salary
    + post_allowance
    + sales_allowance
    + aria_allowance
    + spe_allowance
    + adjust_allowance
    + sep_allowance
    + com_allowance
)
df_jyusetu_k_6.insert(44, "小計 1", sub_total_1)
overtime_allowance = df_jyusetu_k_6["残業手当    "]
df_jyusetu_k_6.insert(45, "残業手当", overtime_allowance)
vacation_allowance = df_jyusetu_k_6["休出手当    "]
df_jyusetu_k_6.insert(46, "休出手当", vacation_allowance)
night_work_allowance = df_jyusetu_k_6["深夜勤務手当"]
df_jyusetu_k_6.insert(47, "深夜勤務手当　", night_work_allowance)
time_difference_allowance = df_jyusetu_k_6["交替時差手当"]
df_jyusetu_k_6.insert(48, "交替時差手当　", time_difference_allowance)
closed_allowance = df_jyusetu_k_6["休業手当    "]
df_jyusetu_k_6.insert(49, "休業手当", closed_allowance)
closed_deduction = df_jyusetu_k_6["休業控除    "]
df_jyusetu_k_6.insert(50, "休業控除", closed_deduction)
compny_leave_etc = df_jyusetu_k_6["代 休 他    "]
df_jyusetu_k_6.insert(51, "代休他", compny_leave_etc)
abs_early_deduction = df_jyusetu_k_6["欠勤控除    "] + df_jyusetu_k_6["遅早控除    "]
df_jyusetu_k_6.insert(52, "欠勤・遅早控除", abs_early_deduction)
settlement = df_jyusetu_k_6["精 算 分    "]
df_jyusetu_k_6.insert(53, "精算分", settlement)
sub_total_2 = abs_early_deduction + settlement
df_jyusetu_k_6.insert(54, "小計 2", sub_total_2)
total = sub_total_1 - sub_total_2
df_jyusetu_k_6.insert(55, "総支給額", total)
df_jyusetu_k_6.insert(56, "応援時間額", 0)
df_jyusetu_k_6.insert(57, "役員振替", 0)
df_jyusetu_k_6.insert(58, "部門振替", 0)
df_jyusetu_k_6.insert(59, "合計", 0)
# 不要フィールドの削除
df_jyusetu_k_6 = df_jyusetu_k_6.drop("所属2", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("所属3", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("社員ｺｰﾄﾞ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("区分", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("出勤日数    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("有休日数    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("欠勤日数    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("残業時間    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("深夜残業    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("法外休出    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("法定休出    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("代休時間    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("深夜代休    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("６０Ｈ超    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("別居手当    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("通勤手当    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("特別技技手当", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("特殊手当    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("地域手当    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("営業手当    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("役職手当    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("調整手当    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("基 本 給    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("残業手当    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("休出手当    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("深夜勤務手当", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("交替時差手当", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("休業手当    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("代 休 他    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("欠勤控除    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("遅早控除    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("精 算 分    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("支給合計額  ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("休業控除    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("支給額", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("ズレ時間    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("雑費・食事代", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("雑費・衣靴代", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("雑費        ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("受診料・他  ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("雑費・会費等", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("勤務時間    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("遅早時間    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("特休日数    ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.drop("集計区分－２        ", axis=1)
df_jyusetu_k_6 = df_jyusetu_k_6.sum()
df_jyusetu_k_6.to_csv(
    "/content/drive/MyDrive/data/住設/C.csv",
    header=True,
    index=False,
    encoding="shift-jis",
)
