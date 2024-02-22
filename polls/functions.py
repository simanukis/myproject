import pandas as pd
import numpy as np
import sklearn,csv,re,os
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
from django.http import HttpResponse
# from myproject.settings import BASE_DIR

def process_file(data):
    jinji = pd.read_excel(os.path.join('app/static/files/jinji.xls'))
    syain = pd.read_excel(os.path.join('app/static/files/syain.xls'))
    kinsi = pd.read_excel(os.path.join('app/static/files/kinsi.xls'))
    data_kbn = data
    
    # データ処理
    jinji = jinji.drop(columns=['社員名', '集計区分－１        '])
    syain = syain.drop(columns=['氏 名','役職ｺｰﾄﾞ', '役職','所属分類1名','所属分類2名', '所属分類3名', '所 .3','所属分類4名', '所 .4', '所属分類5名','備考', '在籍区分', '在籍', '退職年', '退職月', '退職日'])
    
    # 人事ファイルと社員ファイルの結合してデータフレーム化
    df_m = pd.merge(jinji, syain, how='inner',on='社員ｺｰﾄﾞ')
    
    # 役員データフレームの作成
    df_yakuin_kbn = df_m[(df_m['所 '] == 0) & (df_m['所 .1'] == 0) & (df_m['所 .2'] == 0)]
    df_yakuin_kbn = df_yakuin_kbn.astype('object')
    df_yakuin_kbn['区分'] = '役員1'
    df_yakuin_kbn = df_yakuin_kbn.rename(columns={'所 ': '所属1', '所 .1': '所属2', '所 .2': '所属3'})
    
    # 一般間接1データフレームの作成
    df_ipan_kbn103 = df_m[(df_m['集計区分－２        '] == 103)]
    df_ipan_kbn103 = df_ipan_kbn103.astype('object')
    df_ipan_kbn103['区分'] = '一般間接1'
    
    # 一般間接2データフレームの作成
    df_ipan_kbn104 = df_m[(df_m['集計区分－２        '] == 104)]
    df_ipan_kbn104 = df_ipan_kbn104.astype('object')
    df_ipan_kbn104['区分'] = '一般間接2'
    
    # 一般間接3データフレームの作成
    df_ipan_kbn105 = df_m[(df_m['集計区分－２        '] == 105)]
    df_ipan_kbn105 = df_ipan_kbn105.astype('object')
    df_ipan_kbn105['区分'] = '一般間接3'
    
    # 一般間接6データフレームの作成
    df_ipan_kbn108 = df_m[(df_m['集計区分－２        '] == 108)]
    df_ipan_kbn108 = df_ipan_kbn108.astype('object')
    df_ipan_kbn108['区分'] = '一般間接6'
    
    # 一般販売1データフレームの作成
    df_ipan_kbn109 = df_m[(df_m['集計区分－２        '] == 109)]
    df_ipan_kbn109 = df_ipan_kbn109.astype('object')
    df_ipan_kbn109['区分'] = '一般販売1'
    
    # 一般販売2データフレームの作成
    df_ipan_kbn110 = df_m[(df_m['集計区分－２        '] == 110)]
    df_ipan_kbn110 = df_ipan_kbn110.astype('object')
    df_ipan_kbn110['区分'] = '一般販売2'
    
    # 一般管理データフレームの作成
    df_ipan_kbn = pd.concat(
        [df_ipan_kbn103, df_ipan_kbn104, df_ipan_kbn105, df_ipan_kbn108, df_ipan_kbn109, df_ipan_kbn110],
        axis=0,
        ignore_index=True
    )
    df_ipan_kbn = df_ipan_kbn.rename(columns={'所 ': '所属1', '所 .1': '所属2', '所 .2': '所属3'})
    
    # 鍛造間接1のデータフレーム作成
    df_tanzo_kbn211 = df_m[(df_m['集計区分－２        '] == 211)]
    df_tanzo_kbn211 = df_tanzo_kbn211.astype('object')
    df_tanzo_kbn211['区分'] = '間接1'
    
    # 鍛造間接2のデータフレーム作成
    df_tanzo_kbn212 = df_m[(df_m['集計区分－２        '] == 212)]
    df_tanzo_kbn212 = df_tanzo_kbn212.astype('object')
    df_tanzo_kbn212['区分'] = '間接2'
    
    # 鍛造間接3のデータフレーム作成
    df_tanzo_kbn213 = df_m[(df_m['集計区分－２        '] == 213)]
    df_tanzo_kbn213 = df_tanzo_kbn213.astype('object')
    df_tanzo_kbn213['区分'] = '間接3'
    
    # 鍛造間接4のデータフレーム作成
    df_tanzo_kbn214 = df_m[(df_m['集計区分－２        '] == 214)]
    df_tanzo_kbn214 = df_tanzo_kbn214.astype('object')
    df_tanzo_kbn214['区分'] = '間接5'
    
    # 鍛造間接5のデータフレーム作成
    df_tanzo_kbn215 = df_m[(df_m['集計区分－２        '] == 215)]
    df_tanzo_kbn215 = df_tanzo_kbn215.astype('object')
    df_tanzo_kbn215['区分'] = '間接5'
    
    # 鍛造間接6のデータフレーム作成
    df_tanzo_kbn216 = df_m[(df_m['集計区分－２        '] == 216)]
    df_tanzo_kbn216 = df_tanzo_kbn216.astype('object')
    df_tanzo_kbn216['区分'] = '間接6'
    
    # 鍛造直接1のデータフレーム作成
    df_tanzo_kbn218 = df_m[(df_m['集計区分－２        '] == 218)]
    df_tanzo_kbn218 = df_tanzo_kbn218.astype('object')
    df_tanzo_kbn218['区分'] = '直接1'
    
    # 鍛造直接4のデータフレーム作成
    df_tanzo_kbn221 = df_m[(df_m['集計区分－２        '] == 221)]
    df_tanzo_kbn221 = df_tanzo_kbn221.astype('object')
    df_tanzo_kbn221['区分'] = '直接4'
    
    # 鍛造のデータフレーム作成
    df_tanzo_kbn = pd.concat(
        [df_tanzo_kbn211, df_tanzo_kbn212, df_tanzo_kbn213, df_tanzo_kbn214, df_tanzo_kbn215, df_tanzo_kbn216, df_tanzo_kbn218, df_tanzo_kbn221],
        axis=0,
        ignore_index=True
    )
    df_tanzo_kbn = df_tanzo_kbn.rename(columns={'所 ': '所属1', '所 .1': '所属2', '所 .2': '所属3'})
    
    # 切削間接1のデータフレーム作成
    df_sesaku_kbn311 = df_m[(df_m['集計区分－２        '] == 311)]
    df_sesaku_kbn311 = df_sesaku_kbn311.astype('object')
    df_sesaku_kbn311['区分'] = '間接1'
    
    # 切削間接2のデータフレーム作成
    df_sesaku_kbn312 = df_m[(df_m['集計区分－２        '] == 312)]
    df_sesaku_kbn312 = df_sesaku_kbn312.astype('object')
    df_sesaku_kbn312['区分'] = '間接2'
    
    # 切削間接4のデータフレーム作成
    df_sesaku_kbn314 = df_m[(df_m['集計区分－２        '] == 314)]
    df_sesaku_kbn314 = df_sesaku_kbn314.astype('object')
    df_sesaku_kbn314['区分'] = '間接4'
    
    # 切削間接5のデータフレーム作成
    df_sesaku_kbn315 = df_m[(df_m['集計区分－２        '] == 315)]
    df_sesaku_kbn315 = df_sesaku_kbn315.astype('object')
    df_sesaku_kbn315['区分'] = '間接5'
    
    # 切削間接6のデータフレーム作成
    df_sesaku_kbn316 = df_m[(df_m['集計区分－２        '] == 316)]
    df_sesaku_kbn316 = df_sesaku_kbn316.astype('object')
    df_sesaku_kbn316['区分'] = '間接6'
    
    # 切削直接1のデータフレーム作成
    df_sesaku_kbn318 = df_m[(df_m['集計区分－２        '] == 318)]
    df_sesaku_kbn318 = df_sesaku_kbn318.astype('object')
    df_sesaku_kbn318['区分'] = '直接1'
    
    # 切削直接2のデータフレーム作成
    df_sesaku_kbn319 = df_m[(df_m['集計区分－２        '] == 319)]
    df_sesaku_kbn319 = df_sesaku_kbn319.astype('object')
    df_sesaku_kbn319['区分'] = '直接2'
    
    # 切削直接4のデータフレーム作成
    df_sesaku_kbn321 = df_m[(df_m['集計区分－２        '] == 321)]
    df_sesaku_kbn321 = df_sesaku_kbn321.astype('object')
    df_sesaku_kbn321['区分'] = '直接4'
    
    # 切削のデータフレーム作成
    df_sesaku_kbn = pd.concat(
        [df_sesaku_kbn311, df_sesaku_kbn312, df_sesaku_kbn314, df_sesaku_kbn315, df_sesaku_kbn316, df_sesaku_kbn318, df_sesaku_kbn319, df_sesaku_kbn321],
        axis=0,
        ignore_index=True
    )
    df_sesaku_kbn = df_sesaku_kbn.rename(columns={'所 ': '所属1', '所 .1': '所属2', '所 .2': '所属3'})
    
    # AC間接1のデータフレーム作成
    df_ac_kbn411 = df_m[(df_m['集計区分－２        '] == 411)]
    df_ac_kbn411 = df_ac_kbn411.astype('object')
    df_ac_kbn411['区分'] = '間接1'
    
    # AC間接2のデータフレーム作成
    df_ac_kbn412 = df_m[(df_m['集計区分－２        '] == 412)]
    df_ac_kbn412 = df_ac_kbn412.astype('object')
    df_ac_kbn412['区分'] = '間接2'
    
    # AC間接4のデータフレーム作成
    df_ac_kbn414 = df_m[(df_m['集計区分－２        '] == 414)]
    df_ac_kbn414 = df_ac_kbn414.astype('object')
    df_ac_kbn414['区分'] = '間接4'
    
    # AC間接5のデータフレーム作成
    df_ac_kbn415 = df_m[(df_m['集計区分－２        '] == 415)]
    df_ac_kbn415 = df_ac_kbn415.astype('object')
    df_ac_kbn415['区分'] = '間接5'
    
    # AC直接1のデータフレーム作成
    df_ac_kbn418 = df_m[(df_m['集計区分－２        '] == 418)]
    df_ac_kbn418 = df_ac_kbn418.astype('object')
    df_ac_kbn418['区分'] = '直接1'
    
    # AC直接4のデータフレーム作成
    df_ac_kbn421 = df_m[(df_m['集計区分－２        '] == 421)]
    df_ac_kbn421 = df_ac_kbn421.astype('object')
    df_ac_kbn421['区分'] = '直接4'
    
    # ACのデータフレーム作成
    df_ac_kbn = pd.concat(
        [df_ac_kbn411, df_ac_kbn412, df_ac_kbn414, df_ac_kbn415, df_ac_kbn418, df_ac_kbn421],
        axis=0,
        ignore_index=True
    )
    df_ac_kbn = df_ac_kbn.rename(columns={'所 ': '所属1', '所 .1': '所属2', '所 .2': '所属3'})
    
    # PC間接1のデータフレーム作成
    df_pc_kbn511 = df_m[(df_m['集計区分－２        '] == 511)]
    df_pc_kbn511 = df_pc_kbn511.astype('object')
    df_pc_kbn511['区分'] = '間接1'
    
    # PC間接2のデータフレーム作成
    df_pc_kbn512 = df_m[(df_m['集計区分－２        '] == 512)]
    df_pc_kbn512 = df_pc_kbn512.astype('object')
    df_pc_kbn512['区分'] = '間接2'
    
    # PC間接4のデータフレーム作成
    df_pc_kbn514 = df_m[(df_m['集計区分－２        '] == 514)]
    df_pc_kbn514 = df_pc_kbn514.astype('object')
    df_pc_kbn514['区分'] = '間接4'
    
    # PC間接5のデータフレーム作成
    df_pc_kbn515 = df_m[(df_m['集計区分－２        '] == 515)]
    df_pc_kbn515 = df_pc_kbn515.astype('object')
    df_pc_kbn515['区分'] = '間接5'
    
    # PC間接6のデータフレーム作成
    df_pc_kbn516 = df_m[(df_m['集計区分－２        '] == 516)]
    df_pc_kbn516 = df_pc_kbn516.astype('object')
    df_pc_kbn516['区分'] = '間接6'
    
    # PC直接1のデータフレーム作成
    df_pc_kbn518 = df_m[(df_m['集計区分－２        '] == 518)]
    df_pc_kbn518 = df_pc_kbn518.astype('object')
    df_pc_kbn518['区分'] = '直接1'
    
    # PC直接4のデータフレーム作成
    df_pc_kbn521 = df_m[(df_m['集計区分－２        '] == 521)]
    df_pc_kbn521 = df_pc_kbn521.astype('object')
    df_pc_kbn521['区分'] = '直接4'
    
    # PCのデータフレーム作成
    df_pc_kbn = pd.concat(
        [df_pc_kbn511, df_pc_kbn512, df_pc_kbn514, df_pc_kbn515, df_pc_kbn516, df_pc_kbn518, df_pc_kbn521],
        axis=0,
        ignore_index=True
    )
    df_pc_kbn = df_pc_kbn.rename(columns={'所 ': '所属1', '所 .1': '所属2', '所 .2': '所属3'})
    
    # 宮城間接1のデータフレーム作成
    df_miyagi_kbn611 = df_m[(df_m['集計区分－２        '] == 611)]
    df_miyagi_kbn611 = df_miyagi_kbn611.astype('object')
    df_miyagi_kbn611['区分'] = '間接1'
    
    # 宮城間接2のデータフレーム作成
    df_miyagi_kbn612 = df_m[(df_m['集計区分－２        '] == 612)]
    df_miyagi_kbn612 = df_miyagi_kbn612.astype('object')
    df_miyagi_kbn612['区分'] = '間接2'
    
    # 宮城間接4のデータフレーム作成
    df_miyagi_kbn614 = df_m[(df_m['集計区分－２        '] == 614)]
    df_miyagi_kbn614 = df_miyagi_kbn614.astype('object')
    df_miyagi_kbn614['区分'] = '間接4'
    
    # 宮城間接6のデータフレーム作成
    df_miyagi_kbn616 = df_m[(df_m['集計区分－２        '] == 616)]
    df_miyagi_kbn616 = df_miyagi_kbn616.astype('object')
    df_miyagi_kbn616['区分'] = '間接6'
    
    # 宮城直接1のデータフレーム作成
    df_miyagi_kbn618 = df_m[(df_m['集計区分－２        '] == 618)]
    df_miyagi_kbn618 = df_miyagi_kbn618.astype('object')
    df_miyagi_kbn618['区分'] = '直接1'
    
    # 宮城のデータフレーム作成
    df_miyagi_kbn = pd.concat(
        [df_miyagi_kbn611, df_miyagi_kbn612, df_miyagi_kbn614, df_miyagi_kbn616, df_miyagi_kbn618],
        axis=0,
        ignore_index=True
    )
    df_miyagi_kbn = df_miyagi_kbn.rename(columns={'所 ': '所属1', '所 .1': '所属2', '所 .2': '所属3'})
    
    # 住設間接2のデータフレーム作成
    df_jyusetu_kbn712 = df_m[(df_m['集計区分－２        '] == 712)]
    df_jyusetu_kbn712 = df_jyusetu_kbn712.astype('object')
    df_jyusetu_kbn712['区分'] = '間接2'
    
    # 住設間接4のデータフレーム作成
    df_jyusetu_kbn714 = df_m[(df_m['集計区分－２        '] == 714)]
    df_jyusetu_kbn714 = df_jyusetu_kbn714.astype('object')
    df_jyusetu_kbn714['区分'] = '間接4'
    
    # 住設間接6のデータフレーム作成
    df_jyusetu_kbn716 = df_m[(df_m['集計区分－２        '] == 716)]
    df_jyusetu_kbn716 = df_jyusetu_kbn716.astype('object')
    df_jyusetu_kbn716['区分'] = '間接6'
    
    # 住設のデータフレーム作成
    df_jyusetu_kbn = pd.concat(
        [df_jyusetu_kbn712, df_jyusetu_kbn714, df_jyusetu_kbn716],
        axis=0,
        ignore_index=True
    )
    
    # 給与データファイルのデータフレーム化
    kinsi = kinsi.drop(columns=kinsi.columns[[16, 41, 61]], axis=1)
    kinsi = kinsi.dropna(how='any')
    kinsi.drop(labels=['所属4','所属5','社員名','支給日[年]','支給日[月]','支給日[日]','【 勤怠 】  ','深夜時間    ','休業日数    ','【 支給 】  ','課税支給額  ','【 控除 】  ','健康保険    ','厚生年金保険','雇用保険    ','社会保険計  ','(内)介護保険','課税対象額  ','所得税      ','住民税      ','財形預金    ','生命保険    ','自動車保険  ','社員積立    ','労金        ','その他控除  ','控除合計額  ','【 補助 】  ','職能等級    ','基本給１    ','基本給２    ','【 合計 】  ','差引支給額  ','銀行振込１  ','銀行振込２  ','銀行振込３  ','現金支給額  '], axis=1, inplace=True)
    kinsi_s = kinsi.set_index(['所属1','所属2','所属3'])
    kinsi_s = kinsi_s.rename(columns={'            .1': '支給額'})
    
    # 給与データフレームと各区分データフレームを結合する
    # 役員
    df_yakuin_m = pd.merge(df_yakuin_kbn, kinsi_s,
                        how='inner',on='社員ｺｰﾄﾞ')
    # 一般管理
    df_ipan_m = pd.merge(df_ipan_kbn, kinsi_s,
                        how='inner', on='社員ｺｰﾄﾞ')
    # 鍛造
    df_tanzo_m = pd.merge(df_tanzo_kbn, kinsi_s,
                        how='inner', on='社員ｺｰﾄﾞ')
    # 切削
    df_sesaku_m = pd.merge(df_sesaku_kbn, kinsi_s,
                        how='inner', on='社員ｺｰﾄﾞ')
    # AC
    df_ac_m = pd.merge(df_ac_kbn, kinsi_s,
                        how='inner', on='社員ｺｰﾄﾞ')
    # PC
    df_pc_m = pd.merge(df_pc_kbn, kinsi_s,
                        how='inner', on='社員ｺｰﾄﾞ')
    # 宮城
    df_miyagi_m = pd.merge(df_miyagi_kbn, kinsi_s,
                        how='inner', on='社員ｺｰﾄﾞ')
    # 住設
    df_jyusetu_m = pd.merge(df_jyusetu_kbn, kinsi_s,
                        how='inner', on='社員ｺｰﾄﾞ')

    
    
    
    # データフレームとして返す
    df_result = ...
    
    return df_result

def to_csv(df):
    response = HttpResponse(content_type='text/csv; charset=UTF-8')
    response['Content-Disposition'] = 'attachment; filename="result.csv"'
    
    df.to_csv(path_or_buf = response, encoding = 'utf-8-sig', index=False)
    
    return response