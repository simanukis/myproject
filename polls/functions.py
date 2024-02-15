import pandas as pd
import numpy as np
import sklearn,csv,re,os
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
from django.http import HttpResponse
# from myproject.settings import BASE_DIR

def process_file(data):
    jinji = pd.read_excel(os.path.join('app/static/files/jinji.xls'))
    syain = pd.read_csv(os.path.join('app/static/files/syain.xls'))
    train = pd.read_csv(os.path.join('app/static/files/kinsi.xls'))
    test = data
    
    # データ処理
    
    # データフレームとして返す
    df_result = ...
    
    return df_result

def to_csv(df):
    response = HttpResponse(content_type='text/csv; charset=UTF-8')
    response['Content-Disposition'] = 'attachment; filename="result.csv"'
    
    df.to_csv(path_or_buf = response, encoding = 'utf-8-sig', index=False)
    
    return response