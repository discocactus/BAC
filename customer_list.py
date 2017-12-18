
# coding: utf-8

# # 準備

# In[ ]:

import sys
import codecs
import re
import pandas as pd
import numpy as np
import jaconv


# In[ ]:

# pandas の最大表示列数を設定 (max_rows で表示行数の設定も可能)
pd.set_option('display.max_columns', 30)


# In[ ]:

# ファイルパス
path = r'C:\Users\BestecAudioTecHH\Documents\GitHub\BAC\DO_NOT_PUSH'


# # excel ファイルの読み込み、作業用テーブルの作成

# In[ ]:

customer_excel = pd.read_excel(r'{0}\Bestecaudio得意先リスト20170728更新_local_171214.xlsx'.format(path))


# In[ ]:

customer_excel


# In[ ]:

customer_excel.columns


# In[ ]:

# 作業用テーブルの作成
customer_table = customer_excel[['担当者コード', '担当者名', '得意先', '名カナ', '得意先名1', '得意先名2', '郵便番号', '住所1', '住所2',
       '住所3', '電話番号', 'FAX番号']]


# In[ ]:

customer_table


# # 内容確認

# In[ ]:

customer_table.loc[0, '名カナ']


# In[ ]:

customer_table.loc[0, '得意先名2']


# In[ ]:

customer_table.loc[0, '住所1']


# In[ ]:

customer_table.loc[0, '得意先']


# In[ ]:

customer_table['名カナ'].values


# In[ ]:

customer_table['得意先名2'].values


# # 整形処理

# In[ ]:

# str 型を明示
customer_table = customer_table.astype({'担当者コード':str, '担当者名': str, '得意先': str, '名カナ': str, '得意先名1': str, '得意先名2': str, '郵便番号': str,
                                        '住所1': str, '住所2': str, '住所3': str, '電話番号': str, 'FAX番号': str})

# 型の明示 (作業未到達)
customer_table = customer_table.astype({'担当者コード':int, '担当者名': str, '得意先': int, '名カナ': str, '得意先名1': str, '得意先名2': str, '郵便番号': str,
                                        '住所1': str, '住所2': str, '住所3': str, '電話番号': str, 'FAX番号': str})
# In[ ]:

customer_table.dtypes


# In[ ]:

type(customer_table.loc[1, '担当者コード'])


# In[ ]:

# 'x1f' (ユニット区切り) を除去
customer_table = customer_table.applymap(lambda x: re.sub('\x1f', '', x))


# In[ ]:

customer_table.applymap(lambda x: re.match('\x1f', x)).any()


# In[ ]:

# 半角カナを全角カナに
customer_table = customer_table.applymap(lambda x: jaconv.h2z(x, kana=True, ascii=False, digit=False))


# In[ ]:

# 全角英数を半角英数に
customer_table = customer_table.applymap(lambda x: jaconv.z2h(x, kana=False, ascii=True, digit=True))


# In[ ]:

# 'nan' を除去
customer_table = customer_table.applymap(lambda x: re.sub('nan', '', x))


# In[ ]:

customer_table.applymap(lambda x: re.match('nan', x)).any()


# In[ ]:

# 住所列結合
customer_table['住所'] = customer_table['住所1'] + customer_table['住所2'] + customer_table['住所3']


# In[ ]:

# 結合済み住所列削除
customer_table = customer_table[['担当者コード', '担当者名', '得意先', '名カナ', '得意先名1', '得意先名2', '郵便番号', '住所', '電話番号', 'FAX番号']]


# In[ ]:

# 列名変更
customer_table = customer_table.rename(columns={'得意先': '得意先コード'})


# In[ ]:

customer_table


# # CSV 保存、確認

# In[ ]:

# utf-8 で CSV 保存 (Excel で文字化けする)
customer_table.to_csv(r'{0}\customer_table_utf8.csv'.format(path), encoding='utf-8')


# In[ ]:

# エンコード指定なしで CSV 保存 (そのままでは pandas で読み込めなくなる)
customer_table.to_csv(r'{0}\customer_table.csv'.format(path))


# In[ ]:

# customer_table_utf8.csv
df = pd.read_csv(r'{0}\customer_table_utf8.csv'.format(path), index_col=0)


# In[ ]:

# customer_table.csv
with codecs.open(r'{0}\customer_table.csv'.format(path), mode="r", encoding="Shift-JIS", errors="ignore") as file:
    df = pd.read_table(file, delimiter=",", index_col=0)


# In[ ]:

df


# In[ ]:

df.dtypes


# In[ ]:

help(jaconv)


# In[ ]:



