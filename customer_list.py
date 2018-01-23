
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

# ## データ削減版

# In[ ]:

bestec_excel = pd.read_excel(r'{0}\Bestecaudio得意先リスト20170728更新_local_171214.xlsx'.format(path))


# In[ ]:

# 作業用テーブルの作成
bestec_table = bestec_excel[['担当者コード', '担当者名', '得意先', '名カナ', '得意先名1', '得意先名2', '郵便番号', '住所1', '住所2',
       '住所3', '電話番号', 'FAX番号']]


# In[ ]:

# str 型を明示
bestec_table = bestec_table.astype({'担当者コード':str, '担当者名': str, '得意先': str, '名カナ': str, '得意先名1': str, '得意先名2': str, '郵便番号': str,
                                        '住所1': str, '住所2': str, '住所3': str, '電話番号': str, 'FAX番号': str})


# In[ ]:

# 住所列結合
bestec_table['住所'] = bestec_table['住所1'] + bestec_table['住所2'] + bestec_table['住所3']


# In[ ]:

# 結合済み住所列削除
bestec_table = bestec_table[['担当者コード', '担当者名', '得意先', '名カナ', '得意先名1', '得意先名2', '郵便番号', '住所', '電話番号', 'FAX番号']]


# ## データ現状維持版

# In[ ]:

# version = 171221


# In[ ]:

version = 180111


# In[ ]:

# file_name = r'{0}\Bestecaudio得意先リスト20170728更新_{1}.xlsx'.format(path, version)


# In[ ]:

file_name = r'{0}\得意先リスト20180109更新_{1}.xlsx'.format(path, version)


# ### bestec

# In[ ]:

sheet_1 = 'Bestec得意先リスト'


# In[ ]:

# 読み込み (列の追加や削減、列名の変更があった場合にはこのままでは対応できないので要注意)
bestec_excel = pd.read_excel(file_name, sheet_1,
                               keep_default_na=False,
                               dtype={'FAX番号': str,
                                        '住所1': str,
                                        '住所2': str,
                                        '住所3': str,
                                        '削除区分': str,
                                        '名カナ': str,
                                        '回収支払日1': str,
                                        '回収支払月1': str,
                                        '得意先': str,
                                        '得意先名1': str,
                                        '得意先名2': str,
                                        '担当者コード': str,
                                        '担当者名': str,
                                        '業種': str,
                                        '業種名': str,
                                        '締日グループ名1': str,
                                        '郵便番号': str,
                                        '都道府県': str,
                                        '都道府県名': str,
                                        '電話番号': str})


# In[ ]:

# 列の並び替え (2017年の元ファイルに揃える)
# (列の追加や削減、列名の変更があった場合にはこのままでは対応できないので要注意)
bestec_excel = bestec_excel[['担当者コード',
                              '担当者名',
                              '得意先',
                              '名カナ',
                              '得意先名1',
                              '得意先名2',
                              '郵便番号',
                              '住所1',
                              '住所2',
                              '住所3',
                              '電話番号',
                              'FAX番号',
                              '締日グループ名1',
                              '回収支払月1',
                              '回収支払日1',
                              '業種',
                              '業種名',
                              '都道府県',
                              '都道府県名',
                              '削除区分'
                             ]]


# In[ ]:

bestec_excel.columns


# In[ ]:

bestec_excel


# In[ ]:

bestec_excel.dtypes


# In[ ]:

# 作業用テーブルの作成
bestec_table = bestec_excel.copy()


# In[ ]:

bestec_table


# ### beetech

# In[ ]:

sheet_2 = 'ビーテック得意先リスト'


# In[ ]:

# 読み込み (列の追加や削減、列名の変更があった場合にはこのままでは対応できないので要注意)
beetech_excel = pd.read_excel(file_name, sheet_2,
                               keep_default_na=False,
                               dtype={'FAX番号': str,
                                         '住所1': str,
                                         '住所2': str,
                                         '住所3': str,
                                         '削除区分': str,
                                         '名カナ': str,
                                         '回収支払日1': str,
                                         '回収支払月1': str,
                                         '回収種別名': str,
                                         '得意先': str,
                                         '得意先備考1': str,
                                         '得意先備考2': str,
                                         '得意先備考3': str,
                                         '得意先名1': str,
                                         '得意先名2': str,
                                         '担当者名': str,
                                         '業種': str,
                                         '業種名': str,
                                         '締日グループ名1': str,
                                         '請求先コード': str,
                                         '請求先略称': str,
                                         '郵便番号': str,
                                         '都道府県': str,
                                         '都道府県名': str,
                                         '電話番号': str})


# In[ ]:

# 列の並び替え (2017年の元ファイルに揃える)
# (列の追加や削減、列名の変更があった場合にはこのままでは対応できないので要注意)
beetech_excel = beetech_excel[['名カナ',
                               '得意先',
                               '得意先名1',
                               '得意先名2',
                               '担当者名',
                               '請求先コード',
                               '請求先略称',
                               '郵便番号',
                               '住所1',
                               '住所2',
                               '住所3',
                               '電話番号',
                               'FAX番号',
                               '締日グループ名1',
                               '回収支払月1',
                               '回収支払日1',
                               '回収種別名',
                               '業種',
                               '業種名',
                               '都道府県',
                               '都道府県名',
                               '得意先備考1',
                               '得意先備考2',
                               '得意先備考3',
                               '削除区分'
                              ]]


# In[ ]:

beetech_excel.columns


# In[ ]:

beetech_excel


# In[ ]:

beetech_table.dtypes


# In[ ]:

# 作業用テーブルの作成
beetech_table = beetech_excel.copy()


# In[ ]:

beetech_table


# # 内容確認

# ## bestec

# In[ ]:

bestec_table.loc[0, '名カナ']


# In[ ]:

bestec_table.loc[0, '得意先名2']


# In[ ]:

bestec_table.loc[0, '住所1']


# In[ ]:

bestec_table.loc[0, '得意先']


# In[ ]:

bestec_table['名カナ'].values


# In[ ]:

bestec_table['得意先名2'].values


# In[ ]:

len(bestec_table)


# ## beetech

# In[ ]:

beetech_table.loc[0, '名カナ']


# In[ ]:

beetech_table.loc[0, '得意先名2']


# In[ ]:

beetech_table.loc[0, '住所1']


# In[ ]:

beetech_table.loc[0, '得意先']


# In[ ]:

beetech_table['名カナ'].values


# In[ ]:

beetech_table['得意先名2'].values


# In[ ]:

len(beetech_table)


# # 整形処理

# ## bestec

# In[ ]:

# 'x1f' (ユニット区切り) を除去
bestec_table = bestec_table.applymap(lambda x: re.sub('\x1f', '', x))


# In[ ]:

# 半角カナを全角カナに
bestec_table = bestec_table.applymap(lambda x: jaconv.h2z(x, kana=True, ascii=False, digit=False))


# In[ ]:

# 全角英数を半角英数に
bestec_table = bestec_table.applymap(lambda x: jaconv.z2h(x, kana=False, ascii=True, digit=True))


# In[ ]:

# 'nan' を除去
bestec_table = bestec_table.applymap(lambda x: re.sub('nan', '', x))


# In[ ]:

# 確認
bestec_table.applymap(lambda x: re.match('\x1f', x)).any()


# In[ ]:

# 確認
bestec_table.applymap(lambda x: re.match('nan', x)).any()

# 列名変更
bestec_table = bestec_table.rename(columns={'得意先': '得意先ID'})
# In[ ]:

bestec_table


# In[ ]:

# 得意先コードの重複を確認
bestec_table['得意先'].duplicated().any()


# In[ ]:

# 得意先コードの桁数を確認
bestec_table['得意先'].apply(lambda x: len(x)).min()


# In[ ]:

# 得意先コードの桁数を確認
bestec_table['得意先'].apply(lambda x: len(x)).max()


# ## beetech

# In[ ]:

# 'x1f' (ユニット区切り) を除去
beetech_table = beetech_table.applymap(lambda x: re.sub('\x1f', '', x))


# In[ ]:

# 半角カナを全角カナに
beetech_table = beetech_table.applymap(lambda x: jaconv.h2z(x, kana=True, ascii=False, digit=False))


# In[ ]:

# 全角英数を半角英数に
beetech_table = beetech_table.applymap(lambda x: jaconv.z2h(x, kana=False, ascii=True, digit=True))


# In[ ]:

# 'nan' を除去
beetech_table = beetech_table.applymap(lambda x: re.sub('nan', '', x))


# In[ ]:

# 確認
beetech_table.applymap(lambda x: re.match('\x1f', x)).any()


# In[ ]:

# 確認
beetech_table.applymap(lambda x: re.match('nan', x)).any()

# 列名変更
beetech_table = beetech_table.rename(columns={'得意先': '得意先ID'})
# In[ ]:

beetech_table


# In[ ]:

# 得意先コードの重複を確認
beetech_table['得意先'].duplicated().any()


# In[ ]:

# 得意先コードの桁数を確認
beetech_table['得意先'].apply(lambda x: len(x)).min()


# In[ ]:

# 得意先コードの桁数を確認
beetech_table['得意先'].apply(lambda x: len(x)).max()


# # Excel 保存、再読み込み

# In[ ]:

output_name = r'{0}\Bestecaudio得意先リスト.xlsx'.format(path)


# In[ ]:

# エンコード指定なしで Excel 保存
# 出力後に全セルの書式設定を文字列に変更しておいた方がよさそう
writer = pd.ExcelWriter(output_name)
bestec_table.to_excel(writer, sheet_name=sheet_1, index=False)
beetech_table.to_excel(writer, sheet_name=sheet_2, index=False)
writer.save()


# In[ ]:

# bestec 確認
df = pd.read_excel(output_name, sheet_1, keep_default_na=False,
                   dtype={'FAX番号': str,
                            '住所1': str,
                            '住所2': str,
                            '住所3': str,
                            '削除区分': str,
                            '名カナ': str,
                            '回収支払日1': str,
                            '回収支払月1': str,
                            '得意先': str,
                            '得意先名1': str,
                            '得意先名2': str,
                            '担当者コード': str,
                            '担当者名': str,
                            '業種': str,
                            '業種名': str,
                            '締日グループ名1': str,
                            '郵便番号': str,
                            '都道府県': str,
                            '都道府県名': str,
                            '電話番号': str})


# In[ ]:

df


# In[ ]:

df.dtypes


# In[ ]:

df.applymap(lambda x: re.match('\x1f', x)).any()


# In[ ]:

df.applymap(lambda x: re.match('nan', x)).any()


# In[ ]:

df['電話番号'].values


# In[ ]:

df['名カナ'].values


# In[ ]:

df['得意先名2'].values


# In[ ]:

# beetech 確認
df = pd.read_excel(output_name, sheet_2, keep_default_na=False,
                   dtype={'FAX番号': str,
                             '住所1': str,
                             '住所2': str,
                             '住所3': str,
                             '削除区分': str,
                             '名カナ': str,
                             '回収支払日1': str,
                             '回収支払月1': str,
                             '回収種別名': str,
                             '得意先': str,
                             '得意先備考1': str,
                             '得意先備考2': str,
                             '得意先備考3': str,
                             '得意先名1': str,
                             '得意先名2': str,
                             '担当者名': str,
                             '業種': str,
                             '業種名': str,
                             '締日グループ名1': str,
                             '請求先コード': str,
                             '請求先略称': str,
                             '郵便番号': str,
                             '都道府県': str,
                             '都道府県名': str,
                             '電話番号': str})


# # CSV 保存、確認

# In[ ]:

# utf-8 で CSV 保存 (Excel で文字化けする)
bestec_table.to_csv(r'{0}\bestec_table_utf8.csv'.format(path), encoding='utf-8')


# In[ ]:

# エンコード指定なしで CSV 保存 (そのままでは pandas で読み込めなくなる)
bestec_table.to_csv(r'{0}\bestec_table.csv'.format(path))


# In[ ]:

# bestec_table_utf8.csv
df = pd.read_csv(r'{0}\bestec_table_utf8.csv'.format(path), index_col=0)


# In[ ]:

# bestec_table.csv
with codecs.open(r'{0}\bestec_table.csv'.format(path), mode="r", encoding="Shift-JIS", errors="ignore") as file:
    df = pd.read_table(file, delimiter=",", index_col=0)


# In[ ]:

df


# In[ ]:

df.dtypes


# In[ ]:



