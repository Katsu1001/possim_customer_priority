#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=========================================================================
03_data_utils.py - POSSIM Customer Priority Maker ユーティリティ
=========================================================================

【このファイルの役割】
データ処理で繰り返し使用する関数をまとめたライブラリです。
各関数は独立しているので、他のプログラムからも使えます。

【なぜこうするのか】
同じ処理を何度も書かず、1箇所にまとめることで、
修正が簡単になり、バグが減ります。

例えば、電話番号の整形ルールを変更する場合、
この関数だけを修正すれば、すべての箇所に反映されます。

【作成者】
KATSUMI KURAHASHI

【最終更新日】
2025-11-05
=========================================================================
"""

import pandas as pd
import re
from pathlib import Path

# =========================================================================
# 【1】氏名処理関数
# =========================================================================

def normalize_name(name):
    """
    氏名を正規化する（スペース・記号を削除）

    【やっていること】
    「最上_輝未子」→「最上輝未子」のように、
    スペースや記号を削除して統一します。

    【なぜ必要か】
    同じ人が「最上_輝未子」と「最上 輝未子」のように
    違う形で登録されている場合があるので、
    正規化して同一人物と判定できるようにします。

    【削除する文字】
    - 半角スペース
    - 全角スペース（　）
    - アンダースコア（_）
    - ハイフン（-）
    - 中点（・）

    引数:
        name (str): 氏名（例: "最上_輝未子"）

    戻り値:
        str: 正規化された氏名（例: "最上輝未子"）

    使用例:
        >>> normalize_name("最上_輝未子")
        "最上輝未子"
        >>> normalize_name("最上 輝未子")
        "最上輝未子"
    """
    # 空欄の場合は空文字を返す
    if pd.isna(name) or name == '':
        return ''

    # 文字列に変換
    name_str = str(name)

    # スペース、全角スペース、アンダースコア、ハイフン、中点を削除
    # \s = 半角スペース
    # 　= 全角スペース
    # _ = アンダースコア
    # \- = ハイフン（エスケープが必要）
    # ・ = 中点
    name_str = re.sub(r'[\s　_\-・]', '', name_str)

    # 前後の空白を削除
    return name_str.strip()


def normalize_email(email):
    """
    メールアドレスを正規化する（小文字化・前後空白削除）

    【やっていること】
    Test@Example.com → test@example.com のように、
    小文字に統一して前後の空白を削除します。

    【なぜ必要か】
    メールアドレスは大文字小文字を区別しないので、
    統一して比較できるようにします。

    例: "Test@Example.com" と "test@example.com" は同じメールアドレス

    引数:
        email (str): メールアドレス

    戻り値:
        str: 正規化されたメールアドレス

    使用例:
        >>> normalize_email("Test@Example.com")
        "test@example.com"
        >>> normalize_email("  USER@DOMAIN.COM  ")
        "user@domain.com"
    """
    # 空欄の場合は空文字を返す
    if pd.isna(email) or email == '':
        return ''

    # 小文字に変換して前後の空白を削除
    return str(email).lower().strip()


# =========================================================================
# 【2】住所処理関数
# =========================================================================

def extract_prefecture(address, prefectures):
    """
    住所から都道府県を抽出する

    【やっていること】
    「東京都千代田区...」→「東京都」のように、
    住所文字列から都道府県名だけを取り出します。

    【なぜ必要か】
    地域別の集計をするために、都道府県が必要です。

    【処理の流れ】
    1. 住所文字列に「東京都」が含まれているかチェック
    2. 含まれていれば「東京都」を返す
    3. 含まれていなければ、次の都道府県をチェック
    4. すべての都道府県をチェックして見つからなければNoneを返す

    引数:
        address (str): 住所（例: "東京都千代田区..."）
        prefectures (list): 都道府県リスト

    戻り値:
        str or None: 都道府県名（見つからない場合はNone）

    使用例:
        >>> extract_prefecture("東京都渋谷区代々木...", PREFECTURES)
        "東京都"
        >>> extract_prefecture("横浜市...", PREFECTURES)
        None  # 都道府県名がない
    """
    # 空欄の場合はNoneを返す
    if pd.isna(address):
        return None

    # 文字列に変換
    address = str(address)

    # 都道府県リストを1つずつチェック
    for pref in prefectures:
        if pref in address:
            return pref

    # 見つからない場合はNone
    return None


def create_full_address(row):
    """
    複数の住所列から完全な住所を作成する

    【やっていること】
    都道府県、市区町村、番地を結合して、
    1つの完全な住所を作ります。

    【なぜ必要か】
    HubSpotのデータは都道府県、市区町村、番地が別々の列に
    分かれているため、1つの住所文字列にまとめる必要があります。

    【結合の例】
    都道府県: "東京都"
    市区町村: "渋谷区"
    番地: "代々木1-2-3"
    → 住所: "東京都渋谷区代々木1-2-3"

    引数:
        row (pandas.Series): 1行分のデータ

    戻り値:
        str: 完全な住所

    使用例:
        >>> row = {'都道府県／地域': '東京都', '市区町村': '渋谷区', '番地': '代々木1-2-3'}
        >>> create_full_address(row)
        "東京都渋谷区代々木1-2-3"
    """
    parts = []  # 住所のパーツを入れる箱

    # 都道府県
    if pd.notna(row.get('都道府県／地域', '')):
        parts.append(str(row['都道府県／地域']))

    # 市区町村
    if pd.notna(row.get('市区町村', '')):
        parts.append(str(row['市区町村']))

    # 番地
    if pd.notna(row.get('番地', '')):
        parts.append(str(row['番地']))

    # 結合して返す
    return ''.join(parts)


# =========================================================================
# 【3】電話番号・郵便番号の整形関数
# =========================================================================

def format_phone(phone):
    """
    電話番号を整形する（090-1234-5678形式）

    【やっていること】
    09012345678 → 090-1234-5678 のように、
    ハイフンを入れて見やすくします。

    【対応形式】
    - 11桁: 090-1234-5678（携帯電話）
    - 10桁: 03-1234-5678（固定電話）
    - その他: そのまま返す

    【処理の流れ】
    1. 数字以外の文字を削除（ハイフン、カッコなど）
    2. 桁数をチェック
    3. 11桁なら 090-1234-5678 形式
    4. 10桁なら 03-1234-5678 形式

    引数:
        phone (str): 電話番号

    戻り値:
        str: 整形された電話番号

    使用例:
        >>> format_phone("09012345678")
        "090-1234-5678"
        >>> format_phone("03-1234-5678")
        "03-1234-5678"
        >>> format_phone("(090)1234-5678")
        "090-1234-5678"
    """
    # 空欄の場合は空文字を返す
    if pd.isna(phone) or phone == '':
        return ''

    # 数字以外を削除
    # \D = 数字以外の文字（ハイフン、カッコ、スペースなど）
    digits = re.sub(r'\D', '', str(phone))

    # 11桁の場合（携帯電話など）
    # 例: 09012345678 → 090-1234-5678
    if len(digits) == 11:
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"

    # 10桁の場合（固定電話など）
    # 例: 0312345678 → 03-1234-5678
    elif len(digits) == 10:
        return f"{digits[:2]}-{digits[2:6]}-{digits[6:]}"

    # それ以外はそのまま返す
    return digits


def format_postal(postal):
    """
    郵便番号を整形する（123-4567形式）

    【やっていること】
    1234567 → 123-4567 のように、
    ハイフンを入れて見やすくします。

    【処理の流れ】
    1. 数字以外の文字を削除（ハイフンなど）
    2. 桁数をチェック
    3. 7桁なら 123-4567 形式

    引数:
        postal (str): 郵便番号

    戻り値:
        str: 整形された郵便番号

    使用例:
        >>> format_postal("1234567")
        "123-4567"
        >>> format_postal("123-4567")
        "123-4567"
    """
    # 空欄の場合は空文字を返す
    if pd.isna(postal) or postal == '':
        return ''

    # 数字以外を削除
    digits = re.sub(r'\D', '', str(postal))

    # 7桁の場合
    if len(digits) == 7:
        return f"{digits[:3]}-{digits[3:]}"

    # それ以外はそのまま返す
    return digits


# =========================================================================
# 【4】アンケート処理関数
# =========================================================================

def extract_survey_emails(excel_path, email_column_index=1):
    """
    Excelアンケートファイルからメールアドレスを抽出する

    【やっていること】
    卒業式アンケート（Excel）から、
    回答者のメールアドレスを取得します。

    【なぜ必要か】
    アンケート回答者は優先順位が高いので、
    回答者のメールアドレスを特定する必要があります。

    【処理の流れ】
    1. Excelファイルを読み込む
    2. メールアドレスの列（通常は2列目）を取得
    3. @が含まれているもののみ抽出（有効なメールアドレス）
    4. 小文字に統一して重複を削除
    5. セットにして返す

    引数:
        excel_path (Path): Excelファイルのパス
        email_column_index (int): メールアドレスの列番号（デフォルト: 1 = 2列目）

    戻り値:
        set: メールアドレスのセット

    使用例:
        >>> extract_survey_emails(Path("HCU7期卒業式アンケート.xlsx"))
        {"user1@example.com", "user2@example.com", ...}
    """
    try:
        # Excelファイルを読み込む
        df = pd.read_excel(excel_path)

        # メールアドレスの列を取得
        # 通常、Googleフォームのアンケート結果は、
        # 1列目がタイムスタンプ、2列目がメールアドレス
        email_col = df.columns[email_column_index]

        # メールアドレスを抽出
        emails = df[email_col].dropna().astype(str).str.lower().str.strip()

        # @が含まれているもののみ（有効なメールアドレス）
        valid_emails = emails[emails.str.contains('@', na=False)]

        # セットにして返す（重複を自動的に削除）
        return set(valid_emails)

    except Exception as e:
        # エラーが発生した場合は、警告を表示して空のセットを返す
        print(f"  ⚠️ アンケート読み込みエラー: {excel_path.name} - {e}")
        return set()


def extract_survey_data(excel_path, email_column_index=1, name_column_index=2):
    """
    Excelアンケートファイルから氏名とメールアドレスを抽出する

    【やっていること】
    卒業式アンケート（Excel）から、
    回答者の氏名とメールアドレスを取得します。

    【なぜ必要か】
    氏名とメールアドレスの両方でマッチングを試みるため、
    両方の情報を取得する必要があります。

    【処理の流れ】
    1. Excelファイルを読み込む
    2. 氏名とメールアドレスの列を取得
    3. 正規化して返す（氏名は正規化、メールは小文字化）

    引数:
        excel_path (Path): Excelファイルのパス
        email_column_index (int): メールアドレスの列番号（デフォルト: 1 = 2列目）
        name_column_index (int): 氏名の列番号（デフォルト: 2 = 3列目）

    戻り値:
        dict: {'names': set(氏名セット), 'emails': set(メールセット)}

    使用例:
        >>> extract_survey_data(Path("HCU7期卒業式アンケート.xlsx"))
        {'names': {'最上輝未子', ...}, 'emails': {'user@example.com', ...}}
    """
    try:
        # Excelファイルを読み込む
        df = pd.read_excel(excel_path)

        result = {'names': set(), 'emails': set()}

        # メールアドレスを抽出
        if email_column_index < len(df.columns):
            email_col = df.columns[email_column_index]
            emails = df[email_col].dropna().astype(str).str.lower().str.strip()
            valid_emails = emails[emails.str.contains('@', na=False)]
            result['emails'] = set(valid_emails)

        # 氏名を抽出（氏名列が存在する場合）
        if name_column_index < len(df.columns):
            name_col = df.columns[name_column_index]
            names = df[name_col].dropna().astype(str)
            # 氏名を正規化
            normalized_names = names.apply(normalize_name)
            # 空欄を除外
            valid_names = normalized_names[normalized_names != '']
            result['names'] = set(valid_names)

        return result

    except Exception as e:
        # エラーが発生した場合は、警告を表示して空の辞書を返す
        print(f"  ⚠️ アンケート読み込みエラー: {excel_path.name} - {e}")
        return {'names': set(), 'emails': set()}


# =========================================================================
# 【5】データ検証関数
# =========================================================================

def check_address_duplicates(df, address_column='住所', top_n=10):
    """
    住所の重複をチェックする

    【やっていること】
    同じ住所が複数回出現していないかチェックします。

    【なぜ必要か】
    データ品質を確認するためです。
    同じ住所に複数人が住んでいる場合、
    家族や同じ会社の可能性があります。

    【処理の流れ】
    1. 住所でグループ化してカウント
    2. 2件以上の重複のみ抽出
    3. 上位N件を取得

    引数:
        df (pandas.DataFrame): データフレーム
        address_column (str): 住所の列名
        top_n (int): 上位何件を表示するか

    戻り値:
        tuple: (重複データフレーム, 重複件数)

    使用例:
        >>> df_duplicates, duplicate_count = check_address_duplicates(df, top_n=10)
        >>> print(f"重複住所: {duplicate_count} 件")
    """
    # 住所でグループ化してカウント
    address_counts = df[address_column].value_counts()

    # 2件以上の重複のみ抽出
    duplicates = address_counts[address_counts > 1]

    # 上位N件を取得
    top_duplicates = duplicates.head(top_n)

    # DataFrameに変換
    df_duplicates = pd.DataFrame({
        '順位': [f"{i+1}." for i in range(len(top_duplicates))],
        '住所': top_duplicates.index.tolist(),
        '件数': top_duplicates.values.tolist()
    })

    return df_duplicates, len(duplicates)


def calculate_data_quality(df):
    """
    データ品質を計算する（欠損率など）

    【やっていること】
    電話番号、住所、メールアドレスなどの
    データ保有率を計算します。

    【なぜ必要か】
    データの品質を確認し、問題があれば対処します。

    例えば、電話番号が50%しかない場合、
    データの収集方法に問題がある可能性があります。

    【処理の流れ】
    1. 各項目の件数をカウント（null以外）
    2. 全体件数で割って割合を計算
    3. DataFrameにまとめて返す

    引数:
        df (pandas.DataFrame): データフレーム

    戻り値:
        pandas.DataFrame: データ品質チェック結果

    使用例:
        >>> quality_df = calculate_data_quality(df)
        >>> print(quality_df)
        チェック項目            件数  全体  割合(%)
        電話番号データ件数      1500  2000   75.0
        住所データ件数         1800  2000   90.0
        ...
    """
    total = len(df)  # 全体の件数

    # 各項目の件数をカウント（null以外）
    phone_count = df['電話番号'].notna().sum()
    address_count = df['住所'].notna().sum()
    prefecture_count = df['都道府県'].notna().sum()
    email_count = df['Eメール'].notna().sum()

    # DataFrameに変換
    quality_data = {
        'チェック項目': [
            '電話番号データ件数',
            '住所データ件数',
            '都道府県データ件数',
            'メールアドレスデータ件数'
        ],
        '件数': [phone_count, address_count, prefecture_count, email_count],
        '全体': [total, total, total, total],
        '割合(%)': [
            f"{phone_count / total * 100:.1f}",
            f"{address_count / total * 100:.1f}",
            f"{prefecture_count / total * 100:.1f}",
            f"{email_count / total * 100:.1f}"
        ]
    }

    return pd.DataFrame(quality_data)


# =========================================================================
# 【6】優先順位計算関数
# =========================================================================

def calculate_priority(row, survey_emails):
    """
    顧客の優先順位を計算する

    【やっていること】
    期番号とアンケート回答状況から、
    営業優先順位（1～4）を決定します。

    【優先順位の定義】
    1: 12期生（最優先）
    2: 7～11期 + アンケート回答者（次に優先）
    3: 7～11期 + 未回答（その次）
    4: 1～6期（低優先）

    【なぜこの順番なのか】
    - 12期生は最新の受講生なので、最も優先度が高い
    - 7～11期でアンケート回答者は、関係性が強いと判断
    - 7～11期でも未回答者は、優先度を下げる
    - 1～6期は受講から時間が経っているため、優先度は低い

    【処理の流れ】
    1. 期番号を取得
    2. メールアドレスがアンケート回答者に含まれているかチェック
    3. 条件に応じて優先順位を返す

    引数:
        row (pandas.Series): 1行分のデータ
        survey_emails (set): アンケート回答者のメールアドレスセット

    戻り値:
        int: 優先順位（1～4）

    使用例:
        >>> row = {'期_番号': 12, 'Eメール': 'user@example.com'}
        >>> calculate_priority(row, survey_emails)
        1  # 12期生なので優先順位1
    """
    period_num = row['期_番号']  # 期番号（1～12）
    email = row['Eメール']  # メールアドレス

    # アンケート回答チェック
    is_survey_respondent = email in survey_emails

    # 優先順位を判定
    if period_num == 12:
        # 優先順位1: 12期生
        return 1
    elif 7 <= period_num <= 11 and is_survey_respondent:
        # 優先順位2: 7～11期 + アンケート回答
        return 2
    elif 7 <= period_num <= 11:
        # 優先順位3: 7～11期 + 未回答
        return 3
    else:
        # 優先順位4: 1～6期
        return 4


# =========================================================================
# 【7】郵便番号から都道府県を補完する関数
# =========================================================================

def get_prefecture_from_postal_code(postal_code_str, postal_mapping):
    """
    郵便番号文字列から都道府県を抽出

    【やっていること】
    郵便番号の最初の3桁から都道府県を判定します。

    【使用例】
    get_prefecture_from_postal_code('1001111', POSTAL_CODE_TO_PREFECTURE) → '東京都'
    get_prefecture_from_postal_code('100-1111', POSTAL_CODE_TO_PREFECTURE) → '東京都'
    get_prefecture_from_postal_code('9601111', POSTAL_CODE_TO_PREFECTURE) → '福島県'

    【処理の流れ】
    1. 郵便番号から数字のみを抽出
    2. 最初の3桁を取得
    3. マッピング辞書で検索
    4. マッチした都道府県を返す

    引数:
        postal_code_str (str): 郵便番号文字列（ハイフンあり・なし両対応）
        postal_mapping (dict): 郵便番号→都道府県のマッピング辞書

    戻り値:
        str or None: 都道府県名（見つからない場合はNone）
    """
    # 空欄の場合はNoneを返す
    if pd.isna(postal_code_str) or postal_code_str == '':
        return None

    # 郵便番号から数字のみを抽出
    postal_digits = re.sub(r'\D', '', str(postal_code_str))

    if len(postal_digits) < 3:
        return None

    # 最初の3桁で検索
    prefix_3digit = postal_digits[:3]

    # マッピング辞書で検索
    return postal_mapping.get(prefix_3digit, None)


def fill_prefecture_from_postal_code(df, postal_code_col='郵便番号_フォーマット', prefecture_col='都道府県'):
    """
    【新規機能】郵便番号をもとに都道府県を補完

    【やっていること】
    郵便番号が存在するが、都道府県が空欄のセルに対して、
    郵便番号の最初の3桁をもとに、対応する都道府県を自動補完します。

    例:
    - 郵便番号: 1001234 → 都道府県を「東京都」に補完
    - 郵便番号: 5301234 → 都道府県を「大阪府」に補完
    - 郵便番号: 9601234 → 都道府県を「福島県」に補完

    処理フロー:
    1. 郵便番号→都道府県マッピング辞書を取得（02_config.py から）
    2. df を走査:
       - 郵便番号が存在し、都道府県が空欄のレコードを抽出
       - 郵便番号の最初の3桁から対応する都道府県を検索
       - マッピング結果を都道府県カラムに記入
    3. 補完結果をカウント＆ログ出力

    引数:
        df (DataFrame): 顧客データ
        postal_code_col (str): 郵便番号カラム名
        prefecture_col (str): 都道府県カラム名

    戻り値:
        tuple: (補完済みDF, 補完件数, 補完できなかった件数)
    """
    # 02_config.py から郵便番号マッピング辞書をインポート
    # （この関数が呼ばれる時点で、すでに config_02 モジュールから POSTAL_CODE_TO_PREFECTURE がインポートされている前提）
    from config_02 import POSTAL_CODE_TO_PREFECTURE

    # 補完前の都道府県データ件数を記録
    before_count = df[prefecture_col].notna().sum()

    # 郵便番号が存在し、都道府県が空欄のレコードを抽出
    mask_needs_completion = df[postal_code_col].notna() & df[prefecture_col].isna()
    needs_completion_count = mask_needs_completion.sum()

    print(f"  📍 郵便番号が存在し、都道府県が空欄: {needs_completion_count:,} 件")

    # 補完を実行
    completed_count = 0
    failed_count = 0

    for idx in df[mask_needs_completion].index:
        postal_code = df.at[idx, postal_code_col]
        prefecture = get_prefecture_from_postal_code(postal_code, POSTAL_CODE_TO_PREFECTURE)

        if prefecture:
            df.at[idx, prefecture_col] = prefecture
            completed_count += 1
        else:
            failed_count += 1

    # 補完後の都道府県データ件数
    after_count = df[prefecture_col].notna().sum()

    return df, completed_count, failed_count


# =========================================================================
# 【8】アンケート紐づけ関数（氏名＋メールアドレス）
# =========================================================================

def merge_survey_data_enhanced(df_combined, survey_data_dict):
    """
    【改善版】アンケート回答者をデータに紐づける

    【改善内容】
    1. 氏名での紐づけを実行
    2. 氏名での紐づけに失敗した場合、メールアドレスでも試行
    3. マッチング結果をログに出力

    処理フロー:
    1. df_combined に 'アンケート回答' カラムを初期化（False）
    2. 各調査ファイル（HCU7～11）から氏名とメールを抽出
    3. df_combined を走査して:
       - 第1段階: 氏名で完全一致を確認
       - 第2段階: 氏名未マッチなら、メールで完全一致を確認
       - マッチ結果をカウント
    4. ログに以下を出力:
       ✅ HCU8期: 13件（氏名: 13件, メール: 0件）
       ✅ HCU9期: 53件（氏名: 50件, メール: 3件）
       ...
       📊 総マッチ数: 91件 → ○○件（メールで追加マッチ）

    引数:
        df_combined (DataFrame): 全期間の顧客データ（氏名が正規化済み）
        survey_data_dict (dict): {期番号: アンケートファイルパス}

    戻り値:
        tuple: (アンケート回答フラグが追加されたDF, アンケート回答者のメールセット)
    """
    # アンケート回答フラグを初期化
    df_combined['アンケート回答'] = False

    # 全期の集計用
    total_name_matches = 0
    total_email_matches = 0
    all_survey_emails = set()

    print("\n各期のマッチング結果:")

    # 各期のアンケートファイルを処理
    for period, filepath in survey_data_dict.items():
        try:
            # アンケートデータから氏名とメールアドレスを抽出
            survey_data = extract_survey_data(filepath)
            survey_names = survey_data['names']
            survey_emails = survey_data['emails']

            # この期のマッチング結果をカウント
            name_match_count = 0
            email_match_count = 0

            # 氏名でマッチング
            for name in survey_names:
                # 氏名が完全一致するレコードを探す
                mask = (df_combined['氏名_key'] == name) & (~df_combined['アンケート回答'])
                matched_indices = df_combined[mask].index

                if len(matched_indices) > 0:
                    # マッチしたレコードにフラグを立てる
                    df_combined.loc[matched_indices, 'アンケート回答'] = True
                    name_match_count += len(matched_indices)

            # メールアドレスでマッチング（氏名でマッチしなかったもののみ）
            for email in survey_emails:
                # メールアドレスが完全一致し、まだマッチしていないレコードを探す
                mask = (df_combined['Eメール'] == email) & (~df_combined['アンケート回答'])
                matched_indices = df_combined[mask].index

                if len(matched_indices) > 0:
                    # マッチしたレコードにフラグを立てる
                    df_combined.loc[matched_indices, 'アンケート回答'] = True
                    email_match_count += len(matched_indices)

            # 全期のメールアドレスセットに追加（優先順位計算用）
            all_survey_emails.update(survey_emails)

            # この期の結果をログ出力
            total_matches = name_match_count + email_match_count
            if email_match_count > 0:
                print(f"  ✅ HCU{period}期: {total_matches} 件（氏名: {name_match_count}件, メール: {email_match_count}件★）")
            else:
                print(f"  ✅ HCU{period}期: {total_matches} 件（氏名: {name_match_count}件, メール: {email_match_count}件）")

            # 全期の集計に加算
            total_name_matches += name_match_count
            total_email_matches += email_match_count

        except FileNotFoundError:
            # ファイルが見つからない場合は警告を表示
            print(f"  ⚠️ HCU{period}期: ファイルが見つかりません")
        except Exception as e:
            # その他のエラー
            print(f"  ⚠️ HCU{period}期: エラー - {e}")

    # 総マッチ数を計算
    total_respondents = df_combined['アンケート回答'].sum()
    total_customers = len(df_combined)
    response_rate = total_respondents / total_customers * 100

    print(f"\n  📊 アンケート回答者数: {total_respondents:,} 人 / {total_customers:,} 人 ({response_rate:.1f}%)")
    print(f"     内訳: 氏名マッチ {total_name_matches}件、メールマッチ {total_email_matches}件（新規）")

    return df_combined, all_survey_emails


def generate_survey_response_details(df_combined, survey_data_dict):
    """
    アンケート回答詳細データを生成する

    【目的】
    アンケート回答者の詳細情報を一覧化し、
    マッチング方法（氏名/メール/未マッチ）を明示する

    【出力カラム】
    - 氏名（顧客）
    - メールアドレス（顧客）
    - 氏名（アンケート）
    - メールアドレス（アンケート）
    - マッチング方法
    - アンケート回答期
    - アンケート回答フラグ

    引数:
        df_combined (DataFrame): 全期間の顧客データ（マッチング済み）
        survey_data_dict (dict): {期番号: アンケートファイルパス}

    戻り値:
        DataFrame: アンケート回答詳細データ
    """
    detail_records = []

    # 各期のアンケートファイルを処理
    for period, filepath in survey_data_dict.items():
        try:
            # アンケートデータから氏名とメールアドレスを抽出
            survey_data = extract_survey_data(filepath)
            survey_names = survey_data['names']
            survey_emails = survey_data['emails']

            # アンケート回答者ごとに処理
            # 1. 氏名でマッチング確認
            for survey_name in survey_names:
                mask = df_combined['氏名_key'] == survey_name
                matched_customers = df_combined[mask]

                if len(matched_customers) > 0:
                    for _, customer in matched_customers.iterrows():
                        detail_records.append({
                            '氏名（顧客）': customer.get('氏名', ''),
                            'メールアドレス（顧客）': customer.get('Eメール', ''),
                            '氏名（アンケート）': survey_name,
                            'メールアドレス（アンケート）': '',  # 氏名マッチ時はメール不明
                            'マッチング方法': '氏名',
                            'アンケート回答期': f'HCU{period}',
                            'アンケート回答フラグ': '○'
                        })

            # 2. メールアドレスでマッチング確認（氏名でマッチしなかったもののみ）
            for survey_email in survey_emails:
                mask = (df_combined['Eメール'] == survey_email) & \
                       (~df_combined['氏名_key'].isin(survey_names))
                matched_customers = df_combined[mask]

                if len(matched_customers) > 0:
                    for _, customer in matched_customers.iterrows():
                        detail_records.append({
                            '氏名（顧客）': customer.get('氏名', ''),
                            'メールアドレス（顧客）': customer.get('Eメール', ''),
                            '氏名（アンケート）': '',  # メールマッチ時は氏名不明
                            'メールアドレス（アンケート）': survey_email,
                            'マッチング方法': 'メール',
                            'アンケート回答期': f'HCU{period}',
                            'アンケート回答フラグ': '○'
                        })

        except Exception as e:
            print(f"  ⚠️ HCU{period}期の詳細データ生成エラー: {e}")

    # DataFrameに変換
    df_details = pd.DataFrame(detail_records)

    # 重複を削除（同じ顧客が複数期に回答している場合）
    if len(df_details) > 0:
        df_details = df_details.drop_duplicates(subset=['氏名（顧客）', 'メールアドレス（顧客）'], keep='first')

    return df_details


# =========================================================================
# 【10】住所標準化関数
# =========================================================================

# 市区町村名から都道府県を推定する辞書
CITY_TO_PREFECTURE = {
    # 東京都の特別区
    '千代田区': '東京都', '中央区': '東京都', '港区': '東京都', '新宿区': '東京都',
    '文京区': '東京都', '台東区': '東京都', '墨田区': '東京都', '江東区': '東京都',
    '品川区': '東京都', '目黒区': '東京都', '大田区': '東京都', '世田谷区': '東京都',
    '渋谷区': '東京都', '中野区': '東京都', '杉並区': '東京都', '豊島区': '東京都',
    '北区': '東京都', '荒川区': '東京都', '板橋区': '東京都', '練馬区': '東京都',
    '足立区': '東京都', '葛飾区': '東京都', '江戸川区': '東京都',

    # 東京都の市部（主要）
    '八王子市': '東京都', '立川市': '東京都', '武蔵野市': '東京都', '三鷹市': '東京都',
    '青梅市': '東京都', '府中市': '東京都', '昭島市': '東京都', '調布市': '東京都',
    '町田市': '東京都', '小金井市': '東京都', '小平市': '東京都', '日野市': '東京都',
    '東村山市': '東京都', '国分寺市': '東京都', '国立市': '東京都', '福生市': '東京都',
    '狛江市': '東京都', '東大和市': '東京都', '清瀬市': '東京都', '東久留米市': '東京都',
    '武蔵村山市': '東京都', '多摩市': '東京都', '稲城市': '東京都', '羽村市': '東京都',
    'あきる野市': '東京都', '西東京市': '東京都',

    # 神奈川県の政令指定都市・主要市
    '横浜市': '神奈川県', '川崎市': '神奈川県', '相模原市': '神奈川県',
    '横須賀市': '神奈川県', '平塚市': '神奈川県', '鎌倉市': '神奈川県',
    '藤沢市': '神奈川県', '小田原市': '神奈川県', '茅ヶ崎市': '神奈川県',
    '逗子市': '神奈川県', '三浦市': '神奈川県', '秦野市': '神奈川県',
    '厚木市': '神奈川県', '大和市': '神奈川県', '伊勢原市': '神奈川県',
    '海老名市': '神奈川県', '座間市': '神奈川県', '南足柄市': '神奈川県',
    '綾瀬市': '神奈川県',

    # 埼玉県の政令指定都市・主要市
    'さいたま市': '埼玉県', '川越市': '埼玉県', '熊谷市': '埼玉県',
    '川口市': '埼玉県', '行田市': '埼玉県', '秩父市': '埼玉県',
    '所沢市': '埼玉県', '飯能市': '埼玉県', '加須市': '埼玉県',
    '本庄市': '埼玉県', '東松山市': '埼玉県', '春日部市': '埼玉県',
    '狭山市': '埼玉県', '羽生市': '埼玉県', '鴻巣市': '埼玉県',
    '深谷市': '埼玉県', '上尾市': '埼玉県', '草加市': '埼玉県',
    '越谷市': '埼玉県', '蕨市': '埼玉県', '戸田市': '埼玉県',
    '入間市': '埼玉県', '朝霞市': '埼玉県', '志木市': '埼玉県',
    '和光市': '埼玉県', '新座市': '埼玉県', '桶川市': '埼玉県',
    '久喜市': '埼玉県', '北本市': '埼玉県', '八潮市': '埼玉県',
    '富士見市': '埼玉県', '三郷市': '埼玉県', '蓮田市': '埼玉県',
    '坂戸市': '埼玉県', '幸手市': '埼玉県', '鶴ヶ島市': '埼玉県',
    '日高市': '埼玉県', '吉川市': '埼玉県', 'ふじみ野市': '埼玉県',
    '白岡市': '埼玉県',

    # 千葉県の政令指定都市・主要市
    '千葉市': '千葉県', '銚子市': '千葉県', '市川市': '千葉県',
    '船橋市': '千葉県', '館山市': '千葉県', '木更津市': '千葉県',
    '松戸市': '千葉県', '野田市': '千葉県', '茂原市': '千葉県',
    '成田市': '千葉県', '佐倉市': '千葉県', '東金市': '千葉県',
    '旭市': '千葉県', '習志野市': '千葉県', '柏市': '千葉県',
    '勝浦市': '千葉県', '市原市': '千葉県', '流山市': '千葉県',
    '八千代市': '千葉県', '我孫子市': '千葉県', '鴨川市': '千葉県',
    '鎌ケ谷市': '千葉県', '君津市': '千葉県', '富津市': '千葉県',
    '浦安市': '千葉県', '四街道市': '千葉県', '袖ケ浦市': '千葉県',
    '八街市': '千葉県', '印西市': '千葉県', '白井市': '千葉県',
    '富里市': '千葉県', '南房総市': '千葉県', '匝瑳市': '千葉県',
    '香取市': '千葉県', '山武市': '千葉県', 'いすみ市': '千葉県',
    '大網白里市': '千葉県',

    # その他の政令指定都市
    '札幌市': '北海道', '仙台市': '宮城県', '新潟市': '新潟県',
    '静岡市': '静岡県', '浜松市': '静岡県', '名古屋市': '愛知県',
    '京都市': '京都府', '大阪市': '大阪府', '堺市': '大阪府',
    '神戸市': '兵庫県', '岡山市': '岡山県', '広島市': '広島県',
    '北九州市': '福岡県', '福岡市': '福岡県', '熊本市': '熊本県',
}


def standardize_address_with_prefecture(address, current_prefecture=None):
    """
    住所を都道府県名から始まる標準形式に変換する

    【やっていること】
    「横浜市鶴見区...」→「神奈川県横浜市鶴見区...」のように、
    都道府県名が欠けている住所に都道府県名を補完します。

    【処理の流れ】
    1. 既に都道府県名から始まっている場合はそのまま返す
    2. 市区町村名から都道府県を推定
    3. 推定できた場合は都道府県名を追加
    4. 推定できない場合は元の住所を返す

    引数:
        address (str): 住所文字列
        current_prefecture (str): 既に判明している都道府県（オプション）

    戻り値:
        str: 都道府県名から始まる住所
    """
    if pd.isna(address) or address == '':
        return address

    address = str(address).strip()

    # 既に都道府県名から始まっている場合はそのまま返す
    prefectures = [
        '北海道',
        '青森県', '岩手県', '宮城県', '秋田県', '山形県', '福島県',
        '茨城県', '栃木県', '群馬県', '埼玉県', '千葉県', '東京都', '神奈川県',
        '新潟県', '富山県', '石川県', '福井県', '山梨県', '長野県',
        '岐阜県', '静岡県', '愛知県', '三重県',
        '滋賀県', '京都府', '大阪府', '兵庫県', '奈良県', '和歌山県',
        '鳥取県', '島根県', '岡山県', '広島県', '山口県',
        '徳島県', '香川県', '愛媛県', '高知県',
        '福岡県', '佐賀県', '長崎県', '熊本県', '大分県', '宮崎県', '鹿児島県', '沖縄県'
    ]

    for pref in prefectures:
        if address.startswith(pref):
            return address

    # 市区町村名から都道府県を推定
    for city, pref in CITY_TO_PREFECTURE.items():
        if address.startswith(city):
            return f"{pref}{address}"

    # 既に判明している都道府県があれば使用
    if current_prefecture and pd.notna(current_prefecture):
        return f"{current_prefecture}{address}"

    # 推定できない場合は元の住所を返す
    return address


# ローマ字地名を日本語に変換する辞書
ROMAJI_TO_JAPANESE_PLACE = {
    # 都道府県
    'Tokyo': '東京都', 'Osaka': '大阪府', 'Kyoto': '京都府',
    'Hokkaido': '北海道', 'Fukuoka': '福岡県', 'Nagasaki': '長崎県',
    'Hiroshima': '広島県', 'Yokohama': '神奈川県', 'Nagoya': '愛知県',
    'Sapporo': '北海道', 'Kobe': '兵庫県', 'Sendai': '宮城県',

    # 主要市区町村（東京）
    'Shibuya': '渋谷区', 'Shinjuku': '新宿区', 'Minato': '港区',
    'Chiyoda': '千代田区', 'Chuo': '中央区', 'Setagaya': '世田谷区',
    'Meguro': '目黒区', 'Shinagawa': '品川区', 'Ota': '大田区',
    'Suginami': '杉並区', 'Nakano': '中野区', 'Toshima': '豊島区',
    'Bunkyo': '文京区', 'Taito': '台東区', 'Sumida': '墨田区',
    'Koto': '江東区', 'Kita': '北区', 'Arakawa': '荒川区',
    'Itabashi': '板橋区', 'Nerima': '練馬区', 'Adachi': '足立区',
    'Katsushika': '葛飾区', 'Edogawa': '江戸川区',

    # 主要市（神奈川）
    'Kawasaki': '川崎市', 'Sagamihara': '相模原市', 'Fujisawa': '藤沢市',
    'Kamakura': '鎌倉市', 'Yokosuka': '横須賀市',

    # 主要市（埼玉・千葉）
    'Saitama': 'さいたま市', 'Kawagoe': '川越市', 'Chiba': '千葉市',
    'Funabashi': '船橋市', 'Matsudo': '松戸市', 'Kashiwa': '柏市',

    # 長崎関連
    'Nagasaki': '長崎県', 'Tomachi': '戸町',
}


def convert_romaji_address_to_japanese(address):
    """
    ローマ字住所を日本語に変換する

    【やっていること】
    「NagasakiTomachi 2-20-57」→「長崎県長崎市戸町 2-20-57」のように、
    ローマ字で表記されている地名を日本語に変換します。

    【処理の流れ】
    1. アドレスにローマ字地名が含まれているか確認
    2. 含まれている場合は日本語に置換
    3. 都道府県名を補完

    引数:
        address (str): 住所文字列

    戻り値:
        str: 日本語に変換された住所
    """
    if pd.isna(address) or address == '':
        return address

    address = str(address).strip()
    original_address = address

    # ローマ字地名を日本語に変換
    for romaji, japanese in ROMAJI_TO_JAPANESE_PLACE.items():
        # 大文字小文字を区別せずに検索
        pattern = re.compile(re.escape(romaji), re.IGNORECASE)
        address = pattern.sub(japanese, address)

    # 変換が行われた場合は都道府県名を補完
    if address != original_address:
        # 長崎県の特別処理（Nagasaki→長崎県 となった場合）
        if address.startswith('長崎県') and '長崎市' not in address:
            # 「長崎県戸町」→「長崎県長崎市戸町」
            address = address.replace('長崎県戸町', '長崎県長崎市戸町')

    return address
