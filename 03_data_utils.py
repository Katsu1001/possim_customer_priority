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
Claude Code

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
