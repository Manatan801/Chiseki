"""Accessファイル(.mdb/.accdb)の照合モジュール

mdb-tools (Linux CLI) を使ってテーブル/レコードを読み取り、
2つのAccessファイル間の差分を検出する。
"""

import io
import subprocess
from dataclasses import dataclass, field

import pandas as pd


@dataclass
class TableDiff:
    """1テーブルの差分結果"""
    table_name: str
    only_in_a: pd.DataFrame = field(default_factory=pd.DataFrame)
    only_in_b: pd.DataFrame = field(default_factory=pd.DataFrame)
    modified: pd.DataFrame = field(default_factory=pd.DataFrame)
    columns_only_in_a: list = field(default_factory=list)
    columns_only_in_b: list = field(default_factory=list)


@dataclass
class CompareResult:
    """照合結果全体"""
    tables_only_in_a: list = field(default_factory=list)
    tables_only_in_b: list = field(default_factory=list)
    tables_common: list = field(default_factory=list)
    table_diffs: list = field(default_factory=list)
    has_differences: bool = False


def list_tables(mdb_path: str) -> list[str]:
    """mdb-tablesでテーブル一覧を取得"""
    result = subprocess.run(
        ["mdb-tables", "-1", mdb_path],
        capture_output=True, text=True, timeout=30,
    )
    if result.returncode != 0:
        raise RuntimeError(f"mdb-tables failed: {result.stderr}")
    return [t.strip() for t in result.stdout.strip().split("\n") if t.strip()]


def read_table(mdb_path: str, table_name: str) -> pd.DataFrame:
    """mdb-exportでテーブルをDataFrameとして読み取る"""
    result = subprocess.run(
        ["mdb-export", mdb_path, table_name],
        capture_output=True, text=True, timeout=60,
    )
    if result.returncode != 0:
        raise RuntimeError(f"mdb-export failed for {table_name}: {result.stderr}")
    if not result.stdout.strip():
        return pd.DataFrame()
    return pd.read_csv(io.StringIO(result.stdout), dtype=str, keep_default_na=False)


def _guess_key_columns(df: pd.DataFrame) -> list[str]:
    """キーカラムを推定する"""
    key_patterns = ["id", "ID", "番号", "コード", "No", "NO", "key", "KEY"]
    candidates = []
    for col in df.columns:
        for pat in key_patterns:
            if pat in col:
                candidates.append(col)
                break
    return candidates if candidates else list(df.columns)


def compare_tables(df_a: pd.DataFrame, df_b: pd.DataFrame, table_name: str) -> TableDiff:
    """2つのDataFrameを比較して差分を返す"""
    diff = TableDiff(table_name=table_name)

    cols_a = set(df_a.columns)
    cols_b = set(df_b.columns)
    diff.columns_only_in_a = sorted(cols_a - cols_b)
    diff.columns_only_in_b = sorted(cols_b - cols_a)

    common_cols = sorted(cols_a & cols_b)
    if not common_cols:
        return diff

    a = df_a[common_cols].copy()
    b = df_b[common_cols].copy()

    for col in common_cols:
        a[col] = a[col].astype(str).str.strip()
        b[col] = b[col].astype(str).str.strip()

    key_cols = _guess_key_columns(a)

    if set(key_cols) == set(common_cols):
        a_tuples = set(a.itertuples(index=False, name=None))
        b_tuples = set(b.itertuples(index=False, name=None))
        only_a = a_tuples - b_tuples
        only_b = b_tuples - a_tuples
        if only_a:
            diff.only_in_a = pd.DataFrame(list(only_a), columns=common_cols)
        if only_b:
            diff.only_in_b = pd.DataFrame(list(only_b), columns=common_cols)
    else:
        a_keyed = a.set_index(key_cols, drop=False)
        b_keyed = b.set_index(key_cols, drop=False)
        keys_a = set(a_keyed.index)
        keys_b = set(b_keyed.index)

        if keys_a - keys_b:
            diff.only_in_a = a_keyed.loc[list(keys_a - keys_b)].reset_index(drop=True)
        if keys_b - keys_a:
            diff.only_in_b = b_keyed.loc[list(keys_b - keys_a)].reset_index(drop=True)

        modified_rows = []
        value_cols = [c for c in common_cols if c not in key_cols]
        for key in keys_a & keys_b:
            row_a = a_keyed.loc[key]
            row_b = b_keyed.loc[key]
            if isinstance(row_a, pd.DataFrame):
                row_a = row_a.iloc[0]
            if isinstance(row_b, pd.DataFrame):
                row_b = row_b.iloc[0]
            diffs = {}
            for col in value_cols:
                va = str(row_a[col]).strip()
                vb = str(row_b[col]).strip()
                if va != vb:
                    diffs[col] = f"{va} → {vb}"
            if diffs:
                row_info = {c: str(row_a[c]) for c in key_cols}
                row_info.update(diffs)
                modified_rows.append(row_info)
        if modified_rows:
            diff.modified = pd.DataFrame(modified_rows)

    return diff


def compare_access_files(path_a: str, path_b: str) -> CompareResult:
    """2つのAccessファイルを全テーブル照合する"""
    result = CompareResult()

    tables_a = list_tables(path_a)
    tables_b = list_tables(path_b)

    set_a = set(tables_a)
    set_b = set(tables_b)

    result.tables_only_in_a = sorted(set_a - set_b)
    result.tables_only_in_b = sorted(set_b - set_a)
    result.tables_common = sorted(set_a & set_b)

    if result.tables_only_in_a or result.tables_only_in_b:
        result.has_differences = True

    for table in result.tables_common:
        try:
            df_a = read_table(path_a, table)
            df_b = read_table(path_b, table)
        except RuntimeError:
            continue

        diff = compare_tables(df_a, df_b, table)
        has_diff = (
            len(diff.only_in_a) > 0
            or len(diff.only_in_b) > 0
            or len(diff.modified) > 0
            or diff.columns_only_in_a
            or diff.columns_only_in_b
        )
        if has_diff:
            result.has_differences = True
            result.table_diffs.append(diff)

    return result
