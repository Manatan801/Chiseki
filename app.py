"""地籍調査DX帳票自動生成 - Streamlit Webアプリ

DXFファイルをアップロードすると、結線指示票と交点計算指示書を生成しダウンロードできる。

ローカル起動:
    streamlit run app.py

本番（VPS）起動:
    streamlit run app.py --server.port=8501 --server.address=127.0.0.1 --server.baseUrlPath=/chiseki
"""

import os
import tempfile
from pathlib import Path

import streamlit as st

from src.dxf_parser import parse_dxf
from src.kessen_generator import generate_kessen
from src.kouten_generator import generate_kouten
from src.excel_writer import (
    write_kessen_excel,
    write_kouten_excel,
    KessenResult as ExcelKessenResult,
    IntersectionResult as ExcelIntersectionResult,
)


# ===== 認証 =====
AUTH_USER = os.environ.get("CHISEKI_USER", "kitaibachiseki")
AUTH_PASSWORD = os.environ.get("CHISEKI_PASSWORD", "tankachou")

# ===== テンプレートパス =====
PROJECT_ROOT = Path(__file__).parent
KESSEN_TEMPLATE = PROJECT_ROOT / "data" / "templates" / "結線指示票 - ブランク.xls"
KOUTEN_TEMPLATE = PROJECT_ROOT / "data" / "templates" / "交点計算指示書 - ブランク.xlsx"


def check_auth() -> bool:
    """セッションベースの簡易認証。認証済みならTrueを返す。"""
    if st.session_state.get("authenticated"):
        return True

    st.title("地籍調査DX帳票自動生成")
    st.caption("ログインが必要です")

    with st.form("login"):
        user = st.text_input("ユーザーID")
        password = st.text_input("パスワード", type="password")
        submitted = st.form_submit_button("ログイン")

    if submitted:
        if user == AUTH_USER and password == AUTH_PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("ユーザーIDまたはパスワードが違います")
    return False


def process_dxf(dxf_path: str, parcels_filter, generate_kessen_flag: bool, generate_kouten_flag: bool, block_number):
    """DXFを処理し、生成されたExcelファイルのパスを返す。"""
    out_dir = Path(tempfile.mkdtemp(prefix="chiseki_"))
    results = {"kessen_path": None, "kouten_path": None, "kessen_count": 0, "kouten_count": 0, "kessen_details": [], "kouten_details": []}

    parsed = parse_dxf(dxf_path)

    # 診断情報
    named = parsed.get_stakes_with_numbers()
    intersections = [s for s in named if s.is_intersection]
    results["diagnostics"] = {
        "stakes_total": len(parsed.stakes),
        "stakes_named": len(named),
        "stakes_intersection": len(intersections),
        "lines_total": len(parsed.lines),
        "lines_matched": sum(1 for l in parsed.lines if l.stake1_number and l.stake2_number),
        "parcels": len(parsed.parcels),
        "public_lands": len(parsed.public_lands),
        "entity_counts": parsed.entity_counts,
    }

    if generate_kessen_flag:
        kessen_results = generate_kessen(parsed, parcels_filter)
        if kessen_results:
            excel_results = [
                ExcelKessenResult(
                    parcel_number=r.parcel_number,
                    stake_sequence=r.stake_sequence,
                    land_use=r.land_use,
                    owner=r.owner,
                )
                for r in kessen_results
            ]
            output_path = out_dir / "結線指示票.xls"
            write_kessen_excel(excel_results, str(KESSEN_TEMPLATE), str(output_path))
            results["kessen_path"] = output_path
            results["kessen_count"] = len(kessen_results)
            results["kessen_details"] = [
                f"{r.parcel_number}: {len(r.stake_sequence)}杭" for r in kessen_results
            ]

    if generate_kouten_flag:
        kouten_results = generate_kouten(parsed, block_number=block_number)
        if kouten_results:
            excel_results = [
                ExcelIntersectionResult(
                    intersection_stake=r.intersection_stake,
                    baseline_point1=r.baseline_point1,
                    baseline_point2=r.baseline_point2,
                    extension_point1=r.extension_point1,
                    extension_point2=r.extension_point2,
                )
                for r in kouten_results
            ]
            output_path = out_dir / "交点計算指示書.xlsx"
            write_kouten_excel(excel_results, str(KOUTEN_TEMPLATE), str(output_path))
            results["kouten_path"] = output_path
            results["kouten_count"] = len(kouten_results)
            results["kouten_details"] = [
                f"交点={r.intersection_stake} 基準線=({r.baseline_point1},{r.baseline_point2}) 延長線=({r.extension_point1},{r.extension_point2})"
                for r in kouten_results
            ]

    return results


def main():
    st.set_page_config(page_title="地籍調査DX帳票自動生成", page_icon="📐", layout="centered")

    if not check_auth():
        return

    # ===== 認証済み: メインUI =====
    col_title, col_logout = st.columns([4, 1])
    with col_title:
        st.title("地籍調査DX帳票自動生成")
    with col_logout:
        if st.button("ログアウト"):
            st.session_state.authenticated = False
            st.rerun()

    st.markdown("DXFファイルをアップロードすると、**結線指示票**と**交点計算指示書**を自動生成します。")

    uploaded = st.file_uploader("DXFファイルを選択", type=["dxf", "DXF"])

    with st.expander("オプション設定"):
        col1, col2 = st.columns(2)
        with col1:
            gen_kessen = st.checkbox("結線指示票を生成", value=True)
        with col2:
            gen_kouten = st.checkbox("交点計算指示書を生成", value=True)

        block_number = st.number_input(
            "交点計算ブロック番号（0で全ブロック）",
            min_value=0, value=0, step=1,
        )

    # 「帳票を生成」ボタンが押されたら処理して結果を session_state に保存
    if uploaded and st.button("帳票を生成", type="primary"):
        if not (gen_kessen or gen_kouten):
            st.warning("少なくとも1つの帳票を選択してください")
            return

        # 一時ファイルに保存
        with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf") as tmp:
            tmp.write(uploaded.getvalue())
            tmp_path = tmp.name

        try:
            with st.spinner("DXFを解析中..."):
                results = process_dxf(
                    tmp_path,
                    None,
                    gen_kessen,
                    gen_kouten,
                    block_number if block_number > 0 else None,
                )

            # ファイルパスではなくバイトデータとして保存（一時ファイル削除後も使える）
            if results.get("kessen_path"):
                with open(results["kessen_path"], "rb") as f:
                    results["kessen_bytes"] = f.read()
            if results.get("kouten_path"):
                with open(results["kouten_path"], "rb") as f:
                    results["kouten_bytes"] = f.read()

            st.session_state["results"] = results
            st.session_state["gen_kessen"] = gen_kessen
            st.session_state["gen_kouten"] = gen_kouten

        except Exception as e:
            st.error(f"処理中にエラーが発生しました: {e}")
            st.exception(e)
            st.session_state.pop("results", None)
        finally:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass

    # 結果表示（session_state に結果があれば毎回表示する）
    if "results" in st.session_state:
        results = st.session_state["results"]
        gen_kessen = st.session_state.get("gen_kessen", True)
        gen_kouten = st.session_state.get("gen_kouten", True)

        st.success("生成完了")

        # 診断情報
        diag = results.get("diagnostics", {})
        with st.expander("DXF解析結果（診断情報）", expanded=True):
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("杭(CIRCLE r=0.25)", diag.get("stakes_total", 0))
            c2.metric("杭番号あり", diag.get("stakes_named", 0))
            c3.metric("交点杭", diag.get("stakes_intersection", 0))
            c4.metric("地番", diag.get("parcels", 0))
            c5, c6, c7, c8 = st.columns(4)
            c5.metric("境界線(2点)", diag.get("lines_total", 0))
            c6.metric("杭マッチ済み", diag.get("lines_matched", 0))
            c7.metric("公共用地", diag.get("public_lands", 0))

            ec = diag.get("entity_counts", {})
            if ec:
                st.caption("DXF内のエンティティ種別:")
                st.json(ec)

            if diag.get("stakes_total", 0) == 0:
                st.warning("**杭(CIRCLE)が検出されませんでした。** 杭は半径0.25のCIRCLEで描画する必要があります。")
            elif diag.get("stakes_named", 0) == 0:
                st.warning("**杭番号(MTEXT)が検出されませんでした。**")
            elif diag.get("parcels", 0) == 0:
                st.warning("**地番が検出されませんでした。**")

        if results.get("kessen_bytes"):
            st.subheader(f"結線指示票 ({results['kessen_count']}筆)")
            st.download_button(
                "結線指示票.xls をダウンロード",
                results["kessen_bytes"],
                file_name="結線指示票.xls",
                mime="application/vnd.ms-excel",
                key="dl_kessen",
            )
            with st.expander("生成内容"):
                for d in results["kessen_details"]:
                    st.text(d)
        elif gen_kessen:
            st.info("結線指示票: 対象の地番が見つかりませんでした")

        if results.get("kouten_bytes"):
            st.subheader(f"交点計算指示書 ({results['kouten_count']}交点)")
            st.download_button(
                "交点計算指示書.xlsx をダウンロード",
                results["kouten_bytes"],
                file_name="交点計算指示書.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_kouten",
            )
            with st.expander("生成内容"):
                for d in results["kouten_details"]:
                    st.text(d)
        elif gen_kouten:
            st.info("交点計算指示書: 交点杭が見つかりませんでした")

        if st.button("結果をクリア"):
            st.session_state.pop("results", None)
            st.rerun()


if __name__ == "__main__":
    main()
