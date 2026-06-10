"""地籍調査DX帳票自動生成 - Streamlit Webアプリ

DXFファイルをアップロードすると、結線指示票と交点計算指示書を生成しダウンロードできる。

ローカル起動:
    streamlit run app.py

本番（VPS）起動:
    streamlit run app.py --server.port=8501 --server.address=127.0.0.1 --server.baseUrlPath=/chiseki
"""

import io
import os
import tempfile
from pathlib import Path

from streamlit_folium import st_folium
import folium
from folium.plugins import Draw

import pandas as pd
import streamlit as st

from src.dxf_parser import parse_dxf, get_block_numbers
from src.kessen_generator import generate_kessen
from src.kouten_generator import generate_kouten
from src.excel_writer import (
    write_kessen_excel,
    write_kouten_excel,
    KessenResult as ExcelKessenResult,
    IntersectionResult as ExcelIntersectionResult,
)
from src.access_compare import (
    compare_access_files,
    list_tables,
    read_table,
)


# ===== 認証 =====
AUTH_USER = os.environ.get("CHISEKI_USER", "kitaibachiseki")
AUTH_PASSWORD = os.environ.get("CHISEKI_PASSWORD", "tankachou")

# ===== テンプレートパス =====
PROJECT_ROOT = Path(__file__).parent
KESSEN_TEMPLATE = PROJECT_ROOT / "data" / "templates" / "結線指示票 - ブランク.xls"
KOUTEN_TEMPLATE = PROJECT_ROOT / "data" / "templates" / "交点計算指示書 - ブランク.xlsx"
UPLOADS_DIR = PROJECT_ROOT / "data" / "uploads"
DXF_UPLOADS_DIR = PROJECT_ROOT / "data" / "uploads" / "dxf"


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
    blocks = get_block_numbers(parsed)
    results["diagnostics"] = {
        "stakes_total": len(parsed.stakes),
        "stakes_named": len(named),
        "stakes_intersection": len(intersections),
        "lines_total": len(parsed.lines),
        "lines_matched": sum(1 for l in parsed.lines if l.stake1_number and l.stake2_number),
        "parcels": len(parsed.parcels),
        "public_lands": len(parsed.public_lands),
        "entity_counts": parsed.entity_counts,
        "block_numbers": blocks,
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


def page_dxf():
    """DXF帳票生成ページ"""
    st.title("DXF帳票自動生成")
    st.markdown("DXFファイルをアップロードすると、**結線指示票**と**交点計算指示書**を自動生成します。")

    # --- 保存済みDXFファイル表示 ---
    saved_dxf = _list_saved_dxf()
    if saved_dxf:
        with st.expander(f"保存済みDXFファイル ({len(saved_dxf)}件)", expanded=False):
            for p in saved_dxf:
                col_name, col_size = st.columns([3, 1])
                col_name.text(p.name)
                col_size.text(_format_file_size(p.stat().st_size))

    uploaded = st.file_uploader("DXFファイルを選択", type=["dxf", "DXF"], key="dxf_upload")

    # --- サーバー保存セクション ---
    if uploaded:
        st.divider()
        st.subheader("サーバーに保存")
        default_name = Path(uploaded.name).stem
        save_name = st.text_input(
            "保存ファイル名（変更可能）",
            value=default_name,
            key="dxf_save_name",
        )
        save_password = st.text_input(
            "パスワードを入力してください",
            type="password",
            key="dxf_save_password",
        )

        if st.button("サーバーに保存", key="btn_save_dxf"):
            if not save_name.strip():
                st.warning("ファイル名を入力してください")
            elif save_password != AUTH_PASSWORD:
                st.error("パスワードが正しくありません")
            else:
                saved_path = _save_uploaded_dxf(uploaded, save_name.strip())
                st.success(f"サーバーに保存しました: {saved_path.name}")
                st.rerun()

        st.divider()

    with st.expander("オプション設定"):
        col1, col2 = st.columns(2)
        with col1:
            gen_kessen = st.checkbox("結線指示票を生成", value=True, key="chk_kessen")
        with col2:
            gen_kouten = st.checkbox("交点計算指示書を生成", value=True, key="chk_kouten")

        st.caption("交点計算指示書はブロック番号を自動検出し、ブロックごとにシートを分割します")

    # 「帳票を生成」ボタンが押されたら処理して結果を session_state に保存
    if uploaded and st.button("帳票を生成", type="primary", key="btn_generate"):
        if not (gen_kessen or gen_kouten):
            st.warning("少なくとも1つの帳票を選択してください")
            return

        # 一時ファイルに保存して処理
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
                    None,
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
            c1.metric("杭(CIRCLE)", diag.get("stakes_total", 0))
            c2.metric("杭番号あり", diag.get("stakes_named", 0))
            c3.metric("交点杭", diag.get("stakes_intersection", 0))
            c4.metric("地番", diag.get("parcels", 0))
            c5, c6, c7, c8 = st.columns(4)
            c5.metric("境界線", diag.get("lines_total", 0))
            c6.metric("杭マッチ済み", diag.get("lines_matched", 0))
            c7.metric("公共用地", diag.get("public_lands", 0))

            blocks = diag.get("block_numbers", [])
            if blocks:
                st.caption(f"検出ブロック番号: {', '.join(blocks)}")

            ec = diag.get("entity_counts", {})
            if ec:
                st.caption("DXF内のエンティティ種別:")
                st.json(ec)

            if diag.get("stakes_total", 0) == 0:
                st.warning("**杭(CIRCLE)が検出されませんでした。**")
            elif diag.get("stakes_named", 0) == 0:
                st.warning("**杭番号が検出されませんでした。**")
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

        if st.button("結果をクリア", key="btn_clear"):
            st.session_state.pop("results", None)
            st.rerun()


def _format_file_size(num_bytes: int) -> str:
    """バイト数を人間可読形式に変換"""
    for unit in ["B", "KB", "MB", "GB"]:
        if num_bytes < 1024:
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= 1024
    return f"{num_bytes:.1f} TB"


def _save_uploaded_dxf(uploaded_file, save_name: str) -> Path:
    """アップロードされたDXFファイルをディスクに永続保存する"""
    DXF_UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
    # .dxf拡張子がなければ付与
    if not save_name.lower().endswith(".dxf"):
        save_name += ".dxf"
    dest = DXF_UPLOADS_DIR / save_name
    dest.write_bytes(uploaded_file.getvalue())
    return dest


def _list_saved_dxf() -> list[Path]:
    """保存済みDXFファイルの一覧を返す"""
    if not DXF_UPLOADS_DIR.exists():
        return []
    return sorted(
        p for p in DXF_UPLOADS_DIR.iterdir()
        if p.suffix.lower() == ".dxf"
    )


def _save_uploaded_access(uploaded_file, label: str) -> Path:
    """アップロードされたAccessファイルをディスクに永続保存する"""
    UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
    ext = uploaded_file.name.rsplit(".", 1)[-1].lower()
    safe_name = f"{label}_{uploaded_file.name}"
    dest = UPLOADS_DIR / safe_name
    dest.write_bytes(uploaded_file.getvalue())
    return dest


def _list_saved_files() -> list[Path]:
    """保存済みAccessファイルの一覧を返す"""
    if not UPLOADS_DIR.exists():
        return []
    return sorted(
        p for p in UPLOADS_DIR.iterdir()
        if p.suffix.lower() in (".mdb", ".accdb")
    )


def _build_compare_excel(result, label_a: str, label_b: str) -> bytes:
    """照合結果をExcelファイル(bytes)に変換する。

    シート構成:
      - 概要: テーブル構成の差分サマリー
      - テーブルごと: 差分があるテーブルごとに1シート
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        # --- 概要シート ---
        summary_rows = []
        for t in result.tables_common:
            summary_rows.append({"テーブル名": t, "状態": "両方に存在"})
        for t in result.tables_only_in_a:
            summary_rows.append({"テーブル名": t, "状態": f"{label_a}のみ"})
        for t in result.tables_only_in_b:
            summary_rows.append({"テーブル名": t, "状態": f"{label_b}のみ"})
        if summary_rows:
            pd.DataFrame(summary_rows).to_excel(writer, sheet_name="概要", index=False)

        # --- テーブル別差分シート ---
        for diff in result.table_diffs:
            sheet_name = diff.table_name[:31]  # Excel sheet name max 31 chars
            rows = []

            if diff.columns_only_in_a:
                rows.append({"区分": f"カラム({label_a}のみ)", "内容": ", ".join(diff.columns_only_in_a)})
            if diff.columns_only_in_b:
                rows.append({"区分": f"カラム({label_b}のみ)", "内容": ", ".join(diff.columns_only_in_b)})

            # Aのみのレコード
            if len(diff.only_in_a) > 0:
                df = diff.only_in_a.copy()
                df.insert(0, "区分", f"{label_a}のみ")
                rows_df = df
            else:
                rows_df = pd.DataFrame()

            # Bのみのレコード
            if len(diff.only_in_b) > 0:
                df = diff.only_in_b.copy()
                df.insert(0, "区分", f"{label_b}のみ")
                rows_df = pd.concat([rows_df, df], ignore_index=True) if len(rows_df) > 0 else df

            # 値が異なるレコード
            if len(diff.modified) > 0:
                df = diff.modified.copy()
                df.insert(0, "区分", "値が異なる")
                rows_df = pd.concat([rows_df, df], ignore_index=True) if len(rows_df) > 0 else df

            if len(rows_df) > 0:
                # カラム情報があれば先頭に追記
                if rows:
                    header_df = pd.DataFrame(rows)
                    # 列を揃える
                    for col in rows_df.columns:
                        if col not in header_df.columns:
                            header_df[col] = ""
                    for col in header_df.columns:
                        if col not in rows_df.columns:
                            rows_df[col] = ""
                    combined = pd.concat([header_df, rows_df], ignore_index=True)
                else:
                    combined = rows_df
                combined.to_excel(writer, sheet_name=sheet_name, index=False)
            elif rows:
                pd.DataFrame(rows).to_excel(writer, sheet_name=sheet_name, index=False)

    return buf.getvalue()


def page_terrain():
    """傾斜区分分析ページ"""
    from src.terrain_analysis import merge_dem_tiles, compute_polygon_stats, SLOPE_CLASSES
    import matplotlib.pyplot as plt
    import io as _io

    st.title("🗾 傾斜区分分析")
    st.markdown(
        "調査区域を地図上に描画すると、国土地理院の標高データから"
        "**傾斜区分割合**（平坦地〜急峻地）を自動算出します。"
    )

    st.info(
        "📌 地図上の「ポリゴン描画ツール」（左側ツールバーの五角形アイコン）で"
        "調査区域を手書きしてください。描画後「**傾斜を解析する**」ボタンを押します。"
    )

    # --- 地図 ---
    m = folium.Map(
        location=[36.0, 137.5],
        zoom_start=12,
        tiles="https://cyberjapandata.gsi.go.jp/xyz/pale/{z}/{x}/{y}.png",
        attr="国土地理院",
    )
    Draw(
        export=False,
        draw_options={
            "polyline": False,
            "rectangle": True,
            "polygon": True,
            "circle": False,
            "marker": False,
            "circlemarker": False,
        },
        edit_options={"edit": True, "remove": True},
    ).add_to(m)

    map_data = st_folium(m, width=700, height=500, key="terrain_map")

    # 描画済みポリゴンの取得
    drawn = None
    if map_data and map_data.get("all_drawings"):
        for feature in map_data["all_drawings"]:
            geom = feature.get("geometry", {})
            if geom.get("type") == "Polygon":
                drawn = geom["coordinates"][0]
                break

    if drawn:
        st.success(f"ポリゴンを検出しました（{len(drawn) - 1} 頂点）")
    else:
        st.warning("地図上にポリゴンを描画してください")
        return

    if not st.button("傾斜を解析する", type="primary"):
        return

    with st.spinner("国土地理院から標高データを取得中..."):
        try:
            lons = [c[0] for c in drawn]
            lats = [c[1] for c in drawn]
            bbox_area_deg2 = (max(lats) - min(lats)) * (max(lons) - min(lons))
            if bbox_area_deg2 > 0.25:
                st.error(
                    "選択エリアが広すぎます（約0.25°²以上）。"
                    "取得タイル数が多くなりすぎるため、より狭い範囲を選択してください。"
                )
                return

            dem, bounds = merge_dem_tiles(
                lat_min=min(lats),
                lon_min=min(lons),
                lat_max=max(lats),
                lon_max=max(lons),
            )
        except Exception as e:
            st.error(f"標高データの取得に失敗しました: {e}")
            return

    with st.spinner("傾斜を計算中..."):
        stats = compute_polygon_stats(dem, bounds, drawn)

    # --- 結果テーブル ---
    st.subheader("傾斜区分別 面積・割合")
    df = pd.DataFrame(stats)
    df.columns = ["傾斜区分", "面積 (m²)", "割合 (%)", "色"]
    df["面積 (ha)"] = (df["面積 (m²)"] / 10000).round(3)
    df = df[["傾斜区分", "面積 (ha)", "面積 (m²)", "割合 (%)"]]
    st.dataframe(df, use_container_width=True)

    # --- 円グラフ ---
    st.subheader("傾斜区分 割合グラフ")
    labels = [s["name"] for s in stats if s["percent"] > 0]
    sizes = [s["percent"] for s in stats if s["percent"] > 0]
    colors = [s["color"] for s in stats if s["percent"] > 0]

    if sizes:
        fig, ax = plt.subplots(figsize=(7, 4))
        ax.pie(
            sizes,
            labels=labels,
            colors=colors,
            autopct="%1.1f%%",
            startangle=90,
            pctdistance=0.8,
        )
        ax.axis("equal")
        ax.set_title("傾斜区分割合", fontsize=14)
        buf = _io.BytesIO()
        plt.tight_layout()
        plt.savefig(buf, format="png", dpi=120, bbox_inches="tight")
        buf.seek(0)
        st.image(buf, use_container_width=True)
        plt.close(fig)
    else:
        st.warning("有効なピクセルが見つかりませんでした。対象エリアを確認してください。")

    # --- メタ情報 ---
    north, south, west, east = bounds
    total_area_ha = sum(s["area_m2"] for s in stats) / 10000
    st.caption(
        f"解析エリア: N{north:.4f} S{south:.4f} W{west:.4f} E{east:.4f} | "
        f"DEM解像度: 5m or 10m (国土地理院) | "
        f"集計面積: {total_area_ha:.2f} ha"
    )


def page_access_compare():
    """土地データ照合ページ（Accessファイル差分検出）"""
    st.title("土地データ照合")
    st.markdown(
        "2台のPCで管理している**Accessファイル**をアップロードし、"
        "差分を検出します。アップロードしたファイルはサーバーに保存されます。"
    )

    # --- 保存済みファイル表示 ---
    saved = _list_saved_files()
    if saved:
        with st.expander(f"保存済みファイル ({len(saved)}件)", expanded=False):
            for p in saved:
                col_name, col_size, col_del = st.columns([3, 1, 1])
                col_name.text(p.name)
                col_size.text(_format_file_size(p.stat().st_size))

    # --- アップロード ---
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("ファイルA")
        file_a = st.file_uploader(
            "1台目のAccessファイルを選択",
            type=["mdb", "accdb", "MDB", "ACCDB"],
            key="access_a",
        )
        label_a = st.text_input("ファイルAのラベル（任意）", value="PC-1", key="label_a")

    with col2:
        st.subheader("ファイルB")
        file_b = st.file_uploader(
            "2台目のAccessファイルを選択",
            type=["mdb", "accdb", "MDB", "ACCDB"],
            key="access_b",
        )
        label_b = st.text_input("ファイルBのラベル（任意）", value="PC-2", key="label_b")

    if file_a and file_b:
        st.divider()
        st.subheader("アップロード済みファイル")

        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown(f"**{label_a}**")
            st.text(f"ファイル名: {file_a.name}")
            st.text(f"サイズ: {_format_file_size(file_a.size)}")
        with col_b:
            st.markdown(f"**{label_b}**")
            st.text(f"ファイル名: {file_b.name}")
            st.text(f"サイズ: {_format_file_size(file_b.size)}")

        st.divider()

        if st.button("照合を実行", type="primary", key="btn_compare"):
            # ファイルを保存
            path_a = _save_uploaded_access(file_a, label_a)
            path_b = _save_uploaded_access(file_b, label_b)

            try:
                with st.spinner("照合中..."):
                    result = compare_access_files(str(path_a), str(path_b))
                st.session_state["compare_result"] = result
                st.session_state["compare_label_a"] = label_a
                st.session_state["compare_label_b"] = label_b
                st.session_state["compare_path_a"] = str(path_a)
                st.session_state["compare_path_b"] = str(path_b)
            except Exception as e:
                st.error(f"照合中にエラーが発生しました: {e}")
                st.exception(e)

    elif file_a or file_b:
        st.warning("2つのファイルをアップロードしてください。")

    # --- 照合結果表示 ---
    if "compare_result" in st.session_state:
        result = st.session_state["compare_result"]
        la = st.session_state.get("compare_label_a", "A")
        lb = st.session_state.get("compare_label_b", "B")

        st.divider()
        if not result.has_differences:
            st.success("差分なし: 2つのファイルは同一です")
        else:
            st.warning("差分が検出されました")

        # Excel ダウンロード
        if result.has_differences:
            excel_bytes = _build_compare_excel(result, la, lb)
            st.download_button(
                "差分レポート.xlsx をダウンロード",
                excel_bytes,
                file_name="差分レポート.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="dl_compare_excel",
            )

        # テーブル概要
        st.subheader("テーブル構成")
        c1, c2, c3 = st.columns(3)
        c1.metric("共通テーブル", len(result.tables_common))
        c2.metric(f"{la}のみ", len(result.tables_only_in_a))
        c3.metric(f"{lb}のみ", len(result.tables_only_in_b))

        if result.tables_only_in_a:
            st.caption(f"{la}にのみ存在: {', '.join(result.tables_only_in_a)}")
        if result.tables_only_in_b:
            st.caption(f"{lb}にのみ存在: {', '.join(result.tables_only_in_b)}")

        # テーブルごとの差分
        if result.table_diffs:
            st.subheader("テーブル別の差分")
            for diff in result.table_diffs:
                with st.expander(f"テーブル: {diff.table_name}", expanded=True):
                    if diff.columns_only_in_a:
                        st.caption(f"{la}にのみ存在するカラム: {', '.join(diff.columns_only_in_a)}")
                    if diff.columns_only_in_b:
                        st.caption(f"{lb}にのみ存在するカラム: {', '.join(diff.columns_only_in_b)}")

                    if len(diff.only_in_a) > 0:
                        st.markdown(f"**{la}にのみ存在するレコード** ({len(diff.only_in_a)}件)")
                        st.dataframe(diff.only_in_a, use_container_width=True)

                    if len(diff.only_in_b) > 0:
                        st.markdown(f"**{lb}にのみ存在するレコード** ({len(diff.only_in_b)}件)")
                        st.dataframe(diff.only_in_b, use_container_width=True)

                    if len(diff.modified) > 0:
                        st.markdown(f"**値が異なるレコード** ({len(diff.modified)}件)")
                        st.dataframe(diff.modified, use_container_width=True)

        # 共通テーブルの内容プレビュー
        if result.tables_common and not result.table_diffs:
            st.subheader("共通テーブル一覧")
            for t in result.tables_common:
                st.text(f"  {t}")

        if st.button("照合結果をクリア", key="btn_clear_compare"):
            st.session_state.pop("compare_result", None)
            st.rerun()


def main():
    st.set_page_config(page_title="地籍調査DXツール", page_icon="📐", layout="centered")

    if not check_auth():
        return

    # ===== ヘッダー =====
    col_title, col_logout = st.columns([4, 1])
    with col_title:
        st.markdown("### 📐 地籍調査DXツール")
    with col_logout:
        if st.button("ログアウト"):
            st.session_state.clear()
            st.rerun()

    st.divider()

    # ===== タブで機能切り替え =====
    tab_dxf, tab_compare, tab_terrain = st.tabs([
        "📄 DXF帳票生成", "🔍 土地データ照合", "🗾 傾斜区分分析"
    ])

    with tab_dxf:
        page_dxf()

    with tab_compare:
        page_access_compare()

    with tab_terrain:
        page_terrain()


if __name__ == "__main__":
    main()
