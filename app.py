import pandas as pd
import streamlit as st
import re
import os
import base64
import io
from PIL import Image

st.set_page_config(page_title="ホロライブ食玩図鑑", layout="wide")

# --- CSS設定 ---
st.markdown("""
    <style>
    /* サイドバー（左フレーム）を完全非表示 */
    [data-testid="stSidebar"] {
        display: none !important;
    }
    
    /* メインコンテンツの下部余白 */
    .main .block-container {
        padding-bottom: 120px !important;
        padding-top: 1.5rem !important;
    }
    
    /* ボトムナビゲーション専用コンテナの固定表示 */
    .st-key-bottom_nav_container {
        position: fixed !important;
        bottom: 0 !important;
        left: 0 !important;
        width: 100% !important;
        background-color: rgba(255, 255, 255, 0.98) !important;
        backdrop-filter: blur(10px) !important;
        z-index: 99999 !important;
        border-top: 1px solid rgba(0, 0, 0, 0.1) !important;
        padding: 6px 8px 10px 8px !important;
        box-shadow: 0 -4px 16px rgba(0, 0, 0, 0.08) !important;
    }
    
    /* ボトムナビ内の項目を中央揃えで配置 */
    .st-key-bottom_nav_container div[data-testid="stRadio"] > div {
        display: flex !important;
        justify-content: center !important;
        align-items: center !important;
        gap: 8px !important;
        width: 100% !important;
        max-width: 480px !important;
        margin: 0 auto !important;
    }
    
    /* ボトムナビ内のラジオボタンの丸アイコンを非表示 */
    .st-key-bottom_nav_container div[data-testid="stRadio"] input[type="radio"],
    .st-key-bottom_nav_container div[data-testid="stRadio"] label > div:first-child {
        display: none !important;
    }
    
    /* ボタン風スタイル設定 */
    .st-key-bottom_nav_container div[data-testid="stRadio"] label {
        display: flex !important;
        flex-direction: column !important;
        align-items: center !important;
        justify-content: center !important;
        cursor: pointer !important;
        padding: 6px 4px !important;
        border-radius: 12px !important;
        transition: all 0.2s ease !important;
        margin: 0 !important;
        flex: 1 !important;
        max-width: 85px !important;
        min-width: 60px !important;
        background-color: transparent;
    }
    
    .st-key-bottom_nav_container div[data-testid="stRadio"] label:hover {
        background-color: #f0f2f5;
    }

    /* テキストおよび絵文字アイコンのスタイル */
    .st-key-bottom_nav_container div[data-testid="stRadio"] label p {
        font-size: 11px !important;
        font-weight: 600 !important;
        margin: 0 !important;
        padding: 0 !important;
        line-height: 1.2 !important;
        text-align: center !important;
        color: #555555 !important;
        white-space: pre-line !important;
        transition: color 0.2s ease;
    }
    
    /* 1行目（絵文字アイコン）のサイズ */
    .st-key-bottom_nav_container div[data-testid="stRadio"] label p::first-line {
        font-size: 22px !important;
        line-height: 1.25 !important;
    }
    
    /* 選択中（アクティブ）のメニューデザイン */
    .st-key-bottom_nav_container div[data-testid="stRadio"] label[aria-checked="true"] {
        background-color: #eef6ff !important;
        box-shadow: 0 2px 8px rgba(0, 122, 255, 0.15) !important;
    }

    .st-key-bottom_nav_container div[data-testid="stRadio"] label[aria-checked="true"] p {
        color: #007aff !important;
        font-weight: bold !important;
    }

    /* トレードグリッド内のボタンサイズをカード画像（75px）に固定 */
    div[data-testid="stColumn"] div[data-testid="stButton"] {
        display: flex !important;
        justify-content: center !important;
    }
    div[data-testid="stColumn"] div[data-testid="stButton"] > button {
        padding: 2px 4px !important;
        font-size: 11px !important;
        min-height: 26px !important;
        width: 75px !important;
        max-width: 75px !important;
    }

    /* コンテンツ間隔の微調整 */
    hr {
        margin: 0.8rem 0 !important;
        border-color: #eee !important;
    }
    
    /* 提供リスト（右カラム）の枠線スタイル */
    .offer-container {
        background-color: #f8f9fa;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 8px;
        margin-top: 4px;
    }
    </style>
""", unsafe_allow_html=True)

# セッション状態の初期化
if "offered_cards" not in st.session_state:
    st.session_state.offered_cards = []
if "generated_image_bytes" not in st.session_state:
    st.session_state.generated_image_bytes = None

# データ読み込み
def load_data():
    if not os.path.exists("hololive_cards.xlsx"):
        st.error("⚠️ `hololive_cards.xlsx` が見つかりません。`app.py` と同じフォルダに配置してください。")
        return None
    try:
        with open("hololive_cards.xlsx", "rb") as f:
            data = pd.read_excel(f)
            if "所持数" in data.columns:
                data["所持数"] = data["所持数"].fillna(0).astype(int)
            if "開封済" in data.columns:
                data["開封済"] = data["開封済"].fillna(0).astype(int)
            else:
                data["開封済"] = 0
            return data
    except PermissionError:
        st.error("⚠️ `hololive_cards.xlsx` がExcel等で開かれています。ファイルを閉じてから再読み込みしてください。")
        return None
    except Exception as e:
        st.error(f"⚠️ ファイルの読み込み中にエラーが発生しました: {e}")
        return None

df = load_data()

# 画像を正方形にトリミング＆軽量リサイズ（150x150）してBase64化する関数（スマホメモリ対策）
def get_cropped_square_base64(image_path):
    if image_path and os.path.exists(image_path):
        try:
            with Image.open(image_path) as img:
                img = img.convert("RGB")
                w, h = img.size
                
                # 正方形クロップ
                if w < h:
                    img = img.crop((0, 0, w, w))
                elif w > h:
                    left = (w - h) // 2
                    img = img.crop((left, 0, left + h, h))
                
                # ★スマホ描画制限対策：150x150pxへ軽量リサイズ
                img = img.resize((150, 150), Image.Resampling.LANCZOS)
                
                buffer = io.BytesIO()
                img.save(buffer, format="JPEG", quality=85)
                b64 = base64.b64encode(buffer.getvalue()).decode()
                return f"data:image/jpeg;base64,{b64}"
        except Exception:
            return None
    return None

# 提供リストのカード画像を1枚に結合する関数（元画像の解像度を自動取得して高画質出力）
def create_combined_offer_image(offered_list, cards_per_row=4):
    if not offered_list:
        return None
    
    # 元画像の解像度を自動取得（デフォルト標準値: 500px）
    cell_size = 500
    for item in offered_list:
        img_path = item.get("image_path")
        if img_path and os.path.exists(img_path):
            try:
                with Image.open(img_path) as img:
                    w, h = img.size
                    cell_size = min(w, h)  # 正方形トリミングの基準となる短辺サイズを採用
                    break
            except Exception:
                pass

    # 画像解像度に合わせた余白サイズの自動計算
    padding = max(8, int(cell_size * 0.04))
    
    num_cards = len(offered_list)
    cols = min(num_cards, cards_per_row)
    rows = (num_cards + cols - 1) // cols
    
    bg_w = cols * cell_size + (cols + 1) * padding
    bg_h = rows * cell_size + (rows + 1) * padding
    
    combined_img = Image.new("RGB", (bg_w, bg_h), color=(255, 255, 255))
    
    for idx, item in enumerate(offered_list):
        r = idx // cols
        c = idx % cols
        x = padding + c * (cell_size + padding)
        y = padding + r * (cell_size + padding)
        
        img_path = item.get("image_path")
        card_img = None
        
        if img_path and os.path.exists(img_path):
            try:
                with Image.open(img_path) as img:
                    img = img.convert("RGB")
                    w, h = img.size
                    if w < h:
                        img = img.crop((0, 0, w, w))
                    elif w > h:
                        left = (w - h) // 2
                        img = img.crop((left, 0, left + h, h))
                    card_img = img.resize((cell_size, cell_size), Image.Resampling.LANCZOS)
            except Exception:
                card_img = None
        
        if card_img is None:
            card_img = Image.new("RGB", (cell_size, cell_size), color=(235, 238, 242))
        
        combined_img.paste(card_img, (x, y))
    
    img_byte_arr = io.BytesIO()
    combined_img.save(img_byte_arr, format='JPEG', quality=95)
    return img_byte_arr.getvalue()


# レスポンシブ グリッド表示関数（所持リスト用 HTML描画）
def display_card_list(card_df, show_trade_count=False):
    if card_df.empty:
        st.info("該当するカードがありません。")
        return
    
    series_groups = card_df.groupby('シリーズ名', sort=False)
    
    for series_name, group_df in series_groups:
        total_types = len(group_df)
        owned_types = (group_df["所持数"] > 0).sum()
        collection_rate = (owned_types / total_types * 100) if total_types > 0 else 0.0
        
        st.markdown(f"#### 📦 {series_name} <span style='font-size:14px; color:#555; font-weight:normal;'>（収集率: {collection_rate:.1f}%）</span>", unsafe_allow_html=True)
        
        cards_html = []
        for idx, row in group_df.iterrows():
            full_name = str(row.get('メンバー名', ''))
            disp_name = full_name[:6]
            count = int(row.get('所持数', 0))
            opened_count = int(row.get('開封済', 0))
            
            img_filename = row.get("画像ファイル名")
            image_path = None
            if pd.notna(img_filename) and str(img_filename).strip() != "":
                filename = str(img_filename).strip()
                series_folder_path = os.path.join("images", str(series_name).strip(), filename)
                direct_path = os.path.join("images", filename)
                
                if os.path.exists(series_folder_path):
                    image_path = series_folder_path
                elif os.path.exists(direct_path):
                    image_path = direct_path
            
            img_b64 = get_cropped_square_base64(image_path)
            
            if img_b64:
                img_tag = f'<img src="{img_b64}" style="width:75px; height:75px; object-fit:cover; border-radius:4px; display:block;" />'
            else:
                img_tag = '<div style="width:75px; height:75px; background:#f0f0f0; border-radius:4px; display:flex; align-items:center; justify-content:center; font-size:10px; color:#888;">No Img</div>'
            
            if show_trade_count:
                trade_qty = max(0, count - 1)
                sub_tag = f'<div style="font-size:11px; font-weight:bold; color:#e67e22; margin-top:2px; text-align:center;">{trade_qty}枚可</div>'
            else:
                sub_tag = f'<div style="font-size:10px; color:#666; margin-top:2px; text-align:center;">所持:{count} (開封:{opened_count})</div>'

            card_item = f'<div style="display:flex; flex-direction:column; align-items:center; width:75px; margin-bottom:8px;"><div style="font-weight:bold; font-size:11px; line-height:1.1; margin-bottom:2px; text-align:center; width:75px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="{full_name}">{disp_name}</div>{img_tag}{sub_tag}</div>'
            cards_html.append(card_item)
        
        container_html = f'<div style="display:flex; flex-wrap:wrap; gap:8px; align-items:flex-start;">{"".join(cards_html)}</div>'
        st.markdown(container_html, unsafe_allow_html=True)
        st.markdown("---")


# トレード用カード表示（6列設定＆ボタンサイズ75px）
def display_trade_card_grid(card_df, is_tradeable=True, cols_per_row=6):
    if card_df.empty:
        st.info("該当するカードがありません。")
        return
    
    series_groups = card_df.groupby('シリーズ名', sort=False)
    
    for series_name, group_df in series_groups:
        st.markdown(f"##### 📦 {series_name}")
        
        rows_data = [group_df.iloc[i:i + cols_per_row] for i in range(0, len(group_df), cols_per_row)]
        
        for row_df in rows_data:
            cols = st.columns(cols_per_row)
            for idx, (_, row) in enumerate(row_df.iterrows()):
                with cols[idx]:
                    full_name = str(row.get('メンバー名', ''))
                    disp_name = full_name[:6]
                    count = int(row.get('所持数', 0))
                    trade_qty = max(0, count - 1)
                    img_filename = row.get("画像ファイル名")
                    
                    image_path = None
                    if pd.notna(img_filename) and str(img_filename).strip() != "":
                        filename = str(img_filename).strip()
                        series_folder_path = os.path.join("images", str(series_name).strip(), filename)
                        direct_path = os.path.join("images", filename)
                        
                        if os.path.exists(series_folder_path):
                            image_path = series_folder_path
                        elif os.path.exists(direct_path):
                            image_path = direct_path
                    
                    img_b64 = get_cropped_square_base64(image_path)
                    
                    # メンバー名
                    st.markdown(f'<div style="font-weight:bold; font-size:11px; line-height:1.1; margin:0 auto 2px auto; text-align:center; width:75px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="{full_name}">{disp_name}</div>', unsafe_allow_html=True)
                    
                    # カード画像 (75px * 75px)
                    if img_b64:
                        st.markdown(f'<img src="{img_b64}" style="width:75px; height:75px; object-fit:cover; border-radius:4px; display:block; margin:0 auto;" />', unsafe_allow_html=True)
                    else:
                        st.markdown('<div style="width:75px; height:75px; background:#f0f0f0; border-radius:4px; display:flex; align-items:center; justify-content:center; font-size:10px; color:#888; margin:0 auto;">No Img</div>', unsafe_allow_html=True)
                    
                    # 枚数 & 提供ボタン
                    if is_tradeable:
                        st.markdown(f'<div style="font-size:11px; font-weight:bold; color:#e67e22; margin-top:2px; text-align:center; width:75px; margin-left:auto; margin-right:auto;">{trade_qty}枚可</div>', unsafe_allow_html=True)
                        unique_btn_key = f"add_trade_{series_name}_{full_name}_{row.name}"
                        if st.button("➕提供", key=unique_btn_key, use_container_width=True):
                            card_item = {
                                "series": series_name,
                                "name": full_name,
                                "image_path": image_path,
                                "trade_qty": trade_qty
                            }
                            st.session_state.offered_cards.append(card_item)
                            st.session_state.generated_image_bytes = None
                            st.rerun()
                    else:
                        st.markdown('<div style="font-size:10px; color:#666; margin-top:2px; text-align:center; width:75px; margin-left:auto; margin-right:auto;">未所持</div>', unsafe_allow_html=True)
        st.markdown("---")


# --- 画面下部 固定ナビゲーションメニュー ---
with st.container(key="bottom_nav_container"):
    menu = st.radio(
        "",
        ["🏠\nホーム", "📊\n収集率", "🎴\n所持リスト", "🔄\nトレード"],
        horizontal=True,
        label_visibility="collapsed",
        key="main_bottom_navigation"
    )


# --- メニュー別 画面切替 ---

# 1. ホーム画面
if menu == "🏠\nホーム":
    st.title("📦 ホロライブ食玩 コレクション図鑑")
    st.write("ホロライブ食玩カード・シールの所持管理・検索アプリへようこそ！")
    st.info("👇 画面下のメニューから「📊 収集率」や「🎴 所持リスト」を選択してください。")
    
    if df is not None:
        total_types = len(df)
        owned_types = (df["所持数"] > 0).sum()
        total_cards = df["所持数"].sum()
        total_opened = df["開封済"].sum()
        
        st.markdown("### 📈 コレクション概要")
        col1, col2 = st.columns(2)
        with col1:
            st.metric("収集した種類数", f"{owned_types} / {total_types} 種")
        with col2:
            st.metric("総所持枚数", f"{total_cards} 枚", f"内 開封済 {total_opened} 枚")


# 2. 収集率一覧画面
elif menu == "📊\n収集率":
    st.title("📊 シリーズ別 収集率")
    
    if df is None:
        st.warning("Excelデータが読み込まれていないため、収集率を表示できません。")
    else:
        total_all_types = len(df)
        total_owned_types = (df["所持数"] > 0).sum()
        total_rate = (total_owned_types / total_all_types * 100) if total_all_types > 0 else 0.0
        
        st.markdown(f"### 🌐 全体コンプリート率: **{total_rate:.1f}%** ({total_owned_types}/{total_all_types}種類)")
        st.progress(total_rate / 100.0)
        st.markdown("---")
        
        series_groups = df.groupby('シリーズ名', sort=False)
        for series_name, group_df in series_groups:
            total_types = len(group_df)
            owned_types = (group_df["所持数"] > 0).sum()
            unowned_types = total_types - owned_types
            rate = (owned_types / total_types * 100) if total_types > 0 else 0.0
            
            s_owned_cards = group_df["所持数"].sum()
            s_opened_cards = group_df["開封済"].sum()

            first_row = group_df.iloc[0]
            img_filename = first_row.get("画像ファイル名")
            image_path = None
            if pd.notna(img_filename) and str(img_filename).strip() != "":
                filename = str(img_filename).strip()
                series_folder_path = os.path.join("images", str(series_name).strip(), filename)
                direct_path = os.path.join("images", filename)
                
                if os.path.exists(series_folder_path):
                    image_path = series_folder_path
                elif os.path.exists(direct_path):
                    image_path = direct_path
            
            img_b64 = get_cropped_square_base64(image_path)

            col_img, col_info = st.columns([1, 4])
            with col_img:
                if img_b64:
                    st.markdown(f'<img src="{img_b64}" style="width:75px; height:75px; object-fit:cover; border-radius:6px; display:block; margin-top:4px;" />', unsafe_allow_html=True)
                else:
                    st.markdown('<div style="width:75px; height:75px; background:#f0f0f0; border-radius:6px; display:flex; align-items:center; justify-content:center; font-size:10px; color:#888; margin-top:4px;">No Img</div>', unsafe_allow_html=True)
            
            with col_info:
                st.markdown(f"#### 📦 {series_name}")
                st.markdown(f"**収集率: {rate:.1f}%** （{owned_types} / {total_types} 種類）")
                st.caption(f"所持枚数: {s_owned_cards}枚 (開封:{s_opened_cards}枚) | 未所持: {unowned_types}種類")
                st.progress(rate / 100.0)
            
            st.markdown("---")


# 3. 所持リスト（カード検索画面）
elif menu == "🎴\n所持リスト":
    st.title("🎴 所持リスト・カード検索")
    
    if df is None:
        st.warning("Excelデータが読み込まれていないため、カード検索機能を利用できません。")
    else:
        raw_members = df["メンバー名"].dropna().unique().tolist()
        base_members = sorted(list(set([re.sub(r'[\d①-⑨]+$', '', str(m)).strip() for m in raw_members])))
        member_options = ["-- メンバーを選択してください --"] + base_members

        series_list = sorted(df["シリーズ名"].dropna().unique().tolist())
        series_options = ["-- シリーズを選択してください --"] + series_list

        search_type = st.radio("検索タイプを選択してください", ["キャラ別検索", "シリーズ別検索"], horizontal=True, key="card_search_type")
        
        filtered_df = pd.DataFrame()
        title_text = ""

        if search_type == "キャラ別検索":
            selected_name = st.selectbox("メンバーを選択してください", member_options)
            if selected_name != "-- メンバーを選択してください --":
                filtered_df = df[df["メンバー名"].astype(str).str.contains(selected_name, na=False)].copy()
                title_text = f"「{selected_name}」のカード一覧"

        else:
            selected_series = st.selectbox("シリーズを選択してください", series_options)
            if selected_series != "-- シリーズを選択してください --":
                filtered_df = df[df["シリーズ名"] == selected_series].copy()
                title_text = f"「{selected_series}」のカード一覧"

        if not filtered_df.empty:
            st.subheader(title_text)
            
            owned_df = filtered_df[filtered_df["所持数"] > 0]
            unowned_df = filtered_df[filtered_df["所持数"] == 0]
            tradeable_df = filtered_df[filtered_df["所持数"] >= 2]
            
            total_owned_qty = filtered_df['所持数'].sum()
            total_opened_qty = filtered_df['開封済'].sum()
            
            st.caption(f"全 {len(filtered_df)} 種類 | 所持: {len(owned_df)} 種類 ({total_owned_qty}枚 / 内 開封済: {total_opened_qty}枚) | 未所持: {len(unowned_df)} 種類 | トレード可: {len(tradeable_df)} 種類")

            tab1, tab2, tab3, tab4 = st.tabs([
                "すべて", 
                f"✅ 所持 ({len(owned_df)})", 
                f"❌ 未所持 ({len(unowned_df)})", 
                f"🔄 トレード可 ({len(tradeable_df)})"
            ])
            
            with tab1:
                display_card_list(filtered_df)
            with tab2:
                display_card_list(owned_df)
            with tab3:
                display_card_list(unowned_df)
            with tab4:
                display_card_list(tradeable_df, show_trade_count=True)


# 4. トレード画面
elif menu == "🔄\nトレード":
    st.title("🔄 トレード管理")
    
    if df is None:
        st.warning("Excelデータが読み込まれていないため、トレード機能を利用できません。")
    else:
        # 左（検索エリア）と右（提供リスト）の配置
        col_search, col_offer = st.columns([2.2, 1.0], gap="large")

        # --- 左カラム: 検索エリア ---
        with col_search:
            st.subheader("🔍 カード検索")
            raw_members = df["メンバー名"].dropna().unique().tolist()
            base_members = sorted(list(set([re.sub(r'[\d①-⑨]+$', '', str(m)).strip() for m in raw_members])))
            member_options = ["-- メンバーを選択してください --"] + base_members

            series_list = sorted(df["シリーズ名"].dropna().unique().tolist())
            series_options = ["-- シリーズを選択してください --"] + series_list

            search_type = st.radio("検索タイプを選択してください", ["キャラ別検索", "シリーズ別検索"], horizontal=True, key="trade_search_type")
            
            filtered_df = pd.DataFrame()
            title_text = ""

            if search_type == "キャラ別検索":
                selected_name = st.selectbox("メンバーを選択してください", member_options, key="trade_member_select")
                if selected_name != "-- メンバーを選択してください --":
                    filtered_df = df[df["メンバー名"].astype(str).str.contains(selected_name, na=False)].copy()
                    title_text = f"「{selected_name}」の対象カード"

            else:
                selected_series = st.selectbox("シリーズを選択してください", series_options, key="trade_series_select")
                if selected_series != "-- シリーズを選択してください --":
                    filtered_df = df[df["シリーズ名"] == selected_series].copy()
                    title_text = f"「{selected_series}」の対象カード"

            if not filtered_df.empty:
                st.write(f"### {title_text}")
                
                tradeable_df = filtered_df[filtered_df["所持数"] >= 2].copy()
                unowned_df = filtered_df[filtered_df["所持数"] == 0].copy()

                tab1, tab2 = st.tabs([
                    f"🔄 トレード可能 ({len(tradeable_df)})", 
                    f"❌ 未所持 ({len(unowned_df)})"
                ])
                
                with tab1:
                    st.info("💡 「➕提供」ボタンをクリックすると右側の「提供リスト」に追加されます。")
                    display_trade_card_grid(tradeable_df, is_tradeable=True, cols_per_row=6)
                
                with tab2:
                    st.info("💡 トレードで探している未所持カードの一覧です。")
                    display_trade_card_grid(unowned_df, is_tradeable=False, cols_per_row=6)

        # --- 右カラム: 提供リスト＆画像生成エリア ---
        with col_offer:
            st.subheader("📋 提供リスト")
            offered_list = st.session_state.offered_cards
            
            if offered_list:
                st.caption(f"選択中: **{len(offered_list)}** 枚")
                
                btn_col1, btn_col2 = st.columns(2)
                with btn_col1:
                    if st.button("🖼️ 画像生成", use_container_width=True, key="btn_gen_img"):
                        st.session_state.generated_image_bytes = create_combined_offer_image(offered_list)
                
                with btn_col2:
                    if st.button("🗑️ リスト削除", use_container_width=True, key="clear_offered_cards"):
                        st.session_state.offered_cards = []
                        st.session_state.generated_image_bytes = None
                        st.rerun()

                # 合成画像のプレビュー＆ダウンロード（表示はカラム幅へ縮小フィット）
                if st.session_state.generated_image_bytes is not None:
                    st.markdown("---")
                    st.markdown("##### 📸 出力画像プレビュー")
                    st.image(st.session_state.generated_image_bytes, use_container_width=True)
                    st.download_button(
                        label="💾 画像を保存 (JPG)",
                        data=st.session_state.generated_image_bytes,
                        file_name="trade_offer_cards.jpg",
                        mime="image/jpeg",
                        use_container_width=True
                    )
                    st.markdown("---")

                st.markdown('<div class="offer-container">', unsafe_allow_html=True)
                
                for idx, item in enumerate(offered_list):
                    c_img, c_info, c_del = st.columns([1.2, 1.8, 0.8])
                    
                    with c_img:
                        img_b64 = get_cropped_square_base64(item.get("image_path"))
                        if img_b64:
                            st.markdown(f'<img src="{img_b64}" style="width:75px; height:75px; object-fit:cover; border-radius:4px; display:block;" />', unsafe_allow_html=True)
                        else:
                            st.markdown('<div style="width:75px; height:75px; background:#f0f0f0; border-radius:4px; display:flex; align-items:center; justify-content:center; font-size:10px; color:#888;">No Img</div>', unsafe_allow_html=True)
                    
                    with c_info:
                        st.markdown(f"**{item.get('name')}**")
                        st.caption(f"{item.get('series')}")
                    
                    with c_del:
                        if st.button("❌", key=f"del_offered_{idx}"):
                            st.session_state.offered_cards.pop(idx)
                            st.session_state.generated_image_bytes = None
                            st.rerun()
                    
                    st.markdown("<hr style='margin: 6px 0 !important;'>", unsafe_allow_html=True)
                
                st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.info("👈 左側一覧で「➕提供」を押すと、ここに提供用カードが追加されます。")
