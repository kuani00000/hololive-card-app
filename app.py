import pandas as pd
import streamlit as st
import re
import os

st.set_page_config(page_title="ホロライブ食玩図鑑", layout="wide")

# --- 行間・画像余白を極限までカットするCSS ---
st.markdown("""
    <style>
    /* 全体の縦方向の隙間を最小化 */
    [data-testid="stVerticalBlock"] > div {
        gap: 0px !important;
    }
    /* 列（カラム）要素の上下パディングを削除 */
    [data-testid="column"] {
        padding-top: 0px !important;
        padding-bottom: 0px !important;
    }
    /* 画像要素（stImage）周りの自動余白を完全にゼロ化 */
    [data-testid="stImage"] {
        margin: 0 !important;
        padding: 0 !important;
        line-height: 0 !important;
    }
    [data-testid="stImage"] img {
        margin: 0 !important;
        padding: 0 !important;
        display: block !important;
    }
    /* テキスト要素の余白カット */
    .stMarkdown p {
        margin: 0px !important;
        font-size: 13px !important;
        line-height: 1.1 !important;
    }
    /* 区切り線（hr）の上下余白をわずか1pxに強制作成 */
    hr {
        margin: 1px 0 !important;
        padding: 0 !important;
        border-color: #eee !important;
    }
    </style>
""", unsafe_allow_html=True)

# データ読み込み
def load_data():
    try:
        with open("hololive_cards.xlsx", "rb") as f:
            return pd.read_excel(f)
    except PermissionError:
        st.error("⚠️ `hololive_cards.xlsx` がExcel等で開かれています。ファイルを閉じてから再読み込みしてください。")
        st.stop()
    except FileNotFoundError:
        st.error("⚠️ `hololive_cards.xlsx` が見つかりません。")
        st.stop()

df = load_data()

# 選択肢データの生成
raw_members = df["メンバー名"].dropna().unique().tolist()
base_members = sorted(list(set([re.sub(r'[\d①-⑨]+$', '', str(m)).strip() for m in raw_members])))
member_options = ["-- メンバーを選択してください --"] + base_members

series_list = sorted(df["シリーズ名"].dropna().unique().tolist())
series_options = ["-- シリーズを選択してください --"] + series_list

# 極小行間のリスト表示関数
def display_card_list(card_df):
    if card_df.empty:
        st.info("該当するカードがありません。")
        return
    
    # テーブルヘッダー
    col1, col2, col3, col4, col5 = st.columns([3.5, 1.5, 2, 1.2, 1.5], vertical_alignment="center")
    with col1:
        st.markdown("**シリーズ名**")
    with col2:
        st.markdown("**No**")
    with col3:
        st.markdown("**メンバー名**")
    with col4:
        st.markdown("**画像**")
    with col5:
        st.markdown("**所持**")
    
    st.markdown("---")
    
    # データ行
    for idx, row in card_df.iterrows():
        c1, c2, c3, c4, c5 = st.columns([3.5, 1.5, 2, 1.2, 1.5], vertical_alignment="center")
        
        with c1:
            st.write(str(row.get('シリーズ名', '')))
        with c2:
            st.write(str(row.get('No', '')))
        with c3:
            st.write(str(row.get('メンバー名', '')))
            
        # 4列目: 画像（35px幅＋余白完全除去）
        with c4:
            img_filename = row.get("画像ファイル名")
            image_path = None
            if pd.notna(img_filename) and str(img_filename).strip() != "":
                image_path = os.path.join("images", str(img_filename).strip())
            
            if image_path and os.path.exists(image_path):
                st.image(image_path, width=35)
            else:
                st.caption("No Img")
                
        with c5:
            count = row.get('所持数', 0)
            if count > 0:
                st.markdown(f"✅ **{count}枚**")
            else:
                st.markdown("❌ 未所持")
                
        st.markdown("---")


# --- サイドバー ナビゲーション ---
st.sidebar.title("📌 メニュー")
menu = st.sidebar.radio("移動先を選択", ["ホーム", "🔍 カード検索"])

if menu == "ホーム":
    st.title("📦 ホロライブ食玩 コレクション図鑑")
    st.write("ホロライブ食玩の所持管理・検索アプリへようこそ！")
    st.info("👈 左側のサイドバーメニューから **「🔍 カード検索」** を選択してください。")

elif menu == "🔍 カード検索":
    st.title("🔍 カード検索")
    
    search_type = st.radio("検索タイプを選択してください", ["キャラ別検索", "シリーズ別検索"], horizontal=True)
    
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
        
        st.caption(f"全 {len(filtered_df)} 種類 | 所持: {len(owned_df)} 種類 ({filtered_df['所持数'].sum()}枚) | 未所持: {len(unowned_df)} 種類")

        tab1, tab2, tab3 = st.tabs(["すべて", f"✅ 所持 ({len(owned_df)})", f"❌ 未所持 ({len(unowned_df)})"])
        
        with tab1:
            display_card_list(filtered_df)
        with tab2:
            display_card_list(owned_df)
        with tab3:
            display_card_list(unowned_df)