import pandas as pd
import streamlit as st
import re
import os
import base64

st.set_page_config(page_title="ホロライブ食玩図鑑", layout="wide")

# --- CSSによる余白・スタイルの調整 ---
st.markdown("""
    <style>
    [data-testid="stVerticalBlock"] > div {
        gap: 0.15rem !important;
    }
    hr {
        margin: 0.5rem 0 !important;
        border-color: #eee !important;
    }
    </style>
""", unsafe_allow_html=True)

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

# 画像をBase64に変換してHTML埋め込み可能にする関数
def get_image_base64(image_path):
    if image_path and os.path.exists(image_path):
        try:
            with open(image_path, "rb") as img_file:
                b64 = base64.b64encode(img_file.read()).decode()
                ext = os.path.splitext(image_path)[1].lower().replace('.', '')
                if ext == 'jpg': ext = 'jpeg'
                return f"data:image/{ext};base64,{b64}"
        except Exception:
            return None
    return None

# レスポンシブ（ウィンドウ幅自動調整）グリッド表示関数
def display_card_list(card_df, show_trade_count=False):
    if card_df.empty:
        st.info("該当するカードがありません。")
        return
    
    # シリーズ名ごとにグループ化
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
            
            img_b64 = get_image_base64(image_path)
            
            if img_b64:
                img_tag = f'<img src="{img_b64}" style="width:80px; height:auto; border-radius:4px; display:block;" />'
            else:
                img_tag = '<div style="width:80px; height:80px; background:#f0f0f0; border-radius:4px; display:flex; align-items:center; justify-content:center; font-size:10px; color:#888;">No Img</div>'
            
            if show_trade_count:
                trade_qty = max(0, count - 1)
                sub_tag = f'<div style="font-size:11px; font-weight:bold; color:#e67e22; margin-top:2px; text-align:center;">{trade_qty}枚可</div>'
            else:
                sub_tag = f'<div style="font-size:10px; color:#666; margin-top:2px; text-align:center;">所持:{count} (開封:{opened_count})</div>'

            card_item = f'<div style="display:flex; flex-direction:column; align-items:center; width:80px; margin-bottom:6px;"><div style="font-weight:bold; font-size:12px; line-height:1.1; margin-bottom:2px; text-align:center; width:80px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="{full_name}">{disp_name}</div>{img_tag}{sub_tag}</div>'
            cards_html.append(card_item)
        
        container_html = f'<div style="display:flex; flex-wrap:wrap; gap:6px; align-items:flex-start;">{"".join(cards_html)}</div>'
        st.markdown(container_html, unsafe_allow_html=True)
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
    
    if df is None:
        st.warning("Excelデータが読み込まれていないため、カード検索機能を利用できません。")
    else:
        raw_members = df["メンバー名"].dropna().unique().tolist()
        base_members = sorted(list(set([re.sub(r'[\d①-⑨]+$', '', str(m)).strip() for m in raw_members])))
        member_options = ["-- メンバーを選択してください --"] + base_members

        series_list = sorted(df["シリーズ名"].dropna().unique().tolist())
        series_options = ["-- シリーズを選択してください --"] + series_list

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