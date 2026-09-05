
import streamlit as st
import pandas as pd
import numpy as np
import os
import json
import uuid
from openpyxl import load_workbook
from collections import Counter, defaultdict, deque
from io import BytesIO
from openpyxl.styles import Border
import re

try:
    from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
    AGGRID_AVAILABLE = True
except ImportError:
    AGGRID_AVAILABLE = False


if "bulk_edit_rows" not in st.session_state:
    st.session_state.bulk_edit_rows = 1
if "grid_revision" not in st.session_state:
    st.session_state["grid_revision"] = 0

# =========================
# 税区分マスタを常にロード
# =========================
TAX_CONFIG_FILE = "tax_config.csv"
TAX_RATE_FILE = "tax_rate_config.csv"

if "enabled_tax_list" not in st.session_state:
    if os.path.exists(TAX_CONFIG_FILE):
        df = pd.read_csv(TAX_CONFIG_FILE)
        st.session_state["enabled_tax_list"] = df["税区分"].tolist()
    else:
        st.session_state["enabled_tax_list"] = []

if "enabled_tax_rates" not in st.session_state:
    if os.path.exists(TAX_RATE_FILE):
        df = pd.read_csv(TAX_RATE_FILE)
        st.session_state["enabled_tax_rates"] = df["税率"].tolist()
    else:
        st.session_state["enabled_tax_rates"] = []

if not st.session_state.get("show_settings"):
    st.title("🧮 freee会計 / 収入・支出CSV作成ツール")
# --- 設定フォルダの準備 ---
CONFIG_DIR = "mapping_configs"

# =========================
# 初期設定トグル（最上部）
# =========================
# 初回起動判定


ACCOUNT_MASTER_FILE = "master_accounts.csv"

# 初回起動判定
if "show_settings" not in st.session_state:
    if not os.path.exists(ACCOUNT_MASTER_FILE):
        st.session_state["show_settings"] = True
    else:
        st.session_state["show_settings"] = False
st.sidebar.markdown("### ⚙ マスタ設定（勘定科目・税区分・取引先など）")

toggle_label = "📄 メイン画面へ戻る" if st.session_state["show_settings"] else "⚙ マスタ設定を開く"

if st.sidebar.button(toggle_label):
    st.session_state["show_settings"] = not st.session_state["show_settings"]
    st.rerun()

st.sidebar.markdown("---")
os.makedirs(CONFIG_DIR, exist_ok=True)

# --- ページ設定（最初に一度だけ）---
st.set_page_config(page_title="Excel集計ツール", layout="wide")


if "phase" not in st.session_state:
    st.session_state["phase"] = "input"

st.markdown("""
<style>
div.stDownloadButton > button {
    min-width: 220px;
    white-space: nowrap;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<style>
/* sort_items / multiselect 共通のタグ色（赤→緑） */
div[data-baseweb="tag"] {
    background-color: #c6f6d5 !important;  /* 薄い緑 */
    color: #1f2937 !important;
    border: 1px solid #48bb78 !important;
}

/* ホバー時 */
div[data-baseweb="tag"]:hover {
    background-color: #9ae6b4 !important;
}

/* ドラッグ中（sort_items） */
div[data-baseweb="tag"][aria-grabbed="true"] {
    background-color: #68d391 !important;
}
</style>
""", unsafe_allow_html=True)

# --- サイドバー：マッピング設定永続化 ---
st.sidebar.subheader("🔄 マッピング設定")

# 再実行後も設定読込の完了結果を表示する。
if st.session_state.get("config_load_message"):
    st.sidebar.success(st.session_state.pop("config_load_message"))

# JSONアップロードからの読み込み
mapping_file = st.sidebar.file_uploader("設定JSONをアップロード", type="json", key="upload_json")
if not mapping_file:
    st.session_state.pop("loaded_mapping_file_id", None)

mapping_file_id = f"{mapping_file.name}:{mapping_file.size}" if mapping_file else None
if mapping_file and st.session_state.get("loaded_mapping_file_id") != mapping_file_id:
    try:
        loaded = json.load(mapping_file)
        # 保存時の任意項目とマッピングだけを復元し、前回選択の残りを持ち越さない。
        saved_optional_fields = loaded.pop("__selected_optional_fields__", None)
        if saved_optional_fields is None:
            st.session_state.pop("selected_optional_fields", None)
        else:
            st.session_state["selected_optional_fields"] = saved_optional_fields
        for key in [key for key in st.session_state if key.startswith("map_")]:
            del st.session_state[key]
        saved_mapping_names = loaded.pop("__mapping_header_names__", None)
        if saved_mapping_names is not None:
            st.session_state["pending_mapping_names"] = saved_mapping_names
        else:
            for std_col, sel in loaded.items():
                st.session_state[f"map_{std_col}"] = sel
        st.session_state["loaded_mapping_file_id"] = mapping_file_id
        st.sidebar.success("JSON設定を読み込みました")
    except Exception as e:
        st.sidebar.error(f"JSON読み込み失敗: {e}")

# 既存設定の選択＆削除
config_files = [f for f in os.listdir(CONFIG_DIR) if f.endswith(".json")]
config_names = [""] + [os.path.splitext(f)[0] for f in config_files]



col1, col_load, col_delete = st.sidebar.columns([3, 1, 1], gap="small")
selected_config = col1.selectbox("既存設定を選択", config_names, key="select_config")
if selected_config:
    config_path = os.path.join(CONFIG_DIR, f"{selected_config}.json")
    if col_load.button("読込", key=f"load_{selected_config}"):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            # 保存時の任意項目とマッピングだけを復元し、前回選択の残りを持ち越さない。
            saved_optional_fields = loaded.pop("__selected_optional_fields__", None)
            if saved_optional_fields is None:
                st.session_state.pop("selected_optional_fields", None)
            else:
                st.session_state["selected_optional_fields"] = saved_optional_fields
            for key in [key for key in st.session_state if key.startswith("map_")]:
                del st.session_state[key]
            saved_mapping_names = loaded.pop("__mapping_header_names__", None)
            if saved_mapping_names is not None:
                st.session_state["pending_mapping_names"] = saved_mapping_names
            else:
                for std_col, sel in loaded.items():
                    st.session_state[f"map_{std_col}"] = sel
            # Excelのアップロード状態を失わないよう、再実行せず同じ画面で設定を反映する。
            st.sidebar.success(f"設定 '{selected_config}' を読み込みました")
        except Exception as e:
            st.sidebar.error(f"設定読み込み失敗: {e}")
    if col_delete.button("🗑", key=f"del_{selected_config}", help="この設定を削除"):
        try:
            os.remove(config_path)
            st.sidebar.success(f"設定 '{selected_config}' を削除しました")
            st.rerun()
        except Exception as e:
            st.sidebar.error(f"削除失敗: {e}")

st.sidebar.markdown("---")

# 新規設定名入力＋保存
new_name = st.sidebar.text_input("新規設定名", key="new_config_name")
if st.sidebar.button("設定を保存", key="save_button"):
    if new_name.strip():
        # 列IDではなく列名を保存するため、ヘッダー検出後に保存する。
        st.session_state["pending_config_save"] = new_name.strip()
    else:
        st.sidebar.error("設定名を入力してください")

st.sidebar.markdown("---")

# =========================
# 初期設定画面
# =========================
if st.session_state.get("show_settings"):

    st.title("⚙ マスタ設定")

    st.subheader("📘 (必須)勘定科目マスタの設定");uploaded_account_csv=st.file_uploader("free会計の勘定科目をエクスポートし、CSVをアップロードしてください",type=["csv"],key="account_csv"); 
    if uploaded_account_csv:
        try:
            uploaded_account_csv.seek(0)

            try:
                df_raw = pd.read_csv(
                    uploaded_account_csv,
                    header=None,
                    encoding="utf-8",
                )
            except UnicodeDecodeError:
                # 文字コードエラーの場合だけCP932で再読み込みする。
                uploaded_account_csv.seek(0)
                df_raw = pd.read_csv(
                    uploaded_account_csv,
                    header=None,
                    encoding="cp932",
                )

            if df_raw.empty:
                st.error("CSVが空です")
            else:
                accounts = (
                    df_raw.iloc[1:, 0]
                    .dropna()
                    .astype(str)
                    .str.strip()
                )
                accounts = accounts[accounts != ""].drop_duplicates().tolist()

                if len(accounts) == 0:
                    st.error("A列A2以降にデータがありません")
                else:
                    pd.DataFrame(
                        {"勘定科目": accounts}
                    ).to_csv(
                        "master_accounts.csv",
                        index=False,
                        encoding="utf-8-sig",
                    )
                    st.session_state["account_master"] = accounts
                    st.success(f"勘定科目を{len(accounts)}件 保存しました")

        except pd.errors.EmptyDataError:
            # 0行またはBOM・改行だけのCSVを明示的に処理する。
            st.error("CSVが空です")
        except pd.errors.ParserError as e:
            st.error(f"CSVの形式が正しくありません: {e}")
        except UnicodeDecodeError:
            st.error("CSVの文字コードを読み取れません。UTF-8またはCP932で保存してください")
        except Exception as e:
            st.error(f"読み込み失敗: {e}")
    if os.path.exists("master_accounts.csv"):
        master_df=pd.read_csv("master_accounts.csv");st.subheader("📋 登録済み勘定科目一覧");st.dataframe(master_df,use_container_width=True)

    st.markdown("---")

    st.markdown("""
    <style>
        /* subheader(h3)の上余白を削る */
    h3 {
        margin-bottom: 0.8rem !important;
    }

    /* さらに上のコンテナ余白も削る */
    div[data-testid="stHeadingWithActionElements"] {
        margin-top: 0rem !important;
        padding-top: 0rem !important;
    }

    /* チェックボックスブロックの下に余白 */
    div[data-testid="stCheckbox"] {
        margin-bottom: 20px;
    }

    /* チェックボックス群の直後の見出しを下げる */
    div[data-testid="stCheckbox"] + div h2 {
        margin-top: 2.5rem !important;
    }

    </style>
    """, unsafe_allow_html=True)

    st.subheader("🧾 (必須)税区分マスタ設定")
    st.markdown("#### 使用する税率")

    TAX_RATE_OPTIONS = ["5%","8%","8%(軽)","10%"]
    TAX_RATE_FILE = "tax_rate_config.csv"
    
    # 永続ロード
    if "enabled_tax_rates" not in st.session_state:
        if os.path.exists(TAX_RATE_FILE):
            df_rate = pd.read_csv(TAX_RATE_FILE)
            st.session_state["enabled_tax_rates"] = df_rate["税率"].tolist()
        else:
            st.session_state["enabled_tax_rates"] = []
    cols = st.columns(4)

    for i, rate in enumerate(TAX_RATE_OPTIONS):
        with cols[i]:
            st.checkbox(
                rate,
                key=f"rate_{rate}",
                value=rate in st.session_state["enabled_tax_rates"]
            )

    # 再構築
    selected_rates = [
        r for r in TAX_RATE_OPTIONS
        if st.session_state.get(f"rate_{r}", False)
    ]

    st.session_state["enabled_tax_rates"] = selected_rates

    # 永続保存
    pd.DataFrame({"税率": selected_rates}).to_csv(
        TAX_RATE_FILE, index=False, encoding="utf-8-sig"
    )

    
    st.markdown("#### 使用する税区分")

    # --- 下段：税区分（2列表示：左=チェック/右=説明） ---
    TAX_DATA = [
    ("課税売上","課税取引における売上"),
    ("課対仕入","課税売上に対応する課税仕入"),
    ("未選択","税区分が未定な取引"),
    ("課税","課税取引（暫定登録用の税区分）"),
    ("対象外","消費税と無関係な取引"),
    ("非課税","非課税取引（暫定登録用の税区分）"),
    ("輸出等","輸出等取引（暫定登録用の税区分）"),
    ("課対輸本","課税売上に対応する輸入。課税貨物。"),
    ("課対輸税","課税売上に対応する輸入時に納付した消費税"),
    ("地消貨割","課税貨物の引取に対応する税関に納付した地方消費税"),
    ("課税売倒","課税売上にかかわる貸倒"),
    ("課税売回","回収した貸倒債権の回収"),
    ("有価譲渡","有価証券の譲渡による非課税売上"),
    ("非資売倒","非課税資産売上にかかわる貸倒"),
    ("輸出売倒","輸出免税売上に対応する貸倒"),
    ("非課売倒","非課税売上にかかわる貸倒"),
    ("不課税","不課税取引（暫定登録用の税区分）"),
    ("輸出売上","輸出免税取引における売上"),
    ("非課売上","非課税取引における売上"),
    ("非資売上","非課税資産の輸出による売上"),
    ("対外売上","課税対象外の売上"),
    ("課税売返","課税売上にかかわる対価の返還"),
    ("輸出売返","輸出免税売上にかかわる対価の返還"),
    ("非課売返","非課税売上にかかわる対価の返還"),
    ("非資売返","非課税資産売上にかかわる対価の返還"),
    ("非対輸本","非課税売上に対応する輸入課税"),
    ("共対輸本","課税・非課税売上に共通対応する輸入課税"),
    ("非対輸税","非課税売上に共通対応する輸入時に納付した消費税"),
    ("共対輸税","課税・非課税売上に共通対応する輸入時に納付した消費税"),
    ("非対仕入","非課税売上に対応する課税仕入"),
    ("共対仕入","課税・非課税売上に共通対応する課税仕入"),
    ("非課仕入","非課税取引における仕入"),
    ("対外仕入","課税対象外取引における仕入"),
    ("課対仕返","課税売上に対応する課税仕入の返還"),
    ("非対仕返","非課税売上に対応する課税仕入の返還"),
    ("共対仕返","共通対応する課税仕入の返還"),
    ("非課仕返","非課税仕入の返還"),
    ("課売上一","第一種事業による課税売上"),
    ("課売上二","第二種事業による課税売上"),
    ("課売上三","第三種事業による課税売上"),
    ("課売上四","第四種事業による課税売上"),
    ("課売上五","第五種事業による課税売上"),
    ("課対輸返","課税売上に対応する輸入返還"),
    ("非対輸返","非課税売上に対応する輸入返還"),
    ("共対輸返","共通対応する輸入返還"),
    ("課売返一","第一種事業の売上返還"),
    ("課売返二","第二種事業の売上返還"),
    ("課売返三","第三種事業の売上返還"),
    ("課売返四","第四種事業の売上返還"),
    ("課売返五","第五種事業の売上返還"),
    ("課売上六","第六種事業による課税売上"),
    ("課売返六","第六種事業の売上返還"),
    ]

    TAX_CONFIG_FILE = "tax_config.csv"

    # 初回ロード（永続化）
    if "enabled_tax_list" not in st.session_state:
        if os.path.exists(TAX_CONFIG_FILE):
            saved_df = pd.read_csv(TAX_CONFIG_FILE)
            st.session_state["enabled_tax_list"] = saved_df["税区分"].tolist()
        else:
            st.session_state["enabled_tax_list"] = []

    enabled_list = st.session_state["enabled_tax_list"]
    new_enabled = []
    # 2列に分割
    left_col, right_col = st.columns(2)

    half = len(TAX_DATA) // 2
    left_data = TAX_DATA[:half]
    right_data = TAX_DATA[half:]

    def render_tax_column(data):
        selected = []
        for code, desc in data:
            col1, col2 = st.columns([1, 4])
            with col1:
                checked = st.checkbox(
                    code,
                    value=code in enabled_list,
                    key=f"tax_{code}"
                )
            with col2:
                st.caption(desc)

            if checked:
                selected.append(code)
        return selected

    new_enabled = []

    with left_col:
        new_enabled += render_tax_column(left_data)

    with right_col:
        new_enabled += render_tax_column(right_data)

    # 保存（永続化）
    st.session_state["enabled_tax_list"] = new_enabled
    pd.DataFrame({"税区分": new_enabled}).to_csv(
        TAX_CONFIG_FILE, index=False, encoding="utf-8-sig"
    )

    
    st.markdown("### 📊 現在の設定状況")

    # ---- 税率 ----
    enabled_rates = st.session_state.get("enabled_tax_rates", [])

    st.markdown("**🧾 有効な税率**")

    if enabled_rates:
        cols = st.columns(len(enabled_rates))
        for i, rate in enumerate(enabled_rates):
            cols[i].info(rate)
    else:
        st.caption("選択されていません")

    # ---- 税区分 ----
    enabled_taxes = st.session_state.get("enabled_tax_list", [])

    st.markdown("**📌 有効な税区分**")

    if enabled_taxes:
        st.success(f"{len(enabled_taxes)} 件選択中")
        st.write(" / ".join(enabled_taxes))
    else:
        st.caption("選択されていません")


    st.stop()

    # --- メインUI ---
    st.title("🧮 Excel集計ツール")

trade_mode = st.radio(
    "集計する表の種類を選んでください。",
    ["収入のみ", "支出のみ", "収支混在"],
    horizontal=True
)

uploaded_files = st.file_uploader(
    "📁 Excelファイルをアップロード",
    type=["xlsx"],
    accept_multiple_files=True,
)
header_search_limit = st.number_input("🔍 ヘッダー探索行数", min_value=1, max_value=20, value=10)


is_vertical_mode = st.toggle(
    "📐 縦表（項目名が左に並ぶ表）として処理する",
    value=False
)

# ============================
# freee 固定スキーマ定義
# ============================
FREEE_REQUIRED_COLS = [
    "収支区分",
    "発生日",
    "勘定科目",
    "税区分",
    "金額",
]

FREEE_OPTIONAL_COLS = [
    "税計算区分",
    "税額",
    "管理番号",
    "決済期日",
    "取引先",
    "取引先コード",
    "備考",
    "品目",
    "部門",
    "メモタグ（複数指定可、カンマ区切り）",
    "決済日",
    "決済口座",
    "決済金額",
    "セグメント1",
    "セグメント2",
    "セグメント3",
]

# --- 必須項目（横並び固定） ---
st.subheader("🔒 必須項目")

# freee公式仕様では、税計算区分は任意。税額は税計算区分が「外税」の場合のみ必須。
required_fields = ["収支区分", "発生日", "勘定科目", "税区分", "金額"]
cols = st.columns(len(required_fields))
for col, f in zip(cols, required_fields):
    with col:
        st.checkbox(f, value=True, disabled=True)

# --- 任意項目（空欄でもOK） ---
st.subheader("➕ 任意項目")
optional_fields = ["税計算区分", "税額", "管理番号","決済期日","取引先","取引先コード","備考","品目","部門","メモタグ（複数指定可、カンマ区切り）","決済日","決済口座","決済金額","セグメント1","セグメント2","セグメント3"]
# 旧形式の設定でも、保存済みの任意マッピングから表示対象を復元する。
if "selected_optional_fields" not in st.session_state:
    st.session_state["selected_optional_fields"] = [
        field for field in optional_fields if f"map_{field}" in st.session_state
    ]
selected_optional_fields = st.multiselect(
    "使用する追加項目を選択",
    optional_fields,
    key="selected_optional_fields",
)

# 以降のマッピング対象（必須＋任意）
logical_columns = [
    c for c in (required_fields + selected_optional_fields)
    if c != "収支区分"
]


# ヘッダー検出関連関数
def normalize_vectors(vectors):
    norms = np.linalg.norm(vectors, axis=1, keepdims=True)
    norms[norms == 0] = 1
    return vectors / norms

def cosine_similarity_matrix(vectors):
    return np.dot(vectors, vectors.T)

def row_type_vector(row):
    types = {"str":0,"int":0,"float":0,"date":0,"other":0}
    for cell in row:
        try:
            pd.to_datetime(cell, errors='raise')
            types["date"] += 1
        except:
            try:
                val = float(cell)
                if val.is_integer(): types["int"] += 1
                else: types["float"] += 1
            except:
                if isinstance(cell, str): types["str"] += 1
                else: types["other"] += 1
    return [types[k] for k in types]

def detect_header_candidates(df, max_rows=10, top_n=2,
                               empty_weight=0.7, unique_weight=0.3, diff_weight=0.5,
                               empty_threshold=0.5):
    vecs = np.array([row_type_vector(df.iloc[i]) for i in range(len(df))])
    norm = normalize_vectors(vecs)
    sim = cosine_similarity_matrix(norm)
    mean_sim = sim.mean(axis=1)

    scores = []
    for i in range(min(max_rows, len(df))):
        row = df.iloc[i].fillna("").astype(str)
        tot = len(row)
        empty_curr = sum(1 for v in row if v.strip() == "") / tot if tot else 0
        if empty_curr > empty_threshold:
            continue
        prev_empty = 0.0 if i == 0 else sum(1 for v in df.iloc[i - 1].fillna("").astype(str) if v.strip() == "") / tot
        unique = len(set(v for v in row if v.strip())) / tot if tot else 0
        diff = 1 - mean_sim[i]
        position_penalty = i / max_rows  # 下に行くほど大きくなる
        score = (
            empty_weight * prev_empty +
            unique_weight * unique +
            diff_weight * diff
            - 0.8 * position_penalty      # ★ これを追加
        )
        scores.append((score, i))

    scores_sorted = sorted(scores, reverse=True)[:top_n]
    return [idx for score, idx in scores_sorted]



def make_column_names_unique(cols):
    cnt={}; new=[]
    for c in cols:
        if c in cnt: cnt[c]+=1; new.append(f"{c}_{cnt[c]}")
        else: cnt[c]=0; new.append(c)
    return new



@st.cache_data(show_spinner=False)
def find_bordered_ranges(xlsx_path, sheet_name):
    """
    枠線で囲まれたセル群を連結成分ごとに分離し、
    それぞれを表ブロックとして返す
    """
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb[sheet_name]

    # --- 1) 枠線セル抽出 ---
    bordered = set()

    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            b = cell.border
            if b and (
                (b.top and b.top.style) or
                (b.bottom and b.bottom.style) or
                (b.left and b.left.style) or
                (b.right and b.right.style)
            ):
                # 罫線の各辺がNoneのExcelでも、存在する辺だけを安全に確認する。
                bordered.add((r, c))

    if not bordered:
        return []

    # --- 2) 連結成分探索（BFS） ---
    visited = set()
    blocks = []

    for start in bordered:
        if start in visited:
            continue

        queue = deque([start])
        visited.add(start)

        rows = []
        cols = []

        while queue:
            r, c = queue.popleft()
            rows.append(r)
            cols.append(c)

            for dr, dc in [(-1,0), (1,0), (0,-1), (0,1)]:
                nxt = (r + dr, c + dc)
                if nxt in bordered and nxt not in visited:
                    visited.add(nxt)
                    queue.append(nxt)

        blocks.append({
            "min_row": min(rows),
            "max_row": max(rows),
            "min_col": min(cols),
            "max_col": max(cols),
        })

    return blocks

def is_empty_cell(v):
    return v is None or str(v).strip() == ""

def is_empty_row(row):
    return not any(
        str(v).strip() not in ("", "None", "nan")
        for v in row
    )

def is_empty_col(col):
    return all(is_empty_cell(v) for v in col)

def is_surrounded_by_empty(df, top, bottom, left, right):
    eff_top, eff_bottom, eff_left, eff_right = find_effective_bounds(df)

    # --- 上 ---
    if top > eff_top:
        if not is_empty_row(df.iloc[top-1, left:right+1]):
            return False

    # --- 下 ---
    if bottom < eff_bottom:
        if not is_empty_row(df.iloc[bottom+1, left:right+1]):
            return False

    # --- 左 ---
    if left > eff_left:
        if not is_empty_col(df.iloc[top:bottom+1, left-1]):
            return False

    # --- 右 ---
    if right < eff_right:
        if not is_empty_col(df.iloc[top:bottom+1, right+1]):
            return False

    return True


def looks_like_header_row(row):
    # 文字列が多ければヘッダーとみなす（超現実的）
    non_empty = [v for v in row if not is_empty_cell(v)]
    if not non_empty:
        return False
    str_ratio = sum(isinstance(v, str) for v in non_empty) / len(non_empty)
    return str_ratio >= 0.6

def looks_like_data_row(row):
    # 数値 or 日付が1つでもあればデータ行
    for v in row:
        try:
            float(v)
            return True
        except:
            try:
                pd.to_datetime(v)
                return True
            except:
                pass
    return False

    
def find_table_blocks_by_data(df, header_search_limit=10):
    """
    ヘッダー行を起点に帳票ブロックを検出する（帳票向け）
    """
    blocks = []

    header_rows = detect_header_candidates(
        df,
        max_rows=header_search_limit,
        top_n=2
    )

    for hr in header_rows:
        # --- 横方向：ヘッダー文字がある列 ---
        header_cells = [
            c for c in range(df.shape[1])
            if str(df.iat[hr, c]).strip() != ""
        ]

        if len(header_cells) < 2:
            continue

        left = min(header_cells)
        right = max(header_cells)

        # --- 下方向：データ行が続く限り ---
        bottom = hr
        for r in range(hr + 1, df.shape[0]):
            row = df.iloc[r, left:right + 1]
            if looks_like_data_row(row):
                bottom = r
            else:
                break

        blocks.append({
            "min_row": hr,
            "max_row": bottom,
            "min_col": left,
            "max_col": right
        })

    return blocks


def header_string_ratio(row):
    vals = [v for v in row if str(v).strip() != ""]
    if not vals:
        return 0.0
    str_like = 0
    for v in vals:
        s = str(v)
        # 数値っぽいものを除外
        try:
            float(s.replace(",", ""))
            continue
        except:
            pass
        # 日付っぽいものも除外
        try:
            pd.to_datetime(s)
            continue
        except:
            pass
        str_like += 1
    return str_like / len(vals)

def find_effective_bounds(df):
    """
    実データが存在する最小・最大 row / col を返す
    """
    rows, cols = df.shape

    top = None
    bottom = None
    left = None
    right = None

    # row
    for r in range(rows):
        if not is_empty_row(df.iloc[r]):
            top = r
            break

    for r in reversed(range(rows)):
        if not is_empty_row(df.iloc[r]):
            bottom = r
            break

    # col
    for c in range(cols):
        if not is_empty_col(df.iloc[:, c]):
            left = c
            break

    for c in reversed(range(cols)):
        if not is_empty_col(df.iloc[:, c]):
            right = c
            break

    return top, bottom, left, right

def is_empty_row(row):
    return all(v is None or str(v).strip() == "" for v in row)

def find_row_blocks(df):
    """
    空行で区切られた「連続したデータ行」をブロックとして抽出
    """
    blocks = []
    start = None

    for i in range(len(df)):
        if not is_empty_row(df.iloc[i]):
            if start is None:
                start = i
        else:
            if start is not None:
                blocks.append((start, i - 1))
                start = None

    if start is not None:
        blocks.append((start, len(df) - 1))

    return blocks
def detect_block_columns(df, row_start, row_end):
    """
    行ブロック内で、実データが存在する列範囲を決定
    """
    sub = df.iloc[row_start:row_end + 1]

    non_empty_cols = [
        j for j in range(df.shape[1])
        if not is_empty_row(sub.iloc[:, j])
    ]

    if not non_empty_cols:
        return None

    return min(non_empty_cols), max(non_empty_cols)

def find_table_blocks_by_rows(df):
    """
    空行で区切られた表ブロックを検出（推奨版）
    """
    blocks = []

    row_blocks = find_row_blocks(df)

    for r0, r1 in row_blocks:
        col_range = detect_block_columns(df, r0, r1)
        if not col_range:
            continue

        c0, c1 = col_range

        # 最低条件（1行表・1列表を除外）
        if r1 - r0 < 1 or c1 - c0 < 1:
            continue

        blocks.append((r0, r1, c0, c1))

    return blocks

def looks_like_continuous_data(df, i):
    if i == 0 or i >= len(df) - 1:
        return False
    return (
        looks_like_data_row(df.iloc[i]) and
        looks_like_data_row(df.iloc[i + 1])
    )

def looks_like_table(block, header_row=0):

    # --- 行ごとの非空セル数の安定性 ---
    row_non_empty = block.notna().sum(axis=1)
    if row_non_empty.std() > 1.5:
        return False

    # --- 数値っぽい列が1つ以上あるか ---
    numeric_cols = 0
    for c in range(block.shape[1]):
        ratio = (
            block.iloc[:, c]
            .fillna("")
            .astype(str)
            .str.contains(r"\d")
            .mean()
        )
        if ratio > 0.3:
            numeric_cols += 1

    if numeric_cols == 0:
        return False

    # --- ★ ヘッダー vs データ行の構造差判定 ---
    def type_vector(row):
        v = {"str": 0, "num": 0, "date": 0}
        for x in row:
            s = str(x).strip()
            if s == "":
                continue
            try:
                float(s.replace(",", "").replace("円", ""))
                v["num"] += 1
                continue
            except:
                pass
            try:
                pd.to_datetime(s)
                v["date"] += 1
                continue
            except:
                pass
            v["str"] += 1
        return v

    header = block.iloc[header_row]
    data = block.iloc[header_row + 1]

    h = type_vector(header)
    d = type_vector(data)

    diff = (
        abs(h["str"] - d["str"]) +
        abs(h["num"] - d["num"]) +
        abs(h["date"] - d["date"])
    )

    # 差が小さい＝文章の続きっぽい
    if diff < 2:
        return False

    return True

def find_table_blocks_by_4dir_empty(df):
    """
    上下左右が空行・空列（または端）で囲まれた表ブロックを検出する
    戻り値: [(s, e, c0, c1), ...]
    """

    # 非空セル判定
    # 非空セル判定
    not_empty = df.apply(lambda col: col.notna() & col.astype("string").str.strip().ne(""))

    visited = set()
    blocks = []

    max_r, max_c = df.shape

    for r in range(max_r):
        for c in range(max_c):
            if not not_empty.iloc[r, c]:
                continue
            if (r, c) in visited:
                continue

            # --- BFSで連結成分（島）を取得 ---
            stack = [(r, c)]
            cells = []

            while stack:
                cr, cc = stack.pop()
                if (cr, cc) in visited:
                    continue
                if not not_empty.iloc[cr, cc]:
                    continue

                visited.add((cr, cc))
                cells.append((cr, cc))

                for dr, dc in [(-1,0),(1,0),(0,-1),(0,1)]:
                    nr, nc = cr + dr, cc + dc
                    if 0 <= nr < max_r and 0 <= nc < max_c:
                        stack.append((nr, nc))

            # --- 矩形化 ---
            rows = [x[0] for x in cells]
            cols = [x[1] for x in cells]

            s, e = min(rows), max(rows)
            c0, c1 = min(cols), max(cols)

            # --- 4方向が空か確認 ---
            # 上
            if s > 0 and not_empty.iloc[s-1, c0:c1+1].any():
                continue
            # 下
            if e < max_r-1 and not_empty.iloc[e+1, c0:c1+1].any():
                continue
            # 左
            if c0 > 0 and not_empty.iloc[s:e+1, c0-1].any():
                continue
            # 右
            if c1 < max_c-1 and not_empty.iloc[s:e+1, c1+1].any():
                continue

            blocks.append((s, e, c0, c1))

    return blocks


def normalize_header_name(s: str) -> str:
    if s is None:
        return ""
    s = str(s)

    # Excel由来の不可視・制御文字
    s = s.replace("_x000D_", "")
    s = re.sub(r"[\r\n\t\u00A0\u200B\u200C\u200D\uFEFF]", "", s)

    # 全角スペース → 半角
    s = s.replace("　", " ")

    # 前後空白除去 & 連続空白を1つに
    s = " ".join(s.strip().split())

    return s

def normalize_japanese_date(s):
    if pd.isna(s):
        return None
    s = str(s).strip()

    # 和暦（令和のみ対応：業務的に十分）
    m = re.match(r"令和\s*(\d+)\s*年\s*(\d+)\s*月\s*(\d+)\s*日", s)
    if m:
        y = int(m.group(1)) + 2018
        return f"{y}/{int(m.group(2)):02d}/{int(m.group(3)):02d}"

    # 年月日（日本語）
    s = re.sub(r"[年月]", "/", s).replace("日", "")
    s = s.replace("-", "/")

    # Excel日時型・時刻付き文字列は、発生日として日付部分だけを使用する。
    m = re.match(r"^(\d{4})/(\d{1,2})/(\d{1,2})(?:[ T].*)?$", s)
    if m:
        return f"{int(m.group(1)):04d}/{int(m.group(2)):02d}/{int(m.group(3)):02d}"

    return s


def validate_date_with_errors(df, col):
    raw = df[col].fillna("").astype(str).str.strip()
    fixed = raw.apply(normalize_japanese_date)
    parsed = pd.to_datetime(fixed, errors="coerce", format="%Y/%m/%d")

    error_rows = parsed[parsed.isna()].index.tolist()
    # 不正値は画面で確認できるよう元の文字列を残す。
    normalized = raw.astype("object")
    valid = parsed.notna()
    normalized.loc[valid] = parsed.loc[valid].dt.strftime("%Y/%m/%d")
    df[col] = normalized
    return error_rows

MAX_AMOUNT = 9223372036854775807

def validate_amount_with_errors(df, col):
    raw = df[col].fillna("").astype(str).str.strip()
    nums = pd.to_numeric(
        raw.str.replace(",", "", regex=False).str.replace("円", "", regex=False),
        errors="coerce",
    )
    invalid = (
        raw.eq("")
        | nums.isna()
        | (nums % 1 != 0)
        | (nums < 1)
        | (nums.abs() > MAX_AMOUNT)
    )

    error_rows = df.index[invalid].tolist()
    # 不正値は画面で確認できるよう元の文字列を残す。
    normalized = raw.astype("object")
    valid = ~invalid
    normalized.loc[valid] = nums.loc[valid].round().astype("int64")
    df[col] = normalized
    return error_rows


def validate_required_text_with_errors(df, col, allowed_values=None):
    """必須文字列と、指定値のみ許可される項目を検証する。"""
    values = df[col].fillna("").astype(str).str.strip()
    invalid = values.eq("")
    if allowed_values is not None:
        invalid |= ~values.isin(allowed_values)

    df[col] = values
    return df.index[invalid].tolist()


def validate_optional_choice_with_errors(df, col, allowed_values):
    """空欄を許可しつつ、入力値が許可値に含まれるかを検証する。"""
    values = df[col].fillna("").astype(str).str.strip()
    invalid = values.ne("") & ~values.isin(allowed_values)
    df[col] = values
    return df.index[invalid].tolist()


def validate_optional_tax_amount_with_errors(df, col, amount_col):
    """任意の税額を円単位で検証し、金額と符号が矛盾しないかを確認する。"""
    raw = df[col].fillna("").astype(str).str.strip()
    nums = pd.to_numeric(
        raw.str.replace(",", "", regex=False).str.replace("円", "", regex=False),
        errors="coerce",
    )
    invalid = raw.ne("") & (
        nums.isna() | (nums % 1 != 0) | (nums.abs() > MAX_AMOUNT)
    )

    amounts = pd.to_numeric(df[amount_col], errors="coerce")
    invalid |= (
        raw.ne("")
        & nums.notna()
        & amounts.notna()
        & (nums != 0)
        & ((nums * amounts) < 0)
    )

    # 不正値は画面で確認できるよう元の文字列を残す。
    normalized = raw.astype("object")
    valid = ~invalid & raw.ne("")
    normalized.loc[valid] = nums.loc[valid].round().astype("int64")
    df[col] = normalized
    return df.index[invalid].tolist()


def validate_optional_date_with_errors(df, col):
    """任意の日付を検証し、入力値がある行だけYYYY/MM/DDへ正規化する。"""
    raw = df[col].fillna("").astype(str).str.strip()
    fixed = raw.apply(normalize_japanese_date)
    parsed = pd.to_datetime(fixed, errors="coerce", format="%Y/%m/%d")
    invalid = raw.ne("") & parsed.isna()

    normalized = raw.astype("object")
    valid = raw.ne("") & parsed.notna()
    normalized.loc[valid] = parsed.loc[valid].dt.strftime("%Y/%m/%d")
    df[col] = normalized
    return df.index[invalid].tolist()


def validate_optional_signed_amount_with_errors(df, col):
    """任意の決済金額を円単位の整数として検証する。"""
    raw = df[col].fillna("").astype(str).str.strip()
    nums = pd.to_numeric(
        raw.str.replace(",", "", regex=False).str.replace("円", "", regex=False),
        errors="coerce",
    )
    invalid = raw.ne("") & (
        nums.isna() | (nums % 1 != 0) | (nums.abs() > MAX_AMOUNT)
    )

    normalized = raw.astype("object")
    valid = raw.ne("") & ~invalid
    normalized.loc[valid] = nums.loc[valid].round().astype("int64")
    df[col] = normalized
    return df.index[invalid].tolist()


def validate_settlement_fields_with_errors(df):
    """決済情報を一部だけ入力した行は、決済日・口座・金額をすべて要求する。"""
    columns = ["決済日", "決済口座", "決済金額"]
    if not any(column in df.columns for column in columns):
        return {}

    values = {}
    for column in columns:
        if column in df.columns:
            raw = df[column].astype("object")
            values[column] = raw.where(raw.notna(), "").astype(str).str.strip()
        else:
            values[column] = pd.Series("", index=df.index)
    has_settlement = pd.DataFrame(values).ne("").any(axis=1)
    return {
        column: df.index[has_settlement & values[column].eq("")].tolist()
        for column in columns
    }


def validate_max_length_with_errors(df, col, max_length, split_by_comma=False):
    """freee公式ヘルプに記載された文字数上限を検証する。"""
    values = df[col].fillna("").astype(str).str.strip()
    if split_by_comma:
        invalid = values.apply(
            lambda value: any(len(tag.strip()) > max_length for tag in value.split(",") if tag.strip())
        )
    else:
        invalid = values.str.len() > max_length
    return df.index[invalid].tolist()


def validate_freee_import_data(df, validate_trade=True, validate_all_required=True):
    """freee取引インポート用の必須項目・形式を検証し、エラー位置を返す。"""
    required_columns = [
        "収支区分", "発生日", "勘定科目", "税区分", "金額", "税計算区分", "税額"
    ]
    for column in required_columns:
        if column not in df.columns:
            df[column] = ""

    error_map = {}

    def add_errors(rows, std_col):
        for row_index in rows:
            sheet = df.at[row_index, "シート名"] if "シート名" in df.columns else "不明"
            source_column = (
                df.at[row_index, f"__src__{std_col}"]
                if f"__src__{std_col}" in df.columns
                # 元列がない項目は行エラーではなく、マッピング外として扱う。
                else None
            )
            # 複数シート結合時のNaNは、元列がない状態として扱う。
            if pd.isna(source_column):
                source_column = None
            key = (sheet, row_index, std_col, source_column)
            error_map.setdefault(key, []).append(std_col)

    if validate_trade:
        add_errors(
            validate_required_text_with_errors(df, "収支区分", {"収入", "支出"}),
            "収支区分",
        )

    add_errors(validate_date_with_errors(df, "発生日"), "発生日")
    if validate_all_required:
        add_errors(validate_required_text_with_errors(df, "勘定科目"), "勘定科目")
        add_errors(validate_required_text_with_errors(df, "税区分"), "税区分")
    add_errors(validate_amount_with_errors(df, "金額"), "金額")
    add_errors(
        validate_optional_choice_with_errors(df, "税計算区分", {"", "内税", "外税", "税込"}),
        "税計算区分",
    )
    add_errors(validate_optional_tax_amount_with_errors(df, "税額", "金額"), "税額")

    for column in ["決済期日", "決済日"]:
        if column in df.columns:
            add_errors(validate_optional_date_with_errors(df, column), column)
    if "決済金額" in df.columns:
        add_errors(validate_optional_signed_amount_with_errors(df, "決済金額"), "決済金額")
    for column, rows in validate_settlement_fields_with_errors(df).items():
        add_errors(rows, column)

    for column, max_length, split_by_comma in [
        ("取引先", 255, False),
        ("品目", 255, False),
        ("勘定科目", 255, False),
        ("メモタグ（複数指定可、カンマ区切り）", 30, True),
        ("部門", 30, False),
    ]:
        if column in df.columns:
            add_errors(
                validate_max_length_with_errors(df, column, max_length, split_by_comma),
                column,
            )

    return error_map

def get_excel_row_number(df, index):
    if "__excel_row__" not in df.columns:
        return index + 1

    excel_row = df.at[index, "__excel_row__"]
    return int(excel_row) if pd.notna(excel_row) else index + 1


def build_error_display_dataframe(df, error_map, view_edited_error_keys=None):
    """取込元の不整合はシート・行、画面入力の不整合は表示ビュー内で案内する。"""
    rows = []
    unmapped_counts = {}
    amount_errors = {}
    view_error_counts = {}
    view_edited_error_keys = view_edited_error_keys or set()

    for (sheet, idx, col, src_col), columns in error_map.items():
        # 表示ビューで変更した値は、元Excel列ではなく編集後の入力エラーとして示す。
        if (sheet, idx, col, src_col) in view_edited_error_keys:
            view_error_counts[col] = view_error_counts.get(col, 0) + 1
            continue

        # 元列がある金額エラーは、同じシート・列ごとに件数と行番号をまとめる。
        if col == "金額" and src_col is not None and str(src_col).strip():
            key = (sheet, str(src_col).strip())
            amount_errors.setdefault(key, []).append(get_excel_row_number(df, idx))
            continue

        if src_col is None or not str(src_col).strip():
            key = (sheet, col)
            unmapped_counts[key] = unmapped_counts.get(key, 0) + 1
            continue

        rows.append({
            "場所": f"{sheet}シートの{get_excel_row_number(df, idx)}行目",
            "エラー列": f"{src_col}（{', '.join(columns)}）",
        })

    for (sheet, src_col), excel_rows in amount_errors.items():
        row_text = "・".join(map(str, sorted(set(excel_rows))))
        rows.append({
            "場所": f"{sheet}シートの{row_text}行目",
            "エラー列": f"{src_col}（金額：入力が不正です：{len(excel_rows)}件）",
        })

    for (sheet, col), count in unmapped_counts.items():
        rows.append({
            "場所": f"{sheet}シート（マッピング未設定）",
            "エラー列": f"{col}（{count}行が未設定）",
        })

    for col, count in view_error_counts.items():
        rows.append({
            "場所": "表示ビュー内",
            "エラー列": f"{col}（入力が不正です：{count}件）",
        })

    return pd.DataFrame(rows, columns=["場所", "エラー列"])


def get_view_edited_error_keys(original_df, current_df, error_map):
    """初回集計値から変更されたセルのエラーを、表示ビュー入力として識別する。"""
    edited_keys = set()
    if original_df is None:
        return edited_keys

    for key in error_map:
        _, idx, col, _ = key
        if (
            idx not in original_df.index
            or idx not in current_df.index
            or col not in original_df.columns
            or col not in current_df.columns
        ):
            continue

        original_value = original_df.at[idx, col]
        current_value = current_df.at[idx, col]
        original_text = "" if pd.isna(original_value) else str(original_value).strip()
        current_text = "" if pd.isna(current_value) else str(current_value).strip()
        if original_text != current_text:
            edited_keys.add(key)

    return edited_keys


def contains_date_or_number(value):
    """取込対象のデータ行に日付または数値が含まれるかを判定する。"""
    if value is None or pd.isna(value):
        return False

    text = str(value).strip()
    if not text:
        return False

    numeric_text = text.replace(",", "").replace("円", "")
    if pd.notna(pd.to_numeric(numeric_text, errors="coerce")):
        return True

    normalized_date = normalize_japanese_date(text)
    return pd.notna(pd.to_datetime(normalized_date, errors="coerce"))


def is_importable_table_block(block, header_row):
    """注釈・タイトルを取込対象の表ブロックから除外する。"""
    header = block.iloc[header_row]
    header_count = sum(
        value is not None and not pd.isna(value) and str(value).strip() != ""
        for value in header
    )
    if header_count < 2:
        return False

    data = block.iloc[header_row + 1:]
    non_empty_data_rows = data[
        data.apply(
            lambda row: any(
                value is not None and not pd.isna(value) and str(value).strip() != ""
                for value in row
            ),
            axis=1,
        )
    ]
    if non_empty_data_rows.empty:
        return False

    return any(
        contains_date_or_number(value)
        for value in non_empty_data_rows.to_numpy().flatten()
    )


def find_importable_header_row(block):
    """ヘッダーらしさの順位と直下のデータ構造を組み合わせて見出し行を探す。"""
    if len(block) < 2:
        return None

    search_limit = min(int(header_search_limit), len(block))
    # 既存の空欄率・型分布・ユニーク率による候補順位を先に使う。
    scored_rows = detect_header_candidates(
        block,
        max_rows=search_limit,
        top_n=min(5, search_limit),
    )
    # スコア候補が外れた場合も、構造条件を満たす行を見落とさない。
    candidate_rows = scored_rows + [
        row_index for row_index in range(len(block) - 1)
        if row_index not in scored_rows
    ]

    for row_index in candidate_rows:
        header = block.iloc[row_index]
        header_count = sum(
            value is not None and not pd.isna(value) and str(value).strip() != ""
            for value in header
        )
        if header_count < 2:
            continue

        # 見出し直下の最初の非空行を確認し、注意書きではなくデータ行であることを判定する。
        following_rows = block.iloc[row_index + 1:]
        next_data_index = next(
            (
                index for index in range(len(following_rows))
                if any(
                    value is not None and not pd.isna(value) and str(value).strip() != ""
                    for value in following_rows.iloc[index]
                )
            ),
            None,
        )
        if next_data_index is None:
            continue

        next_data_row = following_rows.iloc[next_data_index]
        next_data_count = sum(
            value is not None and not pd.isna(value) and str(value).strip() != ""
            for value in next_data_row
        )
        if next_data_count < 2 or not any(contains_date_or_number(value) for value in next_data_row):
            continue

        if is_importable_table_block(block, row_index):
            return row_index
    return None


def find_vertical_table_start_column(block):
    """縦表を転置した後、実データが続く列から表本体を開始する。"""
    if block.empty or len(block) < 2:
        return None

    # 注意書き列は、見出しの下にレコード値が続かないため候補から外す。
    record_columns = []
    for column_index in range(block.shape[1]):
        column = block.iloc[:, column_index]
        has_header = column.iloc[0] is not None and not pd.isna(column.iloc[0]) and str(column.iloc[0]).strip() != ""
        has_data = any(
            value is not None and not pd.isna(value) and str(value).strip() != ""
            for value in column.iloc[1:]
        )
        record_columns.append(has_header and has_data)

    # 2列以上連続するレコード列だけを表開始候補とし、共通のヘッダー判定で確定する。
    run_start = None
    for column_index, is_record_column in enumerate(record_columns + [False]):
        if is_record_column and run_start is None:
            run_start = column_index
            continue
        if not is_record_column and run_start is not None:
            if column_index - run_start >= 2:
                candidate = block.iloc[:, run_start:].reset_index(drop=True)
                header_row = find_importable_header_row(candidate)
                if header_row is not None:
                    # 縦表は複数レコードを想定し、転置された横表を除外する。
                    data_rows = candidate.iloc[header_row + 1:]
                    data_row_count = sum(
                        sum(
                            value is not None and not pd.isna(value) and str(value).strip() != ""
                            for value in row
                        ) >= 2 and any(contains_date_or_number(value) for value in row)
                        for _, row in data_rows.iterrows()
                    )
                    if data_row_count >= 2:
                        return run_start
            run_start = None

    return None


# ヘッダー候補収集
# 同名列でも、シート・表ブロック・列位置が異なれば別候補として扱う。
header_candidates = {}
header_candidate_records = []
header_candidate_names = {}
header_indices = {}
header_debug_records = []

if uploaded_files:
    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        uploaded_file.seek(0)
        xls = pd.ExcelFile(uploaded_file)

        for sheet in xls.sheet_names:
            uploaded_file.seek(0)
            df_raw = pd.read_excel(uploaded_file, sheet_name=sheet, header=None, dtype=str)

            uploaded_file.seek(0)
            border_blocks_raw = find_bordered_ranges(uploaded_file, sheet)
            border_blocks = []
            for b in border_blocks_raw:
                if is_vertical_mode:
                    border_blocks.append((b["min_col"] - 1, b["max_col"] - 1, b["min_row"] - 1, b["max_row"] - 1))
                else:
                    border_blocks.append((b["min_row"] - 1, b["max_row"] - 1, b["min_col"] - 1, b["max_col"] - 1))

            if is_vertical_mode:
                df_raw = df_raw.T.reset_index(drop=True)

            usable_border_blocks = []
            for (s, e, c0, c1) in border_blocks:
                if s == e and e + 1 < df_raw.shape[0]:
                    next_row = df_raw.iloc[e + 1, c0:c1 + 1]
                    if next_row.notna().any():
                        continue
                usable_border_blocks.append((s, e, c0, c1))

            blocks = usable_border_blocks or find_table_blocks_by_4dir_empty(df_raw)

            for (s, e, c0, c1) in blocks:
                block = df_raw.iloc[s:e + 1, c0:c1 + 1].reset_index(drop=True)

                if is_vertical_mode:
                    # 項目名に依存せず、横表と共通の「直下にデータがある」判定で開始列を決める。
                    item_column = find_vertical_table_start_column(block)
                    if item_column is None:
                        header_debug_records.append({
                            "ファイル": file_name,
                            "シート": sheet,
                            "ブロック範囲": f"行 {s + 1}〜{e + 1} / 列 {c0 + 1}〜{c1 + 1}",
                            "判定": "除外",
                            "ヘッダー行": "-",
                            "検出ヘッダー": "縦表の開始列を特定できません",
                        })
                        continue
                    block = block.iloc[:, item_column:].reset_index(drop=True)
                    c0 += item_column

                header_row = find_importable_header_row(block)
                if header_row is None:
                    header_debug_records.append({
                        "ファイル": file_name,
                        "シート": sheet,
                        "ブロック範囲": f"行 {s + 1}〜{e + 1} / 列 {c0 + 1}〜{c1 + 1}",
                        "判定": "除外",
                        "ヘッダー行": "-",
                        "検出ヘッダー": "表として成立するヘッダー行を特定できません",
                    })
                    continue

                data_start_row = header_row + 1
                cols = make_column_names_unique(block.iloc[header_row].fillna("").astype(str).tolist())
                block_key = (file_name, sheet, s, e, c0, c1)
                column_candidate_ids = {}

                for column_index, column_name in enumerate(cols):
                    normalized_name = normalize_header_name(column_name)
                    if not normalized_name:
                        continue

                    candidate_id = f"{file_name}|{sheet}|{s}|{e}|{c0}|{c1}|{column_index}"
                    column_candidate_ids[column_index] = candidate_id
                    header_candidate_names[candidate_id] = normalized_name
                    header_candidate_records.append({
                        "id": candidate_id,
                        "name": normalized_name,
                        "sheet": sheet,
                        "file_name": file_name,
                    })

                header_indices[block_key] = {
                    "file_name": file_name,
                    "uploaded_file": uploaded_file,
                    "header_row": header_row,
                    "data_start_row": data_start_row,
                    "columns": cols,
                    "column_candidate_ids": column_candidate_ids,
                }
                # 検出した各表ブロックと採用ヘッダーを、複数表の原因確認に使う。
                header_debug_records.append({
                    "ファイル": file_name,
                    "シート": sheet,
                    "ブロック範囲": f"行 {s + 1}〜{e + 1} / 列 {c0 + 1}〜{c1 + 1}",
                    "判定": "採用",
                    "ヘッダー行": f"ブロック内 {header_row + 1} 行目",
                    "検出ヘッダー": " / ".join(cols),
                })

if header_debug_records:
    with st.expander("🔧 表ブロック・ヘッダー判定", expanded=False):
        st.caption("採用された表だけでなく、除外された表ブロックと理由も確認できます。")
        st.dataframe(pd.DataFrame(header_debug_records), hide_index=True, use_container_width=True)

# 同じヘッダー名は連番を付け、どのシートの列か分かるように表示する。
header_name_counts = Counter(record["name"] for record in header_candidate_records)
header_name_numbers = Counter()
for record in header_candidate_records:
    name = record["name"]
    header_name_numbers[name] += 1
    number = header_name_numbers[name]
    display_name = f"{name}{number}" if header_name_counts[name] > 1 else name
    header_candidates[record["id"]] = f"{display_name}（{record['file_name']} / {record['sheet']}シート）"

# 保存済みの列名を今回検出した候補IDへ変換し、同じExcelを再読込しても復元できるようにする。
pending_mapping_names = st.session_state.pop("pending_mapping_names", None)
if pending_mapping_names:
    for std_col, saved_names in pending_mapping_names.items():
        saved_name_set = set(saved_names if isinstance(saved_names, list) else [])
        st.session_state[f"map_{std_col}"] = [
            candidate_id
            for candidate_id, header_name in header_candidate_names.items()
            if header_name in saved_name_set
        ]

for k in list(st.session_state.keys()):
    if k.startswith("map_"):
        st.session_state[k] = [
            v for v in st.session_state[k]
            if v in header_candidates
        ]


# マッピングUI
st.subheader("🧩 マッピング")

user_mapping = {}

for std in logical_columns:
    if std == "収支区分":
        continue
    user_mapping[std] = st.multiselect(
        f"「{std}」→列名選択",
        options=sorted(header_candidates, key=header_candidates.get),
        format_func=lambda candidate_id: header_candidates[candidate_id],
        key=f"map_{std}"
    )

pending_config_save = st.session_state.pop("pending_config_save", None)
if pending_config_save:
    # 列IDではなく列名を保存し、表ブロック位置が変わっても設定を復元できるようにする。
    mapping_data = {
        "__selected_optional_fields__": selected_optional_fields,
        "__mapping_header_names__": {
            std_col: [
                header_candidate_names[candidate_id]
                for candidate_id in selected
                if candidate_id in header_candidate_names
            ]
            for std_col, selected in user_mapping.items()
        },
    }
    save_path = os.path.join(CONFIG_DIR, f"{pending_config_save}.json")
    try:
        with open(save_path, "w", encoding="utf-8") as f:
            json.dump(mapping_data, f, ensure_ascii=False, indent=2)
        st.sidebar.success(f"設定 '{pending_config_save}' を保存しました")
    except Exception as e:
        st.sidebar.error(f"設定保存失敗: {e}")

run = st.button("📥 集計実行")

if run:
    all_data = []
    errors = []
    selected_header_names = {
        std: {
            header_candidate_names[candidate_id]
            for candidate_id in selected
            if candidate_id in header_candidate_names
        }
        for std, selected in user_mapping.items()
    }

    # 複数ファイルでは、選択済みの元列が存在しないファイルを集計対象から除外する。
    file_header_names = defaultdict(set)
    for info in header_indices.values():
        file_header_names[info["file_name"]].update(
            normalize_header_name(column_name)
            for column_name in info["columns"]
            if normalize_header_name(column_name)
        )

    incompatible_files = {}
    for file_name, file_headers in file_header_names.items():
        missing_columns = [
            std for std, names in selected_header_names.items()
            if names and not (names & file_headers)
        ]
        if missing_columns:
            incompatible_files[file_name] = missing_columns

    for file_name, missing_columns in incompatible_files.items():
        st.warning(
            f"{file_name} は共通マッピングに必要な列（{'、'.join(missing_columns)}）がないため、集計対象から除外しました。"
        )

    for (file_name, sheet, s, e, c0, c1), info in header_indices.items():
        if file_name in incompatible_files:
            continue

        try:
            # --- 元データ読み込み ---
            uploaded_file = info["uploaded_file"]
            uploaded_file.seek(0)
            df0 = pd.read_excel(uploaded_file, sheet_name=sheet, header=None, dtype=str)

            # --- 縦スイッチONなら全体を転置 ---
            if is_vertical_mode:
                df0 = df0.T.reset_index(drop=True)

            # --- 表ブロック切り出し ---
            blk = df0.iloc[s:e + 1, c0:c1 + 1].reset_index(drop=True)

            header = info["columns"]
            data = blk.iloc[info["data_start_row"]:].reset_index(drop=True)

            # 転置後の縦表・横表はいずれも、ヘッダー行の下を通常のレコードとして扱う。
            df_block = data.copy()
            df_block.columns = make_column_names_unique(header)

            col_groups = {std: [] for std in logical_columns}
            for std in logical_columns:
                selected = user_mapping.get(std, [])
                selected_candidate_ids = set(selected)
                selected_names = selected_header_names.get(std, set())

                for column_index, col in enumerate(df_block.columns):
                    candidate_id = info["column_candidate_ids"].get(column_index)
                    # 重複ヘッダーを区別するため、画面で選択した列候補IDだけを採用する。
                    if candidate_id in selected_candidate_ids:
                        col_groups[std].append(col)

            # 選択したマッピングに一致しない表ブロックは、空行として結果へ追加しない。
            if not any(col_groups.values()):
                continue

            new_df = pd.DataFrame(index=df_block.index)
            for std in logical_columns:
                cols = col_groups.get(std, [])
                if cols:
                    series = df_block[cols[0]].copy()
                    for c in cols[1:]:
                        series = series.combine_first(df_block[c])
                    new_df[std] = series
                    new_df[f"__src__{std}"] = cols[0]
                else:
                    # マッピングされていない列も空で作る
                    new_df[std] = ""

            if new_df.empty:
                errors.append(f"{file_name} / {sheet}@{s}-{e}")
                continue

            df_block = new_df


            # --- 収支区分をモードで自動付与 ---
            if trade_mode == "収入のみ":
                df_block["収支区分"] = "収入"
            elif trade_mode == "支出のみ":
                df_block["収支区分"] = "支出"
            else:
                df_block["収支区分"] = ""  # 混在モード

            if not is_vertical_mode:
                df_block["__excel_row__"] = [
                s + info["data_start_row"] + row_offset + 1
                for row_offset in range(len(df_block))]

            # --- シート名付与 & 追加 ---
            df_block["シート名"] = f"{file_name} / {sheet}"
            all_data.append(df_block)

        except Exception:
            errors.append(f"{file_name} / {sheet}@{s}-{e}")

    # =====================================================
    # ↓↓↓ ここから下は【一切変更していません】 ↓↓↓
    # =====================================================
    if all_data:
        res = pd.concat(all_data, ignore_index=True)

        # ★ 必須：結果を保存
        st.session_state["result_df"] = res

        # 収支混在モードでは、集計直後の収支区分は一括編集で設定する前提とする。
        error_map = validate_freee_import_data(
            res,
            validate_trade=(trade_mode != "収支混在"),
            validate_all_required=False,
        )

        st.session_state["error_map"] = error_map

        st.session_state["work_df"] = res.copy()
        # 任意項目の選択状態を、集計後の一覧・一括編集・出力で一貫して使用する。
        st.session_state["active_optional_fields"] = selected_optional_fields.copy()
        st.session_state["selected_ids"] = set()
        st.session_state["grid_revision"] += 1
        st.session_state["phase"] = "result"

        # freee想定の必須列は必ず作る（空OK）
        required_cols = [
            "収支区分",
            "発生日",
            "勘定科目",
            "税区分",
            "金額",
            "税計算区分",
            "税額",
        ]

        for col in required_cols:
            if col not in res.columns:
                res[col] = ""

    else:
        st.warning("集計できるデータがありません")

    if errors:
        st.error("集計失敗ブロック:")
        st.write(errors)
    
if st.session_state.get("phase") == "result":
    error_map = st.session_state.get("error_map", {})
    active_optional_fields = set(st.session_state.get("active_optional_fields", []))
    show_tax_calculation = "税計算区分" in active_optional_fields
    show_tax_amount = "税額" in active_optional_fields

    # 表示ビューで編集した後の最終エラーも、該当セルを赤く表示する。
    preview_out = st.session_state["work_df"].copy()
    preview_final_error_map = validate_freee_import_data(preview_out, validate_trade=True)
    view_edited_error_keys = get_view_edited_error_keys(
        st.session_state.get("result_df"),
        preview_out,
        preview_final_error_map,
    )
    # 集計時のエラーも、表示ビューで修正済みなら赤表示から外す。
    display_error_map = {
        error_key: messages
        for error_key, messages in error_map.items()
        if error_key in preview_final_error_map
    }
    for error_key in view_edited_error_keys:
        display_error_map[error_key] = preview_final_error_map[error_key]

    st.subheader("📋 集計データ（直接編集 / フィルタ / 一括編集）")

    # --- 作業用DF 初期化 ---
    if "work_df" not in st.session_state:
        st.session_state["work_df"] = st.session_state["result_df"].copy()
    if "selected_ids" not in st.session_state:
        st.session_state["selected_ids"] = set()

    df = st.session_state["work_df"].copy()
    df["__idx__"] = df.index

    # =========================
    # フィルタUI
    # =========================
    with st.expander("🔎 フィルタ", expanded=True):
        col_search, col_restore = st.columns([4,1])

        with col_search:
            q = st.text_input("全体検索（含む）", value="", key="filter_q")

        with col_restore:
            search_word = q.strip()
            if search_word and "search_history" in st.session_state:
                if search_word in st.session_state["search_history"]:
                    if st.button("前回履歴を自動入力"):
                        history = st.session_state["search_history"][search_word]

                        st.session_state["bulk_trade"] = history.get("bulk_trade","")
                        st.session_state["bulk_account"] = history.get("bulk_account","")
                        st.session_state["bulk_tax"] = history.get("bulk_tax","")
                        st.session_state["bulk_tax_rate"] = history.get("bulk_tax_rate","")
                        st.session_state["bulk_rows"] = history.get("bulk_rows",1)

                        for k,v in history.items():
                            if k.startswith("bulk_"):
                                st.session_state[k] = v

                        st.rerun()
        fcols = st.multiselect(
            "列フィルタ（値で絞る）",
            options=[
                c for c in df.columns
                if not c.startswith("__")
                and c not in ["シート名"]
                and (c != "税計算区分" or show_tax_calculation)
                and (c != "税額" or show_tax_amount)
                and c != "__idx__"
            ],
            key="filter_cols"
        )

        selected_vals = {}
        for c in fcols:
            vals = df[c].dropna().astype(str).unique().tolist()
            vals = sorted(vals)[:300]
            selected_vals[c] = st.multiselect(
                f"{c} の値",
                options=vals,
                key=f"filter_vals_{c}"
            )
        

    fdf = df.copy()
    # -------------------------
    # エラー行のみ表示ボタン
    # -------------------------
    show_error_only = st.checkbox("⚠ エラー行のみ表示", key="show_error_only")

    if show_error_only:
        error_rows = {idx for (_, idx, _, _) in display_error_map.keys()}

        if error_rows:
            fdf = fdf.loc[fdf.index.intersection(error_rows)]
        else:
            fdf = fdf.iloc[0:0]  # 空表示

    if q.strip():
        qq = q.strip()
        mask = fdf.astype(str).apply(
            lambda s: s.str.contains(qq, na=False)
        ).any(axis=1)
        fdf = fdf[mask]

    for c, vs in selected_vals.items():
        if vs:
            fdf = fdf[fdf[c].astype(str).isin(vs)]

    # =========================
    # 一覧（直接編集・行選択）
    # =========================
    row_ids = fdf.index.tolist()
    grid_key = f"grid_editor_{st.session_state['grid_revision']}"
    view = fdf.copy()
    view.insert(0, "選択", [idx in st.session_state["selected_ids"] for idx in fdf.index])
    view = view.drop(columns=[ "シート名"], errors="ignore")
    view = view.drop(columns=["__idx__"], errors="ignore")
    hidden_tax_columns = []
    if not show_tax_calculation:
        hidden_tax_columns.append("税計算区分")
    if not show_tax_amount:
        hidden_tax_columns.append("税額")
    view = view.drop(columns=hidden_tax_columns, errors="ignore")
    view = view.drop(
    columns=[c for c in view.columns if c.startswith("__")],
    errors="ignore",)

    def highlight_error_cells(data):
        styles = pd.DataFrame("", index=data.index, columns=data.columns)
        for (_, idx, col, _) in display_error_map.keys():
            if idx in styles.index and col in styles.columns:
                styles.at[idx, col] = "background-color:#ffe6e6; color:#b30000; font-weight:bold;"
        return styles

    # ★ 初回だけ「選択」列を追加（state からは復元しない）
    if "選択" not in view.columns:
        view.insert(0, "選択", False)

    # 表の左上のチェックを切り替えた時だけ、表示中の行を全選択・全選択解除する。
    def toggle_all_rows():
        st.session_state["selected_ids"] = set(row_ids) if st.session_state["select_all_rows"] else set()
        st.session_state["grid_revision"] += 1

    st.checkbox("全選択 / 全選択解除", key="select_all_rows", on_change=toggle_all_rows)

    if AGGRID_AVAILABLE:
        grid_view = view.copy()
        # グリッド側の数値変換によるオーバーフローを防ぐため、表示値は文字列で渡す。
        for column in [column for column in grid_view.columns if column != "選択"]:
            grid_view[column] = grid_view[column].map(
                lambda value: "" if pd.isna(value) else str(value)
            )
        error_columns_by_row = defaultdict(set)
        for (_, idx, column, _) in display_error_map.keys():
            error_columns_by_row[idx].add(column)
        # 赤く表示する列名だけをグリッド内部へ渡し、出力データには含めない。
        grid_view["__error_columns__"] = [
            "|".join(sorted(error_columns_by_row.get(idx, set())))
            for idx in grid_view.index
        ]

        error_cell_style = JsCode(
            """
            function(params) {
                const errorColumns = String(params.data.__error_columns__ || "").split("|");
                if (errorColumns.includes(params.colDef.field)) {
                    return {backgroundColor: "#ffe6e6", color: "#b30000", fontWeight: "bold"};
                }
                return null;
            }
            """
        )
        grid_builder = GridOptionsBuilder.from_dataframe(grid_view)
        grid_builder.configure_default_column(editable=True, resizable=True, filterable=False, sortable=False)
        grid_builder.configure_column(
            "選択",
            editable=True,
            cellEditor="agCheckboxCellEditor",
            cellRenderer="agCheckboxCellRenderer",
            width=80,
        )
        grid_builder.configure_column("__error_columns__", hide=True, editable=False)
        for column in [column for column in view.columns if column != "選択"]:
            column_width = 320 if column in ["備考", "メモタグ（複数指定可、カンマ区切り）"] else 200
            grid_builder.configure_column(
                column,
                editable=True,
                cellStyle=error_cell_style,
                width=column_width,
                minWidth=140,
                wrapText=False,
                autoHeight=False,
                wrapHeaderText=False,
                autoHeaderHeight=False,
                tooltipField=column,
                # 行高を一定に保ち、長い値は列幅・横スクロール・ツールチップで確認する。
            )

        grid_response = AgGrid(
            grid_view,
            gridOptions=grid_builder.build(),
            update_on=["cellValueChanged"],
            allow_unsafe_jscode=True,
            theme="streamlit",
            height=min(max(220, 38 * (len(grid_view) + 2)), 600),
            key=grid_key,
        )
        edited = grid_response.data if hasattr(grid_response, "data") else grid_response["data"]
        edited = edited[view.columns]
        comparison_view = grid_view[view.columns]
    else:
        styled_view = view.style.apply(highlight_error_cells, axis=None)
        st.warning("直接編集と赤セル表示には streamlit-aggrid のインストールが必要です。")
        edited = st.data_editor(
            styled_view,
            use_container_width=True,
            num_rows="fixed",
            key=grid_key
        )
        comparison_view = view
    # =========================
    # エラー表示（表の下）
    # =========================
    if display_error_map:
        st.error("入力エラーがあります。赤く表示されたセルを修正してください。")
        st.caption("「⚠ エラー行のみ表示」をONにすると、エラーのある行だけを確認できます。")
        st.dataframe(
            build_error_display_dataframe(
                st.session_state["work_df"],
                display_error_map,
                view_edited_error_keys,
            ),
            use_container_width=True,
        )

    # ★ data_editor の結果だけを state に保存（次フレーム用）
    selected_ids = {
        row_ids[i]
        for i, v in enumerate(edited["選択"])
        if v
    }
    st.session_state["selected_ids"] = selected_ids

    editable_columns = [column for column in view.columns if column != "選択"]
    if not edited[editable_columns].equals(comparison_view[editable_columns]):
        # 表示ビューでの直接編集を作業データへ反映し、最終バリデーションへ渡す。
        st.session_state["work_df"].loc[row_ids, editable_columns] = edited[editable_columns].to_numpy()
        # 再検証後の赤セル表示を確実に反映するため、グリッドを作り直す。
        st.session_state["grid_revision"] += 1
        st.rerun()

    st.divider()

    # =========================
    # 下段：上段に一括編集、下段にエクスポートを縦に表示する。
    # =========================
    col_edit = st.container()
    col_dl = st.container()

    # -------- 一括編集 --------
    with col_edit:

        # --- 収支区分の一括編集（プルダウン）---
        col_title, col_reset = st.columns([4,1])
        with col_title:
            st.subheader("✏️ 収支区分（選択行に適用）")

        with col_reset:
            if st.button("🔄 リセット"):
                st.session_state["bulk_trade"] = ""
                st.session_state["bulk_account"] = ""
                st.session_state["bulk_tax"] = ""
                st.session_state["bulk_tax_rate"] = ""
                st.session_state["bulk_tax_calc"] = ""

                rows = st.session_state.get("bulk_rows", 1)

                for i in range(rows):
                    st.session_state[f"bulk_targets_{i}"] = []
                    st.session_state[f"bulk_single_{i}"] = ""
                    st.session_state[f"bulk_income_{i}"] = ""
                    st.session_state[f"bulk_expense_{i}"] = ""

                st.session_state["bulk_rows"] = 1

                st.rerun()
        bulk_trade = st.selectbox(" ", ["", "収入", "支出"], index=0, key="bulk_trade") 

        # =========================
        # 勘定科目（選択行に適用）
        # =========================
        st.subheader("📘 勘定科目（選択行に適用）")

        account_options = []

        if os.path.exists("master_accounts.csv"):
            account_df = pd.read_csv("master_accounts.csv")
            account_options = account_df["勘定科目"].dropna().astype(str).tolist()

        bulk_account = st.selectbox(
            " ",
            [""] + account_options,
            index=0,
            key="bulk_account"
        )

        # =========================
        # 税区分（選択行に適用）
        # =========================
        st.subheader("🧾 税区分（選択行に適用）")

        tax_options = st.session_state.get("enabled_tax_list", [])
        tax_rates = st.session_state.get("enabled_tax_rates", [])

        col_tax, col_rate = st.columns(2)

        with col_tax:
            bulk_tax = st.selectbox(
                "税区分",
                [""] + tax_options,
                key="bulk_tax"
            )

        with col_rate:
            bulk_tax_rate = st.selectbox(
                "税率",
                [""] + tax_rates,
                key="bulk_tax_rate"
            )

        bulk_tax_calc = ""
        if show_tax_calculation:
            st.subheader("🧮 税計算区分（選択行に適用）")
            bulk_tax_calc = st.selectbox(
                " ",
                ["", "内税", "外税", "税込"],
                key="bulk_tax_calc"
            )

        st.subheader("✏️ 選択行の一括編集")

        if "bulk_rows" not in st.session_state: st.session_state.bulk_rows = 1  # ← 追加

        for i in range(st.session_state.bulk_rows):

            targets = st.multiselect(
                f"変更する列 ({i+1})",
                options=[
                    c for c in df.columns
                    if not c.startswith("__")
                    and c not in ["__idx__", "シート名"]
                    and (c != "税計算区分" or show_tax_calculation)
                    and (c != "税額" or show_tax_amount)
                ],
                key=f"bulk_targets_{i}"
            )

            is_mixed = False
            if "収支区分" in st.session_state["work_df"].columns and st.session_state.get("selected_ids"):
                selected_df = st.session_state["work_df"].loc[list(st.session_state["selected_ids"])]
                if selected_df["収支区分"].nunique() > 1:
                    is_mixed = True

            if is_mixed:
                st.text_input("収入行に適用する値", key=f"bulk_income_{i}")
                st.text_input("支出行に適用する値", key=f"bulk_expense_{i}")
            else:
                st.text_input("適用する値", key=f"bulk_single_{i}")

        if st.button("＋ 項目を追加"):
            st.session_state.bulk_rows += 1
            st.rerun()



        if st.button("🧨 選択行に適用", key="apply_bulk", use_container_width=True):

            ids = list(st.session_state.get("selected_ids", []))
            if ids:
                selected_df = st.session_state["work_df"].loc[ids]
            else:
                selected_df = pd.DataFrame()

            if not ids:
                st.warning("行が選択されていません")

            any_targets = any(
                st.session_state.get(f"bulk_targets_{i}", [])
                for i in range(st.session_state.get("bulk_rows", 1))
            )

            if not any_targets and not bulk_trade and not bulk_account and not bulk_tax and not bulk_tax_rate and not bulk_tax_calc:
                st.warning("変更する項目を選択してください")

            # 警告時は一括編集のみ中止し、下段のエクスポート欄は描画を継続する。
            can_apply = bool(ids) and (
                any_targets or bulk_trade or bulk_account or bulk_tax or bulk_tax_rate or bulk_tax_calc
            )

            base = st.session_state["work_df"].copy()

            # --- ① 収支区分（ボタン押したときだけ反映） ---
            if bulk_trade:
                for idx in ids:
                    base.at[idx, "収支区分"] = bulk_trade

            # --- ② 通常列（複数セット対応） ---
            for i in range(st.session_state.get("bulk_rows", 1)):

                targets_i = st.session_state.get(f"bulk_targets_{i}", [])

                if not targets_i:
                    continue

                if is_mixed:
                    income_value_i = st.session_state.get(f"bulk_income_{i}", "")
                    expense_value_i = st.session_state.get(f"bulk_expense_{i}", "")

                    for idx in ids:
                        for col in targets_i:
                            if base.at[idx, "収支区分"] == "収入":
                                base.at[idx, col] = income_value_i
                            elif base.at[idx, "収支区分"] == "支出":
                                base.at[idx, col] = expense_value_i

                else:
                    bulk_value_i = st.session_state.get(f"bulk_single_{i}", "")
                    for col in targets_i:
                        base[col] = base[col].astype("object")
                        base.loc[ids, col] = bulk_value_i
            
            # --- 検索ワードと編集履歴を保存 ---
            search_word = st.session_state.get("filter_q", "").strip()

            if can_apply and search_word:
                if "search_history" not in st.session_state:
                    st.session_state["search_history"] = {}

                st.session_state["search_history"][search_word] = {
                    "bulk_trade": bulk_trade,
                    "bulk_account": bulk_account,
                    "bulk_tax": bulk_tax,
                    "bulk_tax_rate": bulk_tax_rate,
                    "bulk_rows": st.session_state.get("bulk_rows", 1),
                }

                # 各行の個別入力も保存
                for i in range(st.session_state.get("bulk_rows", 1)):
                    st.session_state["search_history"][search_word][f"bulk_targets_{i}"] = st.session_state.get(f"bulk_targets_{i}", [])
                    st.session_state["search_history"][search_word][f"bulk_single_{i}"] = st.session_state.get(f"bulk_single_{i}", "")


            # --- 勘定科目 ---
            if bulk_account:
                for idx in ids:
                    base.at[idx, "勘定科目"] = bulk_account

            # --- 税区分 ---
            if bulk_tax:
                final_tax = bulk_tax

                if bulk_tax_rate:
                    final_tax = f"{bulk_tax}{bulk_tax_rate}"

                for idx in ids:
                    base.at[idx, "税区分"] = final_tax
            
            if bulk_tax_calc:
                for idx in ids:
                    base.at[idx, "税計算区分"] = bulk_tax_calc

            # 編集結果を保存してから画面を再実行する。
            if can_apply:
                st.session_state["work_df"] = base
                st.session_state["selected_ids"] = set()
                # 一括編集後はグリッドを作り直し、表示の古い状態を残さない。
                st.session_state["grid_revision"] += 1
                st.success(f"{len(ids)} 行を更新しました")
                st.rerun()

    # -------- エクスポート --------
    with col_dl:
        st.divider()
        st.subheader("📤 エクスポート（編集後）")

        out = st.session_state["work_df"].copy()
        # 出力直前に再検証し、一括編集後の値もfreee取込条件に合わせて確認する。
        final_error_map = validate_freee_import_data(out, validate_trade=True)
        view_edited_error_keys = get_view_edited_error_keys(
            st.session_state.get("result_df"),
            out,
            final_error_map,
        )
        # 集計ビューは集計時のエラーだけを表示し、最終チェック結果で上書きしない。

        if final_error_map:
            st.error("freee取込に必要な項目または形式にエラーがあります。修正後に出力してください。")
            st.dataframe(
                build_error_display_dataframe(out, final_error_map, view_edited_error_keys),
                use_container_width=True,
            )
        else:
            # 取込元の追跡情報は画面内だけで使い、freee取込用の出力には含めない。
            out = out.drop(
                columns=[c for c in out.columns if c.startswith("__")] + ["シート名"],
                errors="ignore",
            )
            # 未選択の任意項目は全行空欄なら出力列ごと省き、通常のCSVを簡潔にする。
            for optional_column in ["税計算区分", "税額"]:
                if (
                    optional_column in out.columns
                    and optional_column not in active_optional_fields
                    and out[optional_column].fillna("").astype(str).str.strip().eq("").all()
                ):
                    out = out.drop(columns=[optional_column])

            st.download_button(
                "📄 CSVダウンロード",
                out.to_csv(index=False, encoding="utf-8-sig"),
                "result_edited.csv",
                "text/csv",
                use_container_width=True
            )

            buf = BytesIO()
            out.to_excel(buf, index=False)
            buf.seek(0)

            wb = load_workbook(buf)
            ws = wb.active
            no_border = Border()
            for row in ws.iter_rows():
                for cell in row:
                    cell.border = no_border

            buf2 = BytesIO()
            wb.save(buf2)
            buf2.seek(0)

            st.download_button(
                "📊 Excelダウンロード",
                buf2,
                "result_edited.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
