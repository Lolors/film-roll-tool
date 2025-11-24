import os
import math
import re
import statistics
import pandas as pd
import streamlit as st

# =========================================
# 설정
# =========================================
BOM_FILE = "BOM.xlsx"           # 품번/품명 참조용 파일
BOM_SHEET = "Sheet1"            # 시트 이름
CONFIG_FILE = "film_config.csv" # 품번별 필름 조건 저장
THICKNESS_FILE = "film_thickness.csv"  # 두께 9회 측정 결과 저장


# =========================================
# 데이터 로드 / 저장
# =========================================
@st.cache_data
def load_bom():
    """BOM에서 품번/품명 로드 (C열 품번, D열 품명 사용)"""
    if not os.path.exists(BOM_FILE):
        st.error(f"'{BOM_FILE}' 파일을 찾을 수 없어. 같은 폴더에 BOM 파일을 둬줘.")
        return pd.DataFrame(columns=["품번", "품명"])

    try:
        df = pd.read_excel(BOM_FILE, sheet_name=BOM_SHEET)
    except Exception as e:
        st.error(f"BOM 파일 읽는 중 오류: {e}")
        return pd.DataFrame(columns=["품번", "품명"])

    # C열: 품번, D열: 품명(D열 헤더가 '품명.1'이라고 가정)
    if "품번" not in df.columns or "품명.1" not in df.columns:
        st.error("BOM 파일에서 '품번'과 '품명.1' 컬럼을 찾지 못했어.")
        return pd.DataFrame(columns=["품번", "품명"])

    bom = df[["품번", "품명.1"]].dropna(subset=["품번"])
    bom = bom.drop_duplicates(subset=["품번"])
    bom = bom.rename(columns={"품명.1": "품명"})
    return bom[["품번", "품명"]]


def load_config():
    """저장된 필름 조건 로드"""
    if not os.path.exists(CONFIG_FILE):
        return pd.DataFrame(columns=[
            "품번", "품명",
            "필름두께_mm", "지관외경_cm",
            "아이마크세트길이_cm", "세트당라벨수"
        ])
    try:
        df = pd.read_csv(CONFIG_FILE, encoding="utf-8-sig")
    except Exception:
        df = pd.read_csv(CONFIG_FILE)
    return df


def save_config(df: pd.DataFrame):
    """필름 조건 저장"""
    df.to_csv(CONFIG_FILE, index=False, encoding="utf-8-sig")
    st.session_state["config_df"] = df


def load_thickness():
    """두께 9회 측정 데이터 로드"""
    if not os.path.exists(THICKNESS_FILE):
        return pd.DataFrame(columns=[
            "품번", "품명", "거래처",
            "측정1", "측정2", "측정3",
            "측정4", "측정5", "측정6",
            "측정7", "측정8", "측정9",
            "평균", "표준편차"
        ])
    try:
        df = pd.read_csv(THICKNESS_FILE, encoding="utf-8-sig")
    except Exception:
        df = pd.read_csv(THICKNESS_FILE)
    return df


def save_thickness(df: pd.DataFrame):
    """두께 9회 측정 데이터 저장"""
    df.to_csv(THICKNESS_FILE, index=False, encoding="utf-8-sig")
    st.session_state["thick_df"] = df


# =========================================
# 계산 함수 (엑셀 수식 그대로)
# INT((PI()*(((E/100)^2 - (F/100)^2)/(4*(D/1000)))) / (G/100)) * H
# =========================================
def calc_labels_per_roll(thickness_mm, roll_diam_cm, core_diam_cm,
                         mark_set_cm, labels_per_set):
    if (thickness_mm is None or thickness_mm <= 0 or
        roll_diam_cm is None or roll_diam_cm <= 0 or
        core_diam_cm is None or core_diam_cm <= 0 or
        mark_set_cm is None or mark_set_cm <= 0 or
        labels_per_set is None or labels_per_set <= 0):
        return 0

    if roll_diam_cm <= core_diam_cm:
        return 0

    try:
        film_length_m = math.pi * (((roll_diam_cm / 100) ** 2 - (core_diam_cm / 100) ** 2) /
                                   (4 * (thickness_mm / 1000)))
        sets = film_length_m / (mark_set_cm / 100)
        labels = int(sets) * int(labels_per_set)
        return int(labels)
    except Exception:
        return 0


# =========================================
# Streamlit 앱
# =========================================
st.set_page_config(page_title="필름 관리 도구", layout="wide")
st.title("🎞 필름 관리 도구")

bom_df = load_bom()
if bom_df.empty:
    st.stop()

if "config_df" not in st.session_state:
    st.session_state["config_df"] = load_config()
if "thick_df" not in st.session_state:
    st.session_state["thick_df"] = load_thickness()

config_df = st.session_state["config_df"]
thick_df = st.session_state["thick_df"]

품번_list = bom_df["품번"].astype(str).sort_values().tolist()

tab1, tab2 = st.tabs(["1롤 수량 계산", "필름 두께 측정/평균"])

# =========================================
# TAB 1 : 1롤 수량 계산기
# =========================================
with tab1:
    st.markdown("### 1️⃣ 품번 선택")

    selected_pumbun = st.selectbox("BOM에서 품번 선택", 품번_list, key="tab1_pumbun")
    row = bom_df[bom_df["품번"].astype(str) == str(selected_pumbun)]
    품명 = row["품명"].iloc[0] if not row.empty else ""
    st.write(f"**품명:** {품명}")

    st.markdown("### 2️⃣ 이 품번의 필름 조건 설정")

    # 기존 설정 불러오기
    exist = config_df[config_df["품번"].astype(str) == str(selected_pumbun)]
    if not exist.empty:
        default_thickness = float(exist["필름두께_mm"].iloc[0])
        default_core_d = float(exist["지관외경_cm"].iloc[0])
        default_mark_set = float(exist["아이마크세트길이_cm"].iloc[0])
        default_labels_per_set = int(exist["세트당라벨수"].iloc[0])
    else:
        default_thickness = 0.135
        default_core_d = 9.0
        default_mark_set = 11.45
        default_labels_per_set = 5

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        thickness_mm = st.number_input(
            "필름 두께 (mm)",
            min_value=0.001,
            step=0.001,
            format="%.3f",
            value=default_thickness,
            key=f"thk_{selected_pumbun}",
        )
    with c2:
        core_diam_cm = st.number_input(
            "지관 외경 (cm)",
            min_value=0.1,
            step=0.1,
            format="%.1f",
            value=default_core_d,
            key=f"core_{selected_pumbun}",
        )
    with c3:
        mark_set_cm = st.number_input(
            "아이마크 세트 길이 (cm)",
            min_value=0.1,
            step=0.01,
            format="%.2f",
            value=default_mark_set,
            key=f"mark_{selected_pumbun}",
        )
    with c4:
        labels_per_set = st.number_input(
            "세트당 라벨 개수 (장)",
            min_value=1,
            step=1,
            value=default_labels_per_set,
            key=f"lps_{selected_pumbun}",
        )

    if st.button("💾 이 품번 설정 저장하기", key="save_cfg"):
        new_row = {
            "품번": selected_pumbun,
            "품명": 품명,
            "필름두께_mm": thickness_mm,
            "지관외경_cm": core_diam_cm,
            "아이마크세트길이_cm": mark_set_cm,
            "세트당라벨수": labels_per_set,
        }

        if not exist.empty:
            idx = config_df[config_df["품번"].astype(str) == str(selected_pumbun)].index
            config_df.loc[idx, :] = new_row
        else:
            config_df = pd.concat([config_df, pd.DataFrame([new_row])], ignore_index=True)

        save_config(config_df)
        st.success("이 품번의 필름 설정을 저장했어!")

    st.markdown("### 3️⃣ 실물 직경별 1롤 수량 계산")

    st.caption("쉼표(,)나 줄바꿈으로 여러 개 입력할 수 있어. 예: `29.9, 29.8, 26.8`")

    diam_raw = st.text_area(
        "실물 직경 목록 (cm)",
        height=100,
        placeholder="예) 29.9, 29.8, 26.8",
    )

    diam_list = []
    if diam_raw.strip():
        tokens = re.split(r"[,\s]+", diam_raw.strip())
        for t in tokens:
            if not t:
                continue
            try:
                d = float(t)
                diam_list.append(d)
            except ValueError:
                st.warning(f"숫자로 인식할 수 없는 값이라 무시했어: {t}")

    if (diam_list and thickness_mm > 0 and core_diam_cm > 0
            and mark_set_cm > 0 and labels_per_set > 0):
        rows = []
        for d in diam_list:
            qty = calc_labels_per_roll(
                thickness_mm, d, core_diam_cm, mark_set_cm, labels_per_set
            )
            rows.append({
                "실물 직경 (cm)": d,
                "1롤 수량 (개)": qty,
            })

        result_df = pd.DataFrame(rows)
        st.dataframe(result_df, use_container_width=True)
    else:
        st.info("직경 목록을 입력하면 이 아래에 직경별 1롤 수량이 계산돼.")

    with st.expander("📁 저장된 품번별 필름 조건 보기"):
        if config_df.empty:
            st.write("아직 저장된 설정이 없어.")
        else:
            st.dataframe(config_df, use_container_width=True)


# =========================================
# TAB 2 : 필름 두께 9회 측정 / 평균
# =========================================
with tab2:
    st.markdown("### 1️⃣ 품번 선택 및 기본 정보")

    selected_pumbun2 = st.selectbox(
        "BOM에서 품번 선택",
        품번_list,
        key="tab2_pumbun"
    )
    row2 = bom_df[bom_df["품번"].astype(str) == str(selected_pumbun2)]
    품명2 = row2["품명"].iloc[0] if not row2.empty else ""
    st.write(f"**필름명:** {품명2}")

    거래처 = st.text_input("거래처", value="", placeholder="예) (주)아이제이팩")

    # 기존 측정값 있으면 불러오기
    exist_t = thick_df[thick_df["품번"].astype(str) == str(selected_pumbun2)]
    if not exist_t.empty:
        base_vals = [
            exist_t["측정1"].iloc[0],
            exist_t["측정2"].iloc[0],
            exist_t["측정3"].iloc[0],
            exist_t["측정4"].iloc[0],
            exist_t["측정5"].iloc[0],
            exist_t["측정6"].iloc[0],
            exist_t["측정7"].iloc[0],
            exist_t["측정8"].iloc[0],
            exist_t["측정9"].iloc[0],
        ]
        base_vals = [float(v) if pd.notna(v) else 0.0 for v in base_vals]
        base_vendor = exist_t["거래처"].iloc[0]
        if not 거래처:
            거래처 = base_vendor
    else:
        base_vals = [0.0] * 9

    st.markdown("### 2️⃣ 두께 9회 측정값 입력 (mm)")

    inputs = []
    labels = ["1차측정", "2차측정", "3차측정",
              "4차측정", "5차측정", "6차측정",
              "7차측정", "8차측정", "9차측정"]

    # 3개씩 나눠서 입력 (3열 × 3행)
    idx = 0
    for _ in range(3):
        cols = st.columns(3)
        for c in cols:
            val = c.number_input(
                labels[idx],
                min_value=0.0,
                step=0.001,
                format="%.3f",
                value=base_vals[idx],
                key=f"t_{selected_pumbun2}_{idx}",
            )
            inputs.append(val)
            idx += 1

    # 0보다 큰 값만 유효 측정으로 간주
    valid_vals = [v for v in inputs if v > 0]

    if valid_vals:
        avg = sum(valid_vals) / len(valid_vals)
        if len(valid_vals) > 1:
            std = statistics.stdev(valid_vals)   # 샘플 표준편차
        else:
            std = 0.0
    else:
        avg = 0.0
        std = 0.0

    st.markdown("### 3️⃣ 결과")
    st.write(f"**평균 두께:** {avg:.3f} mm")
    st.write(f"**표준편차:** {std:.6f} mm")

    if st.button("💾 이 품번의 두께 측정값 저장하기", key="save_thickness"):
        new_row_t = {
            "품번": selected_pumbun2,
            "품명": 품명2,
            "거래처": 거래처,
            "측정1": inputs[0],
            "측정2": inputs[1],
            "측정3": inputs[2],
            "측정4": inputs[3],
            "측정5": inputs[4],
            "측정6": inputs[5],
            "측정7": inputs[6],
            "측정8": inputs[7],
            "측정9": inputs[8],
            "평균": avg,
            "표준편차": std,
        }

        if not exist_t.empty:
            idx_t = thick_df[thick_df["품번"].astype(str) == str(selected_pumbun2)].index
            thick_df.loc[idx_t, :] = new_row_t
        else:
            thick_df = pd.concat([thick_df, pd.DataFrame([new_row_t])], ignore_index=True)

        save_thickness(thick_df)
        st.success("이 품번의 두께 측정 정보를 저장했어!")

    st.markdown("### 4️⃣ 저장된 두께 측정 결과")

    if thick_df.empty:
        st.write("아직 저장된 두께 측정 데이터가 없어.")
    else:
        st.dataframe(thick_df, use_container_width=True)
