from __future__ import annotations

import json
import math
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

import pandas as pd
import streamlit as st


APP_NAME = "Kadastro Harc Hesaplama"
YEAR = 2026
MIN_TALEBE_BAGLI = 1504.0
MIN_KONTROLLUK = 2660.0


def resource_path(relative: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative


def user_data_dir() -> Path:
    root = os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")
    path = Path(root) / "KadastroHarc"
    path.mkdir(parents=True, exist_ok=True)
    return path


OVERRIDES_FILE = user_data_dir() / "ayarlar.json"


def tr_money(value: float) -> str:
    text = f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{text} TL"


def ceil_to_lira(value: float) -> float:
    return float(round(value))


def normalize_area(value: float) -> float:
    return max(float(value or 0), 0.0)


def stepped_area(area_m2: float, brackets: list[tuple[float | None, float, str]]) -> tuple[float, list[str]]:
    area = normalize_area(area_m2)
    previous = 0.0
    total = 0.0
    details: list[str] = []
    for upper, unit_price, unit in brackets:
        limit = area if upper is None else min(area, upper)
        amount = max(limit - previous, 0.0)
        if amount <= 0:
            previous = upper or previous
            continue
        divisor = 10000.0 if unit == "ha" else 1000.0
        line_total = amount / divisor * unit_price
        total += line_total
        label = "ha" if unit == "ha" else "1.000 m2"
        details.append(f"{amount:,.2f} m2 / {label} x {tr_money(unit_price)} = {tr_money(line_total)}")
        previous = limit
        if upper is not None and area <= upper:
            break
    return total, details


def parselasyon_fee(area_m2: float, yk: float, geriye_donus: bool = False, yeni_uygulama: bool = False) -> tuple[float, list[str]]:
    base, details = stepped_area(
        area_m2,
        [
            (100000, 3177.0, "ha"),
            (500000, 2678.0, "ha"),
            (1000000, 2347.0, "ha"),
            (1500000, 1354.0, "ha"),
            (None, 338.0, "ha"),
        ],
    )
    if geriye_donus and not yeni_uygulama:
        details.append("Geriye dönüş: hesaplanan bedelin %25'i")
        base *= 0.25
    elif geriye_donus and yeni_uygulama:
        details.append("Yeni uygulama ile geriye dönüş: kontrollük bedeline ek %25")
        base *= 1.25
    return base * yk, details


def imar_degisiklik_fee(area_m2: float, yk: float, parsel_count: int, mera_25: bool) -> tuple[float, list[str]]:
    area = normalize_area(area_m2)
    if area <= 3000:
        base = 2106.0
        details = [f"İlk dilim maktu = {tr_money(base)}"]
    else:
        base, details = stepped_area(
            area,
            [
                (3000, 2106.0 / 3.0, "1000"),
                (5000, 665.0, "1000"),
                (10000, 498.0, "1000"),
                (100000, 420.0, "ha"),
                (2000000, 338.0, "ha"),
            ],
        )
    extra = max(int(parsel_count or 0) - 2, 0) * 461.0
    if extra:
        details.append(f"İkiden fazla parsel ilavesi = {tr_money(extra)}")
        base += extra
    if mera_25:
        details.append("Mera Kanunu özel durumu: %25 tahsil")
        base *= 0.25
    return base * yk, details


def kadastro_degisiklik_fee(area_m2: float, yk: float, parsel_count: int, mera_25: bool) -> tuple[float, list[str]]:
    area = normalize_area(area_m2)
    if area <= 20000:
        base = 1772.0
        details = [f"İlk dilim maktu = {tr_money(base)}"]
    else:
        base, details = stepped_area(
            area,
            [
                (20000, 1772.0 / 2.0, "ha"),
                (100000, 240.0, "ha"),
                (4000000, 143.0, "ha"),
            ],
        )
    extra = max(int(parsel_count or 0) - 2, 0) * 461.0
    if extra:
        details.append(f"İkiden fazla parsel ilavesi = {tr_money(extra)}")
        base += extra
    if mera_25:
        details.append("Mera Kanunu özel durumu: %25 tahsil")
        base *= 0.25
    return base * yk, details


def kamulastirma_serit_fee(km: float, include_docs: bool) -> tuple[float, list[str]]:
    effective_km = max(float(km or 0), 1.0)
    control = effective_km * 4028.0
    total = control
    details = [f"Şeritvari kontrol: {effective_km:g} km x {tr_money(4028)} = {tr_money(control)}"]
    if include_docs:
        docs = effective_km * 892.0
        total += docs
        details.append(f"Harita bilgi ve belge: {effective_km:g} km x {tr_money(892)} = {tr_money(docs)}")
    return total, details


def kamulastirma_hektar_fee(area_m2: float, yk: float, include_docs: bool, mera_25: bool) -> tuple[float, list[str]]:
    ha = max(normalize_area(area_m2) / 10000.0, 1.0)
    ha = min(ha, 100.0)
    control = ha * 892.0
    total = control
    details = [f"Hektar bazlı kontrol: {ha:g} ha x {tr_money(892)} = {tr_money(control)}"]
    if include_docs:
        docs = ha * 338.0
        total += docs
        details.append(f"Harita bilgi ve belge: {ha:g} ha x {tr_money(338)} = {tr_money(docs)}")
    if mera_25:
        details.append("Mera Kanunu özel durumu: %25 tahsil")
        total *= 0.25
    return total * yk, details


@dataclass(frozen=True)
class TechnicalItem:
    label: str
    unit: str
    price: float


TECHNICAL_FIXED = [
    TechnicalItem("Ölçü krokileri", "Sayfa", 307.0),
    TechnicalItem("Takeometrik ölçüm karneleri", "Sayfa", 143.0),
    TechnicalItem("Aplikasyon krokisi", "Sayfa", 143.0),
    TechnicalItem("Ebatlı kroki ve röperli kroki", "Sayfa", 143.0),
    TechnicalItem("Mat kopya (50x70)", "Adet", 518.0),
    TechnicalItem("Mat kopya (70x100)", "Adet", 720.0),
    TechnicalItem("Şeffaf kopya (50x70)", "Adet", 720.0),
    TechnicalItem("Şeffaf kopya (70x100)", "Adet", 955.0),
    TechnicalItem("TUTGA noktası", "Adet", 1773.0),
    TechnicalItem("C1 derece AGA noktası", "Adet", 1773.0),
    TechnicalItem("C2 derece SGA noktası", "Adet", 1033.0),
    TechnicalItem("C3 derece ASN noktası", "Adet", 778.0),
    TechnicalItem("C4 poligon/fotogrametrik nokta", "Adet", 338.0),
    TechnicalItem("ED-50 ana nirengi noktası", "Adet", 665.0),
    TechnicalItem("ED-50 ara nirengi noktası", "Adet", 338.0),
]


def corner_coordinate_fee(count: int) -> tuple[float, float]:
    if count <= 0:
        return 0.0, 0.0
    if count <= 500:
        return count * 29.0, 29.0
    if count <= 1000:
        return count * 24.0, 24.0
    return count * 11.0, 11.0


@st.cache_data(show_spinner=False)
def load_base_rows() -> list[dict]:
    return json.loads(resource_path("katsayilar.json").read_text(encoding="utf-8"))


def load_settings() -> dict:
    if OVERRIDES_FILE.exists():
        try:
            return json.loads(OVERRIDES_FILE.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return {}
    return {}


def save_settings(settings: dict) -> None:
    OVERRIDES_FILE.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")


def build_rows(settings: dict) -> list[dict]:
    rows = load_base_rows()
    custom = settings.get("custom_rows", [])
    return rows + custom


def make_df(settings: dict) -> pd.DataFrame:
    df = pd.DataFrame(build_rows(settings))
    df["il"] = df["il"].str.upper()
    df["ilce"] = df["ilce"].str.upper()
    return df.sort_values(["il", "ilce"]).reset_index(drop=True)


def apply_theme(theme: str) -> None:
    if theme == "Karanlık":
        css = """
        <style>
        .stApp { background: #101418; color: #eef2f6; }
        [data-testid="stSidebar"], .stTabs [data-baseweb="tab-list"] { background: #151b22; }
        div[data-testid="stMetric"] { background:#17202a; border:1px solid #2b3948; border-radius:8px; padding:14px; }
        .result-card { background:#17202a; border:1px solid #2b3948; border-radius:8px; padding:16px; }
        </style>
        """
    else:
        css = """
        <style>
        .stApp { background: #f6f8fb; }
        div[data-testid="stMetric"] { background:#ffffff; border:1px solid #d8e0ea; border-radius:8px; padding:14px; }
        .result-card { background:#ffffff; border:1px solid #d8e0ea; border-radius:8px; padding:16px; }
        </style>
        """
    st.markdown(css, unsafe_allow_html=True)


def result_panel(title: str, raw_fee: float, final_fee: float, details: list[str], minimum: float | None = None) -> None:
    st.markdown('<div class="result-card">', unsafe_allow_html=True)
    st.subheader(title)
    col1, col2 = st.columns(2)
    col1.metric("Hesaplanan bedel", tr_money(ceil_to_lira(raw_fee)))
    col2.metric("Tahsil edilecek bedel", tr_money(ceil_to_lira(final_fee)))
    if minimum and final_fee > raw_fee:
        st.warning(f"Asgari ücret sınırı uygulandı: {tr_money(minimum)}")
    with st.expander("Hesap detayı", expanded=True):
        for line in details:
            st.write("- " + line)
    st.markdown("</div>", unsafe_allow_html=True)


st.set_page_config(page_title=APP_NAME, page_icon="📐", layout="wide")

settings = load_settings()
theme = st.sidebar.radio("Tema", ["Aydınlık", "Karanlık"], index=0 if settings.get("theme", "Aydınlık") == "Aydınlık" else 1)
settings["theme"] = theme
save_settings(settings)
apply_theme(theme)

st.title("Kadastro Döner Sermaye Harç Hesaplama")
st.caption("2026 yılı tarife cetveli ve yöresel katsayı listesine göre yardımcı hesaplama uygulaması.")

df = make_df(settings)
default_il = settings.get("default_il")

tab_teknik, tab_kontrol, tab_admin = st.tabs(["Teknik Bilgi Belge", "Kontrollük", "Admin / Ayarlar"])

with tab_teknik:
    st.header("Teknik Bilgi ve Belge Hesaplama")
    st.write("Yöresel katsayı uygulanmayan maktu ücret kalemleri.")
    total = 0.0
    rows = []
    for item in TECHNICAL_FIXED:
        c1, c2, c3, c4 = st.columns([4, 1.4, 1.5, 1.7])
        c1.write(item.label)
        c2.write(f"{tr_money(item.price)} / {item.unit}")
        qty = c3.number_input(item.unit, min_value=0, step=1, key=f"tech_{item.label}")
        line_total = qty * item.price
        c4.write(tr_money(line_total))
        total += line_total
        if qty:
            rows.append({"Kalem": item.label, "Miktar": qty, "Birim": item.unit, "Tutar": tr_money(line_total)})

    st.divider()
    c1, c2, c3, c4 = st.columns([4, 1.4, 1.5, 1.7])
    c1.write("Parsel köşe koordinatları")
    corner_count = c3.number_input("Nokta", min_value=0, step=1, key="corner_count", label_visibility="collapsed")
    corner_total, corner_unit = corner_coordinate_fee(corner_count)
    c2.write(f"{tr_money(corner_unit)} / Nokta" if corner_unit else "Dilime göre")
    c4.write(tr_money(corner_total))
    total += corner_total
    if corner_count:
        rows.append({"Kalem": "Parsel köşe koordinatları", "Miktar": corner_count, "Birim": "Nokta", "Tutar": tr_money(corner_total)})

    st.metric("Toplam teknik bilgi belge bedeli", tr_money(ceil_to_lira(total)))
    if rows:
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

with tab_kontrol:
    st.header("Kontrollük Hesaplama")
    iller = sorted(df["il"].unique().tolist())
    default_index = iller.index(default_il) if default_il in iller else 0
    col_il, col_ilce, col_yk = st.columns([2, 2, 1])
    il = col_il.selectbox("İl", iller, index=default_index)
    ilceler = sorted(df.loc[df["il"] == il, "ilce"].unique().tolist())
    ilce = col_ilce.selectbox("İlçe", ilceler)
    yk_row = df[(df["il"] == il) & (df["ilce"] == ilce)].iloc[0]
    yk = float(yk_row["kadastro"])
    col_yk.metric("Kadastro YK", f"{yk:g}")

    calc_type = st.selectbox(
        "İşlem türü",
        [
            "Parselasyon planlarının kontrolü",
            "Değişiklik işlemleri - oluşmuş imar parselleri",
            "Değişiklik işlemleri - kadastro parselleri",
            "Kamulaştırma haritalarının kontrolü",
        ],
    )

    if calc_type == "Parselasyon planlarının kontrolü":
        area = st.number_input("Düzenleme sahası yüzölçümü (m2)", min_value=1.0, value=10000.0, step=100.0)
        col_a, col_b = st.columns(2)
        geriye = col_a.checkbox("Geriye dönüş işlemi")
        yeni = col_b.checkbox("Geriye dönüşle birlikte yeni uygulama")
        if st.button("Kontrollük bedelini hesapla", type="primary"):
            raw, details = parselasyon_fee(area, yk, geriye, yeni)
            final = max(raw, MIN_KONTROLLUK)
            result_panel("Parselasyon sonucu", raw, final, details + [f"Yöresel katsayı: {yk:g}"], MIN_KONTROLLUK)

    elif calc_type.startswith("Değişiklik işlemleri"):
        area = st.number_input("İşleme esas yüzölçümü (m2)", min_value=1.0, value=1000.0, step=100.0)
        parsel_count = st.number_input("İfraz sonucu oluşacak parsel sayısı", min_value=1, value=2, step=1)
        mera = st.checkbox("4342 sayılı Mera Kanunu kapsamı (%25)")
        if st.button("Değişiklik bedelini hesapla", type="primary"):
            calculator: Callable[[float, float, int, bool], tuple[float, list[str]]]
            calculator = imar_degisiklik_fee if "imar" in calc_type else kadastro_degisiklik_fee
            raw, details = calculator(area, yk, parsel_count, mera)
            final = max(raw, MIN_KONTROLLUK)
            result_panel("Değişiklik işlemi sonucu", raw, final, details + [f"Yöresel katsayı: {yk:g}"], MIN_KONTROLLUK)

    else:
        sub_type = st.radio("Kamulaştırma türü", ["Şeritvari bazlı", "Hektar bazlı", "Hektar bazlı - MERA"], horizontal=True)
        include_docs = st.checkbox("Harita bilgi ve belge bedelini dahil et", value=True)
        if sub_type == "Şeritvari bazlı":
            km = st.number_input("Uzunluk (km)", min_value=0.01, value=1.0, step=0.1)
            if st.button("Kamulaştırma bedelini hesapla", type="primary"):
                raw, details = kamulastirma_serit_fee(km, include_docs)
                final = max(raw, MIN_KONTROLLUK)
                result_panel("Şeritvari kamulaştırma sonucu", raw, final, details, MIN_KONTROLLUK)
        else:
            area = st.number_input("Kamulaştırmaya tabi yüzölçümü (m2)", min_value=1.0, value=10000.0, step=100.0)
            if st.button("Kamulaştırma bedelini hesapla", type="primary"):
                raw, details = kamulastirma_hektar_fee(area, yk, include_docs, sub_type.endswith("MERA"))
                final = max(raw, MIN_KONTROLLUK)
                result_panel("Hektar bazlı kamulaştırma sonucu", raw, final, details + [f"Yöresel katsayı: {yk:g}"], MIN_KONTROLLUK)

with tab_admin:
    st.header("Admin / Ayarlar")
    st.write("Varsayılan il seçimi ve özel katsayı kayıtları buradan yönetilir.")
    selected_default = st.selectbox("Açılışta seçilecek il", sorted(df["il"].unique().tolist()), index=default_index)
    if st.button("Varsayılan ili kaydet"):
        settings["default_il"] = selected_default
        save_settings(settings)
        st.success("Varsayılan il kaydedildi.")

    st.subheader("Katsayı ekle / güncelle")
    col1, col2, col3, col4 = st.columns(4)
    new_bolge = col1.text_input("Bölge adı", value=selected_default)
    new_il = col2.text_input("İl adı", value=selected_default)
    new_ilce = col3.text_input("İlçe adı")
    new_yk = col4.number_input("Kadastro YK", min_value=0.1, max_value=5.0, value=1.0, step=0.05)
    new_tapu = st.number_input("Tapu YK", min_value=0.1, max_value=5.0, value=1.0, step=0.05)
    if st.button("Katsayı kaydını ekle / güncelle", type="primary"):
        if not new_il.strip() or not new_ilce.strip():
            st.error("İl ve ilçe alanları boş bırakılamaz.")
        else:
            custom = settings.setdefault("custom_rows", [])
            new_row = {
                "sira": 10000 + len(custom),
                "bolge": new_bolge.strip().upper(),
                "il": new_il.strip().upper(),
                "ilce": new_ilce.strip().upper(),
                "tapu": float(new_tapu),
                "kadastro": float(new_yk),
            }
            custom[:] = [r for r in custom if not (r.get("il") == new_row["il"] and r.get("ilce") == new_row["ilce"])]
            custom.append(new_row)
            save_settings(settings)
            st.success("Kayıt kaydedildi. Liste yenilendiğinde yeni katsayı kullanılacak.")

    with st.expander("Katsayı listesi"):
        st.dataframe(df[["bolge", "il", "ilce", "tapu", "kadastro"]], use_container_width=True, hide_index=True)

    st.info(f"Ayar dosyası: {OVERRIDES_FILE}")
