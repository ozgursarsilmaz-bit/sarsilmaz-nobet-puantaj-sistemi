import base64
import calendar
import math
import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from ortools.sat.python import cp_model
import pandas as pd
import streamlit as st

def dagit_acil_nobet(personeller, nobetler, sabit_acil, gun_saatleri, devirler, arama_suresi):
    """Normal nöbetleri değiştirmeden ortak havuzda günde tek acil seçer."""
    gunler = sorted(gun_saatleri)
    secimler = {p: set() for p in personeller}
    sabit_gunler = {}
    for p in personeller:
        for d in set(sabit_acil.get(p, ())):
            if d not in gun_saatleri or d not in nobetler.get(p, ()):
                raise ValueError(f"{p} — {d}. gün: sabit acil için 24 saatlik nöbet gerekli.")
            if d in sabit_gunler and sabit_gunler[d] != p:
                raise ValueError(f"{d}. gün: birden fazla sabit acil nöbetçi seçilmiş.")
            sabit_gunler[d] = p
            secimler[p].add(d)
    # Her zaman geçerli bir başlangıç: süre dolarsa bu liste korunur.
    saatler = {p: sum(gun_saatleri[d] for d in secimler[p]) for p in personeller}
    for d in gunler:
        if d in sabit_gunler: continue
        adaylar = [p for p in personeller if d in nobetler.get(p, ())]
        if not adaylar:
            raise ValueError(f"{d}. gün: acil için uygun 24 saatlik nöbetçi yok.")
        p = min(adaylar, key=lambda p: (bool(secimler[p]), int(devirler.get(p, 0)), saatler[p], p))
        secimler[p].add(d)
        saatler[p] += gun_saatleri[d]

    model = cp_model.CpModel()
    a = {(p, d): model.NewBoolVar(f"acil_{p}_{d}")
         for d in gunler for p in personeller if d in nobetler.get(p, ())}
    for d in gunler:
        model.Add(sum(a[(p, d)] for p in personeller if (p, d) in a) == 1)
    for d, p in sabit_gunler.items(): model.Add(a[(p, d)] == 1)
    gorev_aldi, bu_ay, toplam = {}, {}, {}
    saat_ust = sum(gun_saatleri.values())
    devir_min = min(int(devirler.get(p, 0)) for p in personeller)
    devir_max = max(int(devirler.get(p, 0)) for p in personeller)
    for p in personeller:
        vars_p = [a[(p, d)] for d in gunler if (p, d) in a]
        gorev_aldi[p] = model.NewBoolVar(f"acil_aldi_{p}")
        if vars_p: model.AddMaxEquality(gorev_aldi[p], vars_p)
        else: model.Add(gorev_aldi[p] == 0)
        bu_ay[p] = model.NewIntVar(0, saat_ust, f"acil_saat_{p}")
        model.Add(bu_ay[p] == sum(a[(p, d)] * gun_saatleri[d] for d in gunler if (p, d) in a))
        toplam[p] = model.NewIntVar(devir_min, devir_max + saat_ust, f"acil_toplam_{p}")
        model.Add(toplam[p] == int(devirler.get(p, 0)) + bu_ay[p])
    def saat_farki(degerler, alt, ust, ad):
        lo = model.NewIntVar(alt, ust, ad + "_min")
        hi = model.NewIntVar(alt, ust, ad + "_max")
        model.AddMinEquality(lo, list(degerler))
        model.AddMaxEquality(hi, list(degerler))
        return hi - lo
    hedefler = [
        ("Acil verilen farklı personel sayısı", len(personeller) - sum(gorev_aldi.values())),
        ("Görev alacak personelde düşük devir önceliği", sum(int(devirler.get(p, 0)) * gorev_aldi[p] for p in personeller)),
        ("Bu ay acil saat dengesi", saat_farki(bu_ay.values(), 0, saat_ust, "buay")),
        ("Devir + bu ay acil saat dengesi", saat_farki(toplam.values(), devir_min, devir_max + saat_ust, "devir")),
    ]
    rapor = []
    for ad, hedef in hedefler:
        model.ClearHints()
        for (p, d), var in a.items(): model.AddHint(var, int(d in secimler[p]))
        model.Minimize(hedef)
        solver = cp_model.CpSolver()
        solver.parameters.num_search_workers = 1
        solver.parameters.max_time_in_seconds = float(arama_suresi)
        status = solver.Solve(model)
        if status not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
            rapor.append({"Hedef": ad, "Durum": "Süre içinde yeni sonuç yok; geçerli seçim korundu"})
            break
        secimler = {p: {d for d in gunler if (p, d) in a and solver.Value(a[(p, d)])} for p in personeller}
        deger = int(solver.Value(hedef))
        rapor.append({"Hedef": ad, "Değer": deger, "En iyi sonuç kanıtlandı": status == cp_model.OPTIMAL})
        model.Add(hedef == deger)
        if status != cp_model.OPTIMAL:
            # Kapsama/devir önceliğinin daha iyisi araştırılmadan alt hedefe geçilmez.
            break
    return secimler, rapor


# --- SAYFA YAPILANDIRMASI ---
st.set_page_config(
    page_title="Mikrobiyoloji Laboratuvarı Yönetim Sistemi",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- SESSION STATE (OTURUM HAFIZASI) ---
if "hesaplanan_sonuc" not in st.session_state:
    st.session_state.hesaplanan_sonuc = None

# --- TAM PERSONEL VERİTABANI ---
TUM_PERSONEL_VERISI = {
    "ÖZGÜR SARSILMAZ": {"birim": "Mikro", "muaf": False},
    "AYSEL BOZAN": {"birim": "Mikro", "muaf": True},
    "AYŞEGÜL SARUHAN": {"birim": "Kültür", "muaf": True},
    "BELGİN PARİN": {"birim": "Mikro", "muaf": True},
    "BELGİN UYSAL": {"birim": "PCR", "muaf": False},
    "BİRSEL KAYA": {"birim": "Kültür", "muaf": False},
    "DENİZ YAMAN": {"birim": "Mikro", "muaf": True},
    "DURMUŞ AKTÜRK": {"birim": "Mikro", "muaf": False},
    "ESRA ÇORAK": {"birim": "Mikro", "muaf": True},
    "GÜLÇİN HORDACI": {"birim": "Kültür", "muaf": False},
    "HACER BÜKÜM": {"birim": "Kültür", "muaf": False},
    "HİCRAN ATA AYHAN": {"birim": "Kültür", "muaf": False},
    "HÜMEYRA YEŞİLTAŞ": {"birim": "Mikro", "muaf": False},
    "İBRAHİM ER": {"birim": "Kültür", "muaf": False},
    "KADRİYE ARDIÇ": {"birim": "Mikro", "muaf": False},
    "KEVSER DUMLU": {"birim": "PCR", "muaf": False},
    "MEHMET BAĞCI": {"birim": "Mikro", "muaf": False},
    "MELEK AKKAŞ": {"birim": "Mikro", "muaf": False},
    "MELEK BARUT": {"birim": "Kültür", "muaf": False},
    "MUHAMMED ATABERK YILDIZ": {"birim": "Kültür", "muaf": False},
    "MURAT GENCER": {"birim": "PCR", "muaf": False},
    "MUSTAFA HARTOĞLU": {"birim": "Kültür", "muaf": False},
    "MUSTAFA TOSUN": {"birim": "Mikro", "muaf": False},
    "MÜCELLA MANTAR": {"birim": "Mikro", "muaf": False},
    "ÖZTÜRK MAVİŞ": {"birim": "Mikro", "muaf": False},
    "REMZİYE AKSOY": {"birim": "Mikro", "muaf": False},
    "SEÇİL EMİR YILDIRAK": {"birim": "PCR", "muaf": False},
    "SEMRA KUMRUOĞLU": {"birim": "Kültür", "muaf": True},
    "SUNA SARSILMAZ": {"birim": "PCR", "muaf": False},
    "SÜLEYMAN POLAT": {"birim": "Kültür", "muaf": False},
    "ŞADUMAN YALÇIN": {"birim": "PCR", "muaf": False},
    "ŞENEL TAŞ": {"birim": "PCR", "muaf": False},
    "ŞENGÜL URLU": {"birim": "Mikro", "muaf": False},
    "ŞÜKRAN KAYA": {"birim": "Mikro", "muaf": False},
    "ZOZAN ATLI": {"birim": "Kültür", "muaf": True},
}

# --- ÖZEL CSS TASARIMI ---
st.markdown(
    """
<style>
    .main .block-container { padding-top: 0.25rem !important; padding-bottom: 2rem !important; }
    [data-testid="stMainBlockContainer"], [data-testid="stAppViewContainer"] .main > div, [data-testid="stAppViewContainer"] .main .block-container { padding-top: 0.25rem !important; }
    [data-testid="stAppViewContainer"] { padding-top: 0 !important; }
    section.main { padding-top: 0 !important; }
    /* Üst çubuğu gizlemek, menü kapandığında açma okunu da gizler. */
    [data-testid="stHeader"], .stAppHeader {
        display: flex !important;
        background: transparent !important;
        pointer-events: none;
    }
    [data-testid="stHeader"] button, .stAppHeader button {
        pointer-events: auto;
    }
    [data-testid="stSidebarCollapsedControl"] {
        display: flex !important;
        visibility: visible !important;
        position: fixed;
        top: 0.5rem;
        left: 0.5rem;
        z-index: 100001;
        pointer-events: auto;
    }
    .main { background-color: #F8F9FA; }
    .header-box {
        background: linear-gradient(135deg, #0d6efd 0%, #0a58ca 100%); color: white; padding: 12px 20px;
        border-radius: 10px; box-shadow: 0 3px 8px rgba(13, 110, 253, 0.15); margin-bottom: 12px; position: relative;
    }
    .header-box h1 { margin: 0; font-size: 1.35rem !important; font-weight: 700; line-height: 1.2; }
    .header-box p { margin: 2px 0 0 0; opacity: 0.88; font-size: 0.82rem !important; }
    .header-imza { position: absolute; bottom: 6px; right: 15px; font-size: 0.72rem; font-weight: 600; opacity: 0.75; letter-spacing: 0.5px; font-style: italic; }
    div[data-testid="stMetric"] { background-color: #FFFFFF; border: 1px solid #E9ECEF; padding: 10px 14px; border-radius: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.03); }
    .section-title { font-size: 1.1rem; font-weight: 700; color: #212529; margin-bottom: 10px; display: flex; align-items: center; gap: 6px; }
    .stButton>button {
        width: 100%; background: linear-gradient(135deg, #198754 0%, #146c43 100%); color: white; border: none;
        padding: 10px 20px; font-size: 1rem; font-weight: 600; border-radius: 8px; box-shadow: 0 4px 10px rgba(25, 135, 84, 0.2); transition: all 0.3s ease;
    }
    .stButton>button:hover { background: linear-gradient(135deg, #146c43 0%, #0f5132 100%); transform: translateY(-1px); box-shadow: 0 6px 14px rgba(25, 135, 84, 0.3); }
    .direct-download-btn {
        display: inline-block; width: 100%; text-align: center; background: linear-gradient(135deg, #0d6efd 0%, #0a58ca 100%);
        color: white !important; text-decoration: none !important; padding: 12px 20px; font-size: 1.05rem; font-weight: 700;
        border-radius: 8px; box-shadow: 0 4px 10px rgba(13, 110, 253, 0.25); transition: all 0.3s ease;
    }
    .direct-download-btn:hover { background: linear-gradient(135deg, #0a58ca 0%, #084298 100%); transform: translateY(-1px); box-shadow: 0 6px 14px rgba(13, 110, 253, 0.35); }
</style>
""",
    unsafe_allow_html=True,
)

tr_gunler = {
    0: "Pazartesi", 1: "Salı", 2: "Çarşamba", 3: "Perşembe", 4: "Cuma", 5: "Cumartesi", 6: "Pazar",
}

def tr_norm(text):
    s = str(text).strip()
    replacements = [
        ("İ", "i"), ("I", "ı"), ("ı", "i"), ("ğ", "g"), ("Ğ", "g"),
        ("ü", "u"), ("Ü", "u"), ("ş", "s"), ("Ş", "s"), ("ö", "o"),
        ("Ö", "o"), ("ç", "c"), ("Ç", "c"),
    ]
    for old, new in replacements:
        s = s.replace(old, new)
    return s.lower()

# ==========================================
# MODÜL SEÇİMİ (SIDEBAR NAVİGASYON)
# ==========================================
st.sidebar.title("🏥 Mikrobiyoloji Lab.")
secilen_modul = st.sidebar.radio(
    "📌 Lütfen İşlem Yapılacak Modülü Seçiniz:",
    [
        "1. Personel Nöbet & Puantaj",
        "2. Uzman Dr. Çalışma Listesi",
        "3. Asistan Dr. Çalışma Listesi",
        "4. Eğitici Destekleme Puan Çizelgesi"
    ]
)
st.sidebar.markdown("---")


# ==========================================
# 1. MODÜL: PERSONEL NÖBET & PUANTAJ
# ==========================================
if secilen_modul == "1. Personel Nöbet & Puantaj":
    st.markdown(
        """
    <div class="header-box">
        <h1>🏥 Mikrobiyoloji Laboratuvarı Nöbet Dağılım ve Puantaj Sistemi</h1>
        <p>Tüm Günler Dengeli Dağılım & Adil Saat Optimizasyonu & Çalışma Listesi_Ay Sonu İstatistik_Puantaj Oluşturma</p>
        <div class="header-imza">✍️ Özgür SARSILMAZ</div>
    </div>
    """,
        unsafe_allow_html=True,
    )

    col_yil, col_ay = st.sidebar.columns(2)

    with col_yil:
        yil = st.number_input("Yıl", value=2026, min_value=2024, max_value=2030)
    with col_ay:
        ay = st.selectbox("Ay", list(range(1, 13)), index=9)

    _, gun_sayisi = calendar.monthrange(yil, ay)
    gun_secenekleri = list(range(1, gun_sayisi + 1))

    gecmis_ay_alani = st.sidebar.container()


    resmi_tatil_gunleri = st.sidebar.multiselect(
        "Tam Gün Tatil / Resmi Günler:",
        options=gun_secenekleri,
        default=[],
        help="Tam gün resmi/idari tatil günlerini seçiniz.",
    )
    yarim_gun_tatil_gunleri = st.sidebar.multiselect(
        "Yarım Gün / Arife Günleri:",
        options=[g for g in gun_secenekleri if g not in resmi_tatil_gunleri],
        default=[],
        help="Arife veya yarım gün tatil günlerini seçiniz.",
    )

    kultur_8s_tatil_gunleri = st.sidebar.multiselect(
        "Kültür 8s Tatil Vardiyası Günleri:",
        options=resmi_tatil_gunleri,
        default=[],
        help="Resmi tatillerde Kültür biriminden 1 kişinin 8 saatlik vardiyaya geleceği ek tatil günlerini seçiniz. Cumartesiler otomatiktir.",
    )

    st.sidebar.markdown("---")
    birim_secimi = st.sidebar.selectbox(
        "🔬 Çalışma Grubu / Birim:",
        ["Mikro", "Kültür", "PCR", "Tüm Laboratuvar (Birleşik)"],
        index=3,
    )

    varsayilan_liste = [
        p_adi
        for p_adi, p_info in TUM_PERSONEL_VERISI.items()
        if birim_secimi == "Tüm Laboratuvar (Birleşik)"
        or p_info["birim"] == birim_secimi
    ]
    personel_input = st.sidebar.text_area(
        "Personel Listesi (Her satıra bir isim):",
        value="\n".join(varsayilan_liste),
        height=220,
    )

    tum_girilen_personeller = [
        p.strip().upper() for p in personel_input.split("\n") if p.strip()
    ]
    nobetci_personeller = [
        p
        for p in tum_girilen_personeller
        if not TUM_PERSONEL_VERISI.get(p, {}).get("muaf", False)
    ]
    muaf_personeller = [
        p
        for p in tum_girilen_personeller
        if TUM_PERSONEL_VERISI.get(p, {}).get("muaf", False)
    ]

    st.sidebar.markdown("---")
    st.sidebar.subheader("🛡️️ Genel Kural Kurulumu")
    gunluk_nobetci = st.sidebar.number_input(
        "Birim Başı Günlük Nöbetçi İhtiyacı:",
        min_value=1,
        max_value=10,
        value=1,
        step=1,
    )
    dinlenme_gun_sayisi = st.sidebar.number_input(
        "Genel Nöbet Arası Min. Dinlenme (Gün):",
        min_value=0,
        max_value=10,
        value=3,
    )
    persembe_pazar_yasagi = st.sidebar.checkbox(
        "Genel Perşembe - Pazar Yasağı", value=True
    )
    cuma_haftasonu_siki_kural = st.sidebar.checkbox(
        "📌 Cuma / Cmts / Pzr tekrarını azalt (esnek tercih)", value=True
    )
    cozum_arama_suresi = st.sidebar.number_input(
        "Dağıtım aşaması başına arama süresi (saniye)", min_value=5, max_value=120, value=15,
        help="En iyi denge kanıtlanamazsa süreyi artırabilirsiniz. Beş aşama sırayla çalışır.",
    )
    esnek_personel = st.sidebar.multiselect(
        "🔓 Özel Esneklik Tanınacak Personel(ler):",
        options=nobetci_personeller,
        default=[],
    )

    st.sidebar.markdown("---")
    gecmis_ay_alani.subheader("📊 Geçmiş Ay Rotasyonu (Önceki Ay Dosyası)")
    uploaded_file = gecmis_ay_alani.file_uploader(
        "Önceki Ayın Excel Dosyası:",
        type=["xlsx", "xls"],
        help="Sistemin ürettiği 3 sekmeli Excel dosyasını yükleyin.",
    )

    gecmis_istatistik = {}
    gecmis_acil_istatistik = {}
    gecmis_kultur8_istatistik = {}
    otomatik_son_gun_normal = []
    otomatik_son_gun_acil = []
    otomatik_kultur_ni = []

    if uploaded_file is not None:
        try:
            xls = pd.ExcelFile(uploaded_file)
            if "İstatistik & Mesai Yükü" in xls.sheet_names:
                df_gecmis = pd.read_excel(xls, sheet_name="İstatistik & Mesai Yükü")
                gunler_sira = [
                    "Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"
                ]

                for r_idx in range(1, len(df_gecmis)):
                    row = df_gecmis.iloc[r_idx]
                    p_name = str(row.iloc[0]).strip().upper()
                    if not p_name or p_name in ["NAN", "NONE", "AD SOYAD", "ADI SOYADI", "PERSONEL"]:
                        continue

                    p_dict = {}
                    offset = 1 if "Birim" in str(df_gecmis.iloc[0]).strip() or len(df_gecmis.columns) > 28 else 0
                    for g_idx, g_name in enumerate(gunler_sira):
                        toplam_col_idx = (3 + offset) + (g_idx * 3)
                        try:
                            val = int(row.iloc[toplam_col_idx])
                        except (ValueError, TypeError, IndexError):
                            val = 0
                        p_dict[g_name] = val
                    gecmis_istatistik[p_name] = p_dict

                    try:
                        acil_col_idx = 27 if offset == 1 else 26
                        acil_devir_val = int(row.iloc[acil_col_idx])
                    except (ValueError, TypeError, IndexError):
                        acil_devir_val = 0
                    gecmis_acil_istatistik[p_name] = acil_devir_val

                    try:
                        kultur8_col_idx = 30 if offset == 1 else 29
                        k8_devir_val = int(row.iloc[kultur8_col_idx])
                    except (ValueError, TypeError, IndexError):
                        k8_devir_val = 0
                    gecmis_kultur8_istatistik[p_name] = k8_devir_val

            if "Aylık Görev Listesi" in xls.sheet_names:
                df_goreg = pd.read_excel(xls, sheet_name="Aylık Görev Listesi", skiprows=2)
                if not df_goreg.empty:
                    last_row = df_goreg.iloc[-1]
                    for col_name in ["PCR", "Mikro", "Kültür"]:
                        if col_name in df_goreg.columns and pd.notna(last_row[col_name]):
                            names = str(last_row[col_name]).split(",")
                            for n in names:
                                is_acil = "(Acil)" in n
                                clean_n = n.replace("(Acil)", "").strip().upper()
                                if clean_n in nobetci_personeller:
                                    if is_acil:
                                        otomatik_son_gun_acil.append(clean_n)
                                    else:
                                        otomatik_son_gun_normal.append(clean_n)

            if "Kültür Nİ Devir" in xls.sheet_names:
                df_kultur_devir = pd.read_excel(xls, sheet_name="Kültür Nİ Devir")
                if "Personel" in df_kultur_devir.columns:
                    otomatik_kultur_ni = [str(p).strip().upper() for p in df_kultur_devir["Personel"].dropna()
                                         if str(p).strip().upper() in nobetci_personeller]
            elif "Aylık Görev Listesi" in xls.sheet_names and not df_goreg.empty:
                # Eski dosyalarda son gün Kültür 8s vardiyasını aktar.
                if "Kültür (8s Vardiya)" in df_goreg.columns and pd.notna(last_row["Kültür (8s Vardiya)"]):
                    otomatik_kultur_ni = [p.strip().upper() for p in str(last_row["Kültür (8s Vardiya)"]).split(",")
                                         if p.strip().upper() in nobetci_personeller]

            gecmis_ay_alani.success(f"✅ {len(gecmis_istatistik)} personelin devir verileri aktarıldı!")
        except Exception as e:
            gecmis_ay_alani.error(f"❌ Hata: Yüklenen Excel okunurken sorun oluştu ({e}).")

    gecmis_ay_son_gun_normal = gecmis_ay_alani.multiselect(
        "🌙 Geçmiş Ay Son Günü NORMAL Nöbetçileri:",
        options=nobetci_personeller,
        default=list(set(otomatik_son_gun_normal)),
        help="Önceki ayın son günü normal nöbet tutan personeller.",
    )

    gecmis_ay_son_gun_acil = gecmis_ay_alani.multiselect(
        "🚨 Geçmiş Ay Son Günü ACİL Nöbetçileri:",
        options=[p for p in nobetci_personeller if p not in gecmis_ay_son_gun_normal],
        default=list(set(otomatik_son_gun_acil)),
        help="Önceki ayın son günü acil nöbet tutan personeller.",
    )

    gecmis_ay_kultur_ni = gecmis_ay_alani.multiselect(
        "🧫 Önceki Aydan Kültür 8s Nİ Devri:",
        options=nobetci_personeller, default=list(set(otomatik_kultur_ni)),
        help="Kültür 8s vardiyasının karşılığı olan Nİ bu ayın ilk mesai gününe kalan personeller.",
    )
    ilk_mesai_gunu = next((d for d in range(1, gun_sayisi + 1)
                          if datetime.date(yil, ay, d).weekday() < 5 and d not in resmi_tatil_gunleri), None)

    gecmis_ay_son_gun_nobetcileri = list(set(gecmis_ay_son_gun_normal + gecmis_ay_son_gun_acil))

    def get_prev(p_name, category):
        p_norm = tr_norm(p_name)
        for key, val in gecmis_istatistik.items():
            if tr_norm(key) == p_norm:
                return val.get(category, 0)
        return 0

    def get_prev_acil(p_name):
        p_norm = tr_norm(p_name)
        for key, val in gecmis_acil_istatistik.items():
            if tr_norm(key) == p_norm:
                return val
        return 0

    def get_prev_kultur8(p_name):
        p_norm = tr_norm(p_name)
        for key, val in gecmis_kultur8_istatistik.items():
            if tr_norm(key) == p_norm:
                return val
        return 0

    # Satır kimlikleri değişmez: aradan silinen satır diğer girişleri etkilemez.
    def satir_grubunu_hazirla(grup, eski_sayac):
        ids_key, next_key = grup + "_ids", grup + "_next"
        if ids_key not in st.session_state:
            st.session_state[ids_key] = list(range(st.session_state.get(eski_sayac, 1)))
        if next_key not in st.session_state:
            st.session_state[next_key] = max(st.session_state[ids_key], default=-1) + 1

    def satir_ekle(grup):
        row_id = st.session_state[grup + "_next"]
        st.session_state[grup + "_ids"].append(row_id)
        st.session_state[grup + "_next"] = row_id + 1

    def satir_sil(grup, row_id, alanlar):
        st.session_state[grup + "_ids"] = [i for i in st.session_state[grup + "_ids"] if i != row_id]
        for alan in alanlar:
            st.session_state.pop(f"{alan}_{row_id}", None)

    satir_grubunu_hazirla("mazeret_satir", "mazeret_satir_sayisi")
    satir_grubunu_hazirla("acil_satir", "acil_satir_sayisi")
    satir_grubunu_hazirla("kisi_kisit", "kisi_kisit_sayisi")

    st.markdown('<div class="section-title">📋 Personel Mazeret ve Sabit Nöbet Girişleri</div>', unsafe_allow_html=True)
    izinler = {p: [] for p in nobetci_personeller}
    sabit_nobetler = {p: [] for p in nobetci_personeller}
    toplam_izin_sayisi, toplam_sabit_sayisi = 0, 0
    for m_idx in st.session_state.mazeret_satir_ids:
        c1, c2, c3, c4 = st.columns([1.2, 1.8, 1.8, 0.35], vertical_alignment="bottom")
        with c1: p_secilen = st.selectbox("Personel:", options=["Seçiniz..."] + nobetci_personeller, key=f"m_personel_{m_idx}")
        with c2: selected_days = st.multiselect("Mazeret / İzin Günleri:", options=gun_secenekleri, default=[], key=f"m_leave_{m_idx}")
        with c3: selected_sabit_days = st.multiselect("Sabit Nöbet Günleri:", options=gun_secenekleri, default=[], key=f"m_forced_{m_idx}")
        with c4: st.button("🗑️", key=f"m_delete_{m_idx}", help="Bu personel giriş satırını sil", on_click=satir_sil, args=("mazeret_satir", m_idx, ("m_personel", "m_leave", "m_forced")))
        if p_secilen != "Seçiniz...":
            izinler[p_secilen].extend([d - 1 for d in selected_days])
            sabit_nobetler[p_secilen].extend([d - 1 for d in selected_sabit_days])
            toplam_izin_sayisi += len(selected_days)
            toplam_sabit_sayisi += len(selected_sabit_days)
    st.button("➕ Mazeret / Sabit Nöbet Ekle", on_click=satir_ekle, args=("mazeret_satir",))

    st.markdown('<div class="section-title">🚨 Acil Nöbetçi Girişleri (Opsiyonel)</div>', unsafe_allow_html=True)
    sabit_acil_nobetler = {p: [] for p in nobetci_personeller}
    toplam_sabit_acil_sayisi = 0
    for a_idx in st.session_state.acil_satir_ids:
        c1, c2, c3 = st.columns([1.5, 3.3, 0.35], vertical_alignment="bottom")
        with c1: p_acil_secilen = st.selectbox("Acil Nöbetçi:", options=["Seçiniz..."] + nobetci_personeller, key=f"a_personel_{a_idx}")
        with c2: selected_acil_days = st.multiselect("Sabit Acil Nöbet Günleri:", options=gun_secenekleri, default=[], key=f"a_forced_{a_idx}")
        with c3: st.button("🗑️", key=f"a_delete_{a_idx}", help="Bu acil giriş satırını sil", on_click=satir_sil, args=("acil_satir", a_idx, ("a_personel", "a_forced")))
        if p_acil_secilen != "Seçiniz...":
            sabit_acil_nobetler[p_acil_secilen].extend(selected_acil_days)
            toplam_sabit_acil_sayisi += len(selected_acil_days)
    st.button("➕ Sabit Acil Nöbet Ekle", on_click=satir_ekle, args=("acil_satir",))

    st.markdown('<div class="section-title">🤝 Kişiler Arası Nöbet Mesafe ve Çakışma Yasağı Kuralları</div>', unsafe_allow_html=True)
    kisi_kisitlari = []
    for k_idx in st.session_state.kisi_kisit_ids:
        c1, c2, c3, c4 = st.columns([1.2, 1, 2, 0.35], vertical_alignment="bottom")
        with c1: p_ana = st.selectbox("Ana Personel:", options=["Seçiniz..."] + nobetci_personeller, key=f"p_ana_{k_idx}")
        with c2: min_aralik = st.number_input("Min. Mesafe (Gün):", min_value=0, max_value=15, value=1, key=f"min_aralik_{k_idx}")
        with c3: p_yasakli_list = st.multiselect("Birlikte/Yakın Nöbet Tutamayacağı Kişiler:", options=[p for p in nobetci_personeller if p != p_ana], key=f"p_yasakli_{k_idx}")
        with c4: st.button("🗑️", key=f"k_delete_{k_idx}", help="Bu kişiler arası kısıt satırını sil", on_click=satir_sil, args=("kisi_kisit", k_idx, ("p_ana", "min_aralik", "p_yasakli")))
        if p_ana != "Seçiniz..." and p_yasakli_list:
            kisi_kisitlari.append({"ana": p_ana, "aralik": min_aralik, "yasaklilar": p_yasakli_list})
    st.button("➕ Yeni Kısıt Ekle", on_click=satir_ekle, args=("kisi_kisit",))

    def is_day_off(yil, ay, day, resmi_tatil_gunleri):
        dt = datetime.date(yil, ay, day)
        return (dt.weekday() in [5, 6]) or (day in resmi_tatil_gunleri)

    def calculate_aylik_calisma_saati(yil, ay, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri):
        toplam_saat = 0
        for d in range(1, gun_sayisi + 1):
            if not is_day_off(yil, ay, d, resmi_tatil_gunleri):
                if d in yarim_gun_tatil_gunleri: toplam_saat += 5
                else: toplam_saat += 8
        return toplam_saat

    def calculate_nobet_hakedis(yil, ay, day, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri):
        dt = datetime.date(yil, ay, day)
        next_dt = dt + datetime.timedelta(days=1)
        next_off = next_dt.weekday() >= 5 or (
            next_dt.month == ay and next_dt.day in resmi_tatil_gunleri)
        if day in yarim_gun_tatil_gunleri:
            return 12, 7
        if day == gun_sayisi and not next_off:
            return 12, 4
        if is_day_off(yil, ay, day, resmi_tatil_gunleri):
            return (12, 12) if next_off else (12, 4)
        if dt.weekday() < 4:
            # Ertesi gün yarım gün mesai (5s) ise Nİ karşılığı 5s olur.
            # Olağan 8s mesaiye göre kalan 3s normal hakedişe eklenir.
            if next_dt.month == ay and next_dt.day in yarim_gun_tatil_gunleri and not next_off:
                return 8, 3
            return 8, 0
        return 12, 4

    def kultur_ni_gunleri(yil, ay, gun_sayisi, vardiyalar, resmi_tatil_gunleri):
        ni_gunleri = set()
        for vardiya in vardiyalar:
            for day in range(vardiya + 1, gun_sayisi + 1):
                if not is_day_off(yil, ay, day, resmi_tatil_gunleri):
                    ni_gunleri.add(day)
                    break
        return ni_gunleri

    def calculate_personel_puantaj_metrikleri(p_name, p_row_dict, yil, ay, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri, acil_days_set, kultur_8s_tatil_gunleri, gecmis_ay_son_gun_normal, gecmis_ay_son_gun_acil, kultur_8s_vardiyalari=None):
        aylik_hedef_saat = calculate_aylik_calisma_saati(yil, ay, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri)

        toplam_calisma = 0
        norm_gece = 0
        norm_normal = 0
        risk_gece = 0
        risk_normal = 0

        # 1. Ham Nöbet Saatlerinin Dağıtılması
        for d in range(1, gun_sayisi + 1):
            val = str(p_row_dict.get(str(d), "")).strip()

            # Kültür 8s dahil çalışılan saatler toplama girer.
            # İlk mesai günündeki Nİ=0; ayrıca 8 saat düşülmez.
            if val in ["8", "5", "16", "19", "11", "24"]:
                toplam_calisma += int(val)

            dt = datetime.date(yil, ay, d)
            weekday = dt.weekday()
            is_off = is_day_off(yil, ay, d, resmi_tatil_gunleri)
            is_next_day_off = is_day_off(yil, ay, d + 1, resmi_tatil_gunleri) if d < gun_sayisi else False
            is_risk = d in acil_days_set

            if val == "24":
                g_saat, n_saat = calculate_nobet_hakedis(
                    yil, ay, d, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri)

                if is_risk:
                    risk_gece += g_saat
                    risk_normal += n_saat
                else:
                    norm_gece += g_saat
                    norm_normal += n_saat

        # Kültür 8s karşılığı Nİ sonraki aya kalıyorsa, bu 8s mevcut
        # ayın ödeme hesabına girmez. Günlük hücrede 8 olarak kalır.
        for vardiya in set(kultur_8s_vardiyalari or ()):
            if str(p_row_dict.get(str(vardiya), "")).strip() == "8" and not kultur_ni_gunleri(
                    yil, ay, gun_sayisi, {vardiya}, resmi_tatil_gunleri):
                toplam_calisma -= 8

        # Önceki ayın ödeme hesabından ayrılan Kültür 8s, bu ay Nİ
        # kullanıldığında mesai karşılığı olarak eklenir; nöbet ücreti değildir.
        if p_name in gecmis_ay_kultur_ni and ilk_mesai_gunu is not None:
            if str(p_row_dict.get(str(ilk_mesai_gunu), "")).strip() == "Nİ":
                toplam_calisma += 8

        # 2. Akıllı Devir Nİ Mahsuplaşması (1. Gün Nİ Alan Personel)
        is_devir_normal = p_name in gecmis_ay_son_gun_normal
        is_devir_acil = p_name in gecmis_ay_son_gun_acil

        if is_devir_normal or is_devir_acil:
            def mahsup_et(gece, normal, kalan):
                # Önce en fazla 4 normal + 4 gece; eksik normal payı
                # geceden, eksik gece payı normalden tamamlanır.
                normal_dusum = min(4, normal, kalan)
                normal -= normal_dusum
                kalan -= normal_dusum

                gece_dusum = min(4, gece, kalan)
                gece -= gece_dusum
                kalan -= gece_dusum

                ek_gece = min(gece, kalan)
                gece -= ek_gece
                kalan -= ek_gece

                ek_normal = min(normal, kalan)
                normal -= ek_normal
                kalan -= ek_normal
                return gece, normal, kalan

            dusulecek_saat = 8
            if is_devir_acil:
                # Acil devir: riskli hakedişler, ardından normal hakedişler.
                risk_gece, risk_normal, dusulecek_saat = mahsup_et(
                    risk_gece, risk_normal, dusulecek_saat)
                norm_gece, norm_normal, dusulecek_saat = mahsup_et(
                    norm_gece, norm_normal, dusulecek_saat)
            else:
                # Normal devir: normal hakedişler, ardından riskli hakedişler.
                norm_gece, norm_normal, dusulecek_saat = mahsup_et(
                    norm_gece, norm_normal, dusulecek_saat)
                risk_gece, risk_normal, dusulecek_saat = mahsup_et(
                    risk_gece, risk_normal, dusulecek_saat)

        fazla_nobet = max(0, toplam_calisma - aylik_hedef_saat)

        return {
            "Toplam Çalışma Saati": toplam_calisma,
            "Aylık Çalışma Saati": aylik_hedef_saat,
            "Fazla Nöbet Saati": fazla_nobet,
            "Normal_Gece": norm_gece,
            "Normal_Normal": norm_normal,
            "Riskli_Gece": risk_gece,
            "Riskli_Normal": risk_normal,
        }

    def generate_3_tab_excel(yil, ay, nobetci_personeller, tum_girilen_personeller, nobet_dict, kultur_8s_dict, acil_nobet_dict, gecmis_istatistik, gecmis_acil_istatistik, gecmis_kultur8_istatistik, gun_sayisi, birim_secimi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri, kultur_8s_tatil_gunleri, gecmis_ay_son_gun_normal, gecmis_ay_son_gun_acil):
        wb = openpyxl.Workbook()
        font_title = Font(name="Calibri", size=12, bold=True)
        font_header = Font(name="Calibri", size=10, bold=True)
        font_body = Font(name="Calibri", size=9)
        font_bold = Font(name="Calibri", size=9, bold=True)
        font_ni = Font(name="Calibri", size=9, bold=True)
        font_24 = Font(name="Calibri", size=9, bold=True)

        fill_header = PatternFill(start_color="C5D9A4", end_color="C5D9A4", fill_type="solid")
        fill_green_bg = PatternFill(start_color="D8E4BC", end_color="D8E4BC", fill_type="solid")
        fill_grey_weekend = PatternFill(start_color="A6A6A6", end_color="A6A6A6", fill_type="solid")
        fill_bright_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

        thin_side = Side(style="thin", color="000000")
        border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")
        align_vertical_header = Alignment(horizontal="center", vertical="center", text_rotation=90, wrap_text=True)

        # 1. SEKME
        ws1 = wb.active
        ws1.title = "Aylık Görev Listesi"
        ws1.views.sheetView[0].showGridLines = True
        ws1.cell(row=1, column=1, value=f"{yil} yılı {ay}. Ay Mikrobiyoloji Laboratuvarı Nöbet Çizelgesi").font = font_title

        headers1 = ["Tarih", "Gün", "PCR", "Mikro", "Kültür", "Kültür (8s Vardiya)"]
        for c_idx, h in enumerate(headers1, 1):
            cell = ws1.cell(row=3, column=c_idx, value=h)
            cell.font, cell.fill, cell.alignment, cell.border = font_header, fill_header, align_center, border_cell

        for d in range(1, gun_sayisi + 1):
            r = d + 3
            tarih = datetime.date(yil, ay, d)
            pcr_list, mikro_list, kultur_list, kultur_8s_list = [], [], [], []
            for p in nobetci_personeller:
                p_unit = TUM_PERSONEL_VERISI.get(p, {}).get("birim", "Mikro")
                is_24 = d in nobet_dict.get(p, set())
                is_8s = d in kultur_8s_dict.get(p, set())
                is_acil = d in acil_nobet_dict.get(p, set())

                if is_24:
                    p_text = f"{p} (Acil)" if is_acil else p
                    if p_unit == "PCR": pcr_list.append(p_text)
                    elif p_unit == "Mikro": mikro_list.append(p_text)
                    elif p_unit == "Kültür": kultur_list.append(p_text)
                elif is_8s: kultur_8s_list.append(p)

            c1 = ws1.cell(row=r, column=1, value=tarih.strftime("%d.%m.%Y"))
            c2 = ws1.cell(row=r, column=2, value=tr_gunler[tarih.weekday()])
            c3 = ws1.cell(row=r, column=3, value=", ".join(pcr_list))
            c4 = ws1.cell(row=r, column=4, value=", ".join(mikro_list))
            c5 = ws1.cell(row=r, column=5, value=", ".join(kultur_list))
            c6 = ws1.cell(row=r, column=6, value=", ".join(kultur_8s_list))

            for c in [c1, c2, c3, c4, c5, c6]:
                c.font, c.border = font_body, border_cell
                c.alignment = align_center if c in [c1, c2] else align_left
                if is_day_off(yil, ay, d, resmi_tatil_gunleri): c.fill = fill_grey_weekend

        for c_let, w in zip(["A", "B", "C", "D", "E", "F"], [13, 13, 25, 25, 25, 25]):
            ws1.column_dimensions[c_let].width = w

        # 2. SEKME
        ws2 = wb.create_sheet("İstatistik & Mesai Yükü")
        ws2.views.sheetView[0].showGridLines = True
        gunler_listesi = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]

        ws2.merge_cells("A1:A2")
        cell_a = ws2.cell(row=1, column=1, value="AD SOYAD")
        cell_a.font, cell_a.fill, cell_a.alignment, cell_a.border = font_header, fill_header, align_center, border_cell

        ws2.merge_cells("B1:B2")
        cell_b = ws2.cell(row=1, column=2, value="Birim")
        cell_b.font, cell_b.fill, cell_b.alignment, cell_b.border = font_header, fill_header, align_center, border_cell

        col_counter = 3
        for g in gunler_listesi:
            ws2.merge_cells(start_row=1, start_column=col_counter, end_row=1, end_column=col_counter + 2)
            top_cell = ws2.cell(row=1, column=col_counter, value=g)
            top_cell.font, top_cell.fill, top_cell.alignment = font_header, fill_header, align_center
            for idx, sub in enumerate(["Devir", "Bu Ay", "Toplam"]):
                sub_cell = ws2.cell(row=2, column=col_counter + idx, value=sub)
                sub_cell.font, sub_cell.fill, sub_cell.alignment, sub_cell.border = font_header, fill_green_bg, align_center, border_cell
            col_counter += 3

        ws2.merge_cells(start_row=1, start_column=col_counter, end_row=2, end_column=col_counter)
        c_tn = ws2.cell(row=1, column=col_counter, value="T.NöbetSaati")
        c_tn.font, c_tn.fill, c_tn.alignment, c_tn.border = font_header, fill_header, align_center, border_cell

        col_counter += 1
        ws2.merge_cells(start_row=1, start_column=col_counter, end_row=2, end_column=col_counter)
        c_tnb = ws2.cell(row=1, column=col_counter, value="TOPLAM NÖBET")
        c_tnb.font, c_tnb.fill, c_tnb.alignment, c_tnb.border = font_header, fill_header, align_center, border_cell

        col_counter += 1
        ws2.merge_cells(start_row=1, start_column=col_counter, end_row=1, end_column=col_counter + 2)
        top_acil = ws2.cell(row=1, column=col_counter, value="ACİL NÖBET (SAAT)")
        top_acil.font, top_acil.fill, top_acil.alignment = font_header, fill_header, align_center
        for idx, sub in enumerate(["Devir", "Bu Ay", "Toplam"]):
            sub_cell = ws2.cell(row=2, column=col_counter + idx, value=sub)
            sub_cell.font, sub_cell.fill, sub_cell.alignment, sub_cell.border = font_header, fill_green_bg, align_center, border_cell
        col_counter += 3

        ws2.merge_cells(start_row=1, start_column=col_counter, end_row=1, end_column=col_counter + 2)
        top_k8 = ws2.cell(row=1, column=col_counter, value="KÜLTÜR 8S VARDİYA")
        top_k8.font, top_k8.fill, top_k8.alignment = font_header, fill_header, align_center
        for idx, sub in enumerate(["Devir", "Bu Ay", "Toplam"]):
            sub_cell = ws2.cell(row=2, column=col_counter + idx, value=sub)
            sub_cell.font, sub_cell.fill, sub_cell.alignment, sub_cell.border = font_header, fill_green_bg, align_center, border_cell

        for r_idx in [1, 2]:
            for c_idx in range(1, col_counter + 3):
                ws2.cell(row=r_idx, column=c_idx).border = border_cell

        for p_idx, p in enumerate(nobetci_personeller, 3):
            p_birim = TUM_PERSONEL_VERISI.get(p, {}).get("birim", birim_secimi)
            ws2.cell(row=p_idx, column=1, value=p).alignment = align_left
            ws2.cell(row=p_idx, column=1).font, ws2.cell(row=p_idx, column=1).border = font_body, border_cell
            ws2.cell(row=p_idx, column=2, value=p_birim).alignment = align_center
            ws2.cell(row=p_idx, column=2).font, ws2.cell(row=p_idx, column=2).border = font_body, border_cell

            bu_ay_gunler = {g: 0 for g in gunler_listesi}
            bu_ay_toplam_nobet = len(nobet_dict.get(p, set()))
            bu_ay_saat = sum(sum(calculate_nobet_hakedis(yil, ay, d, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri)) for d in nobet_dict.get(p, set()))
            for d in nobet_dict.get(p, set()):
                w = datetime.date(yil, ay, d).weekday()
                bu_ay_gunler[tr_gunler[w]] += 1

            c_i = 3
            for g in gunler_listesi:
                devir, bu_ay = int(get_prev(p, g)), int(bu_ay_gunler[g])
                for v in [devir, bu_ay, devir + bu_ay]:
                    cell = ws2.cell(row=p_idx, column=c_i, value=int(v))
                    cell.font, cell.alignment, cell.border = font_body, align_center, border_cell
                    c_i += 1

            cell_saat = ws2.cell(row=p_idx, column=c_i, value=int(bu_ay_saat))
            cell_saat.font, cell_saat.alignment, cell_saat.border = font_bold, align_center, border_cell
            c_i += 1

            cell_nobet = ws2.cell(row=p_idx, column=c_i, value=int(bu_ay_toplam_nobet))
            cell_nobet.font, cell_nobet.alignment, cell_nobet.border = font_bold, align_center, border_cell
            c_i += 1

            bu_ay_acil_saat = sum(sum(calculate_nobet_hakedis(yil, ay, d, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri)) for d in acil_nobet_dict.get(p, set()))
            acil_devir = int(get_prev_acil(p))
            for idx, v in enumerate([acil_devir, bu_ay_acil_saat, acil_devir + bu_ay_acil_saat]):
                cell = ws2.cell(row=p_idx, column=c_i + idx, value=int(v))
                cell.font = font_bold if idx == 2 else font_body
                cell.alignment, cell.border = align_center, border_cell
            c_i += 3

            bu_ay_k8_sayisi = len(kultur_8s_dict.get(p, set()))
            k8_devir = int(get_prev_kultur8(p))
            for idx, v in enumerate([k8_devir, bu_ay_k8_sayisi, k8_devir + bu_ay_k8_sayisi]):
                cell = ws2.cell(row=p_idx, column=c_i + idx, value=int(v))
                cell.font = font_bold if idx == 2 else font_body
                cell.alignment, cell.border = align_center, border_cell

        ws2.column_dimensions["A"].width = 25
        ws2.column_dimensions["B"].width = 12

        # 3. SEKME
        ws3 = wb.create_sheet("Puantaj Tablosu")
        ws3.views.sheetView[0].showGridLines = True
        ws3.row_dimensions[5].height = 110

        ws3.cell(row=5, column=1, value="Adı Soyadı").fill = fill_white
        ws3.cell(row=5, column=1).font, ws3.cell(row=5, column=1).alignment, ws3.cell(row=5, column=1).border = font_header, align_left, border_cell

        ws3.cell(row=5, column=2, value="Birim").fill = fill_white
        ws3.cell(row=5, column=2).font, ws3.cell(row=5, column=2).alignment, ws3.cell(row=5, column=2).border = font_header, align_center, border_cell

        off_days = set()
        for day in range(1, gun_sayisi + 1):
            col_idx = day + 2
            if is_day_off(yil, ay, day, resmi_tatil_gunleri): off_days.add(day)
            cell = ws3.cell(row=5, column=col_idx, value=int(day))
            cell.font, cell.alignment, cell.border = font_header, align_center, border_cell
            cell.fill = fill_grey_weekend if day in off_days else fill_white

        ek_basliklar = [
            "Toplam çalışma saati", "Aylık Çalışma Saati", "Fazla nöbet saati",
            "Normal Nöbet - Artırımlı (Gece)", "Normal Nöbet - Artırımsız (Normal)",
            "Riskli Nöbet - Artırımlı (Gece)", "Riskli Nöbet - Artırımsız (Normal)",
        ]

        start_col = gun_sayisi + 3
        for idx, b_adi in enumerate(ek_basliklar):
            c_i = start_col + idx
            cell = ws3.cell(row=5, column=c_i, value=b_adi)
            cell.font, cell.fill, cell.alignment, cell.border = font_header, fill_green_bg, align_vertical_header, border_cell

        gecmis_ay_son_gun_nobetcileri = list(set(gecmis_ay_son_gun_normal + gecmis_ay_son_gun_acil))

        for idx, p in enumerate(tum_girilen_personeller):
            r = 6 + idx
            ws3.row_dimensions[r].height = 18
            p_birim = TUM_PERSONEL_VERISI.get(p, {}).get("birim", birim_secimi)
            is_muaf = TUM_PERSONEL_VERISI.get(p, {}).get("muaf", False)

            c_name = ws3.cell(row=r, column=1, value=p)
            c_name.fill, c_name.font, c_name.alignment, c_name.border = fill_white, font_body, align_left, border_cell

            c_unit = ws3.cell(row=r, column=2, value=p_birim)
            c_unit.fill, c_unit.font, c_unit.alignment, c_unit.border = fill_white, font_body, align_center, border_cell

            p_shifts = nobet_dict.get(p, set())
            p_k8_shifts = kultur_8s_dict.get(p, set())
            p_k8_ni = kultur_ni_gunleri(yil, ay, gun_sayisi, p_k8_shifts, resmi_tatil_gunleri)
            p_acils = acil_nobet_dict.get(p, set())
            p_row_dict = {}

            for day in range(1, gun_sayisi + 1):
                col_idx = day + 2
                cell = ws3.cell(row=r, column=col_idx)
                cell.border, cell.alignment = border_cell, align_center
                cell.fill = fill_grey_weekend if day in off_days else fill_white
                dt = datetime.date(yil, ay, day)

                if is_muaf:
                    if day in off_days: cell.value = "T"
                    elif day in yarim_gun_tatil_gunleri: cell.value, cell.font = 5, font_body
                    else: cell.value, cell.font = 8, font_body
                else:
                    if p in gecmis_ay_kultur_ni and day == ilk_mesai_gunu:
                        cell.value, cell.font = "Nİ", font_ni
                    elif day == 1 and p in gecmis_ay_son_gun_nobetcileri:
                        cell.value = "T" if day in off_days else "Nİ"
                        if cell.value == "Nİ": cell.font = font_ni
                    elif day in p_shifts:
                        cell.value, cell.font = 24, font_24
                        if day in p_acils: cell.fill = fill_bright_yellow
                    elif day in p_k8_shifts: cell.value, cell.font = 8, font_bold
                    elif (day - 1) in p_shifts:
                        cell.value = "T" if day in off_days else "Nİ"
                        if cell.value == "Nİ": cell.font = font_ni
                    elif day in p_k8_ni:
                        cell.value, cell.font = "Nİ", font_ni
                    else:
                        if day in off_days: cell.value = "T"
                        elif day in yarim_gun_tatil_gunleri: cell.value, cell.font = 5, font_body
                        else: cell.value, cell.font = 8, font_body

                p_row_dict[str(day)] = str(cell.value)

            m = calculate_personel_puantaj_metrikleri(p, p_row_dict, yil, ay, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri, p_acils, kultur_8s_tatil_gunleri, gecmis_ay_son_gun_normal, gecmis_ay_son_gun_acil, p_k8_shifts)
            metrik_values = [
                int(m["Toplam Çalışma Saati"]), int(m["Aylık Çalışma Saati"]), int(m["Fazla Nöbet Saati"]),
                int(m["Normal_Gece"]), int(m["Normal_Normal"]), int(m["Riskli_Gece"]), int(m["Riskli_Normal"]),
            ]

            for m_idx, val in enumerate(metrik_values):
                col_idx = start_col + m_idx
                cell = ws3.cell(row=r, column=col_idx, value=val)
                cell.font = font_bold if m_idx < 3 else font_body
                cell.alignment, cell.border = align_center, border_cell

        summary_r = 6 + len(tum_girilen_personeller)
        ws3.row_dimensions[summary_r].height = 20

        c_sum_name = ws3.cell(row=summary_r, column=1, value="Hesaplanan Nöbet Saati")
        c_sum_name.fill, c_sum_name.font, c_sum_name.alignment, c_sum_name.border = fill_green_bg, font_bold, align_left, border_cell

        c_sum_unit = ws3.cell(row=summary_r, column=2, value="")
        c_sum_unit.fill, c_sum_unit.font, c_sum_unit.alignment, c_sum_unit.border = fill_green_bg, font_bold, align_center, border_cell

        for day in range(1, gun_sayisi + 1):
            col_idx = day + 2
            cell = ws3.cell(row=summary_r, column=col_idx, value=sum(calculate_nobet_hakedis(yil, ay, day, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri)))
            cell.font, cell.alignment, cell.border, cell.fill = font_bold, align_center, border_cell, fill_green_bg

        for m_idx in range(len(ek_basliklar)):
            c_i = start_col + m_idx
            cell = ws3.cell(row=summary_r, column=c_i, value="")
            cell.fill, cell.border = fill_green_bg, border_cell

        ws3.column_dimensions["A"].width = 28
        ws3.column_dimensions["B"].width = 12
        for day in range(1, gun_sayisi + 1):
            col_letter = get_column_letter(day + 2)
            ws3.column_dimensions[col_letter].width = 4.5

        for m_idx in range(len(ek_basliklar)):
            col_letter = get_column_letter(start_col + m_idx)
            ws3.column_dimensions[col_letter].width = 6.5

        ws_devir = wb.create_sheet("Kültür Nİ Devir")
        ws_devir.append(["Personel"])
        for p, vardiyalar in kultur_8s_dict.items():
            if any(not kultur_ni_gunleri(yil, ay, gun_sayisi, {d}, resmi_tatil_gunleri) for d in vardiyalar):
                ws_devir.append([p])
        ws_devir.sheet_state = "hidden"

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    # HESAPLAMA BUTONU
    if st.button("🚀 Otomatik ve Adil Nöbet Listesini Oluştur"):
        st.session_state.hesaplanan_sonuc = None
        hata_listesi = []
        pcr_nobetcileri = [p for p in nobetci_personeller if TUM_PERSONEL_VERISI.get(p, {}).get("birim") == "PCR"]
        mikro_nobetcileri = [p for p in nobetci_personeller if TUM_PERSONEL_VERISI.get(p, {}).get("birim") == "Mikro"]
        kultur_nobetcileri = [p for p in nobetci_personeller if TUM_PERSONEL_VERISI.get(p, {}).get("birim") == "Kültür"]

        for d in range(gun_sayisi):
            if birim_secimi == "Tüm Laboratuvar (Birleşik)":
                m_pcr = sum(1 for p in pcr_nobetcileri if d not in izinler[p])
                m_mikro = sum(1 for p in mikro_nobetcileri if d not in izinler[p])
                m_kultur = sum(1 for p in kultur_nobetcileri if d not in izinler[p])
                for birim, personeller, musait in [("PCR", pcr_nobetcileri, m_pcr), ("Mikro", mikro_nobetcileri, m_mikro), ("Kültür", kultur_nobetcileri, m_kultur)]:
                    if musait < gunluk_nobetci:
                        tarih_str = datetime.date(yil, ay, d + 1).strftime("%d.%m.%Y")
                        engeller = ", ".join(p for p in personeller if d in izinler[p]) or "kadro yetersiz"
                        hata_listesi.append(f"❌ {tarih_str} — {birim}: {gunluk_nobetci} kişi gerekli, {musait} kişi müsait. İzin/mazeret: {engeller}.")
            else:
                musait_sayisi = sum(1 for p in nobetci_personeller if d not in izinler[p])
                if musait_sayisi < gunluk_nobetci:
                    tarih_str = datetime.date(yil, ay, d + 1).strftime("%d.%m.%Y")
                    hata_listesi.append(f"❌ {tarih_str} — {birim_secimi}: {gunluk_nobetci} kişi gerekli, {musait_sayisi} kişi müsait. İzin/mazeret: {', '.join(p for p in nobetci_personeller if d in izinler[p]) or 'kadro yetersiz'}.")

        for p in nobetci_personeller:
            ortak = set(izinler[p]).intersection(set(sabit_nobetler[p]))
            for d in ortak:
                tarih_str = datetime.date(yil, ay, d + 1).strftime("%d.%m.%Y")
                hata_listesi.append(f"❌ **{p}**, **{tarih_str}** tarihinde hem 'İzinli' hem 'Sabit Nöbetçi'!")

        for d in range(1, gun_sayisi + 1):
            sabit_acil_kisiler = [p for p, gunler in sabit_acil_nobetler.items() if d in gunler]
            if len(sabit_acil_kisiler) > 1:
                tarih_str = datetime.date(yil, ay, d).strftime("%d.%m.%Y")
                hata_listesi.append(f"❌ {tarih_str}: günde tek acil nöbetçi olabilir. Çakışan sabit aciller: {', '.join(sabit_acil_kisiler)}.")

        if hata_listesi:
            for err in hata_listesi: st.error(err)
        else:
            model = cp_model.CpModel()
            x, k8 = {}, {}
            kisit_aciklamalari = {}
            def tarih(d):
                return datetime.date(yil, ay, d + 1).strftime("%d.%m.%Y")
            def zorunlu_kisit(ifade, aciklama):
                lit = model.NewBoolVar(f"kisit_{len(kisit_aciklamalari)}")
                model.Add(ifade).OnlyEnforceIf(lit)
                model.AddAssumption(lit)
                kisit_aciklamalari[lit.Index()] = aciklama


            for p in nobetci_personeller:
                for d in range(gun_sayisi): x[(p, d)] = model.NewBoolVar(f"x_{p}_{d}")
            for p in kultur_nobetcileri:
                for d in range(gun_sayisi): k8[(p, d)] = model.NewBoolVar(f"k8_{p}_{d}")

            if birim_secimi == "Tüm Laboratuvar (Birleşik)":
                for d in range(gun_sayisi):
                    zorunlu_kisit(sum(x[(p, d)] for p in pcr_nobetcileri) == gunluk_nobetci, f"{tarih(d)} — PCR: {gunluk_nobetci} nöbetçi gerekli")
                    zorunlu_kisit(sum(x[(p, d)] for p in mikro_nobetcileri) == gunluk_nobetci, f"{tarih(d)} — Mikro: {gunluk_nobetci} nöbetçi gerekli")
                    zorunlu_kisit(sum(x[(p, d)] for p in kultur_nobetcileri) == gunluk_nobetci, f"{tarih(d)} — Kültür: {gunluk_nobetci} nöbetçi gerekli")
            else:
                for d in range(gun_sayisi):
                    zorunlu_kisit(sum(x[(p, d)] for p in nobetci_personeller) == gunluk_nobetci, f"{tarih(d)} — {birim_secimi}: {gunluk_nobetci} nöbetçi gerekli")

            kultur_8s_gun_indeksleri = [d for d in range(gun_sayisi) if datetime.date(yil, ay, d + 1).weekday() == 5 or (d + 1) in kultur_8s_tatil_gunleri]
            for d in range(gun_sayisi):
                if d in kultur_8s_gun_indeksleri and kultur_nobetcileri:
                    zorunlu_kisit(sum(k8[(p, d)] for p in kultur_nobetcileri) == 1, f"{tarih(d)} — Kültür 8s: 1 vardiyacı gerekli")
                else:
                    for p in kultur_nobetcileri: zorunlu_kisit(k8[(p, d)] == 0, f"{p} — {tarih(d)}: Kültür 8s vardiya günü değil")

            for p in kultur_nobetcileri:
                for d in kultur_8s_gun_indeksleri:
                    ni_days = kultur_ni_gunleri(yil, ay, gun_sayisi, {d + 1}, resmi_tatil_gunleri)
                    for ni_day in ni_days:
                        zorunlu_kisit(k8[(p, d)] + x[(p, ni_day - 1)] <= 1, f"{p} — {tarih(d)} Kültür 8s sonrası {tarih(ni_day - 1)} Nİ: nöbet verilemez")
                        zorunlu_kisit(k8[(p, d)] + k8[(p, ni_day - 1)] <= 1, f"{p} — {tarih(d)} Kültür 8s sonrası {tarih(ni_day - 1)} Nİ: vardiya verilemez")

            for p in nobetci_personeller:
                for d in izinler[p]:
                    zorunlu_kisit(x[(p, d)] == 0, f"{p} — {tarih(d)}: izin/mazeret nedeniyle nöbet verilemez")
                    if p in kultur_nobetcileri: zorunlu_kisit(k8[(p, d)] == 0, f"{p} — {tarih(d)}: izin/mazeret nedeniyle Kültür 8s verilemez")

            for p in nobetci_personeller:
                for d in sabit_nobetler[p]: zorunlu_kisit(x[(p, d)] == 1, f"{p} — {tarih(d)}: sabit nöbet")

            for p, gunler in sabit_acil_nobetler.items():
                for d in set(gunler):
                    zorunlu_kisit(x[(p, d - 1)] == 1, f"{p} — {tarih(d - 1)}: sabit acil için 24 saatlik nöbet zorunlu")

            if ilk_mesai_gunu is not None:
                for p in gecmis_ay_kultur_ni:
                    zorunlu_kisit(x[(p, ilk_mesai_gunu - 1)] == 0, f"{p} — {tarih(ilk_mesai_gunu - 1)}: önceki aydan Kültür Nİ devri")
                    if p in kultur_nobetcileri:
                        zorunlu_kisit(k8[(p, ilk_mesai_gunu - 1)] == 0, f"{p} — {tarih(ilk_mesai_gunu - 1)}: Kültür Nİ devrinde vardiya verilemez")

            for p in gecmis_ay_son_gun_nobetcileri:
                if p in nobetci_personeller:
                    zorunlu_kisit(x[(p, 0)] == 0, f"{p} — {tarih(0)}: önceki ay son gün nöbeti sonrası dinlenme")
                    if p in kultur_nobetcileri: zorunlu_kisit(k8[(p, 0)] == 0, f"{p} — {tarih(0)}: önceki ay son gün nöbeti sonrası vardiya verilemez")

            for p in nobetci_personeller:
                p_dinlenme = 1 if p in esnek_personel else max(1, int(dinlenme_gun_sayisi))
                for d in range(gun_sayisi):
                    if p in kultur_nobetcileri:
                        zorunlu_kisit(x[(p, d)] + k8[(p, d)] <= 1, f"{p} — {tarih(d)}: nöbet ve Kültür 8s çakışamaz")
                        if d > 0: zorunlu_kisit(x[(p, d - 1)] + k8[(p, d)] <= 1, f"{p} — {tarih(d - 1)} nöbeti sonrası {tarih(d)} Kültür 8s verilemez")
                        if d < gun_sayisi - 1: zorunlu_kisit(k8[(p, d)] + x[(p, d + 1)] <= 1, f"{p} — {tarih(d)} Kültür 8s sonrası {tarih(d + 1)} nöbet verilemez")

                    for k in range(1, p_dinlenme + 1):
                        if d + k < gun_sayisi:
                            zorunlu_kisit(x[(p, d)] + x[(p, d + k)] <= 1, f"{p} — {tarih(d)} / {tarih(d + k)}: en az {p_dinlenme} tam gün dinlenme")
                            if p in kultur_nobetcileri:
                                zorunlu_kisit(x[(p, d)] + k8[(p, d + k)] <= 1, f"{p} — {tarih(d)} nöbet / {tarih(d + k)} Kültür 8s: {p_dinlenme} gün dinlenme")
                                zorunlu_kisit(k8[(p, d)] + x[(p, d + k)] <= 1, f"{p} — {tarih(d)} Kültür 8s / {tarih(d + k)} nöbet: {p_dinlenme} gün dinlenme")

            if persembe_pazar_yasagi:
                for d in range(gun_sayisi - 3):
                    if datetime.date(yil, ay, d + 1).weekday() == 3:
                        for p in nobetci_personeller:
                            if p not in esnek_personel:
                                zorunlu_kisit(x.get((p, d), 0) + x.get((p, d + 3), 0) <= 1, f"{p} — {tarih(d)} / {tarih(d + 3)}: Perşembe–Pazar yasağı")

            for rule in kisi_kisitlari:
                p1, aralik = rule["ana"], rule["aralik"]
                for p2 in rule["yasaklilar"]:
                    for d1 in range(gun_sayisi):
                        for d2 in range(max(0, d1 - aralik), min(gun_sayisi, d1 + aralik + 1)):
                            p1_gorev = x[(p1, d1)] + (k8[(p1, d1)] if p1 in kultur_nobetcileri else 0)
                            p2_gorev = x[(p2, d2)] + (k8[(p2, d2)] if p2 in kultur_nobetcileri else 0)
                            zorunlu_kisit(p1_gorev + p2_gorev <= 1, f"{p1} ({tarih(d1)}) / {p2} ({tarih(d2)}): kişiler arası {aralik} gün mesafe/çakışma yasağı")

            def gun_kategorisi_indeksleri(w_list):
                return [d for d in range(gun_sayisi) if datetime.date(yil, ay, d + 1).weekday() in w_list]

            cuma_indeksleri = gun_kategorisi_indeksleri([4])
            cumartesi_indeksleri = gun_kategorisi_indeksleri([5])
            pazar_indeksleri = gun_kategorisi_indeksleri([6])

            birimler_listesi = ["PCR", "Mikro", "Kültür"] if birim_secimi == "Tüm Laboratuvar (Birleşik)" else [birim_secimi]
            saat_farklari, adet_farklari, kategori_farklari = [], [], []
            gunluk_saatler = [sum(calculate_nobet_hakedis(yil, ay, d + 1, gun_sayisi,
                              resmi_tatil_gunleri, yarim_gun_tatil_gunleri)) for d in range(gun_sayisi)]
            def fark_hedefi(degerler, alt, ust, ad):
                en_az = model.NewIntVar(alt, ust, f"min_{ad}")
                en_cok = model.NewIntVar(alt, ust, f"max_{ad}")
                model.AddMinEquality(en_az, degerler)
                model.AddMaxEquality(en_cok, degerler)
                return en_cok - en_az

            for birim in birimler_listesi:
                personeller = [p for p in nobetci_personeller
                               if TUM_PERSONEL_VERISI.get(p, {}).get("birim") == birim
                               or birim_secimi != "Tüm Laboratuvar (Birleşik)"]
                if not personeller: continue
                saatler, adetler = [], []
                for p in personeller:
                    saat = model.NewIntVar(0, sum(gunluk_saatler), f"hakedis_{p}")
                    adet = model.NewIntVar(0, gun_sayisi, f"adet_{p}")
                    model.Add(saat == sum(x[(p, d)] * gunluk_saatler[d] for d in range(gun_sayisi)))
                    model.Add(adet == sum(x[(p, d)] for d in range(gun_sayisi)))
                    saatler.append(saat)
                    adetler.append(adet)
                saat_farki = fark_hedefi(saatler, 0, sum(gunluk_saatler), birim + "_saat")
                adet_farki = fark_hedefi(adetler, 0, gun_sayisi, birim + "_adet")
                # Günlük ihtiyaç sabit: toplamın kişi sayısına bölünmesinden
                # gelen zorunlu alt sınırlar aramayı hızlandırır.
                saat_adimi = math.gcd(*gunluk_saatler)
                if (sum(gunluk_saatler) * gunluk_nobetci // saat_adimi) % len(personeller):
                    model.Add(saat_farki >= saat_adimi)
                if (gun_sayisi * gunluk_nobetci) % len(personeller):
                    model.Add(adet_farki >= 1)
                saat_farklari.append(saat_farki)
                adet_farklari.append(adet_farki)
                for weekday in range(7):
                    gun_adi = tr_gunler[weekday]
                    indices = gun_kategorisi_indeksleri([weekday])
                    devirler = [int(get_prev(p, gun_adi)) for p in personeller]
                    alt, ust = min(devirler), max(devirler) + len(indices)
                    degerler = []
                    for p, devir in zip(personeller, devirler):
                        v = model.NewIntVar(alt, ust, f"gun_{p}_{weekday}")
                        model.Add(v == devir + sum(x[(p, d)] for d in indices))
                        degerler.append(v)
                    kategori_farklari.append(fark_hedefi(degerler, alt, ust, birim + gun_adi))

            # 2 hafta sonu ve tekrarlar tercih; hiçbiri çözümü engellemez.
            hafta_gruplari, haftasonu_gruplari = {}, {}
            for d in range(gun_sayisi):
                dt = datetime.date(yil, ay, d + 1)
                hafta = dt - datetime.timedelta(days=dt.weekday())
                hafta_gruplari.setdefault(hafta, []).append(d)
                if dt.weekday() >= 4:
                    haftasonu_gruplari.setdefault(hafta, []).append(d)
            yayilim_cezalari = []
            for p in nobetci_personeller:
                for hafta, indices in hafta_gruplari.items():
                    fazla = model.NewIntVar(0, gun_sayisi, f"hafta_tekrar_{p}_{hafta}")
                    model.AddMaxEquality(fazla, [0, sum(x[(p, d)] for d in indices) - 1])
                    yayilim_cezalari.append(fazla)
                weekend_vars = []
                for hafta, indices in haftasonu_gruplari.items():
                    var = model.NewBoolVar(f"haftasonu_{p}_{hafta}")
                    model.AddMaxEquality(var, [x[(p, d)] for d in indices])
                    weekend_vars.append(var)
                fazla = model.NewIntVar(0, len(weekend_vars), f"haftasonu_fazla_{p}")
                model.AddMaxEquality(fazla, [0, sum(weekend_vars) - 2])
                yayilim_cezalari.append(fazla)
                if cuma_haftasonu_siki_kural:
                    for weekday in [4, 5, 6]:
                        indices = gun_kategorisi_indeksleri([weekday])
                        tekrar = model.NewIntVar(0, gun_sayisi, f"tur_tekrar_{p}_{weekday}")
                        model.AddMaxEquality(tekrar, [0, sum(x[(p, d)] for d in indices) - 1])
                        yayilim_cezalari.append(tekrar)

            kultur_hedefleri = []
            if kultur_nobetcileri:
                devirler = [int(get_prev_kultur8(p)) for p in kultur_nobetcileri]
                alt, ust = min(devirler), max(devirler) + len(kultur_8s_gun_indeksleri)
                degerler = []
                for p, devir in zip(kultur_nobetcileri, devirler):
                    var = model.NewIntVar(alt, ust, f"kultur_adet_{p}")
                    model.Add(var == devir + sum(k8[(p, d)] for d in kultur_8s_gun_indeksleri))
                    degerler.append(var)
                kultur_hedefleri.append(fark_hedefi(degerler, alt, ust, "kultur8"))

            hedefler = [
                ("T.NöbetSaati dengesi", sum(saat_farklari)),
                ("Nöbet adedi dengesi", sum(adet_farklari)),
                ("Gün türü devir dengesi", sum(kategori_farklari)),
                ("Haftalara yayılım", sum(yayilim_cezalari)),
                ("Kültür 8s adet dengesi", sum(kultur_hedefleri)),
            ]
            solver = None
            status = cp_model.UNKNOWN
            optimizasyon_ozeti = []
            ilerleme = st.empty()
            for asama, (ad, hedef) in enumerate(hedefler, 1):
                ilerleme.info(f"Dağıtım kontrolü {asama}/5: {ad}")
                model.Minimize(hedef)
                aday = cp_model.CpSolver()
                aday.parameters.num_search_workers = 1
                aday.parameters.max_time_in_seconds = float(cozum_arama_suresi)
                aday_status = aday.Solve(model)
                if aday_status not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
                    if solver is None:
                        solver, status = aday, aday_status
                    else:
                        st.warning(f"{ad} için arama süresi içinde yeni sonuç alınamadı; önceki geçerli liste korundu.")
                    break
                solver, status = aday, aday_status
                deger = int(solver.Value(hedef))
                optimizasyon_ozeti.append({"Öncelik": asama, "Hedef": ad,
                    "Fark / ceza": deger, "En iyi sonuç kanıtlandı": status == cp_model.OPTIMAL})
                if status != cp_model.OPTIMAL:
                    if asama == 1:
                        st.warning(f"{ad}: geçerli liste bulundu, ancak süre içinde en küçük fark kanıtlanamadı. Saat önceliği nedeniyle alt hedeflere geçilmedi; arama süresini artırabilirsiniz.")
                        break
                    st.warning(f"{ad}: en küçük fark süre içinde kanıtlanamadı. Bulunan değer korunarak sonraki öncelik değerlendirilecek.")
                # Önceki hedefin elde edilmiş değeri değiştirilemez.
                model.Add(hedef == deger)
                model.ClearHints()
                for var in list(x.values()) + list(k8.values()):
                    model.AddHint(var, solver.Value(var))
            ilerleme.empty()
            if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
                nobet_dict = {p: set() for p in nobetci_personeller}
                kultur_8s_dict = {p: set() for p in nobetci_personeller}
                acil_nobet_dict = {p: set() for p in nobetci_personeller}

                for d in range(gun_sayisi):
                    for p in nobetci_personeller:
                        if solver.Value(x[(p, d)]) == 1: nobet_dict[p].add(d + 1)
                        if p in kultur_nobetcileri and solver.Value(k8[(p, d)]) == 1: kultur_8s_dict[p].add(d + 1)

                with st.spinner("Ortak havuzda acil nöbet dağılımı dengeleniyor..."):
                    acil_nobet_dict, acil_dagitim_ozeti = dagit_acil_nobet(
                        nobetci_personeller, nobet_dict, sabit_acil_nobetler,
                        {d: gunluk_saatler[d - 1] for d in range(1, gun_sayisi + 1)},
                        {p: int(get_prev_acil(p)) for p in nobetci_personeller}, cozum_arama_suresi)

                liste_data = []
                for d in range(1, gun_sayisi + 1):
                    tarih = datetime.date(yil, ay, d)
                    pcr_list, mikro_list, kultur_list, kultur_8s_list = [], [], [], []
                    for p in nobetci_personeller:
                        p_unit = TUM_PERSONEL_VERISI.get(p, {}).get("birim", "Mikro")
                        is_24 = d in nobet_dict[p]
                        is_8s = d in kultur_8s_dict[p]
                        is_acil = d in acil_nobet_dict[p]

                        if is_24:
                            p_text = f"{p} (Acil)" if is_acil else p
                            if p_unit == "PCR": pcr_list.append(p_text)
                            elif p_unit == "Mikro": mikro_list.append(p_text)
                            elif p_unit == "Kültür": kultur_list.append(p_text)
                        elif is_8s: kultur_8s_list.append(p)

                    liste_data.append({
                        "Tarih": tarih.strftime("%d.%m.%Y"), "Gün": tr_gunler[tarih.weekday()],
                        "PCR": ", ".join(pcr_list), "Mikro": ", ".join(mikro_list),
                        "Kültür": ", ".join(kultur_list), "Kültür (8s Vardiya)": ", ".join(kultur_8s_list),
                    })
                df_liste = pd.DataFrame(liste_data)

                gunler_listesi = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
                columns_tuples = [("AD SOYAD", ""), ("Birim", "")]
                for g in gunler_listesi: columns_tuples.extend([(g, "Devir"), (g, "Bu Ay"), (g, "Toplam")])
                columns_tuples.extend([("T.NöbetSaati", ""), ("TOPLAM NÖBET", "")])
                columns_tuples.extend([("ACİL NÖBET (SAAT)", "Devir"), ("ACİL NÖBET (SAAT)", "Bu Ay"), ("ACİL NÖBET (SAAT)", "Toplam")])
                columns_tuples.extend([("KÜLTÜR 8S VARDİYA", "Devir"), ("KÜLTÜR 8S VARDİYA", "Bu Ay"), ("KÜLTÜR 8S VARDİYA", "Toplam")])

                multi_cols = pd.MultiIndex.from_tuples(columns_tuples)
                istatistik_rows = []

                for p in nobetci_personeller:
                    p_birim = TUM_PERSONEL_VERISI.get(p, {}).get("birim", birim_secimi)
                    bu_ay_gunler = {g: 0 for g in gunler_listesi}
                    bu_ay_toplam_nobet = len(nobet_dict[p])
                    bu_ay_saat = sum(sum(calculate_nobet_hakedis(yil, ay, d, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri)) for d in nobet_dict.get(p, set()))
                    for d in nobet_dict[p]:
                        w = datetime.date(yil, ay, d).weekday()
                        bu_ay_gunler[tr_gunler[w]] += 1

                    row_dict = {("AD SOYAD", ""): p, ("Birim", ""): p_birim}
                    for g in gunler_listesi:
                        devir_val = get_prev(p, g)
                        bu_ay_val = bu_ay_gunler[g]
                        row_dict[(g, "Devir")] = devir_val
                        row_dict[(g, "Bu Ay")] = bu_ay_val
                        row_dict[(g, "Toplam")] = devir_val + bu_ay_val

                    row_dict[("T.NöbetSaati", "")] = bu_ay_saat
                    row_dict[("TOPLAM NÖBET", "")] = bu_ay_toplam_nobet

                    bu_ay_acil_saat_val = sum(sum(calculate_nobet_hakedis(yil, ay, d, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri)) for d in acil_nobet_dict.get(p, set()))
                    acil_devir = get_prev_acil(p)
                    row_dict[("ACİL NÖBET (SAAT)", "Devir")] = acil_devir
                    row_dict[("ACİL NÖBET (SAAT)", "Bu Ay")] = bu_ay_acil_saat_val
                    row_dict[("ACİL NÖBET (SAAT)", "Toplam")] = acil_devir + bu_ay_acil_saat_val

                    bu_ay_k8_val = len(kultur_8s_dict[p])
                    k8_devir = get_prev_kultur8(p)
                    row_dict[("KÜLTÜR 8S VARDİYA", "Devir")] = k8_devir
                    row_dict[("KÜLTÜR 8S VARDİYA", "Bu Ay")] = bu_ay_k8_val
                    row_dict[("KÜLTÜR 8S VARDİYA", "Toplam")] = k8_devir + bu_ay_k8_val

                    istatistik_rows.append(row_dict)

                df_istatistik = pd.DataFrame(istatistik_rows, columns=multi_cols)

                puantaj_rows = []
                for p in tum_girilen_personeller:
                    p_birim = TUM_PERSONEL_VERISI.get(p, {}).get("birim", birim_secimi)
                    is_muaf = TUM_PERSONEL_VERISI.get(p, {}).get("muaf", False)
                    p_row = {"Adı Soyadı": p, "Birim": p_birim}
                    p_shifts = nobet_dict.get(p, set())
                    p_k8_shifts = kultur_8s_dict.get(p, set())
                    p_k8_ni = kultur_ni_gunleri(yil, ay, gun_sayisi, p_k8_shifts, resmi_tatil_gunleri)

                    for d in range(1, gun_sayisi + 1):
                        day_is_off = is_day_off(yil, ay, d, resmi_tatil_gunleri)
                        dt = datetime.date(yil, ay, d)

                        if is_muaf:
                            if day_is_off: p_row[str(d)] = "T"
                            elif d in yarim_gun_tatil_gunleri: p_row[str(d)] = "5"
                            else: p_row[str(d)] = "8"
                        else:
                            if p in gecmis_ay_kultur_ni and d == ilk_mesai_gunu: p_row[str(d)] = "Nİ"
                            elif d == 1 and p in gecmis_ay_son_gun_nobetcileri: p_row[str(d)] = "T" if day_is_off else "Nİ"
                            elif d in p_shifts: p_row[str(d)] = "24"
                            elif d in p_k8_shifts: p_row[str(d)] = "8"
                            elif (d - 1) in p_shifts: p_row[str(d)] = "T" if day_is_off else "Nİ"
                            elif d in p_k8_ni: p_row[str(d)] = "Nİ"
                            else:
                                if day_is_off: p_row[str(d)] = "T"
                                elif d in yarim_gun_tatil_gunleri: p_row[str(d)] = "5"
                                else: p_row[str(d)] = "8"

                    m = calculate_personel_puantaj_metrikleri(p, p_row, yil, ay, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri, acil_nobet_dict.get(p, set()), kultur_8s_tatil_gunleri, gecmis_ay_son_gun_normal, gecmis_ay_son_gun_acil, p_k8_shifts)
                    p_row["Toplam çalışma saati"] = m["Toplam Çalışma Saati"]
                    p_row["Aylık Çalışma Saati"] = m["Aylık Çalışma Saati"]
                    p_row["Fazla nöbet saati"] = m["Fazla Nöbet Saati"]
                    p_row["Normal Nöbet - Artırımlı (Gece)"] = m["Normal_Gece"]
                    p_row["Normal Nöbet - Artırımsız (Normal)"] = m["Normal_Normal"]
                    p_row["Riskli Nöbet - Artırımlı (Gece)"] = m["Riskli_Gece"]
                    p_row["Riskli Nöbet - Artırımsız (Normal)"] = m["Riskli_Normal"]
                    puantaj_rows.append(p_row)

                df_puantaj = pd.DataFrame(puantaj_rows)
                excel_bytes = generate_3_tab_excel(
                    yil, ay, nobetci_personeller, tum_girilen_personeller,
                    nobet_dict, kultur_8s_dict, acil_nobet_dict,
                    gecmis_istatistik, gecmis_acil_istatistik, gecmis_kultur8_istatistik,
                    gun_sayisi, birim_secimi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri,
                    kultur_8s_tatil_gunleri, gecmis_ay_son_gun_normal, gecmis_ay_son_gun_acil,
                )

                st.session_state.hesaplanan_sonuc = {
                    "optimizasyon_ozeti": optimizasyon_ozeti,
                    "acil_dagitim_ozeti": acil_dagitim_ozeti,
                    "df_liste": df_liste, "df_istatistik": df_istatistik,
                    "df_puantaj": df_puantaj, "excel_bytes": excel_bytes,
                    "birim_secimi": birim_secimi, "yil": yil, "ay": ay,
                }
                st.balloons()
                st.success("✨ Nöbet çizelgesi ve Puantaj tablosu başarıyla oluşturuldu!")

            else:
                if status == cp_model.INFEASIBLE:
                    st.error("❌ Zorunlu kısıtlar birlikte uygulanınca çözüm yok. Esnek dağıtım tercihleri bunun nedeni değildir.")
                    tanilama_modeli = model.Clone()
                    tanilama_modeli.ClearObjective()
                    tanilama_modeli.ClearHints()
                    tani_solver = cp_model.CpSolver()
                    tani_solver.parameters.num_search_workers = 1
                    tani_solver.parameters.max_time_in_seconds = 3.0
                    tani_status = tani_solver.Solve(tanilama_modeli)
                    core = list(tani_solver.SufficientAssumptionsForInfeasibility()) if tani_status == cp_model.INFEASIBLE else list(solver.SufficientAssumptionsForInfeasibility())
                    # En fazla 20 kısa deneme; minimal küme olduğu iddia edilmez.
                    for lit in core[:20]:
                        aday_core = [i for i in core if i != lit]
                        tanilama_modeli.ClearAssumptions()
                        tanilama_modeli.AddAssumptions([tanilama_modeli.GetBoolVarFromProtoIndex(i) for i in aday_core])
                        tani_solver.parameters.max_time_in_seconds = 0.2
                        if tani_solver.Solve(tanilama_modeli) == cp_model.INFEASIBLE:
                            core = aday_core
                    mesajlar = list(dict.fromkeys(kisit_aciklamalari.get(i, f"Kısıt {i}") for i in core))
                    if mesajlar:
                        st.write("Birlikte çözümü engelleyen kısıt kümesi (tek tek hepsinin kaldırılması gerekmez):")
                        for mesaj in mesajlar:
                            st.error(mesaj)
                    else:
                        st.error("Kısıt çakışması doğrulandı; çözücü kişi/tarih içeren bir kısıt kümesi döndürmedi.")
                elif status == cp_model.MODEL_INVALID:
                    st.error("❌ Dağıtım modeli geçersiz: " + model.Validate())
                else:
                    st.warning("⏱️ Arama süresi içinde geçerli liste bulunamadı. Bu, kuralların çeliştiğinin kanıtı değildir. Süreyi artırarak tekrar deneyin.")

    # SONUÇLARI EKRANDA GÖSTER
    if st.session_state.hesaplanan_sonuc is not None:
        sonuc = st.session_state.hesaplanan_sonuc
        if sonuc.get("optimizasyon_ozeti"):
            with st.expander("Dağıtım öncelikleri ve kontrol sonuçları"):
                st.caption("İlk üç hedefteki değerler birim içi farkların toplamıdır. Son iki hedefte tekrar/yayılım cezası ve Kültür adet farkı gösterilir. En iyi sonuç kanıtlanmadıysa arama süresini artırabilirsiniz.")
                st.dataframe(pd.DataFrame(sonuc["optimizasyon_ozeti"]), use_container_width=True)
        if sonuc.get("acil_dagitim_ozeti"):
            with st.expander("Ortak havuz acil dağıtımı kontrol sonuçları"):
                st.caption("Önce farklı personele görev ve düşük devir önceliği; ardından aylık ve birikimli hakediş saat dengesi. Normal nöbet çizelgesi değiştirilmez.")
                st.dataframe(pd.DataFrame(sonuc["acil_dagitim_ozeti"]), use_container_width=True)
        b64 = base64.b64encode(sonuc["excel_bytes"].getvalue()).decode()
        file_name = f"Nobet_ve_Puantaj_Listesi_{sonuc['birim_secimi']}_{sonuc['yil']}_{sonuc['ay']}.xlsx"
        href_link = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{file_name}" class="direct-download-btn">📥 3 Sekmeli Resmi Excel Dosyasını İndir (.xlsx)</a>'

        tab1, tab2, tab3, tab4 = st.tabs(["📅 Aylık Çizelge", "📊 İstatistik & Mesai", "📋 Puantaj Matrisi", "📥 Excel İndir"])
        with tab1:
            st.subheader("🗓️ Birim Bazlı Aylık Görev Listesi (PCR | Mikro | Kültür)")
            st.dataframe(sonuc["df_liste"], use_container_width=True, height=450)
        with tab2:
            st.subheader("📈 Personel Mesai Yükü & Acil / Kültür 8s İstatistiği")
            st.dataframe(sonuc["df_istatistik"], use_container_width=True)
        with tab3:
            st.subheader("📋 Resmi Puantaj Tablosu Önizleme (Sarı Hücreler = Acil Nöbet)")
            st.dataframe(sonuc["df_puantaj"], use_container_width=True)
        with tab4:
            st.subheader("📥 Excel Dosyasını İndir")
            st.write("Aşağıdaki butona tıkladığınızda dosyanız doğrudan bilgisayarınıza indirilecektir:")
            st.markdown(href_link, unsafe_allow_html=True)


# ==========================================
# 2. MODÜL: UZMAN DOKTOR ÇALIŞMA LİSTESİ
# ==========================================
elif secilen_modul == "2. Uzman Dr. Çalışma Listesi":
    st.markdown(
        """
    <div class="header-box">
        <h1>👨‍⚕️ Uzman Doktor Çalışma ve Nöbet Kodlama Sistemi</h1>
        <p>Klinik Bazlı Çalışma Çizelgesinden Hekim Bazlı L-Kodlu (L-1 - L-20) Çalışma Listesi Oluşturma</p>
        <div class="header-imza">✍️ Özgür SARSILMAZ</div>
    </div>
    """,
        unsafe_allow_html=True,
    )

    uzman_doc_mapping = {
        'Ali Osman ŞEKERCİOĞLU': 'A.O.Şekercioğlu',
        'ALPER KANDİŞER': 'A.Kandişer',
        'Aydan KARAGÜL': 'A.Karagül',
        'Ayşe SARI': 'A.Sarı',
        'C.Aylin ERMAN DALOĞLU': 'C.A.E.Daloğlu',
        'Çiğdem YILDIRIM': 'Ç.Yıldırım',
        'Ebru KANDIRALI DUYGUN': 'E.K.Duygun',
        'Gül AYDIN TIĞLI': 'G.A.Tığlı',
        'H.Nevgün ÖZEN': 'H.N.Özen',
        'Halil ER': 'H.Er',
        'Halil MANSUROĞLU': 'H.Mansuroğlu',
        'Koray ÖNCEL': 'K.Öncel',
        'Nilgün GÜR': 'N.Gür',
        'Özgül ÇETİNKAYA': 'Ö.Çetinkaya',
        'Özgür DOĞAN': 'Ö.Doğan',
        'Özlem KOCA': 'Ö.Koca',
        'Yeşim ÇEKİN': 'Y.Çekin',
        'Zübeyde ERES SARITAŞ': 'Z.E.Sarıtaş'
    }

    st.subheader("📂 Uzman Doktor Çalışma Listesi Excel Dosyası Yükleme")
    uploaded_uzman_file = st.file_uploader("Uzman Dr. Çalışma Listesi Excel Dosyasını Yükleyin (.xlsx)", type=["xlsx"], key="uzman_uploader")

    if uploaded_uzman_file is not None:
        try:
            wb = openpyxl.load_workbook(uploaded_uzman_file)
            
            if 'LAB.KLİNİK BAZLI ÇALIŞMA LİST.' not in wb.sheetnames or 'LAB.HEKİM BAZLI ÇALIŞMA LİST' not in wb.sheetnames:
                st.error("❌ Hata: Yüklenen dosyada 'LAB.KLİNİK BAZLI ÇALIŞMA LİST.' veya 'LAB.HEKİM BAZLI ÇALIŞMA LİST' sekmeleri bulunamadı.")
            else:
                sh_klinik = wb['LAB.KLİNİK BAZLI ÇALIŞMA LİST.']
                sh_hekim = wb['LAB.HEKİM BAZLI ÇALIŞMA LİST']

                uzman_schedule = {day: {} for day in range(1, 32)}

                for day in range(1, 32):
                    r = day + 7
                    for c in range(3, 23):
                        code = sh_klinik.cell(row=6, column=c).value
                        doc = sh_klinik.cell(row=r, column=c).value
                        if doc and str(doc).strip():
                            doc_str = str(doc).strip()
                            if doc_str not in uzman_schedule[day]:
                                uzman_schedule[day][doc_str] = []
                            uzman_schedule[day][doc_str].append(code)

                uzman_result_grid = {}
                for full_name, short_name in uzman_doc_mapping.items():
                    uzman_result_grid[full_name] = {}
                    for day in range(1, 32):
                        codes = uzman_schedule[day].get(short_name, [])
                        if not codes: final_code = ""
                        elif 'L-14' in codes: final_code = 'L-14'
                        elif 'L-15' in codes: final_code = 'L-15'
                        else: final_code = codes[0]
                        uzman_result_grid[full_name][day] = final_code

                for r in range(9, 27):
                    full_name = str(sh_hekim.cell(row=r, column=2).value).strip()
                    if full_name in uzman_result_grid:
                        for day in range(1, 32):
                            col_idx = day + 4
                            val = uzman_result_grid[full_name][day]
                            sh_hekim.cell(row=r, column=col_idx).value = val if val else None

                output_uzman = BytesIO()
                wb.save(output_uzman)
                output_uzman.seek(0)

                uzman_grid_rows = []
                for doc_n in uzman_doc_mapping:
                    row_data = {"Hekim Adı Soyadı": doc_n}
                    for d in range(1, 32):
                        v = uzman_result_grid[doc_n][d]
                        row_data[str(d)] = v if v else "-"
                    uzman_grid_rows.append(row_data)

                df_uzman_preview = pd.DataFrame(uzman_grid_rows)

                st.success("✅ Uzman Doktor çalışma listesi başarıyla L-kodlarına (L-1 - L-20) dönüştürüldü ve kodlandı!")
                
                st.subheader("📊 Kodlanmış Uzman Doktor Çizelgesi Önizleme")
                st.dataframe(df_uzman_preview, use_container_width=True)

                st.markdown("---")
                st.subheader("📥 Kodlanmış Excel Dosyasını İndir")
                b64_uzman = base64.b64encode(output_uzman.getvalue()).decode()
                uzman_file_name = "Uzman_Dr_Calisma_Listesi_10.2026_Kodlanmis.xlsx"
                href_uzman = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_uzman}" download="{uzman_file_name}" class="direct-download-btn">📥 Kodlanmış Uzman Dr. Excel Dosyasını İndir (.xlsx)</a>'
                st.markdown(href_uzman, unsafe_allow_html=True)

        except Exception as e:
            st.error(f"❌ Dosya işlenirken bir hata oluştu: {e}")


# ==========================================
# 3. MODÜL: ASİSTAN DOKTOR ÇALIŞMA LİSTESİ
# ==========================================
elif secilen_modul == "3. Asistan Dr. Çalışma Listesi":
    st.markdown(
        """
    <div class="header-box">
        <h1>🧑‍⚕️ Asistan Doktor Çalışma ve Nöbet Kodlama Sistemi</h1>
        <p>Klinik Bazlı Çalışma Çizelgesinden Hekim Bazlı L-Kodlu (L-1 - L-11) Çalışma Listesi Oluşturma</p>
        <div class="header-imza">✍️ Özgür SARSILMAZ</div>
    </div>
    """,
        unsafe_allow_html=True,
    )

    as_doc_mapping = {
        'YAREN ŞEKERCİOĞLU': ['Yaren Şekercioğlu', 'Yaren Şekercioglu'],
        'KENAN SAĞIR': ['Kenan Sağır', 'Kenan Sagır'],
        'FATMA ÇANKAYA': ['Fatma Çankaya', 'Fatma Cankaya'],
        'BETÜL BEDEL': ['Betül Bedel', 'Betul Bedel'],
        'ŞULE S. YILMAZ': ['Şule S. Yılma', 'Şule S. Yılmaz', 'Şule S.Yılmaz', 'Sule S. Yilmaz'],
        'CANAY DEMİRTAŞ': ['Canay Demirtaş', 'Canay Demirtas'],
        'OĞUZ AKIN': ['Oğuz Akın', 'Oguz Akin'],
        'ASLI HESAPÇI': ['Aslı Hesapçı', 'Asli Hesapci']
    }

    def match_as_doc(raw_name, target_full_name):
        if not raw_name: return False
        raw_name = str(raw_name).strip().lower()
        targets = as_doc_mapping[target_full_name]
        return any(raw_name == t.lower() for t in targets)

    st.subheader("📂 Asistan Doktor Çalışma Listesi Excel Dosyası Yükleme")
    uploaded_as_file = st.file_uploader("As.Dr. Çalışma Listesi Excel Dosyasını Yükleyin (.xlsx)", type=["xlsx"], key="asistan_uploader")

    if uploaded_as_file is not None:
        try:
            wb = openpyxl.load_workbook(uploaded_as_file)
            
            if 'LAB.KLİNİK BAZLI ÇALIŞMA LİST.' not in wb.sheetnames or 'LAB.HEKİM BAZLI ÇALIŞMA LİST' not in wb.sheetnames:
                st.error("❌ Hata: Yüklenen dosyada 'LAB.KLİNİK BAZLI ÇALIŞMA LİST.' veya 'LAB.HEKİM BAZLI ÇALIŞMA LİST' sekmeleri bulunamadı.")
            else:
                sh_klinik_as = wb['LAB.KLİNİK BAZLI ÇALIŞMA LİST.']
                sh_hekim_as = wb['LAB.HEKİM BAZLI ÇALIŞMA LİST']

                as_schedule = {day: {} for day in range(1, 32)}
                
                for day in range(1, 32):
                    r = day + 7
                    for c in range(3, 14):
                        code = sh_klinik_as.cell(row=6, column=c).value
                        doc = sh_klinik_as.cell(row=r, column=c).value
                        if doc and str(doc).strip():
                            doc_str = str(doc).strip()
                            matched_doc = None
                            for full_n in as_doc_mapping:
                                if match_as_doc(doc_str, full_n):
                                    matched_doc = full_n
                                    break
                            if matched_doc:
                                if matched_doc not in as_schedule[day]:
                                    as_schedule[day][matched_doc] = []
                                as_schedule[day][matched_doc].append(code)

                as_result_grid = {}
                for full_name in as_doc_mapping:
                    as_result_grid[full_name] = {}
                    for day in range(1, 32):
                        codes = as_schedule[day].get(full_name, [])
                        if not codes: final_code = ""
                        elif 'L-6' in codes: final_code = 'L-6'
                        elif 'L-7' in codes: final_code = 'L-7'
                        else: final_code = codes[0]
                        as_result_grid[full_name][day] = final_code

                for r in range(9, 17):
                    full_name = str(sh_hekim_as.cell(row=r, column=2).value).strip()
                    if full_name in as_result_grid:
                        for day in range(1, 32):
                            col_idx = day + 4
                            val = as_result_grid[full_name][day]
                            sh_hekim_as.cell(row=r, column=col_idx).value = val if val else None

                title_val = sh_hekim_as.cell(row=1, column=5).value
                if title_val and 'EYLÜL' in str(title_val):
                    sh_hekim_as.cell(row=1, column=5).value = str(title_val).replace('EYLÜL', 'EKİM')

                output_as = BytesIO()
                wb.save(output_as)
                output_as.seek(0)

                grid_rows = []
                for doc_n, days_dict in as_result_grid.items():
                    row_data = {"Hekim Adı Soyadı": doc_n}
                    for d in range(1, 32):
                        row_data[str(d)] = days_dict[d] if days_dict[d] else "-"
                    grid_rows.append(row_data)

                df_as_preview = pd.DataFrame(grid_rows)

                st.success("✅ Asistan Doktor çalışma listesi başarıyla L-kodlarına dönüştürüldü ve kodlandı!")
                
                st.subheader("📊 Kodlanmış Asistan Doktor Çizelgesi Önizleme")
                st.dataframe(df_as_preview, use_container_width=True)

                st.markdown("---")
                st.subheader("📥 Kodlanmış Excel Dosyasını İndir")
                b64_as = base64.b64encode(output_as.getvalue()).decode()
                as_file_name = "As.Dr._Calisma_Listesi_10.2026_Kodlanmis.xlsx"
                href_as = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_as}" download="{as_file_name}" class="direct-download-btn">📥 Kodlanmış Asistan Dr. Excel Dosyasını İndir (.xlsx)</a>'
                st.markdown(href_as, unsafe_allow_html=True)

        except Exception as e:
            st.error(f"❌ Dosya işlenirken bir hata oluştu: {e}")


# ==========================================
# 4. MODÜL: EĞİTİCİ DESTEKLEME PUAN ÇİZELGESİ
# ==========================================
elif secilen_modul == "4. Eğitici Destekleme Puan Çizelgesi":
    st.markdown(
        """
    <div class="header-box">
        <h1>🎓 Eğitici Destekleme Puan Çizelgesi Modülü</h1>
        <p>Haftalara Dengeli Teorik Eğitim (8s) & İzin Durumunda Esnek Gün Dağıtımı & Resmi Şablon Excel Çıktısı</p>
        <div class="header-imza">✍️ Özgür SARSILMAZ</div>
    </div>
    """,
        unsafe_allow_html=True,
    )

    st.sidebar.subheader("📅 Tarih ve Dönem Seçimi")
    col_egit_yil, col_egit_ay = st.sidebar.columns(2)
    with col_egit_yil:
        egit_yil = st.number_input("Yıl", value=2026, min_value=2024, max_value=2030, key="egit_yil_input")
    with col_egit_ay:
        egit_ay = st.selectbox("Ay", list(range(1, 13)), index=8, format_func=lambda x: f"{x}. Ay ({calendar.month_name[x]})", key="egit_ay_select")

    _, egit_gun_sayisi = calendar.monthrange(egit_yil, egit_ay)
    egit_tum_tarihler = [datetime.date(egit_yil, egit_ay, d) for d in range(1, egit_gun_sayisi + 1)]
    egit_mesai_gunleri = [d for d in egit_tum_tarihler if d.weekday() < 5]

    st.subheader("👨‍⚕️ Eğitici Hekim Kadrosu ve İzin Girişi")
    
    default_egiticiler = [
        "Prof.Dr. Yeşim, ÇEKİN",
        "Prof.Dr. Hatice Nevgün, ÖZEN",
        "Doç.Dr. Cemile Aylin, ERMAN DALOĞLU",
        "Doç.DR. Halil, ER",
        "Başasistan Doç.Dr. Özlem, KOCA"
    ]
    
    egitici_input = st.text_area(
        "Eğitici Hekim Listesi (Virgül ile Unvan/Ad ve Soyadı ayırabilirsiniz):",
        value="\n".join(default_egiticiler),
        height=140,
        help="Çift soyadlı hekimlerde soyadı doğru ayrıştırmak için 'Unvan Ad, SOYAD' şeklinde virgül koyabilirsiniz.",
        key="egitici_list_input"
    )
    egitici_listesi = [h.strip() for h in egitici_input.split("\n") if h.strip()]

    st.markdown("---")
    st.markdown('<div class="section-title">🏖️ Hekim Bazlı İzinli Gün Seçimi</div>', unsafe_allow_html=True)
    st.info("💡 Hekimlerin izinli/görevli olduğu günleri seçiniz. Teorik dersler haftalara eşit dağıtılacak ve çıktı orijinal Excel şablonunuzla birebir aynı formatta üretilecektir.")

    egitici_izinler = {}
    for idx, egitici in enumerate(egitici_listesi):
        with st.expander(f"🔴 {egitici} - İzin Günleri Tanımla", expanded=False):
            selected_leaves = st.multiselect(
                f"{egitici} için İzinli / Görevli Olunan Günler:",
                options=egit_mesai_gunleri,
                format_func=lambda d: f"{d.strftime('%d.%m.%Y')} ({tr_gunler[d.weekday()]})",
                key=f"egit_leave_{idx}_{egitici}"
            )
            egitici_izinler[egitici] = set(selected_leaves)

    st.markdown("<br>", unsafe_allow_html=True)

    def generate_egitici_excel(egit_yil, egit_ay, egit_gun_sayisi, egitici_results):
        wb = openpyxl.Workbook()
        default_sheet = wb.active

        ay_adlari_tr = {
            1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan", 5: "Mayıs", 6: "Haziran",
            7: "Temmuz", 8: "Ağustos", 9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık"
        }
        ilgili_ay_str = f"{ay_adlari_tr[egit_ay]} {egit_yil}"

        font_header = Font(name="Calibri", size=11, bold=False)
        font_title = Font(name="Calibri", size=11, bold=True)
        font_red = Font(name="Calibri", size=11, bold=True, color="FF0000")
        font_table_hdr = Font(name="Calibri", size=11, bold=True)
        font_date = Font(name="Times New Roman", size=11)
        font_val = Font(name="Calibri", size=11)
        font_footer = Font(name="Calibri", size=8.5, italic=False)

        fill_weekend = PatternFill(start_color="D4D4D4", end_color="D4D4D4", fill_type="solid")
        align_center = Alignment(horizontal="center", vertical="center")
        align_justify = Alignment(horizontal="justify", vertical="center", wrap_text=True)

        thin_side = Side(style='thin', color='000000')
        thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        for egitici_ad, df_hekim in egitici_results.items():
            clean_name = egitici_ad.replace(",", "").replace("-", "")
            parts = clean_name.split()
            sheet_title = parts[1] if len(parts) > 1 else clean_name[:10]
            sheet_title = sheet_title.replace(".", "").replace("/", "")[:12]

            ws = wb.create_sheet(title=sheet_title)
            ws.views.sheetView[0].showGridLines = True

            ws.merge_cells("A1:C1")
            ws.merge_cells("A2:C2")
            ws.merge_cells("A3:C3")
            ws.merge_cells("A4:C4")
            
            ws.cell(row=1, column=1, value="T.C.").font = font_header
            ws.cell(row=2, column=1, value="ANTALYA VALİLİĞİ").font = font_header
            ws.cell(row=3, column=1, value="İl Sağlık Müdürlüğü").font = font_header
            ws.cell(row=4, column=1, value="Sağlık Bilimleri Üniversitesi Antalya Eğitim ve Araştırma Hastanesi").font = font_header

            for r in range(1, 5):
                ws.cell(row=r, column=1).alignment = align_center

            ws.merge_cells("A6:C6")
            c_title = ws.cell(row=6, column=1, value="EĞİTİCİ DESTEKLEME PUAN ÇİZELGESİ")
            c_title.font = font_title
            c_title.alignment = align_center

            if "," in egitici_ad:
                p_tokens = egitici_ad.split(",")
                unvan_ad = p_tokens[0].strip()
                soyad = p_tokens[1].strip()
            elif "-" in egitici_ad:
                p_tokens = egitici_ad.split("-")
                unvan_ad = p_tokens[0].strip()
                soyad = p_tokens[1].strip()
            elif "ERMAN DALOĞLU" in egitici_ad.upper():
                soyad = "DALOĞLU"
                unvan_ad = egitici_ad.upper().replace("DALOĞLU", "").strip()
            else:
                p_tokens = egitici_ad.split()
                soyad = p_tokens[-1] if len(p_tokens) > 1 else ""
                unvan_ad = " ".join(p_tokens[:-1]) if len(p_tokens) > 1 else egitici_ad

            ws.cell(row=7, column=1, value="ADI :").font = font_header
            ws.cell(row=7, column=2, value=unvan_ad).font = font_header

            ws.cell(row=8, column=1, value="SOYADI :").font = font_header
            ws.cell(row=8, column=2, value=soyad).font = font_header

            ws.cell(row=9, column=1, value="KLİNİĞİ :").font = font_header
            ws.cell(row=9, column=2, value="Tıbbi Mikrobiyoloji").font = font_header

            ws.cell(row=10, column=1, value="İLGİLİ AY :").font = font_header
            c_ay = ws.cell(row=10, column=2, value=ilgili_ay_str)
            c_ay.font = font_red

            ws.cell(row=13, column=1, value="* LABORATUVAR KLİNİKLERİ").font = Font(name="Calibri", size=11, bold=True, underline="single")

            h1 = ws.cell(row=14, column=1, value="GÜN")
            h2 = ws.cell(row=14, column=2, value="PRATİK EĞİTİM ÇALIŞMASI")
            h3 = ws.cell(row=14, column=3, value="TEORİK EĞİTİM")

            for h_cell in [h1, h2, h3]:
                h_cell.font = font_table_hdr
                h_cell.alignment = align_center
                h_cell.border = thin_border

            current_row = 15
            for _, r_data in df_hekim.iterrows():
                dt_val = r_data["Tarih_Obj"]
                pratik_v = r_data["Pratik"]
                teorik_v = r_data["Teorik"]
                is_weekend = dt_val.weekday() >= 5

                ws.row_dimensions[current_row].height = 19

                c_day = ws.cell(row=current_row, column=1, value=dt_val)
                c_day.number_format = 'd.mm.yyyy'
                c_day.font = font_date
                c_day.alignment = align_center
                c_day.border = thin_border

                c_pratik = ws.cell(row=current_row, column=2, value=pratik_v if pratik_v > 0 else None)
                c_pratik.font = font_val
                c_pratik.alignment = align_center
                c_pratik.border = thin_border

                c_teorik = ws.cell(row=current_row, column=3, value=teorik_v if teorik_v > 0 else None)
                c_teorik.font = font_val
                c_teorik.alignment = align_center
                c_teorik.border = thin_border

                if is_weekend:
                    c_day.fill = fill_weekend
                    c_pratik.fill = fill_weekend
                    c_teorik.fill = fill_weekend

                current_row += 1

            footer_row = 46
            ws.row_dimensions[footer_row].height = 30
            ws.merge_cells(start_row=footer_row, start_column=1, end_row=footer_row, end_column=3)
            
            c_ft = ws.cell(
                row=footer_row,
                column=1,
                value="*Laboratuvar klinikleri  için 40 saat pratik eğitim çalışması ve 8 saat teorik asistan eğitim çalışması yapıldığının belgelendirmesi halinde eğitici destekleme puanı verilir,"
            )
            c_ft.font = font_footer
            c_ft.alignment = align_justify

            ws.column_dimensions["A"].width = 18
            ws.column_dimensions["B"].width = 30
            ws.column_dimensions["C"].width = 25

        wb.remove(default_sheet)
        output_e = BytesIO()
        wb.save(output_e)
        output_e.seek(0)
        return output_e

    if st.button("🚀 Eğitici Destekleme Çizelgesini Oluştur ve Dağıt"):
        egitici_results = {}
        hesaplama_hatalari = []

        for egitici in egitici_listesi:
            df_h = pd.DataFrame({"Tarih_Obj": egit_tum_tarihler})
            df_h["Tarih"] = df_h["Tarih_Obj"].apply(lambda d: d.strftime("%d.%m.%Y"))
            df_h["Gün Adı"] = df_h["Tarih_Obj"].apply(lambda d: tr_gunler[d.weekday()])
            df_h["Pratik"] = 0
            df_h["Teorik"] = 0
            egitici_results[egitici] = df_h

        haftalar = {}
        for d in egit_mesai_gunleri:
            w_num = d.isocalendar()[1]
            if w_num not in haftalar:
                haftalar[w_num] = []
            haftalar[w_num].append(d)

        gunluk_teorik_sayisi = {d: 0 for d in egit_mesai_gunleri}
        hekim_haftalik_teorik = {e: {w: 0 for w in haftalar} for e in egitici_listesi}
        hekim_toplam_teorik = {e: 0 for e in egitici_listesi}

        for w_num, w_gunleri in haftalar.items():
            for egitici in egitici_listesi:
                if hekim_toplam_teorik[egitici] >= 8:
                    continue
                if hekim_haftalik_teorik[egitici][w_num] >= 2:
                    continue

                atanan_gun = None
                for d in w_gunleri:
                    if gunluk_teorik_sayisi[d] == 0 and d not in egitici_izinler.get(egitici, set()):
                        atanan_gun = d
                        break

                if atanan_gun is None:
                    musait_gunler = [d for d in w_gunleri if d not in egitici_izinler.get(egitici, set())]
                    if musait_gunler:
                        atanan_gun = min(musait_gunler, key=lambda d: gunluk_teorik_sayisi[d])

                if atanan_gun is not None:
                    df_target = egitici_results[egitici]
                    df_target.loc[df_target["Tarih_Obj"] == atanan_gun, "Teorik"] = 2
                    gunluk_teorik_sayisi[atanan_gun] += 1
                    hekim_haftalik_teorik[egitici][w_num] += 2
                    hekim_toplam_teorik[egitici] += 2

        for egitici, t_saat in hekim_toplam_teorik.items():
            if t_saat < 8:
                kalan = 8 - t_saat
                aktif_g = [d for d in egit_mesai_gunleri if d not in egitici_izinler.get(egitici, set())]
                for d in aktif_g:
                    df_target = egitici_results[egitici]
                    if df_target.loc[df_target["Tarih_Obj"] == d, "Teorik"].values[0] == 0:
                        df_target.loc[df_target["Tarih_Obj"] == d, "Teorik"] = 2
                        gunluk_teorik_sayisi[d] += 1
                        kalan -= 2
                        if kalan <= 0:
                            break

        for egitici in egitici_listesi:
            izinli_gunler = egitici_izinler.get(egitici, set())
            aktif_mesai_gunleri = [d for d in egit_mesai_gunleri if d not in izinli_gunler]

            if len(aktif_mesai_gunleri) < 4:
                hesaplama_hatalari.append(f"❌ **{egitici}**: İzinler sebebiyle aktif gün sayısı yetersiz ({len(aktif_mesai_gunleri)} gün).")
                continue

            num_aktif = len(aktif_mesai_gunleri)
            base_p = 40 // num_aktif
            extra_p = 40 % num_aktif

            df_target = egitici_results[egitici]
            for idx, d in enumerate(aktif_mesai_gunleri):
                assigned_pratik = base_p + (1 if idx < extra_p else 0)
                df_target.loc[df_target["Tarih_Obj"] == d, "Pratik"] = assigned_pratik

        if hesaplama_hatalari:
            for err in hesaplama_hatalari:
                st.error(err)
        else:
            excel_egitici_bytes = generate_egitici_excel(egit_yil, egit_ay, egit_gun_sayisi, egitici_results)
            
            st.session_state.egitici_sonuc = {
                "results": egitici_results,
                "excel_bytes": excel_egitici_bytes,
                "yil": egit_yil,
                "ay": egit_ay
            }
            st.balloons()
            st.success("✨ Tüm eğiticiler için resmi Excel şablonu ile tam uyumlu 40s Pratik / 8s Teorik eğitim çizelgeleri oluşturuldu!")

    if "egitici_sonuc" in st.session_state and st.session_state.egitici_sonuc is not None:
        e_res = st.session_state.egitici_sonuc
        st.markdown("---")
        st.subheader("📊 Eğitici Dağılım Sonuçları Önizleme")

        res_tabs = st.tabs(list(e_res["results"].keys()))
        for idx, (eg_name, df_show) in enumerate(e_res["results"].items()):
            with res_tabs[idx]:
                p_sum = df_show["Pratik"].sum()
                t_sum = df_show["Teorik"].sum()

                m1, m2, m3 = st.columns(3)
                m1.metric("Pratik Eğitim Toplamı", f"{p_sum} / 40 Saat", delta=int(p_sum - 40))
                m2.metric("Teorik Eğitim Toplamı", f"{t_sum} / 8 Saat", delta=int(t_sum - 8))
                m3.metric("Aktif Çalışılan Gün", f"{len(df_show[(df_show['Pratik']>0) | (df_show['Teorik']>0)])} Gün")

                st.dataframe(
                    df_show[["Tarih", "Gün Adı", "Pratik", "Teorik"]],
                    use_container_width=True,
                    height=380
                )

        st.markdown("---")
        st.subheader("📥 Resmi Excel Çizelgesini İndir")
        b64_egit = base64.b64encode(e_res["excel_bytes"].getvalue()).decode()
        egit_file_name = f"Egitici_Destekleme_Puan_Cizelgesi_{e_res['ay']}_{e_res['yil']}.xlsx"
        href_egit = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_egit}" download="{egit_file_name}" class="direct-download-btn">📥 Eğitici Destekleme Resmi Excel Dosyasını İndir (.xlsx)</a>'
        st.markdown(href_egit, unsafe_allow_html=True)
