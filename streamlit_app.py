import base64
import calendar
import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from ortools.sat.python import cp_model
import pandas as pd
import streamlit as st

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
    [data-testid="stHeader"], .stAppHeader { display: none !important; }
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

    st.sidebar.subheader("📅 Tarih ve Birim Seçimi")
    col_yil, col_ay = st.sidebar.columns(2)

    with col_yil:
        yil = st.number_input("Yıl", value=2026, min_value=2024, max_value=2030)
    with col_ay:
        ay = st.selectbox("Ay", list(range(1, 13)), index=9)

    _, gun_sayisi = calendar.monthrange(yil, ay)
    gun_secenekleri = list(range(1, gun_sayisi + 1))

    st.sidebar.markdown("---")
    st.sidebar.subheader("🏖 Resmi & İdari Tatil Günleri")

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

    st.sidebar.markdown("---")
    st.sidebar.subheader("🧫 Kültür 8s Gündüz Vardiyası")
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
    st.sidebar.subheader("🛡 Genel Kural Kurulumu")
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
        "📌 Cuma / Cmts / Pzr Dengeli Dağılım (Maks 1 Gün)", value=True
    )
    esnek_personel = st.sidebar.multiselect(
        "🔓 Özel Esneklik Tanınacak Personel(ler):",
        options=nobetci_personeller,
        default=[],
    )

    st.sidebar.markdown("---")
    st.sidebar.subheader("📊 Geçmiş Ay Rotasyonu (Önceki Ay Dosyası)")
    uploaded_file = st.sidebar.file_uploader(
        "Önceki Ayın Excel Dosyası:",
        type=["xlsx", "xls"],
        help="Sistemin ürettiği 3 sekmeli Excel dosyasını yükleyin.",
    )

    gecmis_istatistik = {}
    gecmis_acil_istatistik = {}
    gecmis_kultur8_istatistik = {}
    otomatik_son_gun_normal = []
    otomatik_son_gun_acil = []

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

            st.sidebar.success(f"✅ {len(gecmis_istatistik)} personelin devir verileri aktarıldı!")
        except Exception as e:
            st.sidebar.error(f"❌ Hata: Yüklenen Excel okunurken sorun oluştu ({e}).")

    gecmis_ay_son_gun_normal = st.sidebar.multiselect(
        "🌙 Geçmiş Ay Son Günü NORMAL Nöbetçileri:",
        options=nobetci_personeller,
        default=list(set(otomatik_son_gun_normal)),
        help="Önceki ayın son günü normal nöbet tutan personeller.",
    )

    gecmis_ay_son_gun_acil = st.sidebar.multiselect(
        "🚨 Geçmiş Ay Son Günü ACİL Nöbetçileri:",
        options=[p for p in nobetci_personeller if p not in gecmis_ay_son_gun_normal],
        default=list(set(otomatik_son_gun_acil)),
        help="Önceki ayın son günü acil nöbet tutan personeller.",
    )

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

    # Mazeret & Sabit Nöbet
    st.markdown('<div class="section-title">📋 Personel Mazeret ve Sabit Nöbet Girişleri</div>', unsafe_allow_html=True)
    if "mazeret_satir_sayisi" not in st.session_state:
        st.session_state.mazeret_satir_sayisi = 1

    def mazeret_satir_ekle(): st.session_state.mazeret_satir_sayisi += 1
    def mazeret_satir_cikar():
        if st.session_state.mazeret_satir_sayisi > 1: st.session_state.mazeret_satir_sayisi -= 1

    izinler = {p: [] for p in nobetci_personeller}
    sabit_nobetler = {p: [] for p in nobetci_personeller}
    toplam_izin_sayisi, toplam_sabit_sayisi = 0, 0

    for m_idx in range(st.session_state.mazeret_satir_sayisi):
        c1, c2, c3 = st.columns([1.2, 1.8, 1.8])
        with c1: p_secilen = st.selectbox(f"Personel #{m_idx+1}:", options=["Seçiniz..."] + nobetci_personeller, key=f"m_personel_{m_idx}")
        with c2: selected_days = st.multiselect(f"Mazeret / İzin Günleri #{m_idx+1}:", options=gun_secenekleri, default=[], key=f"m_leave_{m_idx}")
        with c3: selected_sabit_days = st.multiselect(f"Sabit Nöbet Günleri #{m_idx+1}:", options=gun_secenekleri, default=[], key=f"m_forced_{m_idx}")

        if p_secilen != "Seçiniz...":
            izinler[p_secilen].extend([d - 1 for d in selected_days])
            sabit_nobetler[p_secilen].extend([d - 1 for d in selected_sabit_days])
            toplam_izin_sayisi += len(selected_days)
            toplam_sabit_sayisi += len(selected_sabit_days)

    col_m_btn1, col_m_btn2, _ = st.columns([1.2, 1.2, 3.6])
    with col_m_btn1: st.button("➕ Mazeret / Sabit Nöbet Ekle", on_click=mazeret_satir_ekle)
    with col_m_btn2: st.button("➖ Mazeret Satırı Sil", on_click=mazeret_satir_cikar)

    st.markdown("<br>", unsafe_allow_html=True)

    # Acil Nöbet Girişleri
    st.markdown('<div class="section-title">🚨 Acil Nöbetçi Girişleri (Opsiyonel)</div>', unsafe_allow_html=True)
    if "acil_satir_sayisi" not in st.session_state: st.session_state.acil_satir_sayisi = 1
    def acil_satir_ekle(): st.session_state.acil_satir_sayisi += 1
    def acil_satir_cikar():
        if st.session_state.acil_satir_sayisi > 1: st.session_state.acil_satir_sayisi -= 1

    sabit_acil_nobetler = {p: [] for p in nobetci_personeller}
    toplam_sabit_acil_sayisi = 0

    for a_idx in range(st.session_state.acil_satir_sayisi):
        c1, c2 = st.columns([1.5, 3.3])
        with c1: p_acil_secilen = st.selectbox(f"Acil Nöbetçi #{a_idx+1}:", options=["Seçiniz..."] + nobetci_personeller, key=f"a_personel_{a_idx}")
        with c2: selected_acil_days = st.multiselect(f"Sabit Acil Nöbet Günleri #{a_idx+1}:", options=gun_secenekleri, default=[], key=f"a_forced_{a_idx}")

        if p_acil_secilen != "Seçiniz...":
            sabit_acil_nobetler[p_acil_secilen].extend(selected_acil_days)
            toplam_sabit_acil_sayisi += len(selected_acil_days)

    col_a_btn1, col_a_btn2, _ = st.columns([1.2, 1.2, 3.6])
    with col_a_btn1: st.button("➕ Sabit Acil Nöbet Ekle", on_click=acil_satir_ekle)
    with col_a_btn2: st.button("➖ Acil Satırı Sil", on_click=acil_satir_cikar)

    st.markdown("<br>", unsafe_allow_html=True)

    # Kişiler Arası Kısıtlar
    st.markdown('<div class="section-title">🤝 Kişiler Arası Nöbet Mesafe ve Çakışma Yasağı Kuralları</div>', unsafe_allow_html=True)
    if "kisi_kisit_sayisi" not in st.session_state: st.session_state.kisi_kisit_sayisi = 1
    def kisit_ekle(): st.session_state.kisi_kisit_sayisi += 1
    def kisit_cikar():
        if st.session_state.kisi_kisit_sayisi > 1: st.session_state.kisi_kisit_sayisi -= 1

    kisi_kisitlari = []
    for k_idx in range(st.session_state.kisi_kisit_sayisi):
        c1, c2, c3 = st.columns([1.2, 1, 2])
        with c1: p_ana = st.selectbox(f"Ana Personel #{k_idx+1}:", options=["Seçiniz..."] + nobetci_personeller, key=f"p_ana_{k_idx}")
        with c2: min_aralik = st.number_input(f"Min. Mesafe (Gün) #{k_idx+1}:", min_value=0, max_value=15, value=1, key=f"min_aralik_{k_idx}")
        with c3: p_yasakli_list = st.multiselect(f"Birlikte/Yakın Nöbet Tutamayacağı Kişiler #{k_idx+1}:", options=[p for p in nobetci_personeller if p != p_ana], key=f"p_yasakli_{k_idx}")
        if p_ana != "Seçiniz..." and p_yasakli_list:
            kisi_kisitlari.append({"ana": p_ana, "aralik": min_aralik, "yasaklilar": p_yasakli_list})

    col_btn1, col_btn2, _ = st.columns([1, 1, 4])
    with col_btn1: st.button("➕ Yeni Kısıt Ekle", on_click=kisit_ekle)
    with col_btn2: st.button("➖ Kısıt Sil", on_click=kisit_cikar)

    st.markdown("<br>", unsafe_allow_html=True)
    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("👥 Nöbetçi Kadro", f"{len(nobetci_personeller)} Kişi")
    m2.metric("📅 Ayın Gün Sayısı", f"{gun_sayisi} Gün")
    m3.metric("🎯 Nöbet Slotu", f"{gun_sayisi * gunluk_nobetci * (3 if birim_secimi=='Tüm Laboratuvar (Birleşik)' else 1)} Nöbet")
    m4.metric("🏖️ Kayıtlı İzinler", f"{toplam_izin_sayisi} Gün")
    m5.metric("📌 Sabit Nöbetler", f"{toplam_sabit_sayisi} Gün")
    m6.metric("🚨 Sabit Acil Nöbet", f"{toplam_sabit_acil_sayisi} Gün")
    st.markdown("<br>", unsafe_allow_html=True)

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

    def calculate_personel_puantaj_metrikleri(p_name, p_row_dict, yil, ay, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri, acil_days_set, kultur_8s_tatil_gunleri, gecmis_ay_son_gun_normal, gecmis_ay_son_gun_acil):
        aylik_hedef_saat = calculate_aylik_calisma_saati(yil, ay, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri)

        toplam_calisma = 0
        norm_gece = 0
        norm_normal = 0
        risk_gece = 0
        risk_normal = 0

        # 1. Ham Nöbet Saatlerinin Dağıtılması
        for d in range(1, gun_sayisi + 1):
            val = str(p_row_dict.get(str(d), "")).strip()

            if val in ["8", "5", "16", "19", "11", "24"]:
                toplam_calisma += int(val)

            dt = datetime.date(yil, ay, d)
            weekday = dt.weekday()
            is_off = is_day_off(yil, ay, d, resmi_tatil_gunleri)
            is_next_day_off = is_day_off(yil, ay, d + 1, resmi_tatil_gunleri) if d < gun_sayisi else False
            is_risk = d in acil_days_set

            if val == "24":
                if d == 5 and d in yarim_gun_tatil_gunleri:
                    g_saat, n_saat = 8, 3
                elif d in yarim_gun_tatil_gunleri:
                    g_saat, n_saat = 12, 7
                elif d == gun_sayisi and not is_next_day_off:
                    g_saat, n_saat = 12, 4
                elif is_off:
                    if is_next_day_off:
                        g_saat, n_saat = 12, 12
                    else:
                        g_saat, n_saat = 12, 4
                elif weekday in [0, 1, 2, 3]:
                    g_saat, n_saat = 8, 0
                elif weekday in [4, 6]:
                    g_saat, n_saat = 12, 4
                elif weekday == 5:
                    g_saat, n_saat = 12, 12
                else:
                    g_saat, n_saat = 12, 4

                if is_risk:
                    risk_gece += g_saat
                    risk_normal += n_saat
                else:
                    norm_gece += g_saat
                    norm_normal += n_saat

            elif val == "8" and d in kultur_8s_tatil_gunleri:
                norm_normal += 8

        # 2. Akıllı Devir Nİ Mahsuplaşması
        is_devir_normal = p_name in gecmis_ay_son_gun_normal
        is_devir_acil = p_name in gecmis_ay_son_gun_acil

        if is_devir_normal or is_devir_acil:
            dusulecek_saat = 8

            if is_devir_acil:
                dusulen_norm_gece = min(4, risk_gece)
                risk_gece -= dusulen_norm_gece
                dusulecek_saat -= dusulen_norm_gece

                dusulen_norm_normal = min(4, risk_normal)
                risk_normal -= dusulen_norm_normal
                dusulecek_saat -= dusulen_norm_normal

                if dusulecek_saat > 0 and risk_gece > 0:
                    ek_dusu = min(dusulecek_saat, risk_gece)
                    risk_gece -= ek_dusu
                    dusulecek_saat -= ek_dusu

                if dusulecek_saat > 0:
                    d_gece = min(dusulecek_saat, norm_gece)
                    norm_gece -= d_gece
                    dusulecek_saat -= d_gece

                if dusulecek_saat > 0:
                    d_norm = min(dusulecek_saat, norm_normal)
                    norm_normal -= d_norm
                    dusulecek_saat -= d_norm

            else:
                dusulen_norm_gece = min(4, norm_gece)
                norm_gece -= dusulen_norm_gece
                dusulecek_saat -= dusulen_norm_gece

                dusulen_norm_normal = min(4, norm_normal)
                norm_normal -= dusulen_norm_normal
                dusulecek_saat -= dusulen_norm_normal

                if dusulecek_saat > 0 and norm_gece > 0:
                    ek_dusu = min(dusulecek_saat, norm_gece)
                    norm_gece -= ek_dusu
                    dusulecek_saat -= ek_dusu

                if dusulecek_saat > 0:
                    d_gece = min(dusulecek_saat, risk_gece)
                    risk_gece -= d_gece
                    dusulecek_saat -= d_gece

                if dusulecek_saat > 0:
                    d_norm = min(dusulecek_saat, risk_normal)
                    risk_normal -= d_norm
                    dusulecek_saat -= d_norm

        # Fazla Nöbet Saati = Gerçek Nöbet Hakedişlerinin Toplamı
        fazla_nobet = norm_gece + norm_normal + risk_gece + risk_normal

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
            bu_ay_saat = bu_ay_toplam_nobet * 24
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

            bu_ay_acil_saat = len(acil_nobet_dict.get(p, set())) * 24
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
                    if day == 1 and p in gecmis_ay_son_gun_nobetcileri:
                        cell.value = "T" if day in off_days else "Nİ"
                        if cell.value == "Nİ": cell.font = font_ni
                    elif day in p_shifts:
                        cell.value, cell.font = 24, font_24
                        if day in p_acils: cell.fill = fill_bright_yellow
                    elif day in p_k8_shifts: cell.value, cell.font = 8, font_bold
                    elif (day - 1) in p_shifts:
                        cell.value = "T" if day in off_days else "Nİ"
                        if cell.value == "Nİ": cell.font = font_ni
                    elif dt.weekday() == 0 and (d - 2) in p_k8_shifts and datetime.date(yil, ay, d - 2).weekday() == 5:
                        cell.value, cell.font = "Nİ", font_ni
                    elif (day - 1) in p_k8_shifts and datetime.date(yil, ay, d - 1).weekday() == 5: cell.value = "T"
                    else:
                        if day in off_days: cell.value = "T"
                        elif day in yarim_gun_tatil_gunleri: cell.value, cell.font = 5, font_body
                        else: cell.value, cell.font = 8, font_body

                p_row_dict[str(day)] = str(cell.value)

            m = calculate_personel_puantaj_metrikleri(p, p_row_dict, yil, ay, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri, p_acils, kultur_8s_tatil_gunleri, gecmis_ay_son_gun_normal, gecmis_ay_son_gun_acil)
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
            cell = ws3.cell(row=summary_r, column=col_idx, value="")
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

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    # HESAPLAMA BUTONU
    if st.button("🚀 Otomatik ve Adil Nöbet Listesini Oluştur"):
        hata_listesi = []
        pcr_nobetcileri = [p for p in nobetci_personeller if TUM_PERSONEL_VERISI.get(p, {}).get("birim") == "PCR"]
        mikro_nobetcileri = [p for p in nobetci_personeller if TUM_PERSONEL_VERISI.get(p, {}).get("birim") == "Mikro"]
        kultur_nobetcileri = [p for p in nobetci_personeller if TUM_PERSONEL_VERISI.get(p, {}).get("birim") == "Kültür"]

        for d in range(gun_sayisi):
            if birim_secimi == "Tüm Laboratuvar (Birleşik)":
                m_pcr = sum(1 for p in pcr_nobetcileri if d not in izinler[p])
                m_mikro = sum(1 for p in mikro_nobetcileri if d not in izinler[p])
                m_kultur = sum(1 for p in kultur_nobetcileri if d not in izinler[p])
                if m_pcr < gunluk_nobetci or m_mikro < gunluk_nobetci or m_kultur < gunluk_nobetci:
                    tarih_str = datetime.date(yil, ay, d + 1).strftime("%d.%m.%Y")
                    hata_listesi.append(f"⚠️ **{tarih_str}** ({d+1}. gün): İzinler nedeniyle en az bir birimde yeterli nöbetçi yok!")
            else:
                musait_sayisi = sum(1 for p in nobetci_personeller if d not in izinler[p])
                if musait_sayisi < gunluk_nobetci:
                    tarih_str = datetime.date(yil, ay, d + 1).strftime("%d.%m.%Y")
                    hata_listesi.append(f"⚠️ **{tarih_str}** ({d+1}. gün): En az {gunluk_nobetci} kişi gerekli ancak sadece {musait_sayisi} müsait.")

        for p in nobetci_personeller:
            ortak = set(izinler[p]).intersection(set(sabit_nobetler[p]))
            for d in ortak:
                tarih_str = datetime.date(yil, ay, d + 1).strftime("%d.%m.%Y")
                hata_listesi.append(f"❌ **{p}**, **{tarih_str}** tarihinde hem 'İzinli' hem 'Sabit Nöbetçi'!")

        if hata_listesi:
            for err in hata_listesi: st.error(err)
        else:
            model = cp_model.CpModel()
            x, k8 = {}, {}

            for p in nobetci_personeller:
                for d in range(gun_sayisi): x[(p, d)] = model.NewBoolVar(f"x_{p}_{d}")
            for p in kultur_nobetcileri:
                for d in range(gun_sayisi): k8[(p, d)] = model.NewBoolVar(f"k8_{p}_{d}")

            if birim_secimi == "Tüm Laboratuvar (Birleşik)":
                for d in range(gun_sayisi):
                    model.Add(sum(x[(p, d)] for p in pcr_nobetcileri) == gunluk_nobetci)
                    model.Add(sum(x[(p, d)] for p in mikro_nobetcileri) == gunluk_nobetci)
                    model.Add(sum(x[(p, d)] for p in kultur_nobetcileri) == gunluk_nobetci)
            else:
                for d in range(gun_sayisi):
                    model.Add(sum(x[(p, d)] for p in nobetci_personeller) == gunluk_nobetci)

            kultur_8s_gun_indeksleri = [d for d in range(gun_sayisi) if datetime.date(yil, ay, d + 1).weekday() == 5 or (d + 1) in kultur_8s_tatil_gunleri]
            for d in range(gun_sayisi):
                if d in kultur_8s_gun_indeksleri and kultur_nobetcileri:
                    model.Add(sum(k8[(p, d)] for p in kultur_nobetcileri) == 1)
                else:
                    for p in kultur_nobetcileri: model.Add(k8[(p, d)] == 0)

            for p in nobetci_personeller:
                for d in izinler[p]:
                    model.Add(x[(p, d)] == 0)
                    if p in kultur_nobetcileri: model.Add(k8[(p, d)] == 0)

            for p in nobetci_personeller:
                for d in sabit_nobetler[p]: model.Add(x[(p, d)] == 1)

            for p in gecmis_ay_son_gun_nobetcileri:
                if p in nobetci_personeller:
                    model.Add(x[(p, 0)] == 0)
                    if p in kultur_nobetcileri: model.Add(k8[(p, 0)] == 0)

            for p in nobetci_personeller:
                p_dinlenme = 1 if p in esnek_personel else max(1, int(dinlenme_gun_sayisi))
                for d in range(gun_sayisi):
                    if p in kultur_nobetcileri:
                        model.Add(x[(p, d)] + k8[(p, d)] <= 1)
                        if d > 0: model.Add(x[(p, d - 1)] + k8[(p, d)] <= 1)
                        if d < gun_sayisi - 1: model.Add(k8[(p, d)] + x[(p, d + 1)] <= 1)

                    for k in range(1, p_dinlenme + 1):
                        if d + k < gun_sayisi:
                            model.Add(x[(p, d)] + x[(p, d + k)] <= 1)
                            if p in kultur_nobetcileri:
                                model.Add(x[(p, d)] + k8[(p, d + k)] <= 1)
                                model.Add(k8[(p, d)] + x[(p, d + k)] <= 1)

            if persembe_pazar_yasagi:
                for d in range(gun_sayisi - 3):
                    if datetime.date(yil, ay, d + 1).weekday() == 3:
                        for p in nobetci_personeller:
                            if p not in esnek_personel:
                                model.Add(x.get((p, d), 0) + x.get((p, d + 3), 0) <= 1)

            for rule in kisi_kisitlari:
                p1, aralik = rule["ana"], rule["aralik"]
                for p2 in rule["yasaklilar"]:
                    for d1 in range(gun_sayisi):
                        for d2 in range(max(0, d1 - aralik), min(gun_sayisi, d1 + aralik + 1)):
                            p1_gorev = x[(p1, d1)] + (k8[(p1, d1)] if p1 in kultur_nobetcileri else 0)
                            p2_gorev = x[(p2, d2)] + (k8[(p2, d2)] if p2 in kultur_nobetcileri else 0)
                            model.Add(p1_gorev + p2_gorev <= 1)

            def gun_kategorisi_indeksleri(w_list):
                return [d for d in range(gun_sayisi) if datetime.date(yil, ay, d + 1).weekday() in w_list]

            cuma_indeksleri = gun_kategorisi_indeksleri([4])
            cumartesi_indeksleri = gun_kategorisi_indeksleri([5])
            pazar_indeksleri = gun_kategorisi_indeksleri([6])

            if cuma_haftasonu_siki_kural:
                for p in nobetci_personeller:
                    model.Add(sum(x[(p, d)] for d in cuma_indeksleri) <= 1)
                    model.Add(sum(x[(p, d)] for d in cumartesi_indeksleri) <= 1)
                    model.Add(sum(x[(p, d)] for d in pazar_indeksleri) <= 1)

            kategoriler = {
                "Pazartesi": ([0], 500), "Salı": ([1], 500), "Çarşamba": ([2], 500),
                "Perşembe": ([3], 1000), "Cuma": ([4], 1500), "Cumartesi": ([5], 2500), "Pazar": ([6], 2000),
            }

            birimler_listesi = ["PCR", "Mikro", "Kültür"] if birim_secimi == "Tüm Laboratuvar (Birleşik)" else [birim_secimi]
            saat_farklari, kategori_farklari = [], []

            for b_adi in birimler_listesi:
                b_personelleri = [p for p in nobetci_personeller if TUM_PERSONEL_VERISI.get(p, {}).get("birim") == b_adi or birim_secimi != "Tüm Laboratuvar (Birleşik)"]
                if not b_personelleri: continue

                max_b_saat = model.NewIntVar(0, 1000, f"max_saat_{b_adi}")
                min_b_saat = model.NewIntVar(0, 1000, f"min_saat_{b_adi}")
                for p in b_personelleri:
                    bu_ay_saat = model.NewIntVar(0, 1000, f"saat_{p}")
                    k8_saat_toplam = sum(k8[(p, d)] * 8 for d in kultur_8s_gun_indeksleri) if p in kultur_nobetcileri else 0
                    model.Add(bu_ay_saat == sum(x[(p, d)] * 24 for d in range(gun_sayisi)) + k8_saat_toplam)
                    model.Add(bu_ay_saat <= max_b_saat)
                    model.Add(bu_ay_saat >= min_b_saat)

                b_saat_farki = model.NewIntVar(0, 1000, f"saat_farki_{b_adi}")
                model.Add(b_saat_farki == max_b_saat - min_b_saat)
                saat_farklari.append(b_saat_farki)

                for kat_adi, (w_list, agirlik) in kategoriler.items():
                    day_indices = gun_kategorisi_indeksleri(w_list)
                    max_kat = model.NewIntVar(0, 100, f"max_{b_adi}_{kat_adi}")
                    min_kat = model.NewIntVar(0, 100, f"min_{b_adi}_{kat_adi}")
                    for p in b_personelleri:
                        bu_ay_kat_sayisi = sum(x[(p, d)] for d in day_indices)
                        kumulatif_kat_sayisi = get_prev(p, kat_adi) + bu_ay_kat_sayisi
                        model.Add(kumulatif_kat_sayisi <= max_kat)
                        model.Add(kumulatif_kat_sayisi >= min_kat)
                    fark = model.NewIntVar(0, 100, f"fark_{b_adi}_{kat_adi}")
                    model.Add(fark == max_kat - min_kat)
                    kategori_farklari.append(agirlik * fark)

            kultur8_farklari = []
            if kultur_nobetcileri and kultur_8s_gun_indeksleri:
                max_k8, min_k8 = model.NewIntVar(0, 50, "max_k8"), model.NewIntVar(0, 50, "min_k8")
                for p in kultur_nobetcileri:
                    bu_ay_k8 = sum(k8[(p, d)] for d in kultur_8s_gun_indeksleri)
                    kumulatif_k8 = get_prev_kultur8(p) + bu_ay_k8
                    model.Add(kumulatif_k8 <= max_k8)
                    model.Add(kumulatif_k8 >= min_k8)
                diff_k8 = model.NewIntVar(0, 50, "diff_k8")
                model.Add(diff_k8 == max_k8 - min_k8)
                kultur8_farklari.append(diff_k8)

            yigilma_farklari = []
            yariyil = gun_sayisi // 2
            for p in nobetci_personeller:
                n1 = sum(x[(p, d)] for d in range(0, yariyil))
                n2 = sum(x[(p, d)] for d in range(yariyil, gun_sayisi))
                diff = model.NewIntVar(0, 31, f"y_diff_{p}")
                model.AddAbsEquality(diff, n1 - n2)
                yigilma_farklari.append(diff)

            model.Minimize(
                10000000 * sum(saat_farklari) + 50000 * sum(kultur8_farklari) + sum(kategori_farklari) + 100 * sum(yigilma_farklari)
            )

            solver = cp_model.CpSolver()
            solver.parameters.num_search_workers = 2
            solver.parameters.max_time_in_seconds = 4.0
            status = solver.Solve(model)

            if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
                nobet_dict = {p: set() for p in nobetci_personeller}
                kultur_8s_dict = {p: set() for p in nobetci_personeller}
                acil_nobet_dict = {p: set() for p in nobetci_personeller}

                for d in range(gun_sayisi):
                    for p in nobetci_personeller:
                        if solver.Value(x[(p, d)]) == 1: nobet_dict[p].add(d + 1)
                        if p in kultur_nobetcileri and solver.Value(k8[(p, d)]) == 1: kultur_8s_dict[p].add(d + 1)

                bu_ay_acil_saat = {p: 0 for p in nobetci_personeller}
                kumulatif_acil_saat = {p: get_prev_acil(p) for p in nobetci_personeller}

                for p, g_list in sabit_acil_nobetler.items():
                    for d in g_list:
                        if d in nobet_dict[p]:
                            acil_nobet_dict[p].add(d)
                            bu_ay_acil_saat[p] += 24
                            kumulatif_acil_saat[p] += 24

                for d in range(1, gun_sayisi + 1):
                    mevcut_acil = [p for p in nobetci_personeller if d in acil_nobet_dict[p]]
                    if not mevcut_acil:
                        gun_nobetcileri = [p for p in nobetci_personeller if d in nobet_dict[p]]
                        if gun_nobetcileri:
                            secilen_acil = min(gun_nobetcileri, key=lambda p: (bu_ay_acil_saat[p], kumulatif_acil_saat[p]))
                            acil_nobet_dict[secilen_acil].add(d)
                            bu_ay_acil_saat[secilen_acil] += 24
                            kumulatif_acil_saat[secilen_acil] += 24

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
                    bu_ay_saat = bu_ay_toplam_nobet * 24
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

                    bu_ay_acil_saat_val = len(acil_nobet_dict[p]) * 24
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

                    for d in range(1, gun_sayisi + 1):
                        day_is_off = is_day_off(yil, ay, d, resmi_tatil_gunleri)
                        dt = datetime.date(yil, ay, d)

                        if is_muaf:
                            if day_is_off: p_row[str(d)] = "T"
                            elif d in yarim_gun_tatil_gunleri: p_row[str(d)] = "5"
                            else: p_row[str(d)] = "8"
                        else:
                            if d == 1 and p in gecmis_ay_son_gun_nobetcileri: p_row[str(d)] = "T" if day_is_off else "Nİ"
                            elif d in p_shifts: p_row[str(d)] = "24"
                            elif d in p_k8_shifts: p_row[str(d)] = "8"
                            elif (d - 1) in p_shifts: p_row[str(d)] = "T" if day_is_off else "Nİ"
                            elif dt.weekday() == 0 and (d - 2) in p_k8_shifts and datetime.date(yil, ay, d - 2).weekday() == 5: p_row[str(d)] = "Nİ"
                            elif (d - 1) in p_k8_shifts and datetime.date(yil, ay, d - 1).weekday() == 5: p_row[str(d)] = "T"
                            else:
                                if day_is_off: p_row[str(d)] = "T"
                                elif d in yarim_gun_tatil_gunleri: p_row[str(d)] = "5"
                                else: p_row[str(d)] = "8"

                    m = calculate_personel_puantaj_metrikleri(p, p_row, yil, ay, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri, acil_nobet_dict.get(p, set()), kultur_8s_tatil_gunleri, gecmis_ay_son_gun_normal, gecmis_ay_son_gun_acil)
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
                    "df_liste": df_liste, "df_istatistik": df_istatistik,
                    "df_puantaj": df_puantaj, "excel_bytes": excel_bytes,
                    "birim_secimi": birim_secimi, "yil": yil, "ay": ay,
                }
                st.balloons()
                st.success("✨ Nöbet çizelgesi ve Puantaj tablosu başarıyla oluşturuldu!")

            else:
                st.error("❌ Çözüm bulunamadı! Girilen kısıtlar, izinler veya sabit nöbetler çakışıyor olabilir.")

    # SONUÇLARI EKRANDA GÖSTER
    if st.session_state.hesaplanan_sonuc is not None:
        sonuc = st.session_state.hesaplanan_sonuc
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
            7: "Temmuz", 8: "Ağustos", 9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık"
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
