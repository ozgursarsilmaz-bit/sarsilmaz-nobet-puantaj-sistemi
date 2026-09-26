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
    page_title="Mikrobiyoloji Laboratuvarı Nöbet Dağılım ve Puantaj Sistemi",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- TAM PERSONEL VERİTABANI (Birim ve Muafiyet) ---
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
    .main .block-container {
        padding-top: 0.25rem !important;
        padding-bottom: 2rem !important;
    }
    [data-testid="stMainBlockContainer"],
    [data-testid="stAppViewContainer"] .main > div,
    [data-testid="stAppViewContainer"] .main .block-container {
        padding-top: 0.25rem !important;
    }
    [data-testid="stAppViewContainer"] { padding-top: 0 !important; }
    section.main { padding-top: 0 !important; }
    [data-testid="stHeader"], .stAppHeader { display: none !important; }
    .main { background-color: #F8F9FA; }
    
    .header-box {
        background: linear-gradient(135deg, #0d6efd 0%, #0a58ca 100%);
        color: white;
        padding: 12px 20px;
        border-radius: 10px;
        box-shadow: 0 3px 8px rgba(13, 110, 253, 0.15);
        margin-bottom: 12px;
        position: relative;
    }
    .header-box h1 { margin: 0; font-size: 1.35rem !important; font-weight: 700; line-height: 1.2; }
    .header-box p { margin: 2px 0 0 0; opacity: 0.88; font-size: 0.82rem !important; }
    .header-imza { position: absolute; bottom: 6px; right: 15px; font-size: 0.72rem; font-weight: 600; opacity: 0.75; letter-spacing: 0.5px; font-style: italic; }
    div[data-testid="stMetric"] { background-color: #FFFFFF; border: 1px solid #E9ECEF; padding: 10px 14px; border-radius: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.03); }
    .section-title { font-size: 1.1rem; font-weight: 700; color: #212529; margin-bottom: 10px; display: flex; align-items: center; gap: 6px; }
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #198754 0%, #146c43 100%);
        color: white; border: none; padding: 10px 20px; font-size: 1rem; font-weight: 600; border-radius: 8px;
        box-shadow: 0 4px 10px rgba(25, 135, 84, 0.2); transition: all 0.3s ease;
    }
    .stButton>button:hover { background: linear-gradient(135deg, #146c43 0%, #0f5132 100%); transform: translateY(-1px); box-shadow: 0 6px 14px rgba(25, 135, 84, 0.3); }
    .direct-download-btn {
        display: inline-block; width: 100%; text-align: center; background: linear-gradient(135deg, #0d6efd 0%, #0a58ca 100%);
        color: white !important; text-decoration: none !important; padding: 12px 20px; font-size: 1.05rem; font-weight: 700;
        border-radius: 8px; box-shadow: 0 4px 10px rgba(13, 110, 253, 0.25); transition: all 0.3s ease;
    }
    .direct-download-btn:hover { background: linear-gradient(135deg, #0a58ca 0%, #084298 100%); transform: translateY(-1px); box-shadow: 0 6px 14px rgba(13, 110, 253, 0.35); }
    [data-testid="stHorizontalBlock"] { align-items: center !important; gap: 0.5rem !important; }
    div[data-testid="column"] { padding: 0px !important; }
    .personel-giris-baslik { font-weight: 700; font-size: 0.75rem; color: #495057; padding: 2px 4px; white-space: nowrap; text-transform: uppercase; }
    .personel-giris-adi { min-height: 28px; height: 28px; display: flex; align-items: center; font-weight: 600; font-size: 0.78rem; color: #212529; padding: 0 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .personel-giris-ayirici { margin: 2px 0 !important; padding: 0 !important; height: 1px; border: 0; border-top: 1px solid #EDEFF1; }
    div[data-testid="stMultiSelect"] { margin: 0 !important; padding: 0 !important; }
    div[data-testid="stMultiSelect"] > div { margin: 0 !important; padding: 0 !important; }
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] { min-height: 28px !important; height: 28px !important; border-radius: 6px !important; }
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] > div { max-height: 28px !important; min-height: 28px !important; overflow-y: auto !important; padding: 0px 4px !important; align-content: center; }
    div[data-testid="stMultiSelect"] [data-baseweb="tag"] { font-size: 0.68rem !important; line-height: 16px !important; height: 18px !important; margin: 1px 2px 1px 0 !important; padding: 0 4px !important; }
    div[data-testid="stMultiSelect"] + div { display: none !important; }
</style>
""",
    unsafe_allow_html=True,
)

tr_gunler = {0: "Pazartesi", 1: "Salı", 2: "Çarşamba", 3: "Perşembe", 4: "Cuma", 5: "Cumartesi", 6: "Pazar"}

def tr_norm(text):
    s = str(text).strip()
    replacements = [("İ", "i"), ("I", "ı"), ("ı", "i"), ("ğ", "g"), ("Ğ", "g"), ("ü", "u"), ("Ü", "u"), ("ş", "s"), ("Ş", "s"), ("ö", "o"), ("Ö", "o"), ("ç", "c"), ("Ç", "c")]
    for old, new in replacements: s = s.replace(old, new)
    return s.lower()

# --- BAŞLIK ARAYÜZÜ ---
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

# --- YÖNETİM PANELİ İÇERİĞİ ---
st.sidebar.subheader("📅 Tarih ve Birim Seçimi")
col_yil, col_ay = st.sidebar.columns(2)

with col_yil: yil = st.number_input("Yıl", value=2026, min_value=2024, max_value=2030)
with col_ay: ay = st.selectbox("Ay", list(range(1, 13)), index=9)

_, gun_sayisi = calendar.monthrange(yil, ay)
gun_secenekleri = list(range(1, gun_sayisi + 1))

# --- RESMİ VE İDARİ TATİL SEÇİM PANELİ ---
st.sidebar.markdown("---")
st.sidebar.subheader("🏖️ Resmi & İdari Tatil Günleri")

resmi_tatil_gunleri = st.sidebar.multiselect("Tam Gün Tatil / Resmi Günler:", options=gun_secenekleri, default=[], help="Tam gün resmi/idari tatil günlerini seçiniz.")
yarim_gun_tatil_gunleri = st.sidebar.multiselect("Yarım Gün / Arife Günleri:", options=[g for g in gun_secenekleri if g not in resmi_tatil_gunleri], default=[], help="Arife veya yarım gün tatil günlerini seçiniz.")

# BİRİM SEÇİMİ
st.sidebar.markdown("---")
birim_secimi = st.sidebar.selectbox("🔬 Çalışma Grubu / Birim:", ["Mikro", "Kültür", "PCR", "Tüm Laboratuvar (Birleşik)"], index=3)

varsayilan_liste = [p_adi for p_adi, p_info in TUM_PERSONEL_VERISI.items() if birim_secimi == "Tüm Laboratuvar (Birleşik)" or p_info["birim"] == birim_secimi]
personel_input = st.sidebar.text_area("Personel Listesi (Her satıra bir isim):", value="\n".join(varsayilan_liste), height=220)

tum_girilen_personeller = [p.strip().upper() for p in personel_input.split("\n") if p.strip()]
nobetci_personeller = [p for p in tum_girilen_personeller if not TUM_PERSONEL_VERISI.get(p, {}).get("muaf", False)]
muaf_personeller = [p for p in tum_girilen_personeller if TUM_PERSONEL_VERISI.get(p, {}).get("muaf", False)]

st.sidebar.markdown("---")
st.sidebar.subheader("🛡️ Genel Kural Kurulumu")
gunluk_nobetci = st.sidebar.number_input("Birim Başı Günlük Nöbetçi İhtiyacı:", min_value=1, max_value=10, value=1, step=1)
dinlenme_gun_sayisi = st.sidebar.number_input("Genel Nöbet Arası Min. Dinlenme (Gün):", min_value=0, max_value=10, value=3)
persembe_pazar_yasagi = st.sidebar.checkbox("Genel Perşembe - Pazar Yasağı", value=True)
cuma_haftasonu_siki_kural = st.sidebar.checkbox("📌 Cuma / Cmts / Pzr Dengeli Dağılım (Maks 1 Gün)", value=True)
esnek_personel = st.sidebar.multiselect("🔓 Özel Esneklik Tanınacak Personel(ler):", options=nobetci_personeller, default=[])

# --- 📊 GEÇMİŞ AY ROTASYONU / DEVİR YÜKLEME PANELİ ---
st.sidebar.markdown("---")
st.sidebar.subheader("📊 Geçmiş Ay Rotasyonu (Önceki Ay Dosyası)")
uploaded_file = st.sidebar.file_uploader("Önceki Ayın Excel Dosyası:", type=["xlsx", "xls"], help="Sistemin ürettiği 3 sekmeli Excel dosyasını yükleyin.")

gecmis_istatistik = {}
gecmis_acil_istatistik = {}

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        if "İstatistik & Mesai Yükü" in xls.sheet_names:
            df_gecmis = pd.read_excel(xls, sheet_name="İstatistik & Mesai Yükü")
            gunler_sira = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]

            for r_idx in range(1, len(df_gecmis)):
                row = df_gecmis.iloc[r_idx]
                p_name = str(row.iloc[0]).strip().upper()
                if not p_name or p_name in ["NAN", "NONE", "AD SOYAD", "ADI SOYADI", "PERSONEL"]: continue

                p_dict = {}
                for g_idx, g_name in enumerate(gunler_sira):
                    toplam_col_idx = 3 + (g_idx * 3)
                    try: val = int(row.iloc[toplam_col_idx])
                    except (ValueError, TypeError, IndexError): val = 0
                    p_dict[g_name] = val
                gecmis_istatistik[p_name] = p_dict

                try: acil_devir_val = int(row.iloc[25])
                except (ValueError, TypeError, IndexError): acil_devir_val = 0
                gecmis_acil_istatistik[p_name] = acil_devir_val

        st.sidebar.success(f"✅ {len(gecmis_istatistik)} personelin devir verileri aktarıldı!")
    except Exception as e:
        st.sidebar.error(f"❌ Hata: Yüklenen Excel okunurken sorun oluştu ({e}).")

def get_prev(p_name, category): return gecmis_istatistik.get(p_name, {}).get(category, 0)
def get_prev_acil(p_name): return gecmis_acil_istatistik.get(p_name, 0)

# --- İZİNLİ VE SABİT NÖBET GİRİŞ PANELİ ---
st.markdown('<div class="section-title">📋 Personel Mazeret ve Sabit Nöbet Girişleri</div>', unsafe_allow_html=True)
izinler = {}
sabit_nobetler = {}
toplam_izin_sayisi = 0
toplam_sabit_sayisi = 0

baslik_personel, baslik_izin, baslik_sabit = st.columns([1.1, 1.8, 1.8])
with baslik_personel: st.markdown('<div class="personel-giris-baslik">👤 PERSONEL</div>', unsafe_allow_html=True)
with baslik_izin: st.markdown('<div class="personel-giris-baslik">🏖️ MAZERET / İZİN GÜNLERİ</div>', unsafe_allow_html=True)
with baslik_sabit: st.markdown('<div class="personel-giris-baslik">📌 SABİT / ZORUNLU NÖBET GÜNLERİ</div>', unsafe_allow_html=True)

for idx, p in enumerate(nobetci_personeller):
    col_personel, col_izin, col_sabit = st.columns([1.1, 1.8, 1.8])
    with col_personel: st.markdown(f'<div class="personel-giris-adi" title="{p}">{p}</div>', unsafe_allow_html=True)
    with col_izin:
        selected_days = st.multiselect("İzin Günleri", options=gun_secenekleri, default=[], key=f"leave_{p}", label_visibility="collapsed", placeholder="Gün seçin...")
        izinler[p] = [d - 1 for d in selected_days]
        toplam_izin_sayisi += len(selected_days)
    with col_sabit:
        selected_sabit_days = st.multiselect("Sabit Nöbet Günleri", options=gun_secenekleri, default=[], key=f"forced_{p}", label_visibility="collapsed", placeholder="Gün seçin...")
        sabit_nobetler[p] = [d - 1 for d in selected_sabit_days]
        toplam_sabit_sayisi += len(selected_sabit_days)
    if idx < len(nobetci_personeller) - 1:
        st.markdown('<hr class="personel-giris-ayirici">', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# --- KİŞİLER ARASI ÖZEL NÖBET ARALIĞI PANELİ ---
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

# --- CANLI METRİK DASHBOARD ---
st.markdown("<br>", unsafe_allow_html=True)
m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("👥 Nöbetçi Kadro", f"{len(nobetci_personeller)} Kişi")
m2.metric("📅 Ayın Gün Sayısı", f"{gun_sayisi} Gün")
m3.metric("🎯 Nöbet Slotu", f"{gun_sayisi * gunluk_nobetci * (3 if birim_secimi=='Tüm Laboratuvar (Birleşik)' else 1)} Nöbet")
m4.metric("🏖️ Kayıtlı İzinler", f"{toplam_izin_sayisi} Gün")
m5.metric("📌 Sabit Nöbetler", f"{toplam_sabit_sayisi} Gün")
m6.metric("🛡️ Nöbet Muaf", f"{len(muaf_personeller)} Kişi")
st.markdown("<br>", unsafe_allow_html=True)


# --- DİNAMİK SAAT VE TATİL YARDIMCI FONKSİYONLARI ---
def is_day_off(yil, ay, day, resmi_tatil_gunleri):
    dt = datetime.date(yil, ay, day)
    return (dt.weekday() in [5, 6]) or (day in resmi_tatil_gunleri)

def calculate_shift_hours(yil, ay, d, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri):
    dt = datetime.date(yil, ay, d)
    if d in yarim_gun_tatil_gunleri: return 19
    if (d + 1) in yarim_gun_tatil_gunleri: return 11

    if is_day_off(yil, ay, d, resmi_tatil_gunleri):
        if d < gun_sayisi:
            next_is_off = is_day_off(yil, ay, d + 1, resmi_tatil_gunleri)
            next_is_half = (d + 1) in yarim_gun_tatil_gunleri
            if not next_is_off and not next_is_half: return 16
        return 24
        
    if d < gun_sayisi and is_day_off(yil, ay, d + 1, resmi_tatil_gunleri): return 16
    w = dt.weekday()
    if w in [0, 1, 2, 3]: return 8
    elif w in [4, 6]: return 16
    else: return 24

def calculate_aylik_calisma_saati(yil, ay, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri):
    toplam_saat = 0
    for d in range(1, gun_sayisi + 1):
        if not is_day_off(yil, ay, d, resmi_tatil_gunleri):
            if d in yarim_gun_tatil_gunleri: toplam_saat += 5
            else: toplam_saat += 8
    return toplam_saat

def calculate_personel_puantaj_metrikleri(
    p_row_dict, yil, ay, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri, acil_days_set
):
    """
    TAM YÜZDE 100 MUTABIK KALINAN BAĞIMSIZ NORMAL VE ACİL (SARI) NÖBET HESAPLAMA FONKSİYONU
    - Sarı Nöbetler doğrudan Riskli_Gece ve Riskli_Normal sütunlarına aktarılır.
    - Normal (Beyaz) Nöbetler doğrudan Normal_Gece ve Normal_Normal sütunlarına aktarılır.
    """
    aylik_hedef_saat = calculate_aylik_calisma_saati(yil, ay, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri)
    
    toplam_calisma = 0
    norm_gece = 0
    norm_normal = 0
    risk_gece = 0
    risk_normal = 0

    for d in range(1, gun_sayisi + 1):
        val = str(p_row_dict.get(str(d), "")).strip()
        dt = datetime.date(yil, ay, d)
        w = dt.weekday()

        if val in ["8", "5", "16", "19", "11", "24"]:
            toplam_calisma += int(val)

        if val == "24":
            is_risk = (d in acil_days_set)
            
            # Tam Nöbet Saati Hesaplama (19, 24 vb.)
            n_saat = calculate_shift_hours(yil, ay, d, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri)

            # Gece (Artırımlı) Saat Hesabı
            next_is_holiday = (d in yarim_gun_tatil_gunleri) or ((d + 1) in yarim_gun_tatil_gunleri) or ((d + 1) in resmi_tatil_gunleri)
            if (d in yarim_gun_tatil_gunleri) or (d in resmi_tatil_gunleri) or (w in [4, 5, 6]) or next_is_holiday:
                g_saat = 12
            else:
                g_saat = 8

            gunduz_saat = max(0, n_saat - g_saat)

            # Bağımsız Ayrım Mantığı
            if is_risk:
                risk_gece += g_saat
                risk_normal += gunduz_saat
            else:
                norm_gece += g_saat
                norm_normal += gunduz_saat

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


# --- 3 SEKMELİ EXCEL OLUŞTURUCU ---
def generate_3_tab_excel(
    yil,
    ay,
    nobetci_personeller,
    tum_girilen_personeller,
    nobet_dict,
    acil_nobet_dict,
    gecmis_istatistik,
    gecmis_acil_istatistik,
    gun_sayisi,
    birim_secimi,
    resmi_tatil_gunleri,
    yarim_gun_tatil_gunleri
):
    wb = openpyxl.Workbook()

    font_title = Font(name="Calibri", size=12, bold=True, color="000000")
    font_header = Font(name="Calibri", size=10, bold=True, color="000000")
    font_body = Font(name="Calibri", size=9, bold=False, color="000000")
    font_bold = Font(name="Calibri", size=9, bold=True, color="000000")
    font_ni = Font(name="Calibri", size=9, bold=True, color="C00000")
    font_24 = Font(name="Calibri", size=9, bold=True, color="002060")

    fill_header = PatternFill(start_color="C5D9A4", end_color="C5D9A4", fill_type="solid")
    fill_green_bg = PatternFill(start_color="D8E4BC", end_color="D8E4BC", fill_type="solid")
    fill_grey = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    fill_yellow_acil = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    thin_side = Side(style="thin", color="A6A6A6")
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    
    # 🌟 DİKEY BAŞLIK METİN ALIGNMENT'I (TEXT ROTATION = 90)
    align_vertical_header = Alignment(horizontal="center", vertical="center", text_rotation=90, wrap_text=True)

    # 1. SEKME: GÖREV LİSTESİ (PCR | Mikro | Kültür)
    ws1 = wb.active
    ws1.title = "Aylık Görev Listesi"
    ws1.views.sheetView[0].showGridLines = True

    ws1.cell(row=1, column=1, value=f"{yil} yılı {ay}. Ay Mikrobiyoloji Laboratuvarı Nöbet Çizelgesi").font = font_title

    headers1 = ["Tarih", "Gün", "PCR", "Mikro", "Kültür"]
    for c_idx, h in enumerate(headers1, 1):
        cell = ws1.cell(row=3, column=c_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_cell

    for d in range(1, gun_sayisi + 1):
        r = d + 3
        tarih = datetime.date(yil, ay, d)

        pcr_list, mikro_list, kultur_list = [], [], []
        for p in nobetci_personeller:
            if d in nobet_dict.get(p, set()):
                p_unit = TUM_PERSONEL_VERISI.get(p, {}).get("birim", "Mikro")
                is_acil = (d in acil_nobet_dict.get(p, set()))
                p_text = f"{p} (Acil)" if is_acil else p
                
                if p_unit == "PCR": pcr_list.append(p_text)
                elif p_unit == "Mikro": mikro_list.append(p_text)
                elif p_unit == "Kültür": kultur_list.append(p_text)

        c1 = ws1.cell(row=r, column=1, value=tarih.strftime("%d.%m.%Y"))
        c2 = ws1.cell(row=r, column=2, value=tr_gunler[tarih.weekday()])
        c3 = ws1.cell(row=r, column=3, value=", ".join(pcr_list))
        c4 = ws1.cell(row=r, column=4, value=", ".join(mikro_list))
        c5 = ws1.cell(row=r, column=5, value=", ".join(kultur_list))

        for c in [c1, c2, c3, c4, c5]:
            c.font = font_body
            c.border = border_cell
            c.alignment = align_center if c in [c1, c2] else align_left
            if is_day_off(yil, ay, d, resmi_tatil_gunleri): c.fill = fill_grey

    ws1.column_dimensions["A"].width = 13
    ws1.column_dimensions["B"].width = 13
    ws1.column_dimensions["C"].width = 25
    ws1.column_dimensions["D"].width = 25
    ws1.column_dimensions["E"].width = 25

    # 2. SEKME: İSTATİSTİK & MESAİ YÜKÜ
    ws2 = wb.create_sheet("İstatistik & Mesai Yükü")
    ws2.views.sheetView[0].showGridLines = True

    gunler_listesi = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]

    ws2.merge_cells("A1:A2")
    cell_a = ws2.cell(row=1, column=1, value="AD SOYAD")
    cell_a.font = font_header
    cell_a.fill = fill_header
    cell_a.alignment = align_center
    cell_a.border = border_cell

    col_counter = 2
    for g in gunler_listesi:
        ws2.merge_cells(start_row=1, start_column=col_counter, end_row=1, end_column=col_counter + 2)
        top_cell = ws2.cell(row=1, column=col_counter, value=g)
        top_cell.font = font_header
        top_cell.fill = fill_header
        top_cell.alignment = align_center

        for idx, sub in enumerate(["Devir", "Bu Ay", "Toplam"]):
            sub_cell = ws2.cell(row=2, column=col_counter + idx, value=sub)
            sub_cell.font = font_header
            sub_cell.fill = fill_green_bg
            sub_cell.alignment = align_center
            sub_cell.border = border_cell
        col_counter += 3

    ws2.merge_cells(start_row=1, start_column=col_counter, end_row=2, end_column=col_counter)
    c_tn = ws2.cell(row=1, column=col_counter, value="T.NöbetSaati")
    c_tn.font = font_header
    c_tn.fill = fill_header
    c_tn.alignment = align_center
    c_tn.border = border_cell

    col_counter += 1
    ws2.merge_cells(start_row=1, start_column=col_counter, end_row=2, end_column=col_counter)
    c_tnb = ws2.cell(row=1, column=col_counter, value="TOPLAM NÖBET")
    c_tnb.font = font_header
    c_tnb.fill = fill_header
    c_tnb.alignment = align_center
    c_tnb.border = border_cell

    col_counter += 1
    ws2.merge_cells(start_row=1, start_column=col_counter, end_row=1, end_column=col_counter + 2)
    top_acil = ws2.cell(row=1, column=col_counter, value="ACİL NÖBET (SAAT)")
    top_acil.font = font_header
    top_acil.fill = fill_header
    top_acil.alignment = align_center

    for idx, sub in enumerate(["Devir", "Bu Ay", "Toplam"]):
        sub_cell = ws2.cell(row=2, column=col_counter + idx, value=sub)
        sub_cell.font = font_header
        sub_cell.fill = fill_green_bg
        sub_cell.alignment = align_center
        sub_cell.border = border_cell

    for r_idx in [1, 2]:
        for c_idx in range(1, col_counter + 3):
            ws2.cell(row=r_idx, column=c_idx).border = border_cell

    for p_idx, p in enumerate(nobetci_personeller, 3):
        ws2.cell(row=p_idx, column=1, value=p).alignment = align_left
        ws2.cell(row=p_idx, column=1).font = font_body
        ws2.cell(row=p_idx, column=1).border = border_cell

        bu_ay_gunler = {g: 0 for g in gunler_listesi}
        bu_ay_toplam_nobet = len(nobet_dict.get(p, set()))

        bu_ay_saat = 0
        for d in nobet_dict.get(p, set()):
            w = datetime.date(yil, ay, d).weekday()
            bu_ay_gunler[tr_gunler[w]] += 1
            bu_ay_saat += calculate_shift_hours(yil, ay, d, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri)

        c_i = 2
        for g in gunler_listesi:
            devir = int(gecmis_istatistik.get(p, {}).get(g, 0))
            bu_ay = int(bu_ay_gunler[g])
            toplam = devir + bu_ay

            for v in [devir, bu_ay, toplam]:
                cell = ws2.cell(row=p_idx, column=c_i, value=int(v))
                cell.font = font_body
                cell.alignment = align_center
                cell.border = border_cell
                c_i += 1

        cell_saat = ws2.cell(row=p_idx, column=c_i, value=int(bu_ay_saat))
        cell_saat.font = font_bold
        cell_saat.alignment = align_center
        cell_saat.border = border_cell

        cell_nobet = ws2.cell(row=p_idx, column=c_i + 1, value=int(bu_ay_toplam_nobet))
        cell_nobet.font = font_bold
        cell_nobet.alignment = align_center
        cell_nobet.border = border_cell

        bu_ay_acil_saat = sum(
            calculate_shift_hours(yil, ay, d, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri)
            for d in acil_nobet_dict.get(p, set())
        )
        acil_devir = int(gecmis_acil_istatistik.get(p, 0))
        acil_toplam = acil_devir + bu_ay_acil_saat

        c_i += 2
        for v in [acil_devir, bu_ay_acil_saat, acil_toplam]:
            cell = ws2.cell(row=p_idx, column=c_i, value=int(v))
            cell.font = font_bold if c_i % 3 == 0 else font_body
            cell.alignment = align_center
            cell.border = border_cell
            c_i += 1

    ws2.column_dimensions["A"].width = 25

    # 3. SEKME: PUANTAJ TABLOSU (DİKEY METİNLER İLE EKSİKSİZ)
    ws3 = wb.create_sheet("Puantaj Tablosu")
    ws3.views.sheetView[0].showGridLines = True

    # 🌟 DİKEY BAŞLIKLAR İÇİN HÜCRE YÜKSEKLİĞİ (110 PX)
    ws3.row_dimensions[5].height = 110

    ws3.cell(row=5, column=1, value="Adı Soyadı").fill = fill_green_bg
    ws3.cell(row=5, column=1).font = font_header
    ws3.cell(row=5, column=1).alignment = align_left
    ws3.cell(row=5, column=1).border = border_cell

    ws3.cell(row=5, column=2, value="Birim").fill = fill_green_bg
    ws3.cell(row=5, column=2).font = font_header
    ws3.cell(row=5, column=2).alignment = align_center
    ws3.cell(row=5, column=2).border = border_cell

    off_days = set()
    for day in range(1, gun_sayisi + 1):
        col_idx = day + 2
        if is_day_off(yil, ay, day, resmi_tatil_gunleri): off_days.add(day)

        cell = ws3.cell(row=5, column=col_idx, value=int(day))
        cell.font = font_header
        cell.alignment = align_center
        cell.border = border_cell
        cell.fill = fill_grey if day in off_days else fill_header

    ek_basliklar = [
        "Toplam çalışma saati",
        "Aylık Çalışma Saati",
        "Fazla nöbet saati",
        "Normal Nöbet - Artırımlı (Gece)",
        "Normal Nöbet - Artırımsız (Normal)",
        "Riskli Nöbet - Artırımlı (Gece)",
        "Riskli Nöbet - Artırımsız (Normal)",
    ]

    start_col = gun_sayisi + 3
    for idx, b_adi in enumerate(ek_basliklar):
        c_i = start_col + idx
        cell = ws3.cell(row=5, column=c_i, value=b_adi)
        cell.font = font_header
        cell.fill = fill_green_bg
        
        # 🌟 DİKEY METİN ROTASYONU (TEXT_ROTATION = 90)
        cell.alignment = align_vertical_header
        cell.border = border_cell

    for idx, p in enumerate(tum_girilen_personeller):
        r = 6 + idx
        ws3.row_dimensions[r].height = 18

        p_birim = TUM_PERSONEL_VERISI.get(p, {}).get("birim", birim_secimi)
        is_muaf = TUM_PERSONEL_VERISI.get(p, {}).get("muaf", False)

        c_name = ws3.cell(row=r, column=1, value=p)
        c_name.fill = fill_green_bg
        c_name.font = font_body
        c_name.alignment = align_left
        c_name.border = border_cell

        c_unit = ws3.cell(row=r, column=2, value=p_birim)
        c_unit.fill = fill_green_bg
        c_unit.font = font_body
        c_unit.alignment = align_center
        c_unit.border = border_cell

        p_shifts = nobet_dict.get(p, set())
        p_acils = acil_nobet_dict.get(p, set())
        p_row_dict = {}

        for day in range(1, gun_sayisi + 1):
            col_idx = day + 2
            cell = ws3.cell(row=r, column=col_idx)
            cell.border = border_cell
            cell.alignment = align_center

            if day in off_days: cell.fill = fill_grey

            if is_muaf:
                if day in off_days: cell.value = "T"
                elif day in yarim_gun_tatil_gunleri:
                    cell.value = 5
                    cell.font = font_body
                else:
                    cell.value = 8
                    cell.font = font_body
            else:
                if day in p_shifts:
                    cell.value = 24
                    cell.font = font_24
                    if day in p_acils:
                        cell.fill = fill_yellow_acil
                elif (day - 1) in p_shifts:
                    if is_day_off(yil, ay, day, resmi_tatil_gunleri): cell.value = "T"
                    else:
                        cell.value = "Nİ"
                        cell.font = font_ni
                else:
                    if day in off_days: cell.value = "T"
                    elif day in yarim_gun_tatil_gunleri:
                        cell.value = 5
                        cell.font = font_body
                    else:
                        cell.value = 8
                        cell.font = font_body

            p_row_dict[str(day)] = str(cell.value)

        m = calculate_personel_puantaj_metrikleri(
            p_row_dict, yil, ay, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri, p_acils
        )

        metrik_values = [
            int(m["Toplam Çalışma Saati"]),
            int(m["Aylık Çalışma Saati"]),
            int(m["Fazla Nöbet Saati"]),
            int(m["Normal_Gece"]),
            int(m["Normal_Normal"]),
            int(m["Riskli_Gece"]),
            int(m["Riskli_Normal"]),
        ]

        for m_idx, val in enumerate(metrik_values):
            col_idx = start_col + m_idx
            cell = ws3.cell(row=r, column=col_idx, value=val)
            cell.font = font_bold if m_idx < 3 else font_body
            cell.alignment = align_center
            cell.border = border_cell

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


# --- HESAPLAMA VE OPTİMİZASYON ---
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
        x = {}

        for p in nobetci_personeller:
            for d in range(gun_sayisi):
                x[(p, d)] = model.NewBoolVar(f"x_{p}_{d}")

        if birim_secimi == "Tüm Laboratuvar (Birleşik)":
            for d in range(gun_sayisi):
                model.Add(sum(x[(p, d)] for p in pcr_nobetcileri) == gunluk_nobetci)
                model.Add(sum(x[(p, d)] for p in mikro_nobetcileri) == gunluk_nobetci)
                model.Add(sum(x[(p, d)] for p in kultur_nobetcileri) == gunluk_nobetci)
        else:
            for d in range(gun_sayisi):
                model.Add(sum(x[(p, d)] for p in nobetci_personeller) == gunluk_nobetci)

        for p in nobetci_personeller:
            for d in izinler[p]: model.Add(x[(p, d)] == 0)

        for p in nobetci_personeller:
            for d in sabit_nobetler[p]: model.Add(x[(p, d)] == 1)

        # MİNİMUM DİNLENME SÜRESİ
        for p in nobetci_personeller:
            p_dinlenme = 1 if p in esnek_personel else max(1, int(dinlenme_gun_sayisi))
            for d in range(gun_sayisi - p_dinlenme):
                model.Add(sum(x.get((p, d + k), 0) for k in range(p_dinlenme + 1)) <= 1)

        # PERŞEMBE - PAZAR YASAĞI
        if persembe_pazar_yasagi:
            for d in range(gun_sayisi - 3):
                if datetime.date(yil, ay, d + 1).weekday() == 3:
                    for p in nobetci_personeller:
                        if p not in esnek_personel:
                            model.Add(x.get((p, d), 0) + x.get((p, d + 3), 0) <= 1)

        # KİŞİLER ARASI ÖZEL MESAFE KURALLARI
        for rule in kisi_kisitlari:
            p1 = rule["ana"]
            aralik = rule["aralik"]
            for p2 in rule["yasaklilar"]:
                for d1 in range(gun_sayisi):
                    for d2 in range(max(0, d1 - aralik), min(gun_sayisi, d1 + aralik + 1)):
                        model.Add(x[(p1, d1)] + x[(p2, d2)] <= 1)

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

        gun_saatleri = [calculate_shift_hours(yil, ay, d, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri) for d in range(1, gun_sayisi + 1)]

        max_bu_ay_saat = model.NewIntVar(0, 1000, "max_bu_ay_saat")
        min_bu_ay_saat = model.NewIntVar(0, 1000, "min_bu_ay_saat")

        for p in nobetci_personeller:
            bu_ay_saat = model.NewIntVar(0, 1000, f"saat_{p}")
            model.Add(bu_ay_saat == sum(x[(p, d)] * gun_saatleri[d] for d in range(gun_sayisi)))
            model.Add(bu_ay_saat <= max_bu_ay_saat)
            model.Add(bu_ay_saat >= min_bu_ay_saat)

        saat_farki = model.NewIntVar(0, 1000, "saat_farki")
        model.Add(saat_farki == max_bu_ay_saat - min_bu_ay_saat)

        kategoriler = {
            "Pazartesi": ([0], 500), "Salı": ([1], 500), "Çarşamba": ([2], 500),
            "Perşembe": ([3], 1000), "Cuma": ([1500]), "Cumartesi": ([2500]), "Pazar": ([2000])
        }

        kategori_farklari = []
        for kat_adi, (w_list, agirlik) in kategoriler.items():
            day_indices = gun_kategorisi_indeksleri(w_list)
            max_kat = model.NewIntVar(0, 100, f"max_{kat_adi}")
            min_kat = model.NewIntVar(0, 100, f"min_{kat_adi}")

            for p in nobetci_personeller:
                bu_ay_kat_sayisi = sum(x[(p, d)] for d in day_indices)
                kumulatif_kat_sayisi = get_prev(p, kat_adi) + bu_ay_kat_sayisi
                model.Add(kumulatif_kat_sayisi <= max_kat)
                model.Add(kumulatif_kat_sayisi >= min_kat)

            fark = model.NewIntVar(0, 100, f"fark_{kat_adi}")
            model.Add(fark == max_kat - min_kat)
            kategori_farklari.append(agirlik * fark)

        model.Minimize(100000 * saat_farki + sum(kategori_farklari) + 10 * max_bu_ay_saat)

        solver = cp_model.CpSolver()
        status = solver.Solve(model)

        if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
            st.balloons()
            st.success("✨ Nöbet çizelgesi ve Puantaj tablosu başarıyla oluşturuldu!")

            nobet_dict = {p: set() for p in nobetci_personeller}
            acil_nobet_dict = {p: set() for p in nobetci_personeller}

            for d in range(gun_sayisi):
                for p in nobetci_personeller:
                    if solver.Value(x[(p, d)]) == 1:
                        nobet_dict[p].add(d + 1)

            # ACİL NÖBET SAAT BAZLI ADİL DAĞITIM ALGORİTMASI
            kumulatif_acil_saat = {p: get_prev_acil(p) for p in nobetci_personeller}

            for d in range(1, gun_sayisi + 1):
                gun_nobetcileri = [p for p in nobetci_personeller if d in nobet_dict[p]]
                if gun_nobetcileri:
                    secilen_acil = min(gun_nobetcileri, key=lambda p: kumulatif_acil_saat[p])
                    acil_nobet_dict[secilen_acil].add(d)
                    
                    g_saat = calculate_shift_hours(yil, ay, d, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri)
                    kumulatif_acil_saat[secilen_acil] += g_saat

            # TABLO VE EKRAN HAZIRLIKLARI
            liste_data = []
            for d in range(1, gun_sayisi + 1):
                tarih = datetime.date(yil, ay, d)
                pcr_list, mikro_list, kultur_list = [], [], []

                for p in nobetci_personeller:
                    if d in nobet_dict[p]:
                        p_unit = TUM_PERSONEL_VERISI.get(p, {}).get("birim", "Mikro")
                        is_acil = (d in acil_nobet_dict[p])
                        p_text = f"{p} (Acil)" if is_acil else p

                        if p_unit == "PCR": pcr_list.append(p_text)
                        elif p_unit == "Mikro": mikro_list.append(p_text)
                        elif p_unit == "Kültür": kultur_list.append(p_text)

                liste_data.append({
                    "Tarih": tarih.strftime("%d.%m.%Y"),
                    "Gün": tr_gunler[tarih.weekday()],
                    "PCR": ", ".join(pcr_list),
                    "Mikro": ", ".join(mikro_list),
                    "Kültür": ", ".join(kultur_list),
                })
            df_liste = pd.DataFrame(liste_data)

            gunler_listesi = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
            columns_tuples = [("AD SOYAD", "")]
            for g in gunler_listesi:
                columns_tuples.extend([(g, "Devir"), (g, "Bu Ay"), (g, "Toplam")])

            columns_tuples.extend([("T.NöbetSaati", ""), ("TOPLAM NÖBET", "")])
            columns_tuples.extend([("ACİL NÖBET (SAAT)", "Devir"), ("ACİL NÖBET (SAAT)", "Bu Ay"), ("ACİL NÖBET (SAAT)", "Toplam")])

            multi_cols = pd.MultiIndex.from_tuples(columns_tuples)
            istatistik_rows = []

            for p in nobetci_personeller:
                bu_ay_gunler = {g: 0 for g in gunler_listesi}
                bu_ay_toplam_nobet = len(nobet_dict[p])

                bu_ay_saat = 0
                for d in nobet_dict[p]:
                    w = datetime.date(yil, ay, d).weekday()
                    bu_ay_gunler[tr_gunler[w]] += 1
                    bu_ay_saat += calculate_shift_hours(yil, ay, d, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri)

                row_dict = {("AD SOYAD", ""): p}

                for g in gunler_listesi:
                    devir_val = get_prev(p, g)
                    bu_ay_val = bu_ay_gunler[g]
                    toplam_val = devir_val + bu_ay_val

                    row_dict[(g, "Devir")] = devir_val
                    row_dict[(g, "Bu Ay")] = bu_ay_val
                    row_dict[(g, "Toplam")] = toplam_val

                row_dict[("T.NöbetSaati", "")] = bu_ay_saat
                row_dict[("TOPLAM NÖBET", "")] = bu_ay_toplam_nobet

                bu_ay_acil_saat = sum(
                    calculate_shift_hours(yil, ay, d, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri)
                    for d in acil_nobet_dict[p]
                )
                acil_devir = get_prev_acil(p)
                row_dict[("ACİL NÖBET (SAAT)", "Devir")] = acil_devir
                row_dict[("ACİL NÖBET (SAAT)", "Bu Ay")] = bu_ay_acil_saat
                row_dict[("ACİL NÖBET (SAAT)", "Toplam")] = acil_devir + bu_ay_acil_saat

                istatistik_rows.append(row_dict)

            df_istatistik = pd.DataFrame(istatistik_rows, columns=multi_cols)

            puantaj_rows = []
            for p in tum_girilen_personeller:
                p_birim = TUM_PERSONEL_VERISI.get(p, {}).get("birim", birim_secimi)
                is_muaf = TUM_PERSONEL_VERISI.get(p, {}).get("muaf", False)

                p_row = {"Adı Soyadı": p, "Birim": p_birim}
                for d in range(1, gun_sayisi + 1):
                    day_is_off = is_day_off(yil, ay, d, resmi_tatil_gunleri)

                    if is_muaf:
                        if day_is_off: p_row[str(d)] = "T"
                        elif d in yarim_gun_tatil_gunleri: p_row[str(d)] = "5"
                        else: p_row[str(d)] = "8"
                    else:
                        if d in nobet_dict[p]:
                            p_row[str(d)] = "24"
                        elif (d - 1) in nobet_dict[p]:
                            if day_is_off: p_row[str(d)] = "T"
                            else: p_row[str(d)] = "Nİ"
                        else:
                            if day_is_off: p_row[str(d)] = "T"
                            elif d in yarim_gun_tatil_gunleri: p_row[str(d)] = "5"
                            else: p_row[str(d)] = "8"

                m = calculate_personel_puantaj_metrikleri(
                    p_row, yil, ay, gun_sayisi, resmi_tatil_gunleri, yarim_gun_tatil_gunleri, acil_nobet_dict.get(p, set())
                )
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
                yil,
                ay,
                nobetci_personeller,
                tum_girilen_personeller,
                nobet_dict,
                acil_nobet_dict,
                gecmis_istatistik,
                gecmis_acil_istatistik,
                gun_sayisi,
                birim_secimi,
                resmi_tatil_gunleri,
                yarim_gun_tatil_gunleri
            )

            b64 = base64.b64encode(excel_bytes.getvalue()).decode()
            file_name = f"Nobet_ve_Puantaj_Listesi_{birim_secimi}_{yil}_{ay}.xlsx"
            href_link = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{file_name}" class="direct-download-btn">📥 3 Sekmeli Resmi Excel Dosyasını İndir (.xlsx)</a>'

            tab1, tab2, tab3, tab4 = st.tabs([
                "📅 Aylık Çizelge",
                "📊 İstatistik & Mesai",
                "📋 Puantaj Matrisi",
                "📥 Excel İndir",
            ])

            with tab1:
                st.subheader("🗓️ Birim Bazlı Aylık Görev Listesi (PCR | Mikro | Kültür)")
                st.dataframe(df_liste, use_container_width=True, height=450)

            with tab2:
                st.subheader("📈 Personel Mesai Yükü & Acil İstatistiği")
                st.dataframe(df_istatistik, use_container_width=True)

            with tab3:
                st.subheader("📋 Resmi Puantaj Tablosu Önizleme (Sarı Hücreler = Acil Nöbet)")
                st.dataframe(df_puantaj, use_container_width=True)

            with tab4:
                st.subheader("📥 Excel Dosyasını İndir")
                st.write("Aşağıdaki butona tıkladığınızda dosyanız doğrudan bilgisayarınıza indirilecektir:")
                st.markdown(href_link, unsafe_allow_html=True)

        else:
            st.error("❌ Çözüm bulunamadı! Girilen kısıtlar, izinler veya sabit nöbetler çakışıyor olabilir.")
