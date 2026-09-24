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

# --- ÖZEL CSS TASARIMI (Sıkılaştırılmış Giriş Ekranı) ---
st.markdown(
    """
<style>
    .main { background-color: #F8F9FA; }
    .header-box {
        background: linear-gradient(135deg, #0d6efd 0%, #0a58ca 100%);
        color: white;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(13, 110, 253, 0.15);
        margin-bottom: 15px;
    }
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border: 1px solid #E9ECEF;
        padding: 10px 14px;
        border-radius: 10px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.03);
    }
    .section-title {
        font-size: 1.1rem;
        font-weight: 700;
        color: #212529;
        margin-bottom: 10px;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #198754 0%, #146c43 100%);
        color: white;
        border: none;
        padding: 10px 20px;
        font-size: 1rem;
        font-weight: 600;
        border-radius: 8px;
        box-shadow: 0 4px 10px rgba(25, 135, 84, 0.2);
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #146c43 0%, #0f5132 100%);
        transform: translateY(-1px);
        box-shadow: 0 6px 14px rgba(25, 135, 84, 0.3);
    }

    /* --- KOMPAKT PERSONEL GİRİŞ TABLOSU DÜZENLEMELERİ --- */
    /* Streamlit varsayılan dikey sütun boşluklarını (gap) sıfırla */
    [data-testid="stHorizontalBlock"] {
        align-items: center !important;
        gap: 0.5rem !important;
    }
    div[data-testid="column"] {
        padding: 0px !important;
    }

    .personel-giris-baslik {
        font-weight: 700;
        font-size: 0.75rem;
        color: #495057;
        padding: 2px 4px;
        white-space: nowrap;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .personel-giris-adi {
        min-height: 28px;
        height: 28px;
        display: flex;
        align-items: center;
        font-weight: 600;
        font-size: 0.78rem;
        color: #212529;
        padding: 0 4px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .personel-giris-ayirici {
        margin: 2px 0 !important;
        padding: 0 !important;
        height: 1px;
        border: 0;
        border-top: 1px solid #EDEFF1;
    }

    /* MultiSelect kompakt görünüm ve kaydırma ayarları */
    div[data-testid="stMultiSelect"] {
        margin: 0 !important;
        padding: 0 !important;
    }
    div[data-testid="stMultiSelect"] > div {
        margin: 0 !important;
        padding: 0 !important;
    }
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] {
        min-height: 28px !important;
        height: 28px !important;
        border-radius: 6px !important;
    }
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] > div {
        max-height: 28px !important;
        min-height: 28px !important;
        overflow-y: auto !important;
        overflow-x: hidden !important;
        padding: 0px 4px !important;
        align-content: center;
    }
    div[data-testid="stMultiSelect"] [data-baseweb="tag"] {
        font-size: 0.68rem !important;
        line-height: 16px !important;
        height: 18px !important;
        margin: 1px 2px 1px 0 !important;
        padding: 0 4px !important;
    }
    div[data-testid="stMultiSelect"] [data-baseweb="tag"] span {
        font-size: 0.68rem !important;
    }
    div[data-testid="stMultiSelect"] input {
        font-size: 0.70rem !important;
    }
    /* Widget altındaki gereksiz boşluğu sil */
    div[data-testid="stMultiSelect"] + div {
        display: none !important;
    }
</style>
""",
    unsafe_allow_html=True,
)

tr_gunler = {
    0: "Pazartesi",
    1: "Salı",
    2: "Çarşamba",
    3: "Perşembe",
    4: "Cuma",
    5: "Cumartesi",
    6: "Pazar",
}

gun_saat_haritasi = {
    "Pazartesi": 8,
    "Salı": 8,
    "Çarşamba": 8,
    "Perşembe": 8,
    "Cuma": 16,
    "Cumartesi": 24,
    "Pazar": 16,
}

personel_birimleri = {
    "ÖZGÜR SARSILMAZ": "Mikro",
    "AYŞEGÜL SARUHAN": "Kültür",
    "BELGİN PARİN": "Mikro",
    "BELGİN UYSAL": "PCR",
    "BİRSEL KAYA": "Kültür",
    "DENİZ YAMAN": "Mikro",
    "DURMUŞ AKTÜRK": "Mikro",
    "ESRA ÇORAK": "Mikro",
    "GÜLÇİN HORDACI": "Kültür",
    "HACER BÜKÜM": "Kültür",
    "HİCRAN ATA AYHAN": "Kültür",
    "HÜMEYRA YEŞİLTAŞ": "Mikro",
    "İBRAHİM ER": "Kültür",
    "KADRİYE ARDIÇ": "Mikro",
    "KEVSER DUMLU": "PCR",
    "MEHMET BAĞCI": "Mikro",
    "MELEK AKKAŞ": "Mikro",
    "MELEK BARUT": "Kültür",
    "MUHAMMED ATABERK YILDIZ": "Kültür",
    "MURAT GENCER": "PCR",
    "MUSTAFA HARTOĞLU": "Kültür",
    "MUSTAFA TOSUN": "Mikro",
    "MÜCELLÂ MANTAR": "Mikro",
    "ÖZTÜRK MAVİŞ": "Mikro",
    "REMZİYE AKSOY": "Mikro",
    "SEÇİL YILDIRAK": "PCR",
    "SEMRA KUMRUOĞLU": "Kültür",
    "SUNA SARSILMAZ": "PCR",
    "SÜLEYMAN POLAT": "Kültür",
    "ŞADUMAN YALÇIN": "PCR",
    "ŞENEL TAŞ": "PCR",
    "ŞENGÜL URLU": "Mikro",
    "ŞÜKRAN KAYA": "Kültür",
    "ZOZAN ATLI": "Kültür",
}


def tr_norm(text):
  s = str(text).strip()
  replacements = [
      ("İ", "i"),
      ("I", "ı"),
      ("ı", "i"),
      ("ğ", "g"),
      ("Ğ", "g"),
      ("ü", "u"),
      ("Ü", "u"),
      ("ş", "s"),
      ("Ş", "s"),
      ("ö", "o"),
      ("Ö", "o"),
      ("ç", "c"),
      ("Ç", "c"),
  ]
  for old, new in replacements:
    s = s.replace(old, new)
  return s.lower()


# --- BAŞLIK ARAYÜZÜ ---
st.markdown(
    """
<div class="header-box">
    <h1 style="margin:0; font-size: 1.8rem; font-weight: 800;">🏥 Mikrobiyoloji Laboratuvarı Nöbet Dağılım ve Puantaj Sistemi</h1>
    <p style="margin:3px 0 0 0; opacity: 0.9; font-size: 0.95rem;">Cuma/Cumartesi/Pazar Dengeli Dağılım & Adil Saat Optimizasyonu</p>
</div>
""",
    unsafe_allow_html=True,
)

# --- YÖNETİM PANELİ İÇERİĞİ ---
st.sidebar.subheader("📅 Tarih ve Kadro")
col_yil, col_ay = st.sidebar.columns(2)

with col_yil:
  yil = st.number_input("Yıl", value=2026, min_value=2024, max_value=2030)

with col_ay:
  ay = st.selectbox("Ay", list(range(1, 13)), index=9)

_, gun_sayisi = calendar.monthrange(yil, ay)

personel_input = st.sidebar.text_area(
    "Personel Listesi (Her satıra bir isim):",
    value="ÖZGÜR SARSILMAZ",
    height=200,
)
personeller = [
    p.strip().upper() for p in personel_input.split("\n") if p.strip()
]

st.sidebar.markdown("---")
st.sidebar.subheader("🛡️ Genel Kural Kurulumu")

gunluk_nobetci = st.sidebar.number_input(
    "Günlük Nöbetçi İhtiyacı:",
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
    "Genel Perşembe - Pazar Yasağı",
    value=True,
)

cuma_haftasonu_siki_kural = st.sidebar.checkbox(
    "📌 Cuma / Cmts / Pzr Dengeli Dağılım (Maks 1 Gün)",
    value=True,
    help=(
        "İşaretlendiğinde bir kişiye ayda en fazla 1 Cuma, 1 Cumartesi ve 1"
        " Pazar yazılabilir."
    ),
)

esnek_personel = st.sidebar.multiselect(
    "🔓 Özel Esneklik Tanınacak Personel(ler):",
    options=personeller,
    default=[],
    help=(
        "Burada seçilen kişilere Perşembe-Pazar yasağı UYGULANMAZ ve min."
        " dinlenme süresi 1 güne düşürülür."
    ),
)

st.sidebar.markdown("---")
st.sidebar.subheader("📊 Geçmiş Ay Rotasyonu (Eylül Dosyası)")
uploaded_file = st.sidebar.file_uploader(
    "Önceki Ayın Excel Dosyası:",
    type=["xlsx", "xls"],
    help="Eylül rotasyon.xlsx dosyasını yükleyin.",
)

gecmis_istatistik = {}

if uploaded_file is not None:
  try:
    xls = pd.ExcelFile(uploaded_file)
    df_gecmis = pd.read_excel(xls, sheet_name=0)

    col_mapping = {}
    for col in df_gecmis.columns:
      norm_key = tr_norm(col)
      col_mapping[norm_key] = col

    name_col = df_gecmis.columns[0]

    for _, row in df_gecmis.iterrows():
      p_name = str(row[name_col]).strip().upper()
      if not p_name or p_name in ["AD SOYAD", "ADI SOYADI", "PERSONEL"]:
        continue

      p_dict = {}
      for g in [
          "Pazartesi",
          "Salı",
          "Çarşamba",
          "Perşembe",
          "Cuma",
          "Cumartesi",
          "Pazar",
      ]:
        g_norm = tr_norm(g)
        if g_norm in col_mapping:
          try:
            val = int(row[col_mapping[g_norm]])
          except (ValueError, TypeError):
            val = 0
        else:
          val = 0
        p_dict[g] = val

      gecmis_istatistik[p_name] = p_dict

    st.sidebar.success(
        f"✅ {len(gecmis_istatistik)} personelin devir verileri yüklendi!"
    )
  except Exception as e:
    st.sidebar.error(f"❌ Hata: Yüklenen Excel okunurken sorun oluştu ({e}).")


def get_prev(p_name, category):
  return gecmis_istatistik.get(p_name, {}).get(category, 0)


# --- İZİNLİ VE SABİT NÖBET GİRİŞ PANELİ ---
st.markdown(
    '<div class="section-title">📋 Personel Mazeret ve Sabit Nöbet Girişleri</div>',
    unsafe_allow_html=True,
)

izinler = {}
sabit_nobetler = {}
toplam_izin_sayisi = 0
toplam_sabit_sayisi = 0

# Başlık satırı - Daraltılmış oranlar [1.1, 1.8, 1.8]
baslik_personel, baslik_izin, baslik_sabit = st.columns([1.1, 1.8, 1.8])

with baslik_personel:
  st.markdown(
      '<div class="personel-giris-baslik">👤 PERSONEL</div>',
      unsafe_allow_html=True,
  )

with baslik_izin:
  st.markdown(
      '<div class="personel-giris-baslik">🏖️ MAZERET / İZİN GÜNLERİ</div>',
      unsafe_allow_html=True,
  )

with baslik_sabit:
  st.markdown(
      '<div class="personel-giris-baslik">📌 SABİT / ZORUNLU NÖBET GÜNLERİ</div>',
      unsafe_allow_html=True,
  )

gun_secenekleri = list(range(1, gun_sayisi + 1))

# Satır içi personel girdileri
for idx, p in enumerate(personeller):
  col_personel, col_izin, col_sabit = st.columns([1.1, 1.8, 1.8])

  with col_personel:
    st.markdown(
        f'<div class="personel-giris-adi" title="{p}">{p}</div>',
        unsafe_allow_html=True,
    )

  with col_izin:
    selected_days = st.multiselect(
        "İzin Günleri",
        options=gun_secenekleri,
        default=[],
        key=f"leave_{p}",
        label_visibility="collapsed",
        placeholder="Gün seçin...",
    )
    izinler[p] = [d - 1 for d in selected_days]
    toplam_izin_sayisi += len(selected_days)

  with col_sabit:
    selected_sabit_days = st.multiselect(
        "Sabit Nöbet Günleri",
        options=gun_secenekleri,
        default=[],
        key=f"forced_{p}",
        label_visibility="collapsed",
        placeholder="Gün seçin...",
    )
    sabit_nobetler[p] = [d - 1 for d in selected_sabit_days]
    toplam_sabit_sayisi += len(selected_sabit_days)

  if idx < len(personeller) - 1:
    st.markdown(
        '<hr class="personel-giris-ayirici">',
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)

# --- KİŞİLER ARASI ÖZEL NÖBET ARALIĞI VE ÇAKIŞMA YASAĞI PANELİ ---
st.markdown(
    '<div class="section-title">🤝 Kişiler Arası Nöbet Mesafe ve Çakışma Yasağı'
    ' Kuralları</div>',
    unsafe_allow_html=True,
)

if "kisi_kisit_sayisi" not in st.session_state:
  st.session_state.kisi_kisit_sayisi = 1


def kisit_ekle():
  st.session_state.kisi_kisit_sayisi += 1


def kisit_cikar():
  if st.session_state.kisi_kisit_sayisi > 1:
    st.session_state.kisi_kisit_sayisi -= 1


kisi_kisitlari = []

for k_idx in range(st.session_state.kisi_kisit_sayisi):
  c1, c2, c3 = st.columns([1.2, 1, 2])
  with c1:
    p_ana = st.selectbox(
        f"Ana Personel #{k_idx+1}:",
        options=["Seçiniz..."] + personeller,
        key=f"p_ana_{k_idx}",
    )
  with c2:
    min_aralik = st.number_input(
        f"Min. Mesafe (Gün) #{k_idx+1}:",
        min_value=0,
        max_value=15,
        value=1,
        help=(
            "0: Aynı gün nöbet tutamazlar. 1: Aralarında en az 1 gün boşluk"
            " olmalı."
        ),
        key=f"min_aralik_{k_idx}",
    )
  with c3:
    p_yasakli_list = st.multiselect(
        f"Birlikte/Yakın Nöbet Tutamayacağı Kişiler #{k_idx+1}:",
        options=[p for p in personeller if p != p_ana],
        key=f"p_yasakli_{k_idx}",
    )

  if p_ana != "Seçiniz..." and p_yasakli_list:
    kisi_kisitlari.append({
        "ana": p_ana,
        "aralik": min_aralik,
        "yasaklilar": p_yasakli_list,
    })

col_btn1, col_btn2, _ = st.columns([1, 1, 4])
with col_btn1:
  st.button("➕ Yeni Kısıt Ekle", on_click=kisit_ekle)
with col_btn2:
  st.button("➖ Kısıt Sil", on_click=kisit_cikar)


# --- CANLI METRİK DASHBOARD ---
st.markdown("<br>", unsafe_allow_html=True)
m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("👥 Aktif Personel", f"{len(personeller)} Kişi")
m2.metric("📅 Ayın Gün Sayısı", f"{gun_sayisi} Gün")
m3.metric("🎯 Nöbet Slotu", f"{gun_sayisi * gunluk_nobetci} Nöbet")
m4.metric("🏖️ Kayıtlı İzinler", f"{toplam_izin_sayisi} Gün")
m5.metric("📌 Sabit Nöbetler", f"{toplam_sabit_sayisi} Gün")
m6.metric("🤝 Özel Kısıtlar", f"{len(kisi_kisitlari)} Kural")

st.markdown("<br>", unsafe_allow_html=True)


# --- 3 SEKMELİ ÖZEL EXCEL OLUŞTURUCU ---
def generate_3_tab_excel(
    yil, ay, personeller, nobet_dict, gecmis_istatistik, gun_sayisi
):
  wb = openpyxl.Workbook()

  font_title = Font(name="Calibri", size=12, bold=True, color="000000")
  font_header = Font(name="Calibri", size=10, bold=True, color="000000")
  font_body = Font(name="Calibri", size=9, bold=False, color="000000")
  font_bold = Font(name="Calibri", size=9, bold=True, color="000000")
  font_ni = Font(name="Calibri", size=9, bold=True, color="C00000")
  font_24 = Font(name="Calibri", size=9, bold=True, color="002060")

  fill_header = PatternFill(
      start_color="C5D9A4", end_color="C5D9A4", fill_type="solid"
  )
  fill_green_bg = PatternFill(
      start_color="D8E4BC", end_color="D8E4BC", fill_type="solid"
  )
  fill_grey = PatternFill(
      start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"
  )

  thin_side = Side(style="thin", color="A6A6A6")
  border_cell = Border(
      left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
  )

  align_center = Alignment(horizontal="center", vertical="center")
  align_left = Alignment(horizontal="left", vertical="center")

  # 1. SEKME: GÖREV LİSTESİ
  ws1 = wb.active
  ws1.title = "Aylık Görev Listesi"
  ws1.views.sheetView[0].showGridLines = True

  ws1.cell(
      row=1, column=1, value=f"{yil} yılı {ay}. Ay Nöbet Çizelgesi"
  ).font = font_title

  headers1 = ["Tarih", "Gün", "Nöbetçi Personel"]
  for c_idx, h in enumerate(headers1, 1):
    cell = ws1.cell(row=3, column=c_idx, value=h)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = align_center
    cell.border = border_cell

  for d in range(1, gun_sayisi + 1):
    r = d + 3
    tarih = datetime.date(yil, ay, d)
    nobetciler = [p for p in personeller if d in nobet_dict.get(p, set())]

    c1 = ws1.cell(row=r, column=1, value=tarih.strftime("%d.%m.%Y"))
    c2 = ws1.cell(row=r, column=2, value=tr_gunler[tarih.weekday()])
    c3 = ws1.cell(row=r, column=3, value=", ".join(nobetciler))

    for c in [c1, c2, c3]:
      c.font = font_body
      c.border = border_cell
      c.alignment = align_center if c != c3 else align_left
      if tarih.weekday() in [5, 6]:
        c.fill = fill_grey

  ws1.column_dimensions["A"].width = 15
  ws1.column_dimensions["B"].width = 15
  ws1.column_dimensions["C"].width = 45

  # 2. SEKME: İSTATİSTİK
  ws2 = wb.create_sheet("İstatistik & Mesai Yükü")
  ws2.views.sheetView[0].showGridLines = True

  gunler_listesi = [
      "Pazartesi",
      "Salı",
      "Çarşamba",
      "Perşembe",
      "Cuma",
      "Cumartesi",
      "Pazar",
  ]

  ws2.merge_cells("A1:A2")
  cell_a = ws2.cell(row=1, column=1, value="AD SOYAD")
  cell_a.font = font_header
  cell_a.fill = fill_header
  cell_a.alignment = align_center
  cell_a.border = border_cell

  col_counter = 2
  for g in gunler_listesi:
    ws2.merge_cells(
        start_row=1,
        start_column=col_counter,
        end_row=1,
        end_column=col_counter + 2,
    )
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

  ws2.merge_cells(
      start_row=1, start_column=col_counter, end_row=2, end_column=col_counter
  )
  c_tn = ws2.cell(row=1, column=col_counter, value="T.NöbetSaati")
  c_tn.font = font_header
  c_tn.fill = fill_header
  c_tn.alignment = align_center
  c_tn.border = border_cell

  col_counter += 1
  ws2.merge_cells(
      start_row=1, start_column=col_counter, end_row=2, end_column=col_counter
  )
  c_tnb = ws2.cell(row=1, column=col_counter, value="TOPLAM NÖBET")
  c_tnb.font = font_header
  c_tnb.fill = fill_header
  c_tnb.alignment = align_center
  c_tnb.border = border_cell

  for r_idx in [1, 2]:
    for c_idx in range(1, col_counter + 1):
      ws2.cell(row=r_idx, column=c_idx).border = border_cell

  for p_idx, p in enumerate(personeller, 3):
    ws2.cell(row=p_idx, column=1, value=p).alignment = align_left
    ws2.cell(row=p_idx, column=1).font = font_body
    ws2.cell(row=p_idx, column=1).border = border_cell

    bu_ay_gunler = {g: 0 for g in gunler_listesi}
    bu_ay_toplam_nobet = len(nobet_dict.get(p, set()))

    for d in nobet_dict.get(p, set()):
      w = datetime.date(yil, ay, d).weekday()
      bu_ay_gunler[tr_gunler[w]] += 1

    c_i = 2
    for g in gunler_listesi:
      devir = gecmis_istatistik.get(p, {}).get(g, 0)
      bu_ay = bu_ay_gunler[g]
      toplam = devir + bu_ay

      for v in [devir, bu_ay, toplam]:
        cell = ws2.cell(row=p_idx, column=c_i, value=v)
        cell.font = font_body
        cell.alignment = align_center
        cell.border = border_cell
        c_i += 1

    bu_ay_saat = sum(
        bu_ay_gunler[g] * gun_saat_haritasi[g] for g in gunler_listesi
    )

    cell_saat = ws2.cell(row=p_idx, column=c_i, value=bu_ay_saat)
    cell_saat.font = font_bold
    cell_saat.alignment = align_center
    cell_saat.border = border_cell

    cell_nobet = ws2.cell(row=p_idx, column=c_i + 1, value=bu_ay_toplam_nobet)
    cell_nobet.font = font_bold
    cell_nobet.alignment = align_center
    cell_nobet.border = border_cell

  ws2.column_dimensions["A"].width = 25

  # 3. SEKME: PUANTAJ TABLOSU
  ws3 = wb.create_sheet("Puantaj Tablosu")
  ws3.views.sheetView[0].showGridLines = True

  ws3.row_dimensions[5].height = 20
  ws3.cell(row=5, column=1, value="Adı Soyadı").fill = fill_green_bg
  ws3.cell(row=5, column=1).font = font_header
  ws3.cell(row=5, column=1).alignment = align_left
  ws3.cell(row=5, column=1).border = border_cell

  ws3.cell(row=5, column=2, value="Birim").fill = fill_green_bg
  ws3.cell(row=5, column=2).font = font_header
  ws3.cell(row=5, column=2).alignment = align_center
  ws3.cell(row=5, column=2).border = border_cell

  grey_days = set()
  for day in range(1, gun_sayisi + 1):
    col_idx = day + 2
    dt = datetime.date(yil, ay, day)
    if dt.weekday() in [5, 6]:
      grey_days.add(day)

    cell = ws3.cell(row=5, column=col_idx, value=day)
    cell.font = font_header
    cell.alignment = align_center
    cell.border = border_cell
    cell.fill = fill_grey if day in grey_days else fill_header

  for idx, p in enumerate(personeller):
    r = 6 + idx
    ws3.row_dimensions[r].height = 18

    birim = personel_birimleri.get(p, "Mikro")

    c_name = ws3.cell(row=r, column=1, value=p)
    c_name.fill = fill_green_bg
    c_name.font = font_body
    c_name.alignment = align_left
    c_name.border = border_cell

    c_unit = ws3.cell(row=r, column=2, value=birim)
    c_unit.fill = fill_green_bg
    c_unit.font = font_body
    c_unit.alignment = align_center
    c_unit.border = border_cell

    p_shifts = nobet_dict.get(p, set())

    for day in range(1, gun_sayisi + 1):
      col_idx = day + 2
      cell = ws3.cell(row=r, column=col_idx)
      cell.border = border_cell
      cell.alignment = align_center

      if day in grey_days:
        cell.fill = fill_grey

      if day in p_shifts:
        cell.value = 24
        cell.font = font_24
      elif (day - 1) in p_shifts:
        cell.value = "Nİ"
        cell.font = font_ni
      else:
        if day in grey_days:
          cell.value = ""
        else:
          cell.value = 8
          cell.font = font_body

  ws3.column_dimensions["A"].width = 28
  ws3.column_dimensions["B"].width = 12
  for day in range(1, gun_sayisi + 1):
    col_letter = get_column_letter(day + 2)
    ws3.column_dimensions[col_letter].width = 4.5

  output = BytesIO()
  wb.save(output)
  output.seek(0)
  return output


# --- HESAPLAMA VE OPTİMİZASYON ---
if st.button("🚀 Otomatik ve Adil Nöbet Listesini Oluştur"):
  hata_listesi = []

  for d in range(gun_sayisi):
    musait_sayisi = sum(1 for p in personeller if d not in izinler[p])
    if musait_sayisi < gunluk_nobetci:
      tarih_str = datetime.date(yil, ay, d + 1).strftime("%d.%m.%Y")
      hata_listesi.append(
          f"⚠️ **{tarih_str}** ({d+1}. gün): En az {gunluk_nobetci} kişi gerekli"
          f" ancak izinler nedeniyle sadece {musait_sayisi} kişi müsait."
      )

  for p in personeller:
    ortak = set(izinler[p]).intersection(set(sabit_nobetler[p]))
    for d in ortak:
      tarih_str = datetime.date(yil, ay, d + 1).strftime("%d.%m.%Y")
      hata_listesi.append(
          f"❌ **{p}**, **{tarih_str}** ({d+1}. gün) tarihi için hem 'İzinli'"
          " hem de 'Sabit Nöbetçi' olarak seçilmiş!"
      )

  if hata_listesi:
    for err in hata_listesi:
      st.error(err)
  else:
    model = cp_model.CpModel()
    x = {}

    for p in personeller:
      for d in range(gun_sayisi):
        x[(p, d)] = model.NewBoolVar(f"x_{p}_{d}")

    for d in range(gun_sayisi):
      model.Add(sum(x[(p, d)] for p in personeller) == gunluk_nobetci)

    for p in personeller:
      for d in izinler[p]:
        model.Add(x[(p, d)] == 0)

    for p in personeller:
      for d in sabit_nobetler[p]:
        model.Add(x[(p, d)] == 1)

    # DİNAMİK MİNİMUM DİNLENME SÜRESİ
    for p in personeller:
      p_dinlenme = (
          1 if p in esnek_personel else max(1, int(dinlenme_gun_sayisi))
      )
      for d in range(gun_sayisi - p_dinlenme):
        model.Add(
            sum(x.get((p, d + k), 0) for k in range(p_dinlenme + 1)) <= 1
        )

    # DİNAMİK PERŞEMBE - PAZAR YASAĞI
    if persembe_pazar_yasagi:
      for d in range(gun_sayisi - 3):
        if datetime.date(yil, ay, d + 1).weekday() == 3:  # Perşembe
          for p in personeller:
            if p not in esnek_personel:
              model.Add(x.get((p, d), 0) + x.get((p, d + 3), 0) <= 1)

    # KİŞİLER ARASI ÖZEL MESAFE KURALLARI
    for rule in kisi_kisitlari:
      p1 = rule["ana"]
      aralik = rule["aralik"]
      for p2 in rule["yasaklilar"]:
        for d1 in range(gun_sayisi):
          for d2 in range(
              max(0, d1 - aralik), min(gun_sayisi, d1 + aralik + 1)
          ):
            model.Add(x[(p1, d1)] + x[(p2, d2)] <= 1)

    def gun_kategorisi_indeksleri(w_list):
      return [
          d
          for d in range(gun_sayisi)
          if datetime.date(yil, ay, d + 1).weekday() in w_list
      ]

    cuma_indeksleri = gun_kategorisi_indeksleri([4])
    cumartesi_indeksleri = gun_kategorisi_indeksleri([5])
    pazar_indeksleri = gun_kategorisi_indeksleri([6])

    # CUMA / CUMARTESİ / PAZAR TEKİL MAKSİMUM 1 NÖBET KURALI
    if cuma_haftasonu_siki_kural:
      for p in personeller:
        cuma_sayisi = sum(x[(p, d)] for d in cuma_indeksleri)
        cumartesi_sayisi = sum(x[(p, d)] for d in cumartesi_indeksleri)
        pazar_sayisi = sum(x[(p, d)] for d in pazar_indeksleri)

        model.Add(cuma_sayisi <= 1)
        model.Add(cumartesi_sayisi <= 1)
        model.Add(pazar_sayisi <= 1)

    gun_saatleri = []
    for d in range(gun_sayisi):
      w = datetime.date(yil, ay, d + 1).weekday()
      if w in [0, 1, 2, 3]:
        h = 8
      elif w in [4, 6]:
        h = 16
      else:
        h = 24
      gun_saatleri.append(h)

    max_bu_ay_saat = model.NewIntVar(0, 1000, "max_bu_ay_saat")
    min_bu_ay_saat = model.NewIntVar(0, 1000, "min_bu_ay_saat")

    for p in personeller:
      bu_ay_saat = model.NewIntVar(0, 1000, f"saat_{p}")
      model.Add(
          bu_ay_saat
          == sum(x[(p, d)] * gun_saatleri[d] for d in range(gun_sayisi))
      )
      model.Add(bu_ay_saat <= max_bu_ay_saat)
      model.Add(bu_ay_saat >= min_bu_ay_saat)

    saat_farki = model.NewIntVar(0, 1000, "saat_farki")
    model.Add(saat_farki == max_bu_ay_saat - min_bu_ay_saat)

    kategoriler = {
        "Pazartesi": ([0], 500),
        "Salı": ([1], 500),
        "Çarşamba": ([2], 500),
        "Perşembe": ([3], 1000),
        "Cuma": ([4], 1500),
        "Cumartesi": ([5], 2500),
        "Pazar": ([6], 2000),
    }

    kategori_farklari = []
    for kat_adi, (w_list, agirlik) in kategoriler.items():
      day_indices = gun_kategorisi_indeksleri(w_list)
      max_kat = model.NewIntVar(0, 100, f"max_{kat_adi}")
      min_kat = model.NewIntVar(0, 100, f"min_{kat_adi}")

      for p in personeller:
        bu_ay_kat_sayisi = sum(x[(p, d)] for d in day_indices)
        kumulatif_kat_sayisi = get_prev(p, kat_adi) + bu_ay_kat_sayisi
        model.Add(kumulatif_kat_sayisi <= max_kat)
        model.Add(kumulatif_kat_sayisi >= min_kat)

      fark = model.NewIntVar(0, 100, f"fark_{kat_adi}")
      model.Add(fark == max_kat - min_kat)
      kategori_farklari.append(agirlik * fark)

    # HEDEF FONKSİYONU
    model.Minimize(
        100000 * saat_farki + sum(kategori_farklari) + 10 * max_bu_ay_saat
    )

    solver = cp_model.CpSolver()
    status = solver.Solve(model)

    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
      st.balloons()
      st.success("✨ Nöbet çizelgesi ve Puantaj tablosu başarıyla oluşturuldu!")

      nobet_dict = {p: set() for p in personeller}

      liste_data = []
      for d in range(gun_sayisi):
        tarih = datetime.date(yil, ay, d + 1)
        nobetciler = []
        for p in personeller:
          if solver.Value(x[(p, d)]) == 1:
            nobetciler.append(p)
            nobet_dict[p].add(d + 1)

        liste_data.append({
            "Tarih": tarih.strftime("%d.%m.%Y"),
            "Gün": tr_gunler[tarih.weekday()],
            "Nöbetçi Personel": ", ".join(nobetciler),
        })
      df_liste = pd.DataFrame(liste_data)

      gunler_listesi = [
          "Pazartesi",
          "Salı",
          "Çarşamba",
          "Perşembe",
          "Cuma",
          "Cumartesi",
          "Pazar",
      ]

      columns_tuples = [("AD SOYAD", "")]
      for g in gunler_listesi:
        columns_tuples.extend([(g, "Devir"), (g, "Bu Ay"), (g, "Toplam")])

      columns_tuples.append(("T.NöbetSaati", ""))
      columns_tuples.append(("TOPLAM NÖBET", ""))

      multi_cols = pd.MultiIndex.from_tuples(columns_tuples)
      istatistik_rows = []

      for p in personeller:
        bu_ay_gunler = {g: 0 for g in gunler_listesi}
        bu_ay_toplam_nobet = len(nobet_dict[p])

        for d in nobet_dict[p]:
          w = datetime.date(yil, ay, d).weekday()
          bu_ay_gunler[tr_gunler[w]] += 1

        bu_ay_saat = sum(
            bu_ay_gunler[g] * gun_saat_haritasi[g] for g in gunler_listesi
        )

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

        istatistik_rows.append(row_dict)

      df_istatistik = pd.DataFrame(istatistik_rows, columns=multi_cols)

      puantaj_rows = []
      for p in personeller:
        p_row = {"Adı Soyadı": p, "Birim": personel_birimleri.get(p, "Mikro")}
        for d in range(1, gun_sayisi + 1):
          dt = datetime.date(yil, ay, d)
          is_weekend = dt.weekday() in [5, 6]

          if d in nobet_dict[p]:
            p_row[str(d)] = "24"
          elif (d - 1) in nobet_dict[p]:
            p_row[str(d)] = "Nİ"
          else:
            p_row[str(d)] = "-" if is_weekend else "8"
        puantaj_rows.append(p_row)

      df_puantaj = pd.DataFrame(puantaj_rows)

      tab1, tab2, tab3, tab4 = st.tabs([
          "📅 Aylık Çizelge",
          "📊 İstatistik & Mesai",
          "📋 Puantaj Matrisi",
          "📥 Excel İndir",
      ])

      with tab1:
        st.subheader("🗓️ Aylık Görev Listesi")
        st.dataframe(df_liste, use_container_width=True, height=450)

      with tab2:
        st.subheader("📈 Personel Mesai Yükü İstatistiği")
        st.dataframe(df_istatistik, use_container_width=True)

      with tab3:
        st.subheader("📋 Resmi Puantaj Tablosu Önizleme (8 / 24 / Nİ)")
        st.dataframe(df_puantaj, use_container_width=True)

      with tab4:
        st.subheader("📥 3 Sekmeli Resmi Excel İndirme Paneli")
        st.write(
            "Aşağıdaki butona tıklayarak **Aylık Çizelge**, **Mesai"
            " İstatistikleri** ve **Resmi Puantaj Tablosu**'nu tek bir Excel"
            " dosyasında indirebilirsiniz:"
        )

        excel_bytes = generate_3_tab_excel(
            yil, ay, personeller, nobet_dict, gecmis_istatistik, gun_sayisi
        )

        st.download_button(
            label="📥 3 Sekmeli Tam Excel Dosyasını İndir (.xlsx)",
            data=excel_bytes,
            file_name=f"Nobet_ve_Puantaj_Listesi_{yil}_{ay}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    else:
      st.error(
          "❌ Çözüm bulunamadı! Girilen kısıtlar, izinler veya sabit nöbetler"
          " birbiriyle çakışıyor olabilir."
      )
