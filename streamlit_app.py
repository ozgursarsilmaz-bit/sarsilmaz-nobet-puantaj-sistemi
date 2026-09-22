import calendar
import datetime
from io import BytesIO
import openpyxl
import pandas as pd
from ortools.sat.python import cp_model
import streamlit as st

# --- Sayfa Yapılandırması ---
st.set_page_config(
    page_title="Mikrobiyoloji Laboratuvarı Nöbet Dağılım ve Puantaj Sistemi",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)
# --- ŞİFRE KORUMASI ---
SIFRE = "Sars.2026"  # İstediğiniz şifreyi buraya yazabilirsiniz

if "authenticated" not in st.session_state:
  st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
  st.sidebar.title("🔒 Giriş Yap")
  girilen_sifre = st.sidebar.text_input(
      "Uygulama Şifresi", type="password", key="password_input"
  )

  if st.sidebar.button("Giriş"):
    if girilen_sifre == SIFRE:
      st.session_state["authenticated"] = True
      st.rerun()
    else:
      st.sidebar.error("❌ Hatalı şifre!")

  st.warning("⚠️ Lütfen devam etmek için sol menüden şifrenizi giriniz.")
  st.stop()  # Şifre doğru girilmeden uygulamanın geri kalanı yüklenmez

# --- ŞİFRE DOĞRUYSA SİSTEM BURADAN DEVAM EDER ---

# --- ÖZEL CSS TASARIMI ---
st.markdown(
    """
<style>
    .main {
        background-color: #F8F9FA;
    }
    .header-box {
        background: linear-gradient(135deg, #0d6efd 0%, #0a58ca 100%);
        color: white;
        padding: 24px;
        border-radius: 16px;
        box-shadow: 0 4px 12px rgba(13, 110, 253, 0.15);
        margin-bottom: 25px;
    }
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border: 1px solid #E9ECEF;
        padding: 15px 20px;
        border-radius: 12px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.03);
    }
    .section-title {
        font-size: 1.2rem;
        font-weight: 700;
        color: #212529;
        margin-bottom: 15px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #198754 0%, #146c43 100%);
        color: white;
        border: none;
        padding: 12px 24px;
        font-size: 1.1rem;
        font-weight: 600;
        border-radius: 10px;
        box-shadow: 0 4px 10px rgba(25, 135, 84, 0.2);
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #146c43 0%, #0f5132 100%);
        transform: translateY(-1px);
        box-shadow: 0 6px 14px rgba(25, 135, 84, 0.3);
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

def tr_norm(text):
  s = str(text).strip()
  replacements = [("İ", "i"), ("I", "ı"), ("ı", "i"), ("ğ", "g"), ("Ğ", "g"), 
                  ("ü", "u"), ("Ü", "u"), ("ş", "s"), ("Ş", "s"), ("ö", "o"), 
                  ("Ö", "o"), ("ç", "c"), ("Ç", "c")]
  for old, new in replacements:
    s = s.replace(old, new)
  return s.lower()

# --- BAŞLIK ARAYÜZÜ ---
st.markdown(
    """
<div class="header-box">
    <h1 style="margin:0; font-size: 2rem; font-weight: 800;">🏥 Mikrobiyoloji Laboratuvarı Nöbet Dağılım ve Puantaj Sistemi</h1>
    <p style="margin:5px 0 0 0; opacity: 0.9; font-size: 1rem;">Hafta Sonu Dengeli Dağılım | Kişiler Arası Nöbet Kısıtları & Adil Saat Optimizasyonu</p>
</div>
""",
    unsafe_allow_html=True,
)

# --- AYARLAR (SOL MENÜ) ---
st.sidebar.title("⚙️ Yönetim Paneli")

st.sidebar.subheader("📅 Tarih ve Kadro")
col_yil, col_ay = st.sidebar.columns(2)

with col_yil:
  yil = st.number_input("Yıl", value=2026, min_value=2024, max_value=2030)

with col_ay:
  ay = st.selectbox("Ay", list(range(1, 13)), index=8)  # 8 = Eylül

_, gun_sayisi = calendar.monthrange(yil, ay)

personel_input = st.sidebar.text_area(
    "Personel Listesi (Her satıra bir isim):",
    value="ÖZGÜR SARSILMAZ",
    height=280,
)
personeller = [p.strip().upper() for p in personel_input.split("\n") if p.strip()]

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
    "Kişisel Nöbet Arası Min. Dinlenme (Gün):",
    min_value=0,
    max_value=10,
    value=4,
)

persembe_pazar_yasagi = st.sidebar.checkbox(
    "Perşembe - Pazar Yasağı",
    value=True,
)

hafta_sonu_tek_nobet_siniri = st.sidebar.checkbox(
    "Hafta Sonu Ayda Maks. 1 Nöbet (Cumartesi / Pazar)",
    value=True,
    help="Aktif edildiğinde bir personele ay içinde 1'den fazla Cumartesi veya 1'den fazla Pazar nöbeti yazılmamasına çalışılır."
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
      for g in ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]:
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
col_count = 2

st.markdown(
    '<div class="section-title">🏖️ Personel İzin ve Mazeret Girişleri</div>',
    unsafe_allow_html=True,
)
izinler = {}
cols_izin = st.columns(col_count)
toplam_izin_sayisi = 0

for idx, p in enumerate(personeller):
  with cols_izin[idx % col_count]:
    selected_days = st.multiselect(
        f"**{p}** İzin Günleri:",
        options=list(range(1, gun_sayisi + 1)),
        default=[],
        key=f"leave_{p}",
    )
    izinler[p] = [d - 1 for d in selected_days]
    toplam_izin_sayisi += len(selected_days)

st.markdown("<br>", unsafe_allow_html=True)

st.markdown(
    '<div class="section-title">📌 Personel Sabit / Zorunlu Nöbet Girişleri</div>',
    unsafe_allow_html=True,
)
sabit_nobetler = {}
cols_sabit = st.columns(col_count)
toplam_sabit_sayisi = 0

for idx, p in enumerate(personeller):
  with cols_sabit[idx % col_count]:
    selected_sabit_days = st.multiselect(
        f"**{p}** Sabit Nöbet Günleri:",
        options=list(range(1, gun_sayisi + 1)),
        default=[],
        key=f"forced_{p}",
    )
    sabit_nobetler[p] = [d - 1 for d in selected_sabit_days]
    toplam_sabit_sayisi += len(selected_sabit_days)

st.markdown("<br>", unsafe_allow_html=True)

# --- KİŞİLER ARASI ÖZEL NÖBET ARALIĞI VE ÇAKIŞMA YASAĞI PANELİ ---
st.markdown(
    '<div class="section-title">🤝 Kişiler Arası Nöbet Mesafe ve Çakışma Yasağı Kuralları</div>',
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
        help="0: Aynı gün nöbet tutamazlar. 1: Aralarında en az 1 gün boşluk olmalı.",
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

    if dinlenme_gun_sayisi > 0:
      for p in personeller:
        for d in range(gun_sayisi - dinlenme_gun_sayisi):
          model.Add(
              sum(
                  x.get((p, d + k), 0)
                  for k in range(dinlenme_gun_sayisi + 1)
              )
              <= 1
          )

    if persembe_pazar_yasagi:
      for d in range(gun_sayisi - 3):
        if datetime.date(yil, ay, d + 1).weekday() == 3:
          for p in personeller:
            model.Add(x.get((p, d), 0) + x.get((p, d + 3), 0) <= 1)

    # --- KİŞİLER ARASI ÖZEL MESAFE / ÇAKIŞMA KURALLARI ---
    for rule in kisi_kisitlari:
      p1 = rule["ana"]
      aralik = rule["aralik"]
      for p2 in rule["yasaklilar"]:
        for d1 in range(gun_sayisi):
          for d2 in range(max(0, d1 - aralik), min(gun_sayisi, d1 + aralik + 1)):
            model.Add(x[(p1, d1)] + x[(p2, d2)] <= 1)

    # --- HAFTA SONU MAKSİMUM 1 NÖBET KISITI (CUMARTESİ / PAZAR) ---
    def gun_kategorisi_indeksleri(w_list):
      return [
          d
          for d in range(gun_sayisi)
          if datetime.date(yil, ay, d + 1).weekday() in w_list
      ]

    cumartesi_indeksleri = gun_kategorisi_indeksleri([5])
    pazar_indeksleri = gun_kategorisi_indeksleri([6])

    hafta_sonu_cezasi = []
    if hafta_sonu_tek_nobet_siniri:
      for p in personeller:
        cumartesi_sayisi = sum(x[(p, d)] for d in cumartesi_indeksleri)
        pazar_sayisi = sum(x[(p, d)] for d in pazar_indeksleri)
        
        # Mümkünse maksimum 1 Cumartesi ve 1 Pazar nöbeti yazılır
        cumartesi_fazla = model.NewIntVar(0, 10, f"cmts_fazla_{p}")
        pazar_fazla = model.NewIntVar(0, 10, f"pzr_fazla_{p}")
        
        model.Add(cumartesi_fazla >= cumartesi_sayisi - 1)
        model.Add(pazar_fazla >= pazar_sayisi - 1)
        
        hafta_sonu_cezasi.append(50000 * cumartesi_fazla + 50000 * pazar_fazla)

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

    # --- OPTİMİZASYON: Bu ayki saatleri eşitle ---
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
        100000 * saat_farki
        + sum(hafta_sonu_cezasi)
        + sum(kategori_farklari)
        + 10 * max_bu_ay_saat
    )

    solver = cp_model.CpSolver()
    status = solver.Solve(model)

    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
      st.balloons()
      st.success("✨ Hafta sonu nöbet sınırlaması ve tüm kısıtlar gözetilerek nöbet listesi başarıyla oluşturuldu!")

      # AYLIK ÇİZELGE
      liste_data = []
      for d in range(gun_sayisi):
        tarih = datetime.date(yil, ay, d + 1)
        nobetciler = [p for p in personeller if solver.Value(x[(p, d)]) == 1]
        liste_data.append({
            "Tarih": tarih.strftime("%d.%m.%Y"),
            "Gün": tr_gunler[tarih.weekday()],
            "Nöbetçi Personel": ", ".join(nobetciler),
        })
      df_liste = pd.DataFrame(liste_data)

      # İSTATİSTİK TABLOSU
      gunler_listesi = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
      
      columns_tuples = [("AD SOYAD", "")]
      for g in gunler_listesi:
        columns_tuples.extend([(g, "Devir"), (g, "Bu Ay"), (g, "Toplam")])
      
      columns_tuples.append(("T.NöbetSaati", ""))
      columns_tuples.append(("TOPLAM NÖBET", ""))
      
      multi_cols = pd.MultiIndex.from_tuples(columns_tuples)
      istatistik_rows = []

      for p in personeller:
        bu_ay_gunler = {g: 0 for g in gunler_listesi}
        bu_ay_toplam_nobet = 0
        
        for d in range(gun_sayisi):
          if solver.Value(x[(p, d)]) == 1:
            bu_ay_toplam_nobet += 1
            w = datetime.date(yil, ay, d + 1).weekday()
            bu_ay_gunler[tr_gunler[w]] += 1

        bu_ay_saat = sum(bu_ay_gunler[g] * gun_saat_haritasi[g] for g in gunler_listesi)
        
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

      # --- SEKMELİ SONUÇ EKRANI ---
      tab1, tab2, tab3 = st.tabs(
          ["📅 Aylık Çizelge", "📊 İstatistik & Mesai", "📥 Excel İndir"]
      )

      with tab1:
        st.subheader("🗓️ Aylık Görev Listesi")
        st.dataframe(df_liste, use_container_width=True, height=450)

      with tab2:
        st.subheader("📈 Personel Mesai Yükü İstatistiği")
        st.dataframe(df_istatistik, use_container_width=True)

      with tab3:
        st.subheader("📥 Çıktı Alma Paneli")
        st.write(
            "Hazırlanan nöbet çizelgesini ve istatistik analizini iki sekmeli resmi Excel belgesi olarak indirebilirsiniz:"
        )

        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
          df_liste.to_excel(writer, index=False, sheet_name="Nobet_Listesi")
          df_istatistik.to_excel(writer, sheet_name="Nobet_Istatistigi")

        st.download_button(
            label="📥 Excel Dosyasını İndir (.xlsx)",
            data=output.getvalue(),
            file_name=f"Nobet_Listesi_{yil}_{ay}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    else:
      st.error(
          "❌ Çözüm bulunamadı! Girilen kısıtlar, izinler veya sabit nöbetler birbiriyle çakışıyor olabilir."
      )
