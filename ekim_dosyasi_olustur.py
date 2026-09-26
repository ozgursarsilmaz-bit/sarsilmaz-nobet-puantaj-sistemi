import datetime
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --- 1. EKİM 2026 MEVCUT NÖBET LİSTESİ VERİSİ ---
october_schedule = [
    ("01.10.2026", "Perşembe", "ŞADUMAN YALÇIN (Acil)", "DURMUŞ AKTÜRK", "M.ATABERK YILDIZ"),
    ("02.10.2026", "Cuma", "MURAT GENCER", "ÖZGÜR SARSILMAZ", "GÜLÇİN HORDACI (Acil)"),
    ("03.10.2026", "Cumartesi", "ŞADUMAN YALÇIN", "KADRİYE ARDIÇ (Acil)", "HİCRAN ATA AYHAN"),
    ("04.10.2026", "Pazar", "SUNA SARSILMAZ", "ŞENGÜL URLU (Acil)", "HACER BÜKÜM"),
    ("05.10.2026", "Pazartesi", "ŞADUMAN YALÇIN", "HÜMEYRA YEŞİLTAŞ (Acil)", "BİRSEL KAYA"),
    ("06.10.2026", "Salı", "MURAT GENCER", "REMZİYE AKSOY (Acil)", "İBRAHİM ER"),
    ("07.10.2026", "Çarşamba", "ŞADUMAN YALÇIN", "DURMUŞ AKTÜRK (Acil)", "M.ATABERK YILDIZ"),
    ("08.10.2026", "Perşembe", "SUNA SARSILMAZ (Acil)", "MELEK AKKAŞ", "GÜLÇİN HORDACI"),
    ("09.10.2026", "Cuma", "KEVSER DUMLU (Acil)", "ŞENGÜL URLU", "HACER BÜKÜM"),
    ("10.10.2026", "Cumartesi", "SEÇİL EMİR YILDIRAK", "MÜCELLÂ MANTAR", "BİRSEL KAYA (Acil)"),
    ("11.10.2026", "Pazar", "BELGİN UYSAL", "ÖZTÜRK MAVİŞ", "HİCRAN ATA AYHAN (Acil)"),
    ("12.10.2026", "Pazartesi", "MURAT GENCER", "REMZİYE AKSOY", "SÜLEYMAN POLAT (Acil)"),
    ("13.10.2026", "Salı", "ŞENEL TAŞ (Acil)", "KADRİYE ARDIÇ", "GÜLÇİN HORDACI"),
    ("14.10.2026", "Çarşamba", "KEVSER DUMLU", "MELEK AKKAŞ", "HACER BÜKÜM (Acil)"),
    ("15.10.2026", "Perşembe", "BELGİN UYSAL", "ŞÜKRAN KAYA", "MELEK BARUT (Acil)"),
    ("16.10.2026", "Cuma", "SUNA SARSILMAZ", "ÖZTÜRK MAVİŞ", "M.ATABERK YILDIZ (Acil)"),
    ("17.10.2026", "Cumartesi", "ŞENEL TAŞ", "ÖZGÜR SARSILMAZ (Acil)", "SÜLEYMAN POLAT"),
    ("18.10.2026", "Pazar", "MURAT GENCER", "REMZİYE AKSOY", "İBRAHİM ER (Acil)"),
    ("19.10.2026", "Pazartesi", "KEVSER DUMLU", "MUSTAFA TOSUN (Acil)", "GÜLÇİN HORDACI"),
    ("20.10.2026", "Salı", "SEÇİL EMİR YILDIRAK", "DURMUŞ AKTÜRK", "HİCRAN ATA AYHAN (Acil)"),
    ("21.10.2026", "Çarşamba", "ŞENEL TAŞ", "ÖZTÜRK MAVİŞ (Acil)", "MELEK BARUT"),
    ("22.10.2026", "Perşembe", "MURAT GENCER (Acil)", "MELEK AKKAŞ", "HACER BÜKÜM"),
    ("23.10.2026", "Cuma", "BELGİN UYSAL (Acil)", "HÜMEYRA YEŞİLTAŞ", "İBRAHİM ER"),
    ("24.10.2026", "Cumartesi", "SUNA SARSILMAZ", "MUSTAFA TOSUN", "MUSTAFA HARTOĞLU (Acil)"),
    ("25.10.2026", "Pazar", "ŞENEL TAŞ", "MEHMET BAĞCI (Acil)", "SÜLEYMAN POLAT"),
    ("26.10.2026", "Pazartesi", "SEÇİL EMİR YILDIRAK", "DURMUŞ AKTÜRK (Acil)", "BİRSEL KAYA"),
    ("27.10.2026", "Salı", "KEVSER DUMLU", "MELEK AKKAŞ (Acil)", "M.ATABERK YILDIZ"),
    ("28.10.2026", "Çarşamba", "BELGİN UYSAL", "MÜCELLÂ MANTAR (Acil)", "GÜLÇİN HORDACI"),
    ("29.10.2026", "Perşembe", "ŞENEL TAŞ", "HÜMEYRA YEŞİLTAŞ (Acil)", "İBRAHİM ER"),
    ("30.10.2026", "Cuma", "SEÇİL EMİR YILDIRAK (Acil)", "MEHMET BAĞCI", "MUSTAFA HARTOĞLU"),
    ("31.10.2026", "Cumartesi", "KEVSER DUMLU", "ŞÜKRAN KAYA (Acil)", "MELEK BARUT")
]

wb = openpyxl.Workbook()

# Stiller ve Yazı Tipleri
font_title = Font(name="Calibri", size=12, bold=True)
font_header = Font(name="Calibri", size=10, bold=True)
font_body = Font(name="Calibri", size=9)

fill_header = PatternFill(start_color="C5D9A4", end_color="C5D9A4", fill_type="solid")
fill_grey = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

thin_side = Side(style="thin", color="A6A6A6")
border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
align_center = Alignment(horizontal="center", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")

# --- 1. SEKME: AYLIK GÖREV LİSTESİ ---
ws1 = wb.active
ws1.title = "Aylık Görev Listesi"
ws1.views.sheetView[0].showGridLines = True

ws1.cell(row=1, column=1, value="2026 yılı 10. Ay Mikrobiyoloji Laboratuvarı Nöbet Çizelgesi").font = font_title

headers1 = ["Tarih", "Gün", "PCR", "Mikro", "Kültür"]
for c_idx, h in enumerate(headers1, 1):
    cell = ws1.cell(row=3, column=c_idx, value=h)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = align_center
    cell.border = border_cell

for idx, row in enumerate(october_schedule, 4):
    c1 = ws1.cell(row=idx, column=1, value=row[0])
    c2 = ws1.cell(row=idx, column=2, value=row[1])
    c3 = ws1.cell(row=idx, column=3, value=row[2])
    c4 = ws1.cell(row=idx, column=4, value=row[3])
    c5 = ws1.cell(row=idx, column=5, value=row[4])
    
    is_off = row[1] in ["Cumartesi", "Pazar"] or row[0] == "29.10.2026"
    for c in [c1, c2, c3, c4, c5]:
        c.font = font_body
        c.border = border_cell
        c.alignment = align_center if c in [c1, c2] else align_left
        if is_off:
            c.fill = fill_grey

ws1.column_dimensions["A"].width = 13
ws1.column_dimensions["B"].width = 13
ws1.column_dimensions["C"].width = 28
ws1.column_dimensions["D"].width = 28
ws1.column_dimensions["E"].width = 28

# Dosyayı kaydet
output_filename = "Ekim_2026_Mevcut_Nobet_Listesi.xlsx"
wb.save(output_filename)
print(f"🎉 Dosya başarıyla oluşturuldu: {output_filename}")
