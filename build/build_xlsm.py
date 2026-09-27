from pathlib import Path
import xlsxwriter

ROOT = Path(__file__).resolve().parents[1]
DIST = ROOT / "dist"
DIST.mkdir(exist_ok=True)
VERSION = "0.1.0-dev-10"
OUTPUT = DIST / f"PCD-Excel-Version-{VERSION}.xlsm"
VBA_BIN = ROOT / "build" / "vbaProject.bin"

if not VBA_BIN.exists():
    raise FileNotFoundError("build/vbaProject.bin is missing. Run build/compile_vba_project.py first.")

wb = xlsxwriter.Workbook(str(OUTPUT))
wb.set_properties({"title": "PCD Excel Version", "comments": "PCD development build 10"})
wb.set_vba_name("ThisWorkbook")
main = wb.add_worksheet("Intervention en cours")
magic = wb.add_worksheet("Magic Import")
main.set_vba_name("Sheet1")
magic.set_vba_name("Sheet2")
wb.add_vba_project(str(VBA_BIN))

main.hide_gridlines(2)
main.set_column("A:A", 3)
main.set_column("B:B", 28)
main.set_column("C:E", 15)
main.set_column("F:F", 28)
main.set_column("G:H", 16)

title = wb.add_format({"bold": True, "font_size": 18, "align": "center", "valign": "vcenter", "bg_color": "#D9E1F2", "border": 1})
label = wb.add_format({"bold": True, "font_color": "#666666", "border": 1, "bg_color": "#F7F7F7", "valign": "vcenter"})
value = wb.add_format({"border": 1, "text_wrap": True, "valign": "top"})
link = wb.add_format({"font_color": "#0563C1", "underline": 1, "align": "center", "valign": "vcenter", "border": 1})

main.merge_range("A1:H1", "INTERVENTION EN COURS", title)
main.write_url("F2", "internal:'Magic Import'!B5", link, "Import intelligent")
main.write_url("G2", "https://github.com/milance78/PCD-excel-version-/raw/refs/heads/main/dist/PCD-Excel-Version-latest.xlsm", link, "Mettre à jour")
main.write("H2", VERSION, label)

fields = [
    ("Infrastructure", "B4"), ("Réseau", "F4"),
    ("ID de l'intervention", "B6"), ("OAG ID", "F6"),
    ("Référence SNOW", "B8"), ("ID client", "F8"),
    ("Description d'intervention", "B10"), ("NA", "B12"),
    ("CID", "F12"), ("Adresse principale", "B14"),
    ("LOM key", "F14"), ("Boîte", "B16"), ("Étage", "C16"),
    ("Appt.", "D16"), ("Bloc", "E16"), ("N° de téléphone (GSM)", "F16"),
    ("Nom du client", "B19"), ("Commentaire", "B25"), ("Status", "F25")
]
for text, cell in fields:
    main.write(cell, text, label)
for rng in ["C4:E4", "G4:H4", "C6:E6", "G6:H6", "B10:F10", "B14:E14", "B25:E25"]:
    main.merge_range(rng, "", value)
for cell in ["B4","F4","B6","F6","B8","F8","B10","B12","F12","B14","F14","B16","C16","D16","E16","F16","B19","B25","F25"]:
    if cell not in {"B10", "B14", "B25"}:
        main.write(cell, "", value)
main.write("B28", f"DEV — {VERSION}", label)
main.insert_button("F28", {"macro": "CheckForUpdate", "caption": "Proveri ažuriranje", "width": 145, "height": 28})

magic.hide_gridlines(2)
magic.set_column("A:A", 3)
magic.set_column("B:H", 20)
magic.merge_range("A1:H1", "IMPORT INTELLIGENT", title)
magic.merge_range("A2:H3", "Colle le texte SAFE / NPS / Work Item dans la grande zone ci-dessous.", wb.add_format({"text_wrap": True, "valign": "vcenter", "border": 1, "bg_color": "#FFF2CC"}))
magic.merge_range("B5:H22", "", wb.add_format({"border": 1, "text_wrap": True, "valign": "top"}))
magic.insert_button("B24", {"macro": "ImportMagicFromSheet", "caption": "Importer", "width": 110, "height": 28})
magic.write_url("D24", "internal:'Intervention en cours'!A1", link, "Retour")
magic.write_url("F24", "https://raw.githubusercontent.com/milance78/PCD-excel-version-/main/dist/PCD-Excel-Version-latest.xlsm", link, "Preuzmi ručno najnoviju verziju")
wb.close()
print(OUTPUT)
