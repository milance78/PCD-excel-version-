from pathlib import Path
import xlsxwriter

ROOT = Path(__file__).resolve().parents[1]
DIST = ROOT / "dist"
DIST.mkdir(exist_ok=True)
VERSION = "0.1.0-dev-57"
OUTPUT = DIST / f"PCD-Excel-Version-{VERSION}.xlsm"
VBA_BIN = ROOT / "build" / "vbaProject.bin"

if not VBA_BIN.exists():
    raise FileNotFoundError("build/vbaProject.bin is missing. Run build/compile_vba_project.py first.")

wb = xlsxwriter.Workbook(str(OUTPUT))
wb.set_properties({"title": "PCD Excel Version", "comments": "PCD Smart Import parity build"})
wb.set_vba_name("ThisWorkbook")
main = wb.add_worksheet("Intervention en cours")
magic = wb.add_worksheet("Magic Import")
main.set_vba_name("Sheet1")
magic.set_vba_name("Sheet2")
wb.add_vba_project(str(VBA_BIN))

main.hide_gridlines(2)
main.set_column("A:A", 3)
main.set_column("B:B", 28)
main.set_column("C:C", 18)
main.set_column("D:D", 18)
main.set_column("E:E", 3)
main.set_column("F:F", 28)
main.set_column("G:G", 18)
main.set_column("H:H", 18)

title = wb.add_format({"bold": True, "font_size": 18, "align": "center", "valign": "vcenter", "bg_color": "#D9E1F2", "border": 1})
label = wb.add_format({"bold": True, "font_color": "#666666", "border": 1, "bg_color": "#F7F7F7", "valign": "vcenter"})
value = wb.add_format({"border": 1, "text_wrap": True, "valign": "top"})
link = wb.add_format({"font_color": "#0563C1", "underline": 1, "align": "center", "valign": "vcenter", "border": 1})
section = wb.add_format({"bold": True, "font_size": 12, "bg_color": "#E2F0D9", "border": 1, "valign": "vcenter"})
input_fmt = wb.add_format({"border": 1, "text_wrap": True, "valign": "top", "bg_color": "#FFF2CC"})

main.merge_range("A1:H1", "INTERVENTION EN COURS", title)
main.write_url("F2", "internal:'Magic Import'!B5", link, "Coller WCT/NPS")
main.write_url("G2", "https://proximuscorp-my.sharepoint.com/personal/milan_pavlovic_proximus_com/Documents/Desktop/PCD-Excel-Version-dev.xlsm", link, "Mettre à jour")
main.write("H2", VERSION, label)

main.merge_range("B3:D3", "IDENTIFICATION", section)
main.merge_range("F3:H3", "RÉSEAU / STATUT", section)

pairs = [
    ("Intervention ID", "B4", "OAG ID", "F4"),
    ("Client ID", "B5", "SNOW ID", "F5"),
    ("Nom du client", "B6", "Téléphone", "F6"),
    ("Infrastructure", "B7", "Réseau", "F7"),
    ("Statut", "B8", "", ""),
]
for left, lc, right, rc in pairs:
    main.write(lc, left, label)
    if right:
        main.write(rc, right, label)

main.merge_range("C4:D4", "", value)
main.merge_range("G4:H4", "", value)
main.merge_range("C5:D5", "", value)
main.merge_range("G5:H5", "", value)
main.merge_range("C6:D6", "", value)
main.merge_range("G6:H6", "", value)
main.merge_range("C7:D7", "", value)
main.merge_range("G7:H7", "", value)
main.merge_range("C8:D8", "", value)

main.merge_range("B10:H10", "ADRESSE", section)
address_rows = [
    ("Rue", "B11", "Numéro", "F11"),
    ("Suffixe", "B12", "Code postal", "F12"),
    ("Ville", "B13", "Mailbox", "F13"),
    ("Étage", "B14", "Appartement", "F14"),
    ("Bloc", "B15", "LOM Key", "F15"),
]
for left, lc, right, rc in address_rows:
    main.write(lc, left, label)
    main.write(rc, right, label)
for rng in ["C11:D11","G11:H11","C12:D12","G12:H12","C13:D13","G13:H13","C14:D14","G14:H14","C15:D15","G15:H15"]:
    main.merge_range(rng, "", value)

main.write("B17", "CID / Service ID", label)
main.merge_range("C17:D17", "", value)
main.write("F17", "NA", label)
main.merge_range("G17:H17", "", value)

main.merge_range("B19:H19", "DESCRIPTION D'INTERVENTION", section)
main.merge_range("B20:H22", "", value)
main.merge_range("B24:H24", "COMMENTAIRE", section)
main.merge_range("B25:H28", "", value)

main.write("B30", "Version", label)
main.write("C30", VERSION, value)
main.insert_button("F30", {"macro": "CheckForUpdate", "caption": "Proveri ažuriranje", "width": 145, "height": 28})

magic.hide_gridlines(2)
magic.set_column("A:A", 3)
magic.set_column("B:H", 20)
magic.merge_range("A1:H1", "IMPORT INTELLIGENT", title)
magic.merge_range("A2:H3", "Colle le texte complet SAFE / NPS / Work Item. Le contenu est analysé automatiquement et les mêmes champs que dans « Intervention en cours » sont remplis.", wb.add_format({"text_wrap": True, "valign": "vcenter", "border": 1, "bg_color": "#FFF2CC"}))
magic.merge_range("B5:H22", "", input_fmt)
magic.insert_button("B24", {"macro": "ImportMagicFromSheet", "caption": "Importer", "width": 110, "height": 28})
magic.write_url("D24", "internal:'Intervention en cours'!A1", link, "Retour")
magic.write_url("F24", "https://proximuscorp-my.sharepoint.com/personal/milan_pavlovic_proximus_com/Documents/Desktop/PCD-Excel-Version-dev.xlsm", link, "Preuzmi ručno najnoviju verziju")
wb.close()
print(OUTPUT)
