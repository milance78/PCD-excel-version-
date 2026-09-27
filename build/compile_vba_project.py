from pathlib import Path
import shutil
import uuid

from ms_ovba.vbaProject import VbaProject
from ms_ovba.Models.Entities.doc_module import DocModule
from ms_ovba.Models.Entities.std_module import StdModule
from ms_ovba.Views.project_ole_file import ProjectOleFile

ROOT = Path(__file__).resolve().parents[1]
VBA = ROOT / "vba"
OUT = ROOT / "build" / "vbaProject.bin"
TMP = ROOT / "build" / "vba_compile"

if TMP.exists():
    shutil.rmtree(TMP)
TMP.mkdir(parents=True)

project = VbaProject()
project.project_id = "{9E394C0B-697E-4AEE-9FA6-446F51FB30DC}"


def write_crlf(source: Path, destination: Path) -> None:
    text = source.read_text(encoding="utf-8")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    destination.write_text(text.replace("\n", "\r\n"), encoding="utf-8")


def add_doc(name, source_name, guid, cookie, body=""):
    src = TMP / source_name
    header = (
        "VERSION 1.0 CLASS\r\n"
        "BEGIN\r\n"
        "  MultiUse = -1  'True\r\n"
        "END\r\n"
        f'Attribute VB_Name = "{name}"\r\n'
        "Attribute VB_GlobalNameSpace = False\r\n"
        "Attribute VB_Creatable = False\r\n"
        "Attribute VB_PredeclaredId = True\r\n"
        "Attribute VB_Exposed = True\r\n"
    )
    body = body.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    src.write_text(header + body, encoding="cp1252")

    mod = DocModule(name)
    mod.add_workspace(0, 0, 0, 0, "C")
    mod.cookie = cookie
    mod.add_guid(uuid.UUID(guid))
    mod.add_file(str(src))
    mod.normalize_file()
    project.add_module(mod)


add_doc(
    "ThisWorkbook",
    "ThisWorkbook.cls",
    "0002081900000000C000000000000046",
    0xB81C,
)

add_doc(
    "Sheet1",
    "Sheet1.cls",
    "0002082000000000C000000000000046",
    0x9B9A,
)

add_doc(
    "Sheet2",
    "Sheet2.cls",
    "0002082000000000C000000000000046",
    0x9B9B,
    """
Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range("B5")) Is Nothing Then Exit Sub
    If Len(Trim$(CStr(Me.Range("B5").Value))) < 10 Then Exit Sub
    On Error GoTo CleanFail
    Application.EnableEvents = False
    ParseMagicImportText CStr(Me.Range("B5").Value), False
CleanFail:
    Application.EnableEvents = True
End Sub
""",
)

for module_name in ("PCD_MagicImport", "PCD_Updater"):
    source = VBA / (module_name + ".bas")
    normalized = TMP / (module_name + ".bas")
    write_crlf(source, normalized)

    mod = StdModule("Module1" if module_name == "PCD_MagicImport" else module_name)
    mod.add_file(str(normalized))
    mod.normalize_file()
    project.add_module(mod)

ProjectOleFile.write_file(project)

generated = Path("vbaProject.bin")
if not generated.exists():
    raise FileNotFoundError("MS-OVBA compiler did not create vbaProject.bin")

shutil.copy2(generated, OUT)
print(OUT)
