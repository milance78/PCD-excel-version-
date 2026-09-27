from pathlib import Path
import shutil
import uuid

from ms_ovba.vbaProject import VbaProject
from ms_ovba.Models.Entities.doc_module import DocModule
from ms_ovba.Models.Entities.std_module import StdModule
from ms_ovba.Views.project_ole_file import ProjectOleFile
from ms_ovba.Models.Entities.reference import Reference
from ms_ovba.Models.Entities.reference_registered import ReferenceRegistered
from ms_ovba.Models.Fields.libid_reference import LibidReference

ROOT = Path(__file__).resolve().parents[1]
VBA = ROOT / "vba"
OUT = ROOT / "build" / "vbaProject.bin"
TMP = ROOT / "build" / "vba_compile"

if TMP.exists():
    shutil.rmtree(TMP)
TMP.mkdir(parents=True)

project = VbaProject()
project.project_id = "{9E394C0B-697E-4AEE-9FA6-446F51FB30DC}"

def add_doc(name, source_name, guid, cookie, body=""):
    src = TMP / source_name
    src.write_text(
        "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\n"
        + f'Attribute VB_Name = "{name}"\n'
        + "Attribute VB_GlobalNameSpace = False\n"
        + "Attribute VB_Creatable = False\n"
        + "Attribute VB_PredeclaredId = True\n"
        + "Attribute VB_Exposed = True\n" + body,
        encoding="cp1252",
    )
    mod = DocModule(name)
    mod.add_workspace(0, 0, 0, 0, "C")
    mod.cookie = cookie
    mod.add_guid(uuid.UUID(guid))
    mod.add_file(str(src))
    mod.normalize_file()
    project.add_module(mod)

add_doc("ThisWorkbook", "ThisWorkbook.cls", "0002081900000000C000000000000046", 0xB81C)
add_doc("Sheet1", "Sheet1.cls", "0002082000000000C000000000000046", 0x9B9A)
add_doc("Sheet2", "Sheet2.cls", "0002082000000000C000000000000046", 0x9B9B, """\nPrivate Sub Worksheet_Change(ByVal Target As Range)\n    If Intersect(Target, Me.Range("B5")) Is Nothing Then Exit Sub\n    If Len(Trim$(CStr(Me.Range("B5").Value))) < 10 Then Exit Sub\n    On Error GoTo CleanFail\n    Application.EnableEvents = False\n    ParseMagicImportText CStr(Me.Range("B5").Value), False\nCleanFail:\n    Application.EnableEvents = True\nEnd Sub\n""")

for module_name in ("PCD_MagicImport", "PCD_Updater"):
    src = VBA / (module_name + ".bas")
    mod = StdModule("Module1" if module_name == "PCD_MagicImport" else module_name)
    mod.add_file(str(src))
    mod.normalize_file()
    project.add_module(mod)

stdole = LibidReference(uuid.UUID("0002043000000000C000000000000046"), "2.0", "0", r"C:\Windows\System32\stdole2.tlb", "OLE Automation")
office = LibidReference(uuid.UUID("2DF8D04C5BFA101BBDE500AA0044DE52"), "2.0", "0", r"C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL", "Microsoft Office 16.0 Object Library")
project.add_reference(Reference(ReferenceRegistered(stdole), "stdole"))
project.add_reference(Reference(ReferenceRegistered(office), "Office"))

ProjectOleFile.write_file(project)
generated = Path("vbaProject.bin")
if not generated.exists():
    raise FileNotFoundError("MS-OVBA compiler did not create vbaProject.bin")
shutil.copy2(generated, OUT)
print(OUT)
