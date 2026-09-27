from pathlib import Path
import urllib.request
import importlib.util

ROOT = Path(__file__).resolve().parents[1]
helper = ROOT / "build" / "make_vba_bin.py"
if not helper.exists():
    urllib.request.urlretrieve("https://raw.githubusercontent.com/figaszewskioskar-ux/plik-rejestracji/main/narzedzia/make_vba_bin.py", helper)
spec = importlib.util.spec_from_file_location("make_vba_bin", helper)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
sources = {
    "Module1": (ROOT / "vba" / "PCD_MagicImport.bas").read_text(encoding="utf-8"),
    "PCD_Updater": (ROOT / "vba" / "PCD_Updater.bas").read_text(encoding="utf-8"),
    "Sheet2": (ROOT / "vba" / "Sheet2.bas").read_text(encoding="utf-8"),
}
bin_path = ROOT / "build" / "vbaProject-custom.bin"
bin_path.write_bytes(mod.make_vba_project(sources, {"Sheet2"}))
print(bin_path)
