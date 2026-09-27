from pathlib import Path
import olefile

p = Path("build/vbaProject.bin")
ole = olefile.OleFileIO(str(p), write_mode=False)
for s in ole.listdir():
    try:
        st = ole.get_stream_size(s)
    except Exception:
        st = "?"
    print("/".join(s), st)
ole.close()
