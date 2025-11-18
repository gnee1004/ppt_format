import olefile, os, shutil

def replace_utf16_in_stream(ppt_path, old_str, new_str):
    if not olefile.isOleFile(ppt_path):
        raise RuntimeError("OLE 기반 PPT(.ppt)가 아님")

    stem, ext = os.path.splitext(ppt_path)
    out_path = stem + "_patched" + ext
    shutil.copy2(ppt_path, out_path)

    with olefile.OleFileIO(out_path) as ole_r:
        buf = bytearray(ole_r.openstream("PowerPoint Document").read())

    old_b = old_str.encode("utf-16le")
    new_b = new_str.encode("utf-16le")
    if len(old_b) != len(new_b):
        raise ValueError("치환 문자열은 반드시 길이가 같아야ddd 함!")

    count = buf.count(old_b)
    if count > 0:
        buf = buf.replace(old_b, new_b)
        print(f"[OK] {count}개 치환 완료")
    else:
        print("[WARN] 대상 문자열을 찾지 못함")

    with olefile.OleFileIO(out_path, write_mode=True) as ole_w:
        ole_w.write_stream("PowerPoint Document", bytes(buf))

    print("저장:", out_path)

if __name__ == "__main__":
    ppt_file = r"C:\경로"
    replace_utf16_in_stream(ppt_file, "테스트", "***")

