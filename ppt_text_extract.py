import olefile, struct

def parse_records(data, start=0, end=None, depth=0):
    if end is None:
        end = len(data)
    off = start
    while off + 8 <= end:
        recVer_inst, recType, recLen = struct.unpack_from("<HHI", data, off)
        header_off = off
        off += 8
        recData = data[off:off+recLen]

        indent = "  " * depth
        print(f"{indent}recType={recType}, recLen={recLen}, off={header_off}")

        # 텍스트 노드 찾기
        if recType == 4000:  # TextCharsAtom
            try:
                text = recData.decode("utf-16le", errors="ignore")
                print(f"{indent}  👉 TextCharsAtom: {text}")
            except Exception:
                pass
        elif recType == 4008:  # TextBytesAtom
            try:
                text = recData.decode("latin1", errors="ignore")
                print(f"{indent}  👉 TextBytesAtom: {text}")
            except Exception:
                pass

        # 컨테이너(recType 1000~2000대 등)는 내부를 재귀적으로 파싱
        if recType in (1000, 1016, 1006):  # DocumentContainer, SlideContainer, Notes
            parse_records(recData, 0, len(recData), depth+1)

        off += recLen

# 실행
ppt_path = r"C:\Users\geunh\Desktop\피피티 뜯기\테스트.ppt"
ole = olefile.OleFileIO(ppt_path)

with ole.openstream("PowerPoint Document") as stream:
    data = stream.read()

parse_records(data)
