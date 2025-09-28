import olefile

ppt_path = "test.ppt"
ole = olefile.OleFileIO(ppt_path)

with ole.openstream("PowerPoint Document") as stream:
    data = stream.read(64)  # 처음 64바이트만 확인
    print("앞부분:", data.hex(" "))

ole.close()
