import olefile

ppt_path = "테스트.ppt"
ole = olefile.OleFileIO(ppt_path)

print("PPT 내부 스트림 구조:")
for entry in ole.listdir():
    print("-", "/".join(entry))

ole.close()
