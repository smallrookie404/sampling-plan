import zipfile
import xml.etree.ElementTree as ET

path = r"G:\codex工作台\采样计划\sampling-plan-app\tests\shots\exported.xlsx"
z = zipfile.ZipFile(path)
ok = True
for name in z.namelist():
    if name.endswith(".xml") or name.endswith(".rels"):
        try:
            ET.fromstring(z.read(name))
        except Exception as e:
            ok = False
            print("XML 解析失败:", name, e)
print("ALL_XML_OK" if ok else "XML_PROBLEM")
print("文件条目数:", len(z.namelist()))
