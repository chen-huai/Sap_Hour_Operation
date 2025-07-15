import easyocr
from pathlib import Path
WORK_PATH = Path("C:\\Users\\chen-fr\\Desktop")
WORK_FILE = '1.ssss - 副本.jpg'
reader = easyocr.Reader(['ch_sim', 'en'], gpu = False)
# result = reader.readtext("C:\\Users\\chen-fr\\Desktop\\1.ssss.jpg",detail = 0)
result = reader.readtext(r"C:\\Users\\chen-fr\\Desktop\\1.ssss - 副本.jpg","rb",detail = 0)
# result = reader.readtext((WORK_PATH / WORK_FILE).as_posix(),detail = 0)
print(result)