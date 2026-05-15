
from logic.extracted_pptx_source.write_to_pptx import rebuild_ppt
from logic.tools import get_pptx_source

def get_pptx_source():
    file_path = "C:/Users/zhuan_rzwfs19/Desktop/work/work06/file/② 基本設計書（対応方針）_SR・CR新利率体系_ver1.2.0.pptx"
    page = 1
    data = get_pptx_source(file_path, page)
    rebuild_ppt(data)
