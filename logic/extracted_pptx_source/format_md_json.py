from logic.extracted_pptx_source.get_pptx_json import get_pptx_source


def get_md_json(data):

    md_json = {}
    source_json = get_pptx_source()
    print(source_json)
    for slide in source_json["slides"]:
        slide_index = slide["slide_index"]
        slide = {'slide_index': 1, 'title': '基本設計書（対応方針）', 'elements': [{'id': 60, 'name': 'Rectangle 22', 'type': 'shape', 'bbox': {'left': 674.58, 'top': 510.92, 'width': 105.42, 'height': 29.08}, 'geometry': 'rect', 'text': [{'runs': [{'text': '２０２５年３月１７日'}], 'align': 'right (3)'}, {'runs': [{'text': 'カード基幹システム本部'}], 'align': 'right (3)'}], 'text_plain': '２０２５年３月１７日\nカード基幹システム本部', 'z_index': 0}, {'id': 3, 'name': 'タイトル 2', 'type': 'shape', 'bbox': {'left': 97.5, 'top': 88.37, 'width': 585.0, 'height': 173.77}, 'text': [{'runs': [{'text': '基本設計書（対応方針）'}]}], 'text_plain': '基本設計書（対応方針）', 'z_index': 1}, {'id': 4, 'name': '字幕 3', 'type': 'shape', 'bbox': {'left': 60.08, 'top': 298.45, 'width': 659.83, 'height': 188.46}, 'text': [{'runs': [{'text': '<案件ID：A058411・A061680>'}]}, {'runs': [{'text': '(242235)リボ手数料体系見直し'}]}], 'text_plain': '<案件ID：A058411・A061680>\n(242235)リボ手数料体系見直し', 'z_index': 2}], 'connection_graph': {'description': '图形间逻辑指向关系（edges）', 'edges': []}}

        if slide_index == 1:
            for elements in slide["elements"]:
                if elements["name"] == "字幕 3":
                    text = elements["text_plain"]
                    md_json["P1_V1_HAIKEI"] = text
                if elements["name"] == "タイトル 2":
                    text = elements["text_plain"]
                    md_json["P1_V1_TITLE"] = text
    return md_json
