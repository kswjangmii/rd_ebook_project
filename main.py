# RD JSON 생성
import pandas as pd
import json
import uuid

def gen_id(prefix="to"):
    """위젯, 에셋 등을 위한 고유 ID 생성 (형식: prefix-16진수6자리)"""
    return f"{prefix}-{uuid.uuid4().hex[:6]}"

def gen_page_id():
    """페이지 고유 ID 생성 (형식: page-16진수4자리-16진수4자리)"""
    return f"page-{uuid.uuid4().hex[:4]}-{uuid.uuid4().hex[:4]}"

def create_base_json():
    """원본 JSON과 동일한 메타데이터, 설정(config), 에셋 폴더 구조를 반환"""
    return {
        "id": "project-264141459-089b",
        "version": "5.8.355",
        "title": "Rourke_G1_lbb_4",
        "description": "",
        "metadata": {"image": {}},
        "config": {
            "language": "en-us", "defaultPageWidth": "750", "defaultPageHeight": "1000",
            "userStylesheets": [], "mobile": {"landscapeDoublePage": False},
            "transition": "flip", "doublepage": True, "firstPageNumber": 1,
            "usePageImageALT": False, "print": False, "share": False, "shadow": True,
            "slideshow": False, "pdfUrl": "", "widgetDataManager": {"url": ""},
            "defaultPageSizeType": "default", "convertImageMinSize": "1600",
            "convertFormat": "image", "theme": "audio_book", "disablePageMoveGesture": False,
            "disableZoomGesture": False, "mobileOnly": False, "disablingContentLoadingBar": False,
            "reversePageTurnDirection": False, "setDefaultFitWidth": False, "publishType": "local",
            "fonts": [
                {"group": "b", "displayName": "jostbook", "name": "jostbook", "css": "_fonts.css", "user": True, "foldername": "jost.book"},
                {"group": "b", "displayName": "poppins", "name": "poppinsR", "css": "_fonts.css", "user": True, "foldername": "poppins"},
                {"group": "b", "displayName": "poppinsM", "name": "poppinsM", "css": "_fonts.css", "user": True, "foldername": "poppinsM"}
            ],
            "useWebkitBrowser": False, "logo": {"alt": ""}, "homepageUrl": "",
            "homepageNewWindow": False, "useCoverPage": True, "bgImage": {}, "bgSound": {},
            "tocPageNumber": "", "tocOpenAll": False, "progressPageControl": False,
            "usePageHistory": False, "brokenPageCorrection": False,
            "thumbnail": {"width": 600, "height": 600, "url": "{ASSETS_DIR}thumbnail.jpg"},
            "minZeroPageLabel": False, "minLearnTimeEnable": False, "minLearnTime": "",
            "publishUniqueId": f"publish-264141459-{uuid.uuid4().hex[:4]}"
        },
        "pages": [], "subpages": [],
        "assets": {
            k: {"folders": [{"id": gen_id("as"), "_default": True, "readonly": True, "lock": False, "name": "기본폴더"}], "assets": []}
            for k in ["image", "gif", "audio", "movie", "svg", "flash", "html", "file", "folder", "data"]
        }
    }

def build_timer(target_id, is_sync_player=True):
    """다음 페이지로 넘어가는 타이머 위젯 생성"""
    timer_id = gen_id("to")
    return {
        "id": timer_id, "label": "", "viewType": "WidgetTimer", "__extra__": {},
        "left": 800 if is_sync_player else -60, "top": 740 if is_sync_player else 0,
        "width": 60, "height": 60, "rotate": 0,
        "properties": {
            "visible": True, "opacity": 100, "interval": 0, "count": 1, "label": "",
            "draggable": False, "intervalOffset": 0, "waitForInterval": 1500,
            "autostart": not is_sync_player, "repeat": False
        },
        "style": {},
        "actions": [
            {
                "name": "WidgetAction", "event": {"type": "timer"},
                "properties": {"widgetId": target_id, "widgetActionName": "stop" if is_sync_player else "play"}
            },
            {
                "name": "GotoPage", "event": {"type": "timer"},
                "properties": {"selectPage": "nextPage"}
            }
        ],
        "animations": [], "responsives": [], "initState": None,
        "states": [
            {"value": "Start", "label": "시작", "readonly": True},
            {"value": "Stop", "label": "정지", "readonly": True}
        ],
        "fitWidth": False, "locked": False, "hidden": False, "widget": True,
        "name": "타이머", "contents": [], "children": []
    }

def build_sync_text(text_raw, top_pos, font_size="28px"):
    """오디오 싱크 텍스트 위젯 생성"""
    # 엑셀의 줄바꿈(\n)을 <p> 태그로 변환
    html_lines = "".join([f'<p style="margin:0;">{line.strip()}</p>' for line in str(text_raw).split('\n') if line.strip()])
    
    return {
        "id": gen_id("to"), "label": "", "viewType": "WidgetAudioSyncText", "__extra__": {},
        "left": 40, "top": top_pos, "width": 660, "height": 86, "rotate": 0,
        "properties": {
            "visible": True, "opacity": 100, "wordBreak": "normal", "label": "",
            "draggable": False, "widgetUserProperty": [], "asset": gen_id("as") # 임의의 오디오 싱크 에셋 연결
        },
        "style": {
            "fontFamily": "poppinsM", "fontSize": font_size, "letterSpacing": "0.3px",
            "lineHeight": "1.5", "textShadow": "", "boxShadow": "", "color": "#000000"
        },
        "actions": [], "animations": [], "responsives": [], "initState": None, "states": [],
        "fitWidth": False, "locked": False, "hidden": False, "widget": True,
        "name": "오디오 싱크텍스트",
        "htmlText": html_lines,
        "contents": [], "boost": True, "children": []
    }

def process_page(row, project_json):
    """엑셀 행 데이터를 받아 단일 JSON 페이지 구조로 조립"""
    bg_img = str(row.get('BgImage', ''))
    audio_file = str(row.get('AudioFile', ''))
    page_type = str(row.get('Type', ''))
    texts = str(row.get('Texts', ''))

    page_id = gen_page_id()
    
    # 1. 배경 설정
    img_url = f"{{ASSETS_DIR}}page-images/{bg_img}" if pd.notna(row.get('BgImage')) else ""
    thumb_url = img_url.replace(".jpg", "-thumb.jpg").replace("page-images/", "page-images/thumbs/") if img_url else ""

    page_obj = {
        "id": page_id, "title": "이름 없음", "label": "", "width": 800, "height": 800,
        "background": {"url": img_url, "width": 1600, "height": 1600} if img_url else {},
        "thumbnail": {"url": thumb_url, "width": 250, "height": 250, "src": ""} if thumb_url else {},
        "config": {}, "properties": {}, "_doc": {}, "actions": [], "contents": [], "subpages": []
    }

    # 2. 오디오 에셋 등록
    audio_asset_id = None
    if audio_file and audio_file != 'nan':
        audio_asset_id = gen_id("as")
        project_json["assets"]["audio"]["assets"].append({
            "url": f"{{ASSETS_DIR}}{audio_file}",
            "name": audio_file.split('.')[0],
            "ref": 0,
            "fid": project_json["assets"]["audio"]["folders"][0]["id"],
            "id": audio_asset_id
        })

    # 3. 위젯(Contents) 조립
    if page_type == 'Audio' and audio_asset_id:
        audio_player_id = gen_id("to")
        # 일반 오디오 위젯
        audio_widget = {
            "id": audio_player_id, "label": 1, "viewType": "WidgetAudio", "__extra__": {},
            "left": -60, "top": 0, "width": 60, "height": 60, "rotate": 0,
            "properties": {"visible": True, "opacity": 100, "actionOnActiveAudioChange": "stop", "label": "", "draggable": False, "asset": audio_asset_id, "autoplay": True, "actionOnActiveMediaChange": "stop"},
            "style": {}, "actions": [], "animations": [], "responsives": [], "initState": None,
            "states": [{"value": "Play", "label": "재생", "readonly": True, "checkable": False}, {"value": "Stop", "label": "정지", "readonly": True, "checkable": False}, {"value": "End", "label": "완료", "readonly": True, "checkable": False}],
            "fitWidth": False, "locked": False, "hidden": False, "widget": True, "name": "오디오", "contents": [], "children": []
        }
        timer_widget = build_timer(audio_player_id, is_sync_player=False)
        
        # Audio 재생이 끝나면 타이머를 시작시키는 액션 추가
        audio_widget["actions"].append({
            "name": "WidgetAction", "event": {"type": "changeState"},
            "properties": {"widgetId": timer_widget["id"], "widgetActionName": "start", "states": [{"target": audio_player_id, "value": "End"}], "condition": "if"}
        })
        page_obj["contents"] = [audio_widget, timer_widget]

    elif page_type == 'SyncPlayer' and audio_asset_id:
        player_id = gen_id("to")
        
        # 컨트롤바
        control_bar = {
            "id": gen_id("to"), "label": "", "viewType": "WidgetAudioSyncPlayerControlBar", "__extra__": {},
            "left": 0, "top": 0, "width": None, "height": None, "rotate": 0,
            "properties": {"visible": True, "opacity": 100, "label": "", "draggable": False},
            "style": {}, "actions": [], "animations": [], "responsives": [], "initState": None, "states": [],
            "fitWidth": False, "locked": False, "hidden": False, "children": []
        }

        # 텍스트 위젯들 생성 (|| 로 구분된 문자열 기준)
        text_widgets = [control_bar]
        if texts and texts != 'nan':
            text_blocks = texts.split('||')
            current_top = 60
            for block in text_blocks:
                text_widgets.append(build_sync_text(block, current_top))
                current_top += 40 # 다음 텍스트 박스 위치 간격
                
        timer_widget = build_timer(player_id, is_sync_player=True)

        # 싱크 플레이어 위젯
        sync_player = {
            "id": player_id, "label": 1, "viewType": "WidgetAudioSyncPlayer", "__extra__": {},
            "left": 33, "top": 0, "width": 734, "height": 800, "rotate": 0,
            "properties": {
                "visible": True, "opacity": 100, "theme": "dark", "alert": "none", "activeFontScale": 100,
                "activeFontBold": False, "wheelSpeed": 10, "activeControlbar": False, "label": "",
                "draggable": False, "fixed": False, "autoHeightFit": False, "controlbarBgColor": "", "controlbarIconColor": "",
                "widgetUserProperty": [], "activeFontColor": "#ff00ff", "actionOnActiveAudioChange": "stop", "actionOnActiveMediaChange": "stop"
            },
            "style": {},
            "actions": [{
                "name": "WidgetAction", "event": {"type": "changeState"},
                "properties": {"widgetId": timer_widget["id"], "widgetActionName": "start", "states": [{"target": player_id, "value": "End"}], "condition": "if"}
            }],
            "animations": [], "responsives": [], "initState": None,
            "states": [{"value": "Play", "label": "재생", "readonly": True, "checkable": False}, {"value": "Stop", "label": "정지", "readonly": True, "checkable": False}, {"value": "End", "label": "완료", "readonly": True, "checkable": False}],
            "fitWidth": False, "locked": False, "hidden": False, "container": True, "background": {"color": ""},
            "contents": text_widgets,
            "widget": True, "name": "오디오 싱크플레이어", "thumbnail": {"width": 70, "height": 100, "src": ""},
            "config": {}, "_doc": {}, "subpages": [], "children": []
        }
        page_obj["contents"] = [sync_player, timer_widget]

    return page_obj

def main():
    try:
        # data.xlsx 파일 읽기
        df = pd.read_excel("data.xlsx")
    except FileNotFoundError:
        print("❌ 'data.xlsx' 파일을 찾을 수 없습니다. 엑셀 파일을 먼저 생성해주세요.")
        return

    # 기본 뼈대 JSON 준비
    final_json = create_base_json()

    # 각 행 데이터를 페이지로 변환
    for _, row in df.iterrows():
        page = process_page(row, final_json)
        final_json["pages"].append(page)

    # 완성된 JSON 저장
    with open("output.json", "w", encoding="utf-8") as f:
        json.dump(final_json, f, ensure_ascii=False, indent=3)
        
    print(f"✅ 변환 완료! 총 {len(df)}장의 페이지가 포함된 'output.json' 파일이 생성되었습니다.")

if __name__ == "__main__":
    main()