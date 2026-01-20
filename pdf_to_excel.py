"""
PDF to Excel Converter with PDF Editor
PDF 파일의 표(테이블)를 Excel 파일로 변환하고, PDF 편집 기능을 제공합니다.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser, simpledialog
import pdfplumber
import pandas as pd
from pathlib import Path
import threading
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import io


class PDFToExcelConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Excel Converter & Editor")
        self.root.geometry("1400x850")
        self.root.minsize(1000, 700)
        self.root.resizable(True, True)

        self.pdf_files = []
        self.current_pdf = None
        self.current_pdf_path = None
        self.current_page = 0
        self.total_pages = 0
        self.pdf_images = []
        self.zoom_level = 1.0

        # 편집 도구 상태
        self.current_tool = "select"
        self.current_color = (1, 0, 0)  # RGB (0-1 범위)
        self.fill_color = None
        self.opacity = 1.0
        self.line_width = 2

        # 마우스 드래그 상태
        self.drag_start = None
        self.temp_shape = None

        # 선택된 주석
        self.selected_annot = None
        self.selected_annot_xref = None

        # 이동 관련 상태
        self.is_moving = False
        self.moving_annot_rect = None
        self.move_preview = None

        # 수정된 페이지 추적
        self.modified = False

        # 보기 모드: "page" 또는 "continuous"
        self.view_mode = "page"
        self.page_images = []  # 연속 보기용 이미지 캐시
        self.page_positions = []  # 각 페이지의 Y 위치
        self.continuous_image = None  # 연속 보기 이미지 참조 유지

        # 썸네일 이미지 캐시
        self.thumbnail_images = []

        # 스크롤 경계 감지용
        self.scroll_at_top = True
        self.scroll_at_bottom = False

        self.setup_ui()
        self.bind_shortcuts()

    def setup_ui(self):
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 상단 버튼 바
        top_btn_frame = ttk.Frame(main_frame)
        top_btn_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(top_btn_frame, text="PDF 파일 선택", command=self.select_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(top_btn_frame, text="목록 지우기", command=self.clear_list).pack(side=tk.LEFT)

        # 오른쪽 버튼들
        ttk.Button(top_btn_frame, text="Excel로 변환", command=self.start_conversion).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(top_btn_frame, text="PDF 저장", command=self.save_pdf).pack(side=tk.RIGHT, padx=(5, 0))

        # 3단 분할
        paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # 왼쪽 패널 (탭 구조: 파일 / 썸네일)
        left_frame = ttk.Frame(paned, width=200)
        paned.add(left_frame, weight=1)

        self.setup_left_panel(left_frame)

        # 중앙 패널 (PDF 뷰어)
        center_frame = ttk.Frame(paned)
        paned.add(center_frame, weight=4)

        # 도구 바
        self.setup_toolbar(center_frame)

        # 보기 설정 바
        self.setup_view_settings(center_frame)

        # PDF 캔버스
        viewer_frame = ttk.Frame(center_frame)
        viewer_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(viewer_frame, bg='gray', cursor='arrow')
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 스크롤바
        self.v_scroll = ttk.Scrollbar(viewer_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.config(yscrollcommand=self.on_scroll_changed)

        h_scroll = ttk.Scrollbar(center_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        h_scroll.pack(fill=tk.X)
        self.canvas.config(xscrollcommand=h_scroll.set)

        # 하단 네비게이션 바
        self.setup_navigation_bar(center_frame)

        # 오른쪽 패널 (탭 구조: 주석 / 옵션)
        right_frame = ttk.Frame(paned, width=250)
        paned.add(right_frame, weight=1)

        self.setup_right_panel(right_frame)

        # 마우스 이벤트 바인딩
        self.canvas.bind('<MouseWheel>', self.on_mousewheel)
        self.canvas.bind('<Button-1>', self.on_mouse_press)
        self.canvas.bind('<B1-Motion>', self.on_mouse_drag)
        self.canvas.bind('<ButtonRelease-1>', self.on_mouse_release)

    def setup_left_panel(self, parent):
        """왼쪽 패널 설정 (탭: 파일 / 썸네일)"""
        self.left_notebook = ttk.Notebook(parent)
        self.left_notebook.pack(fill=tk.BOTH, expand=True)

        # 파일 탭
        file_tab = ttk.Frame(self.left_notebook)
        self.left_notebook.add(file_tab, text="파일")

        # 파일 목록
        list_frame = ttk.LabelFrame(file_tab, text="선택된 PDF 파일", padding="5")
        list_frame.pack(fill=tk.BOTH, expand=True)

        self.file_listbox = tk.Listbox(list_frame, selectmode=tk.SINGLE)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.file_listbox.bind('<<ListboxSelect>>', self.on_file_select)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.config(yscrollcommand=scrollbar.set)

        # 썸네일 탭
        thumbnail_tab = ttk.Frame(self.left_notebook)
        self.left_notebook.add(thumbnail_tab, text="썸네일")

        # 썸네일 캔버스 (스크롤 가능)
        thumb_frame = ttk.Frame(thumbnail_tab)
        thumb_frame.pack(fill=tk.BOTH, expand=True)

        self.thumb_canvas = tk.Canvas(thumb_frame, bg='#f0f0f0', width=150)
        self.thumb_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        thumb_scroll = ttk.Scrollbar(thumb_frame, orient=tk.VERTICAL, command=self.thumb_canvas.yview)
        thumb_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.thumb_canvas.config(yscrollcommand=thumb_scroll.set)

        # 썸네일 내부 프레임
        self.thumb_inner_frame = ttk.Frame(self.thumb_canvas)
        self.thumb_canvas.create_window((0, 0), window=self.thumb_inner_frame, anchor=tk.NW)

        self.thumb_inner_frame.bind('<Configure>',
            lambda e: self.thumb_canvas.config(scrollregion=self.thumb_canvas.bbox("all")))

        # 썸네일 캔버스 마우스휠 바인딩
        self.thumb_canvas.bind('<MouseWheel>', lambda e: self.thumb_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

    def setup_right_panel(self, parent):
        """오른쪽 패널 설정 (탭: 주석 / 옵션)"""
        self.right_notebook = ttk.Notebook(parent)
        self.right_notebook.pack(fill=tk.BOTH, expand=True)

        # 주석 탭
        annot_tab = ttk.Frame(self.right_notebook)
        self.right_notebook.add(annot_tab, text="주석")

        self.setup_annotation_panel(annot_tab)

        # 옵션 탭
        option_tab = ttk.Frame(self.right_notebook)
        self.right_notebook.add(option_tab, text="옵션")

        self.setup_option_panel(option_tab)

    def setup_option_panel(self, parent):
        """옵션 패널 설정"""
        # 옵션 프레임
        option_frame = ttk.LabelFrame(parent, text="변환 옵션", padding="5")
        option_frame.pack(fill=tk.X, pady=(0, 10))

        self.all_pages_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(option_frame, text="모든 페이지 변환", variable=self.all_pages_var).pack(anchor=tk.W)

        self.merge_tables_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(option_frame, text="모든 표를 하나의 시트로 합치기", variable=self.merge_tables_var).pack(anchor=tk.W)

        # 진행 상태
        progress_frame = ttk.LabelFrame(parent, text="진행 상태", padding="5")
        progress_frame.pack(fill=tk.X, pady=(0, 10))

        self.progress_var = tk.StringVar(value="대기 중...")
        ttk.Label(progress_frame, textvariable=self.progress_var).pack(anchor=tk.W)

        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=(5, 0))

        # 선택 삭제 버튼
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(btn_frame, text="선택 삭제", command=self.delete_selected_annotation).pack(fill=tk.X, pady=2)

    def setup_view_settings(self, parent):
        """보기 설정 바"""
        view_frame = ttk.LabelFrame(parent, text="보기 설정", padding="5")
        view_frame.pack(fill=tk.X, pady=(5, 5))

        # 보기 모드 라디오 버튼
        mode_frame = ttk.Frame(view_frame)
        mode_frame.pack(side=tk.LEFT, padx=(0, 20))

        self.view_mode_var = tk.StringVar(value="page")
        ttk.Radiobutton(mode_frame, text="페이지", variable=self.view_mode_var,
                        value="page", command=self.on_view_mode_change).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(mode_frame, text="연속", variable=self.view_mode_var,
                        value="continuous", command=self.on_view_mode_change).pack(side=tk.LEFT)

        # 확대율 콤보박스
        zoom_frame = ttk.Frame(view_frame)
        zoom_frame.pack(side=tk.LEFT)

        ttk.Label(zoom_frame, text="확대:").pack(side=tk.LEFT, padx=(0, 5))

        self.zoom_combo = ttk.Combobox(zoom_frame, width=8, state="readonly",
                                        values=["50%", "75%", "100%", "125%", "150%", "200%"])
        self.zoom_combo.set("100%")
        self.zoom_combo.pack(side=tk.LEFT)
        self.zoom_combo.bind('<<ComboboxSelected>>', self.on_zoom_combo_change)

    def setup_navigation_bar(self, parent):
        """하단 네비게이션 바"""
        nav_frame = ttk.Frame(parent)
        nav_frame.pack(fill=tk.X, pady=(5, 0))

        # 중앙 정렬을 위한 내부 프레임
        inner_nav = ttk.Frame(nav_frame)
        inner_nav.pack(anchor=tk.CENTER)

        # 이전 페이지 버튼
        ttk.Button(inner_nav, text="◀", width=3, command=self.prev_page).pack(side=tk.LEFT, padx=2)

        # 페이지 입력 스핀박스
        self.page_spinbox = ttk.Spinbox(inner_nav, from_=1, to=1, width=5,
                                         command=self.on_page_spinbox_change)
        self.page_spinbox.pack(side=tk.LEFT, padx=2)
        self.page_spinbox.bind('<Return>', self.on_page_spinbox_enter)

        # 총 페이지 라벨
        self.total_page_label = ttk.Label(inner_nav, text="/ 0")
        self.total_page_label.pack(side=tk.LEFT, padx=(0, 5))

        # 다음 페이지 버튼
        ttk.Button(inner_nav, text="▶", width=3, command=self.next_page).pack(side=tk.LEFT, padx=2)

        # 구분선
        ttk.Separator(inner_nav, orient=tk.VERTICAL).pack(side=tk.LEFT, padx=10, fill=tk.Y)

        # 줌 버튼
        ttk.Button(inner_nav, text="−", width=3, command=self.zoom_out).pack(side=tk.LEFT, padx=2)
        ttk.Button(inner_nav, text="+", width=3, command=self.zoom_in).pack(side=tk.LEFT, padx=2)

    def setup_toolbar(self, parent):
        """도구 바 설정"""
        toolbar_frame = ttk.LabelFrame(parent, text="편집 도구", padding="5")
        toolbar_frame.pack(fill=tk.X, pady=(0, 5))

        # 첫 번째 줄: 도구 버튼들
        tools_row1 = ttk.Frame(toolbar_frame)
        tools_row1.pack(fill=tk.X, pady=(0, 5))

        self.tool_buttons = {}
        tools = [
            ("select", "선택"),
            ("line", "직선"),
            ("circle", "원"),
            ("rect", "사각형"),
            ("highlight", "하이라이트"),
        ]
        for tool_id, tool_name in tools:
            btn = ttk.Button(tools_row1, text=tool_name, width=8,
                           command=lambda t=tool_id: self.set_tool(t))
            btn.pack(side=tk.LEFT, padx=2)
            self.tool_buttons[tool_id] = btn

        # 두 번째 줄: 추가 도구
        tools_row2 = ttk.Frame(toolbar_frame)
        tools_row2.pack(fill=tk.X, pady=(0, 5))

        tools2 = [
            ("note", "주석"),
            ("text", "텍스트"),
            ("redact", "Redact"),
        ]
        for tool_id, tool_name in tools2:
            btn = ttk.Button(tools_row2, text=tool_name, width=8,
                           command=lambda t=tool_id: self.set_tool(t))
            btn.pack(side=tk.LEFT, padx=2)
            self.tool_buttons[tool_id] = btn

        # 삭제 버튼 (별도)
        ttk.Separator(tools_row2, orient=tk.VERTICAL).pack(side=tk.LEFT, padx=5, fill=tk.Y)
        ttk.Button(tools_row2, text="선택 삭제", width=10,
                  command=self.delete_selected_annotation).pack(side=tk.LEFT, padx=2)

        # 세 번째 줄: 색상 및 옵션
        options_row = ttk.Frame(toolbar_frame)
        options_row.pack(fill=tk.X)

        # 선 색상
        ttk.Label(options_row, text="색상:").pack(side=tk.LEFT, padx=(0, 2))
        self.color_btn = tk.Button(options_row, width=3, bg='#FF0000',
                                   command=self.choose_color)
        self.color_btn.pack(side=tk.LEFT, padx=(0, 10))

        # 채우기 색상
        ttk.Label(options_row, text="채우기:").pack(side=tk.LEFT, padx=(0, 2))
        self.fill_btn = tk.Button(options_row, width=3, bg='white',
                                  command=self.choose_fill_color)
        self.fill_btn.pack(side=tk.LEFT, padx=(0, 10))

        # 채우기 없음 체크박스
        self.no_fill_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_row, text="투명", variable=self.no_fill_var,
                       command=self.toggle_fill).pack(side=tk.LEFT, padx=(0, 10))

        # 투명도
        ttk.Label(options_row, text="투명도:").pack(side=tk.LEFT, padx=(0, 2))
        self.opacity_scale = ttk.Scale(options_row, from_=0.1, to=1.0, length=100,
                                       orient=tk.HORIZONTAL, command=self.set_opacity)
        self.opacity_scale.set(1.0)
        self.opacity_scale.pack(side=tk.LEFT, padx=(0, 10))

        # 선 두께
        ttk.Label(options_row, text="두께:").pack(side=tk.LEFT, padx=(0, 2))
        self.width_spinbox = ttk.Spinbox(options_row, from_=1, to=10, width=5,
                                         command=self.set_line_width)
        self.width_spinbox.set(2)
        self.width_spinbox.pack(side=tk.LEFT)

        # 현재 도구 표시
        self.tool_label = ttk.Label(toolbar_frame, text="현재 도구: 선택")
        self.tool_label.pack(anchor=tk.W, pady=(5, 0))

    def setup_annotation_panel(self, parent):
        """주석 패널 설정"""
        # 주석 목록
        annot_frame = ttk.LabelFrame(parent, text="주석 목록", padding="5")
        annot_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 주석 리스트
        self.annot_listbox = tk.Listbox(annot_frame, selectmode=tk.SINGLE)
        self.annot_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.annot_listbox.bind('<<ListboxSelect>>', self.on_annot_select)

        annot_scroll = ttk.Scrollbar(annot_frame, orient=tk.VERTICAL, command=self.annot_listbox.yview)
        annot_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.annot_listbox.config(yscrollcommand=annot_scroll.set)

        # 주석 상세 정보
        detail_frame = ttk.LabelFrame(parent, text="주석 상세", padding="5")
        detail_frame.pack(fill=tk.X, pady=(0, 10))

        self.annot_detail_var = tk.StringVar(value="선택된 주석 없음")
        self.annot_detail_label = ttk.Label(detail_frame, textvariable=self.annot_detail_var,
                                            wraplength=200, justify=tk.LEFT)
        self.annot_detail_label.pack(anchor=tk.W)

        # 주석 관리 버튼
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X)

        ttk.Button(btn_frame, text="선택 삭제", command=self.delete_selected_annotation).pack(fill=tk.X, pady=2)
        ttk.Button(btn_frame, text="모두 삭제", command=self.delete_all_annotations).pack(fill=tk.X, pady=2)
        ttk.Button(btn_frame, text="새로고침", command=self.refresh_annotation_list).pack(fill=tk.X, pady=2)

    def bind_shortcuts(self):
        """키보드 단축키 바인딩"""
        # 줌 단축키
        self.root.bind('<Control-plus>', lambda e: self.zoom_in())
        self.root.bind('<Control-equal>', lambda e: self.zoom_in())  # Shift 없이 + 키
        self.root.bind('<Control-minus>', lambda e: self.zoom_out())
        self.root.bind('<Control-0>', lambda e: self.reset_zoom())

        # 페이지 네비게이션
        self.root.bind('<Prior>', lambda e: self.prev_page())  # Page Up
        self.root.bind('<Next>', lambda e: self.next_page())  # Page Down
        self.root.bind('<Home>', lambda e: self.go_to_first_page())
        self.root.bind('<End>', lambda e: self.go_to_last_page())

    def on_view_mode_change(self):
        """보기 모드 변경"""
        self.view_mode = self.view_mode_var.get()
        if self.current_pdf:
            if self.view_mode == "continuous":
                self.render_continuous()
            else:
                self.render_page()

    def on_zoom_combo_change(self, event=None):
        """확대율 콤보박스 변경"""
        zoom_text = self.zoom_combo.get()
        zoom_percent = int(zoom_text.replace('%', ''))
        self.zoom_level = zoom_percent / 100
        self.refresh_view()

    def update_zoom_combo(self):
        """확대율 콤보박스 업데이트"""
        percent = int(self.zoom_level * 100)
        zoom_text = f"{percent}%"
        # 정확한 값이 리스트에 있으면 선택, 아니면 가장 가까운 값
        if zoom_text in self.zoom_combo['values']:
            self.zoom_combo.set(zoom_text)
        else:
            self.zoom_combo.set(f"{percent}%")

    def on_page_spinbox_change(self):
        """페이지 스핀박스 값 변경"""
        try:
            page = int(self.page_spinbox.get()) - 1
            if 0 <= page < self.total_pages and page != self.current_page:
                self.go_to_page(page)
        except ValueError:
            pass

    def on_page_spinbox_enter(self, event=None):
        """페이지 스핀박스에서 Enter 키"""
        self.on_page_spinbox_change()

    def go_to_page(self, page):
        """특정 페이지로 이동"""
        if self.current_pdf is None:
            return
        if 0 <= page < self.total_pages:
            self.current_page = page
            self.selected_annot_xref = None
            self.annot_detail_var.set("선택된 주석 없음")

            if self.view_mode == "continuous":
                self.scroll_to_page_in_continuous(page)
            else:
                self.render_page()
                self.refresh_annotation_list()

            self.update_page_display()
            self.update_thumbnail_selection()

    def go_to_first_page(self):
        """첫 페이지로 이동"""
        self.go_to_page(0)

    def go_to_last_page(self):
        """마지막 페이지로 이동"""
        if self.current_pdf:
            self.go_to_page(self.total_pages - 1)

    def reset_zoom(self):
        """100%로 리셋"""
        self.zoom_level = 1.0
        self.update_zoom_combo()
        self.refresh_view()

    def refresh_view(self):
        """현재 보기 모드에 따라 새로고침"""
        if self.current_pdf is None:
            return
        if self.view_mode == "continuous":
            self.render_continuous()
        else:
            self.render_page()

    def on_scroll_changed(self, *args):
        """스크롤 위치 변경 시 호출"""
        self.v_scroll.set(*args)

        if self.view_mode == "continuous" and self.current_pdf:
            # 연속 보기에서 현재 보이는 페이지 업데이트
            self.update_current_page_from_scroll()

    def update_current_page_from_scroll(self):
        """연속 보기 모드에서 스크롤 위치에 따른 현재 페이지 업데이트"""
        if not self.page_positions:
            return

        # 현재 스크롤 위치
        scroll_top = self.canvas.canvasy(0)
        canvas_height = self.canvas.winfo_height()
        center_y = scroll_top + canvas_height / 2

        # 현재 보이는 페이지 찾기
        for i, pos in enumerate(self.page_positions):
            if i + 1 < len(self.page_positions):
                if pos <= center_y < self.page_positions[i + 1]:
                    if self.current_page != i:
                        self.current_page = i
                        self.update_page_display()
                        self.update_thumbnail_selection()
                        self.refresh_annotation_list()
                    return
            else:
                if center_y >= pos:
                    if self.current_page != i:
                        self.current_page = i
                        self.update_page_display()
                        self.update_thumbnail_selection()
                        self.refresh_annotation_list()
                    return

    def scroll_to_page_in_continuous(self, page):
        """연속 보기에서 특정 페이지로 스크롤"""
        if not self.page_positions or page >= len(self.page_positions):
            return

        y_pos = self.page_positions[page]
        # 전체 높이 대비 비율 계산
        total_height = self.canvas.bbox("all")
        if total_height:
            fraction = y_pos / total_height[3]
            self.canvas.yview_moveto(fraction)

    def update_page_display(self):
        """페이지 표시 업데이트"""
        self.page_spinbox.delete(0, tk.END)
        self.page_spinbox.insert(0, str(self.current_page + 1))
        self.total_page_label.config(text=f"/ {self.total_pages}")

    def render_thumbnails(self):
        """썸네일 렌더링"""
        # 기존 썸네일 제거
        for widget in self.thumb_inner_frame.winfo_children():
            widget.destroy()
        self.thumbnail_images = []

        if self.current_pdf is None:
            return

        thumb_width = 100
        for i in range(self.total_pages):
            page = self.current_pdf[i]
            # 썸네일 크기 계산
            page_rect = page.rect
            scale = thumb_width / page_rect.width
            mat = fitz.Matrix(scale, scale)
            pix = page.get_pixmap(matrix=mat)

            img_data = pix.tobytes("ppm")
            img = Image.open(io.BytesIO(img_data))
            photo = ImageTk.PhotoImage(img)
            self.thumbnail_images.append(photo)

            # 썸네일 프레임
            thumb_frame = ttk.Frame(self.thumb_inner_frame)
            thumb_frame.pack(pady=5, padx=5)

            # 페이지 번호 라벨
            label = ttk.Label(thumb_frame, text=f"[{i + 1}]")
            label.pack()

            # 썸네일 라벨
            thumb_label = ttk.Label(thumb_frame, image=photo, relief="solid", borderwidth=1)
            thumb_label.pack()
            thumb_label.bind('<Button-1>', lambda e, p=i: self.on_thumbnail_click(p))

            # 현재 페이지 표시
            if i == self.current_page:
                thumb_label.config(borderwidth=3)

    def update_thumbnail_selection(self):
        """현재 페이지에 해당하는 썸네일 강조"""
        for i, widget in enumerate(self.thumb_inner_frame.winfo_children()):
            # 각 썸네일 프레임의 두 번째 자식 (이미지 라벨)
            children = widget.winfo_children()
            if len(children) >= 2:
                thumb_label = children[1]
                if i == self.current_page:
                    thumb_label.config(borderwidth=3)
                else:
                    thumb_label.config(borderwidth=1)

    def on_thumbnail_click(self, page):
        """썸네일 클릭 시 해당 페이지로 이동"""
        self.go_to_page(page)
        # 썸네일 탭에서 클릭해도 유지 (탭 전환 안함)

    def refresh_annotation_list(self):
        """주석 목록 새로고침"""
        self.annot_listbox.delete(0, tk.END)

        if self.current_pdf is None:
            return

        page = self.current_pdf[self.current_page]
        annot_types = {
            0: "텍스트", 1: "링크", 2: "자유텍스트", 3: "직선",
            4: "사각형", 5: "원", 6: "다각형", 7: "폴리라인",
            8: "하이라이트", 9: "밑줄", 10: "취소선", 11: "스탬프",
            12: "잉크", 13: "팝업", 14: "파일첨부", 15: "소리",
            16: "영화", 17: "위젯", 18: "스크린", 19: "프린터마크",
            20: "트랩넷", 21: "워터마크", 22: "3D", 23: "리다이렉트"
        }

        for i, annot in enumerate(page.annots()):
            annot_type = annot_types.get(annot.type[0], f"타입{annot.type[0]}")
            info = annot.info
            content = info.get("content", "")[:30] if info.get("content") else ""

            display_text = f"[{i+1}] {annot_type}"
            if content:
                display_text += f": {content}"

            self.annot_listbox.insert(tk.END, display_text)

    def on_annot_select(self, event):
        """주석 목록에서 선택"""
        selection = self.annot_listbox.curselection()
        if not selection or self.current_pdf is None:
            return

        idx = selection[0]
        page = self.current_pdf[self.current_page]

        annots = list(page.annots())
        if idx < len(annots):
            annot = annots[idx]
            self.selected_annot_xref = annot.xref

            # 상세 정보 표시
            info = annot.info
            detail = f"타입: {annot.type[1]}\n"
            detail += f"위치: ({annot.rect.x0:.0f}, {annot.rect.y0:.0f})\n"
            detail += f"크기: {annot.rect.width:.0f} x {annot.rect.height:.0f}\n"
            if info.get("content"):
                detail += f"내용: {info['content'][:100]}"

            self.annot_detail_var.set(detail)

            # 캔버스에서 해당 주석 하이라이트
            self.highlight_selected_annot(annot.rect)

    def highlight_selected_annot(self, rect):
        """선택된 주석 하이라이트"""
        self.render_page()

        # 선택 표시 그리기
        zoom = self.zoom_level * 1.5
        x0, y0 = rect.x0 * zoom, rect.y0 * zoom
        x1, y1 = rect.x1 * zoom, rect.y1 * zoom

        self.canvas.create_rectangle(x0 - 3, y0 - 3, x1 + 3, y1 + 3,
                                     outline='blue', width=2, dash=(5, 3),
                                     tags="selection")

    def set_tool(self, tool):
        """도구 선택"""
        self.current_tool = tool
        tool_names = {
            "select": "선택 (클릭: 선택, 드래그: 이동)",
            "line": "직선",
            "circle": "원",
            "rect": "사각형",
            "highlight": "하이라이트",
            "note": "주석",
            "text": "텍스트",
            "redact": "Redact"
        }
        self.tool_label.config(text=f"현재 도구: {tool_names.get(tool, tool)}")

        # 커서 변경
        cursors = {
            "select": "arrow",
            "line": "crosshair",
            "circle": "crosshair",
            "rect": "crosshair",
            "highlight": "crosshair",
            "note": "hand2",
            "text": "xterm",
            "redact": "crosshair"
        }
        self.canvas.config(cursor=cursors.get(tool, "arrow"))

    def choose_color(self):
        """선 색상 선택"""
        color = colorchooser.askcolor(title="선 색상 선택")
        if color[1]:
            self.color_btn.config(bg=color[1])
            r, g, b = color[0]
            self.current_color = (r / 255, g / 255, b / 255)

    def choose_fill_color(self):
        """채우기 색상 선택"""
        color = colorchooser.askcolor(title="채우기 색상 선택")
        if color[1]:
            self.fill_btn.config(bg=color[1])
            r, g, b = color[0]
            self.fill_color = (r / 255, g / 255, b / 255)
            self.no_fill_var.set(False)

    def toggle_fill(self):
        """채우기 토글"""
        if self.no_fill_var.get():
            self.fill_color = None

    def set_opacity(self, value):
        """투명도 설정"""
        self.opacity = float(value)

    def set_line_width(self):
        """선 두께 설정"""
        try:
            self.line_width = int(self.width_spinbox.get())
        except ValueError:
            self.line_width = 2

    def canvas_to_pdf(self, x, y):
        """캔버스 좌표를 PDF 좌표로 변환"""
        canvas_x = self.canvas.canvasx(x)
        canvas_y = self.canvas.canvasy(y)

        zoom = self.zoom_level * 1.5
        pdf_x = canvas_x / zoom
        pdf_y = canvas_y / zoom
        return fitz.Point(pdf_x, pdf_y)

    def canvas_to_pdf_point(self, canvas_coords):
        """캔버스 좌표를 PDF Point로 변환"""
        x, y = canvas_coords
        zoom = self.zoom_level * 1.5
        pdf_x = x / zoom
        pdf_y = y / zoom
        return fitz.Point(pdf_x, pdf_y)

    def on_mouse_press(self, event):
        """마우스 버튼 누름"""
        if self.current_pdf is None:
            return

        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        self.drag_start = (x, y)

        if self.current_tool == "select":
            # 이미 선택된 주석 위를 클릭하면 이동 모드
            if self.selected_annot_xref is not None:
                if self.check_click_on_selected(event.x, event.y):
                    self.is_moving = True
                    self.start_move_annotation()
                    return
            # 아니면 새로 선택
            self.select_annotation_at(event.x, event.y)
        elif self.current_tool == "note":
            self.add_sticky_note(event.x, event.y)
        elif self.current_tool == "text":
            self.add_text_box(event.x, event.y)

    def on_mouse_drag(self, event):
        """마우스 드래그"""
        if self.current_pdf is None or self.drag_start is None:
            return

        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)

        # 이동 모드
        if self.current_tool == "select" and self.is_moving:
            self.preview_move_annotation(x, y)
        elif self.current_tool in ["line", "circle", "rect", "highlight", "redact"]:
            self.draw_temp_shape(x, y)

    def on_mouse_release(self, event):
        """마우스 버튼 놓음"""
        if self.current_pdf is None or self.drag_start is None:
            return

        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)

        # 이동 모드 완료
        if self.current_tool == "select" and self.is_moving:
            self.finish_move_annotation(x, y)
            self.is_moving = False
            self.drag_start = None
            return

        # 최소 크기 체크
        min_size = 5
        if abs(x - self.drag_start[0]) < min_size and abs(y - self.drag_start[1]) < min_size:
            if self.current_tool not in ["select", "note", "text"]:
                if self.temp_shape:
                    self.canvas.delete(self.temp_shape)
                    self.temp_shape = None
                self.drag_start = None
                return

        if self.current_tool == "line":
            self.draw_line_annot(self.drag_start, (x, y))
        elif self.current_tool == "circle":
            self.draw_circle_annot(self.drag_start, (x, y))
        elif self.current_tool == "rect":
            self.draw_rect_annot(self.drag_start, (x, y))
        elif self.current_tool == "highlight":
            self.add_highlight(self.drag_start, (x, y))
        elif self.current_tool == "redact":
            self.add_redaction(self.drag_start, (x, y))

        # 임시 도형 삭제
        if self.temp_shape:
            self.canvas.delete(self.temp_shape)
            self.temp_shape = None

        self.drag_start = None

    def draw_temp_shape(self, x, y):
        """임시 도형 그리기 (미리보기)"""
        if self.temp_shape:
            self.canvas.delete(self.temp_shape)

        x1, y1 = self.drag_start
        color = '#%02x%02x%02x' % (int(self.current_color[0] * 255),
                                    int(self.current_color[1] * 255),
                                    int(self.current_color[2] * 255))

        if self.current_tool == "line":
            self.temp_shape = self.canvas.create_line(x1, y1, x, y,
                                                       fill=color, width=self.line_width,
                                                       dash=(4, 4))
        elif self.current_tool == "circle":
            # 사각형 내접 원
            self.temp_shape = self.canvas.create_oval(x1, y1, x, y,
                                                       outline=color, width=self.line_width,
                                                       dash=(4, 4))
        elif self.current_tool == "rect":
            self.temp_shape = self.canvas.create_rectangle(x1, y1, x, y,
                                                            outline=color, width=self.line_width,
                                                            dash=(4, 4))
        elif self.current_tool == "highlight":
            self.temp_shape = self.canvas.create_rectangle(x1, y1, x, y,
                                                            outline='yellow', fill='yellow',
                                                            stipple='gray50',
                                                            width=1)
        elif self.current_tool == "redact":
            self.temp_shape = self.canvas.create_rectangle(x1, y1, x, y,
                                                            outline='red', fill='black',
                                                            stipple='gray50',
                                                            width=2, dash=(4, 4))

    def draw_line_annot(self, start, end):
        """직선 주석 그리기"""
        page = self.current_pdf[self.current_page]

        p1 = self.canvas_to_pdf_point(start)
        p2 = self.canvas_to_pdf_point(end)

        annot = page.add_line_annot(p1, p2)
        annot.set_colors(stroke=self.current_color)
        annot.set_border(width=self.line_width)
        annot.set_opacity(self.opacity)
        annot.update()

        self.modified = True
        self.render_page()
        self.refresh_annotation_list()

    def draw_circle_annot(self, start, end):
        """원 주석 그리기"""
        page = self.current_pdf[self.current_page]

        p1 = self.canvas_to_pdf_point(start)
        p2 = self.canvas_to_pdf_point(end)

        rect = fitz.Rect(p1.x, p1.y, p2.x, p2.y)
        rect.normalize()

        annot = page.add_circle_annot(rect)
        annot.set_colors(stroke=self.current_color)
        if not self.no_fill_var.get() and self.fill_color:
            annot.set_colors(stroke=self.current_color, fill=self.fill_color)
        annot.set_border(width=self.line_width)
        annot.set_opacity(self.opacity)
        annot.update()

        self.modified = True
        self.render_page()
        self.refresh_annotation_list()

    def draw_rect_annot(self, start, end):
        """사각형 주석 그리기"""
        page = self.current_pdf[self.current_page]

        p1 = self.canvas_to_pdf_point(start)
        p2 = self.canvas_to_pdf_point(end)

        rect = fitz.Rect(p1.x, p1.y, p2.x, p2.y)
        rect.normalize()

        annot = page.add_rect_annot(rect)
        annot.set_colors(stroke=self.current_color)
        if not self.no_fill_var.get() and self.fill_color:
            annot.set_colors(stroke=self.current_color, fill=self.fill_color)
        annot.set_border(width=self.line_width)
        annot.set_opacity(self.opacity)
        annot.update()

        self.modified = True
        self.render_page()
        self.refresh_annotation_list()

    def add_highlight(self, start, end):
        """하이라이트 주석 추가"""
        page = self.current_pdf[self.current_page]

        p1 = self.canvas_to_pdf_point(start)
        p2 = self.canvas_to_pdf_point(end)

        rect = fitz.Rect(p1.x, p1.y, p2.x, p2.y)
        rect.normalize()

        # Quad로 변환 (highlight는 quad 필요)
        quad = rect.quad

        annot = page.add_highlight_annot(quad)
        if annot:
            # 노란색 기본, 또는 선택한 색상
            annot.set_colors(stroke=self.current_color)
            annot.set_opacity(self.opacity)
            annot.update()

        self.modified = True
        self.render_page()
        self.refresh_annotation_list()

    def add_sticky_note(self, x, y):
        """스티키 노트 추가"""
        if self.current_pdf is None:
            return

        text = simpledialog.askstring("주석", "주석 내용을 입력하세요:")
        if text:
            page = self.current_pdf[self.current_page]
            point = self.canvas_to_pdf(x, y)

            annot = page.add_text_annot(point, text, icon="Note")
            if annot:
                annot.set_colors(stroke=self.current_color)
                annot.set_opacity(self.opacity)
                annot.update()

            self.modified = True
            self.render_page()
            self.refresh_annotation_list()

    def add_text_box(self, x, y):
        """텍스트 박스 추가"""
        if self.current_pdf is None:
            return

        text = simpledialog.askstring("텍스트", "텍스트 내용을 입력하세요:")
        if not text:
            return

        font_size = simpledialog.askinteger("폰트 크기", "폰트 크기:", initialvalue=12, minvalue=6, maxvalue=72)
        if font_size is None:
            font_size = 12

        page = self.current_pdf[self.current_page]
        point = self.canvas_to_pdf(x, y)

        # 텍스트 길이에 따른 사각형 크기 계산
        width = len(text) * font_size * 0.6 + 20
        height = font_size + 10

        rect = fitz.Rect(point.x, point.y, point.x + width, point.y + height)

        # 채우기 색상 결정
        fill_col = (1, 1, 1)  # 기본 흰색
        if not self.no_fill_var.get() and self.fill_color:
            fill_col = self.fill_color

        annot = page.add_freetext_annot(
            rect,
            text,
            fontsize=font_size,
            fontname="helv",
            text_color=self.current_color,
            fill_color=fill_col,
        )
        if annot:
            annot.set_opacity(self.opacity)
            annot.update()

        self.modified = True
        self.render_page()
        self.refresh_annotation_list()

    def select_annotation_at(self, x, y):
        """클릭 위치의 주석 선택"""
        if self.current_pdf is None:
            return

        page = self.current_pdf[self.current_page]
        point = self.canvas_to_pdf(x, y)

        # 모든 주석 확인
        for i, annot in enumerate(page.annots()):
            if annot.rect.contains(point):
                self.selected_annot_xref = annot.xref

                # 리스트박스에서도 선택
                self.annot_listbox.selection_clear(0, tk.END)
                self.annot_listbox.selection_set(i)
                self.annot_listbox.see(i)

                # 상세 정보 표시
                info = annot.info
                detail = f"타입: {annot.type[1]}\n"
                detail += f"위치: ({annot.rect.x0:.0f}, {annot.rect.y0:.0f})\n"
                detail += f"크기: {annot.rect.width:.0f} x {annot.rect.height:.0f}\n"
                if info.get("content"):
                    detail += f"내용: {info['content'][:100]}"
                self.annot_detail_var.set(detail)

                # 선택 표시
                self.highlight_selected_annot(annot.rect)
                return

        # 선택된 것 없음
        self.selected_annot_xref = None
        self.annot_detail_var.set("선택된 주석 없음")
        self.canvas.delete("selection")

    def check_click_on_selected(self, x, y):
        """선택된 주석 위를 클릭했는지 확인"""
        if self.current_pdf is None or self.selected_annot_xref is None:
            return False

        page = self.current_pdf[self.current_page]
        point = self.canvas_to_pdf(x, y)

        for annot in page.annots():
            if annot.xref == self.selected_annot_xref:
                return annot.rect.contains(point)
        return False

    def start_move_annotation(self):
        """이동 시작 - 원본 위치 저장"""
        if self.current_pdf is None or self.selected_annot_xref is None:
            return

        page = self.current_pdf[self.current_page]
        for annot in page.annots():
            if annot.xref == self.selected_annot_xref:
                self.moving_annot_rect = annot.rect
                # 커서 변경
                self.canvas.config(cursor='fleur')
                return

    def preview_move_annotation(self, x, y):
        """이동 미리보기"""
        if self.moving_annot_rect is None:
            return

        # 이전 미리보기 삭제
        if self.move_preview:
            self.canvas.delete(self.move_preview)

        # 이동 거리 계산
        dx = x - self.drag_start[0]
        dy = y - self.drag_start[1]

        # 캔버스 좌표로 변환된 새 위치
        zoom = self.zoom_level * 1.5
        x0 = self.moving_annot_rect.x0 * zoom + dx
        y0 = self.moving_annot_rect.y0 * zoom + dy
        x1 = self.moving_annot_rect.x1 * zoom + dx
        y1 = self.moving_annot_rect.y1 * zoom + dy

        # 미리보기 그리기
        self.move_preview = self.canvas.create_rectangle(
            x0, y0, x1, y1,
            outline='green', width=2, dash=(5, 3),
            tags="move_preview"
        )

    def finish_move_annotation(self, x, y):
        """이동 완료"""
        # 미리보기 삭제
        if self.move_preview:
            self.canvas.delete(self.move_preview)
            self.move_preview = None

        if self.moving_annot_rect is None or self.selected_annot_xref is None:
            self.canvas.config(cursor='arrow')
            return

        # 이동 거리 계산 (PDF 좌표계)
        zoom = self.zoom_level * 1.5
        dx = (x - self.drag_start[0]) / zoom
        dy = (y - self.drag_start[1]) / zoom

        # 최소 이동 거리 체크
        if abs(dx) < 1 and abs(dy) < 1:
            self.moving_annot_rect = None
            self.canvas.config(cursor='arrow')
            return

        page = self.current_pdf[self.current_page]

        for annot in page.annots():
            if annot.xref == self.selected_annot_xref:
                # 새 위치 계산
                old_rect = annot.rect
                new_rect = fitz.Rect(
                    old_rect.x0 + dx,
                    old_rect.y0 + dy,
                    old_rect.x1 + dx,
                    old_rect.y1 + dy
                )

                # 주석 이동 (타입에 따라 다르게 처리)
                annot_type = annot.type[0]

                # Line annotation (타입 3)
                if annot_type == 3:
                    # 직선의 경우 vertices 이동
                    vertices = annot.vertices
                    if vertices and len(vertices) >= 2:
                        new_vertices = [
                            fitz.Point(vertices[0].x + dx, vertices[0].y + dy),
                            fitz.Point(vertices[1].x + dx, vertices[1].y + dy)
                        ]
                        annot.set_vertices(new_vertices)
                else:
                    # 다른 주석은 rect 이동
                    annot.set_rect(new_rect)

                annot.update()
                self.modified = True
                break

        self.moving_annot_rect = None
        self.canvas.config(cursor='arrow')
        self.render_page()
        self.refresh_annotation_list()

        # 선택 상태 유지하며 하이라이트
        page = self.current_pdf[self.current_page]
        for annot in page.annots():
            if annot.xref == self.selected_annot_xref:
                self.highlight_selected_annot(annot.rect)
                break

    def delete_selected_annotation(self):
        """선택된 주석 삭제"""
        if self.current_pdf is None or self.selected_annot_xref is None:
            messagebox.showinfo("알림", "삭제할 주석을 선택하세요.")
            return

        page = self.current_pdf[self.current_page]

        for annot in page.annots():
            if annot.xref == self.selected_annot_xref:
                page.delete_annot(annot)
                self.selected_annot_xref = None
                self.modified = True
                self.render_page()
                self.refresh_annotation_list()
                self.annot_detail_var.set("선택된 주석 없음")
                return

        messagebox.showinfo("알림", "주석을 찾을 수 없습니다.")

    def delete_all_annotations(self):
        """현재 페이지의 모든 주석 삭제"""
        if self.current_pdf is None:
            return

        if not messagebox.askyesno("확인", "현재 페이지의 모든 주석을 삭제하시겠습니까?"):
            return

        page = self.current_pdf[self.current_page]
        annots = list(page.annots())

        for annot in annots:
            page.delete_annot(annot)

        self.selected_annot_xref = None
        self.modified = True
        self.render_page()
        self.refresh_annotation_list()
        self.annot_detail_var.set("선택된 주석 없음")

    def add_redaction(self, start, end):
        """Redaction 추가"""
        page = self.current_pdf[self.current_page]

        p1 = self.canvas_to_pdf_point(start)
        p2 = self.canvas_to_pdf_point(end)

        rect = fitz.Rect(p1.x, p1.y, p2.x, p2.y)
        rect.normalize()

        page.add_redact_annot(rect, fill=(0, 0, 0))

        if messagebox.askyesno("Redaction", "Redaction을 지금 적용하시겠습니까?\n(적용하면 해당 영역의 내용이 영구적으로 삭제됩니다)"):
            page.apply_redactions()

        self.modified = True
        self.render_page()
        self.refresh_annotation_list()

    def save_pdf(self):
        """편집된 PDF 저장"""
        if self.current_pdf is None:
            messagebox.showwarning("경고", "저장할 PDF가 없습니다.")
            return

        save_path = filedialog.asksaveasfilename(
            title="PDF 저장",
            defaultextension=".pdf",
            initialfile=f"{Path(self.current_pdf_path).stem}_edited.pdf" if self.current_pdf_path else "edited.pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )

        if save_path:
            try:
                self.current_pdf.save(save_path, garbage=4, deflate=True)
                self.modified = False
                messagebox.showinfo("완료", f"PDF가 저장되었습니다:\n{save_path}")
            except Exception as e:
                messagebox.showerror("오류", f"PDF 저장 실패: {str(e)}")

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="PDF 파일 선택",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        for file in files:
            if file not in self.pdf_files:
                self.pdf_files.append(file)
                self.file_listbox.insert(tk.END, Path(file).name)

        if self.pdf_files and self.file_listbox.curselection() == ():
            self.file_listbox.selection_set(0)
            self.load_pdf(self.pdf_files[0])

    def clear_list(self):
        if self.modified:
            if not messagebox.askyesno("경고", "저장하지 않은 변경사항이 있습니다. 계속하시겠습니까?"):
                return

        self.pdf_files = []
        self.file_listbox.delete(0, tk.END)
        self.progress_var.set("대기 중...")
        self.progress_bar['value'] = 0
        self.current_pdf = None
        self.current_pdf_path = None
        self.canvas.delete("all")
        self.update_page_display()
        self.modified = False
        self.annot_listbox.delete(0, tk.END)
        self.annot_detail_var.set("선택된 주석 없음")
        self.thumbnail_images = []
        for widget in self.thumb_inner_frame.winfo_children():
            widget.destroy()

    def on_file_select(self, event):
        selection = self.file_listbox.curselection()
        if selection:
            idx = selection[0]
            if self.modified:
                if not messagebox.askyesno("경고", "저장하지 않은 변경사항이 있습니다. 다른 파일을 열겠습니까?"):
                    return
            self.load_pdf(self.pdf_files[idx])

    def load_pdf(self, pdf_path):
        """PDF 파일 로드"""
        try:
            if self.current_pdf:
                self.current_pdf.close()
            self.current_pdf = fitz.open(pdf_path)
            self.current_pdf_path = pdf_path
            self.total_pages = len(self.current_pdf)
            self.current_page = 0
            self.modified = False
            self.selected_annot_xref = None

            # 스핀박스 범위 설정
            self.page_spinbox.config(to=self.total_pages)

            self.render_page()
            self.refresh_annotation_list()
            self.update_page_display()
            self.render_thumbnails()
        except Exception as e:
            messagebox.showerror("오류", f"PDF 로드 실패: {str(e)}")

    def render_page(self):
        """현재 페이지 렌더링 (페이지 모드)"""
        if self.current_pdf is None:
            return

        self.canvas.delete("all")

        page = self.current_pdf[self.current_page]
        zoom = self.zoom_level * 1.5
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)

        img_data = pix.tobytes("ppm")
        img = Image.open(io.BytesIO(img_data))
        self.current_image = ImageTk.PhotoImage(img)

        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.current_image)
        self.canvas.config(scrollregion=(0, 0, pix.width, pix.height))

        self.update_page_display()

    def render_continuous(self):
        """연속 보기 렌더링"""
        if self.current_pdf is None:
            return

        self.canvas.delete("all")
        self.page_images = []
        self.page_positions = []

        zoom = self.zoom_level * 1.5
        page_gap = 10  # 페이지 사이 간격
        current_y = 0
        max_width = 0

        # 모든 페이지 렌더링
        for i in range(self.total_pages):
            self.page_positions.append(current_y)

            page = self.current_pdf[i]
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)

            img_data = pix.tobytes("ppm")
            img = Image.open(io.BytesIO(img_data))
            photo = ImageTk.PhotoImage(img)
            self.page_images.append(photo)

            self.canvas.create_image(0, current_y, anchor=tk.NW, image=photo)

            # 페이지 번호 표시
            self.canvas.create_text(10, current_y + 10, anchor=tk.NW,
                                    text=f"Page {i + 1}", fill="blue",
                                    font=("Arial", 10, "bold"))

            current_y += pix.height + page_gap
            max_width = max(max_width, pix.width)

        self.canvas.config(scrollregion=(0, 0, max_width, current_y))
        self.update_page_display()

    def prev_page(self):
        if self.current_pdf and self.current_page > 0:
            self.current_page -= 1
            self.selected_annot_xref = None

            if self.view_mode == "continuous":
                self.scroll_to_page_in_continuous(self.current_page)
            else:
                self.render_page()

            self.refresh_annotation_list()
            self.annot_detail_var.set("선택된 주석 없음")
            self.update_page_display()
            self.update_thumbnail_selection()

    def next_page(self):
        if self.current_pdf and self.current_page < self.total_pages - 1:
            self.current_page += 1
            self.selected_annot_xref = None

            if self.view_mode == "continuous":
                self.scroll_to_page_in_continuous(self.current_page)
            else:
                self.render_page()

            self.refresh_annotation_list()
            self.annot_detail_var.set("선택된 주석 없음")
            self.update_page_display()
            self.update_thumbnail_selection()

    def zoom_in(self):
        self.zoom_level = min(3.0, self.zoom_level + 0.1)
        self.update_zoom_combo()
        self.refresh_view()

    def zoom_out(self):
        self.zoom_level = max(0.5, self.zoom_level - 0.1)
        self.update_zoom_combo()
        self.refresh_view()

    def on_mousewheel(self, event):
        """마우스 휠 이벤트 처리"""
        # Ctrl 키가 눌렸는지 확인
        if event.state & 0x4:  # Ctrl key
            if event.delta > 0:
                self.zoom_in()
            else:
                self.zoom_out()
        else:
            if self.view_mode == "page":
                # 페이지 모드: 스크롤 끝 도달 시 페이지 전환
                self.handle_page_mode_scroll(event)
            else:
                # 연속 모드: 일반 스크롤
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def handle_page_mode_scroll(self, event):
        """페이지 모드에서 스크롤 처리"""
        # 현재 스크롤 위치 확인
        scroll_pos = self.canvas.yview()

        if event.delta > 0:  # 위로 스크롤
            if scroll_pos[0] <= 0:  # 최상단에 도달
                if self.current_page > 0:
                    self.prev_page()
                    # 새 페이지의 하단으로 이동
                    self.root.after(50, lambda: self.canvas.yview_moveto(1.0))
            else:
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        else:  # 아래로 스크롤
            if scroll_pos[1] >= 1:  # 최하단에 도달
                if self.current_page < self.total_pages - 1:
                    self.next_page()
                    # 새 페이지의 상단으로 이동
                    self.canvas.yview_moveto(0)
            else:
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def start_conversion(self):
        if not self.pdf_files:
            messagebox.showwarning("경고", "PDF 파일을 선택해주세요.")
            return

        save_paths = []
        for pdf_path in self.pdf_files:
            pdf_name = Path(pdf_path).stem
            excel_path = filedialog.asksaveasfilename(
                title=f"저장할 파일 이름 선택 - {Path(pdf_path).name}",
                defaultextension=".xlsx",
                initialfile=f"{pdf_name}.xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            if not excel_path:
                return
            save_paths.append((pdf_path, excel_path))

        merge_tables = self.merge_tables_var.get()

        thread = threading.Thread(target=self.convert_files, args=(save_paths, merge_tables))
        thread.start()

    def convert_files(self, save_paths, merge_tables):
        total_files = len(save_paths)

        for idx, (pdf_path, excel_path) in enumerate(save_paths):
            self.progress_var.set(f"변환 중: {Path(pdf_path).name} ({idx + 1}/{total_files})")
            self.progress_bar['value'] = (idx / total_files) * 100
            self.root.update_idletasks()

            try:
                self.convert_single_file(pdf_path, excel_path, merge_tables)
            except Exception as e:
                self.root.after(0, lambda e=e, p=pdf_path: messagebox.showerror(
                    "오류", f"파일 변환 실패: {Path(p).name}\n{str(e)}"
                ))

        self.progress_bar['value'] = 100
        self.progress_var.set("변환 완료!")
        self.root.after(0, lambda: messagebox.showinfo("완료", "변환이 완료되었습니다."))

    def convert_single_file(self, pdf_path, excel_path, merge_tables):
        """단일 PDF 파일을 Excel로 변환"""
        all_tables = []

        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                tables = page.extract_tables()

                for table_idx, table in enumerate(tables):
                    if table and len(table) > 0:
                        header = table[0] if table[0] else [f"Col{i}" for i in range(len(table[1]) if len(table) > 1 else 1)]
                        data = table[1:] if len(table) > 1 else []
                        if data:
                            df = pd.DataFrame(data, columns=header)
                            df['_page'] = page_num + 1
                            df['_table'] = table_idx + 1
                            all_tables.append(df)

        if not all_tables:
            with pdfplumber.open(pdf_path) as pdf:
                text_data = []
                for page_num, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if text:
                        text_data.append({
                            'Page': page_num + 1,
                            'Content': text
                        })

                if text_data:
                    df = pd.DataFrame(text_data)
                    df.to_excel(excel_path, index=False, sheet_name='Text Content', engine='openpyxl')
                else:
                    raise Exception("PDF에서 추출할 내용이 없습니다.")
            return

        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            if merge_tables:
                all_rows = []
                for df in all_tables:
                    df = df.drop(['_page', '_table'], axis=1, errors='ignore')
                    all_rows.append(df.columns.tolist())
                    for _, row in df.iterrows():
                        all_rows.append(row.tolist())
                    all_rows.append([''] * len(df.columns))

                max_cols = max(len(row) for row in all_rows) if all_rows else 1
                for i, row in enumerate(all_rows):
                    if len(row) < max_cols:
                        all_rows[i] = row + [''] * (max_cols - len(row))

                merged_df = pd.DataFrame(all_rows)
                merged_df = merged_df.fillna('')
                merged_df.to_excel(writer, sheet_name='All Tables', index=False, header=False)
            else:
                for idx, df in enumerate(all_tables):
                    page = df['_page'].iloc[0]
                    table = df['_table'].iloc[0]
                    df = df.drop(['_page', '_table'], axis=1, errors='ignore')
                    df = df.fillna('')
                    sheet_name = f"Page{page}_Table{table}"[:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)


def main():
    root = tk.Tk()
    app = PDFToExcelConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()
