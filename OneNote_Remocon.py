# -*- coding: utf-8 -*-
import sys
import json
import os
import time
import uuid
import ctypes
from ctypes import wintypes
from typing import Optional, List, Dict, Any

from PyQt6.QtWidgets import (
    QApplication,
    # QWidget, # QWidget은 이제 직접 상속하지 않으므로 없어도 됩니다.
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QDialog,
    QListWidget,
    QGroupBox,
    QTreeWidget,
    QTreeWidgetItem,
    QToolButton,
    QSplitter,
    QMenu,
    QInputDialog,
    QMessageBox,
    QAbstractItemView,
    QMainWindow,  # 이 부분이 반드시 있어야 합니다.
    QFileDialog,
    QWidget,  # central_widget을 위해 QWidget도 필요합니다.
)
from PyQt6.QtCore import (
    QThread,
    pyqtSignal,
    QTimer,
    Qt,
    QSettings,
    QEvent,
    QRect,
)
from PyQt6.QtGui import QIcon, QAction

# ----------------- 0. 전역 상수 -----------------
SETTINGS_FILE = "OneNote_Remocon_Setting.json"
APP_ICON_PATH = "app_icon.ico"  # 사용할 아이콘 파일 경로 (ICO 형식 권장)

ONENOTE_CLASS_NAME = "ApplicationFrameWindow"  # UWP/Modern OneNote Class Name
SCROLL_STEP_SENSITIVITY = 40

# QTreeWidget 커스텀 데이터 롤
ROLE_TYPE = Qt.ItemDataRole.UserRole + 1  # 'group' | 'section'
ROLE_DATA = Qt.ItemDataRole.UserRole + 2  # dict payload

# ----------------- 0.0 설정 파일 로드/저장 유틸리티 -----------------
DEFAULT_SETTINGS = {
    "window_geometry": {"x": 200, "y": 180, "width": 960, "height": 540},
    "connection_signature": None,
    "favorites": [],
}


def load_settings() -> Dict[str, Any]:
    if not os.path.exists(SETTINGS_FILE):
        return DEFAULT_SETTINGS.copy()
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            # 기본 설정과 병합하여 키 누락 방지
            settings = DEFAULT_SETTINGS.copy()
            settings.update(data)
            return settings
    except Exception as e:
        print(f"[ERROR] 설정 파일 로드 실패: {e}")
        return DEFAULT_SETTINGS.copy()


def save_settings(data: Dict[str, Any]):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[ERROR] 설정 파일 저장 실패: {e}")


# ----------------- 0.1 pywinauto 지연 로딩 -----------------
Desktop = None
WindowNotFoundError = None
ElementNotFoundError = None
TimeoutError = None
UIAWrapper = None
UIAElementInfo = None
mouse = None
keyboard = None

_pwa_ready = False


def ensure_pywinauto():
    global _pwa_ready, Desktop, WindowNotFoundError, ElementNotFoundError, TimeoutError, UIAWrapper, UIAElementInfo, mouse, keyboard
    if _pwa_ready:
        return
    try:
        from pywinauto import (
            Desktop as _Desktop,
            mouse as _mouse,
            keyboard as _keyboard,
        )
        from pywinauto.findwindows import (
            WindowNotFoundError as _WNF,
            ElementNotFoundError as _ENF,
        )
        from pywinauto.timings import TimeoutError as _TO
        from pywinauto.controls.uiawrapper import UIAWrapper as _UIAWrapper
        from pywinauto.uia_element_info import UIAElementInfo as _UIAElementInfo

        Desktop = _Desktop
        WindowNotFoundError = _WNF
        ElementNotFoundError = _ENF
        TimeoutError = _TO
        UIAWrapper = _UIAWrapper
        UIAElementInfo = _UIAElementInfo
        mouse = _mouse
        keyboard = _keyboard
        _pwa_ready = True
    except ImportError:
        # pywinauto가 없으면 자동화 기능은 비활성 상태
        pass


# ----------------- 0.2 Win32 빠른 창 열거 -----------------
_user32 = ctypes.windll.user32


def _win_get_window_text(hwnd):
    length = _user32.GetWindowTextLengthW(hwnd)
    buf = ctypes.create_unicode_buffer(length + 1 if length > 0 else 1)
    _user32.GetWindowTextW(hwnd, buf, len(buf))
    return buf.value


def _win_get_class_name(hwnd):
    # 중복 호출 버그 수정
    buf = ctypes.create_unicode_buffer(256)
    _user32.GetClassNameW(hwnd, buf, 256)
    return buf.value


def enum_windows_fast(filter_title_substr=None):
    if isinstance(filter_title_substr, str):
        filters = [filter_title_substr.lower()]
    elif filter_title_substr:
        filters = [str(s).lower() for s in filter_title_substr]
    else:
        filters = None

    results = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    def _enum_proc(hwnd, lparam):
        try:
            if not _user32.IsWindowVisible(hwnd):
                return True
            title = _win_get_window_text(hwnd)
            if not title:
                return True
            if filters and not any(f in title.lower() for f in filters):
                return True

            cls = _win_get_class_name(hwnd)
            pid = wintypes.DWORD()
            _user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            results.append(
                {
                    "handle": int(hwnd),
                    "title": title,
                    "class_name": cls,
                    "pid": pid.value,
                }
            )
        except Exception:
            pass
        return True

    _user32.EnumWindows(_enum_proc, 0)
    return results


# ----------------- 0.3 리소스 경로 헬퍼 (PyInstaller 호환) -----------------
def resource_path(relative_path):
    """
    PyInstaller에서 묶인 리소스 파일을 찾는 경로를 반환합니다.
    """
    try:
        # PyInstaller 실행 환경: 임시 폴더 경로를 사용
        base_path = sys._MEIPASS
    except Exception:
        # 개발 환경: 현재 디렉토리 경로를 사용
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# ----------------- 1. 프로세스 실행 파일 경로 얻기 -----------------
def get_process_image_path(pid: int) -> Optional[str]:
    if not pid:
        return None
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    kernel32 = ctypes.windll.kernel32
    OpenProcess = kernel32.OpenProcess
    CloseHandle = kernel32.CloseHandle
    QueryFullProcessImageNameW = kernel32.QueryFullProcessImageNameW

    hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not hProcess:
        return None
    try:
        buf_len = wintypes.DWORD(1024)
        buf = ctypes.create_unicode_buffer(buf_len.value)
        if QueryFullProcessImageNameW(hProcess, 0, buf, ctypes.byref(buf_len)):
            return buf.value
        return None
    finally:
        CloseHandle(hProcess)


# ----------------- 1.1 엄격한 OneNote 창 검증 헬퍼 (2번 로직 이식) -----------------
def is_strict_onenote_window(w: Dict[str, Any], my_pid: int) -> bool:
    """주어진 창 정보가 실제로 OneNote 앱 창인지 엄격하게 확인합니다."""
    if w.get("pid") == my_pid:
        return False

    title_lower = w.get("title", "").lower()
    cls = w.get("class_name", "")
    pid = w.get("pid")

    # 1. Classic Desktop (OMain*) - 레거시 OneNote
    if "omain" in (cls or "").lower():
        return True

    # 2. Modern App (ApplicationFrameWindow) + 타이틀 키워드
    if cls == ONENOTE_CLASS_NAME and (
        "onenote" in title_lower or "원노트" in title_lower
    ):
        return True

    # 3. Fallback: 제목에 키워드 + EXE 확인
    if "onenote" in title_lower or "원노트" in title_lower:
        exe_path = get_process_image_path(pid)
        if exe_path:
            exe_name = os.path.basename(exe_path).lower()
            if "onenote.exe" in exe_name or "onenoteim.exe" in exe_name:
                return True

    return False


# ----------------- 4. 짧은 폴링으로 Rect 안정화 대기 -----------------
def _wait_rect_settle(get_rect, timeout=0.3, interval=0.03):
    start = time.perf_counter()
    prev = get_rect()
    while time.perf_counter() - start < timeout:
        time.sleep(interval)
        cur = get_rect()
        if abs(cur.top - prev.top) < 2 and abs(cur.bottom - prev.bottom) < 2:
            break
        prev = cur


# ----------------- 5. 패턴 기반 수직 스크롤 시도 -----------------
def _scroll_vertical_via_pattern(
    container, direction: str, small=True, repeats=1
) -> bool:
    ensure_pywinauto()
    if not _pwa_ready:
        return False
    try:
        iface = getattr(container, "iface_scroll", None)
        if iface is None:
            return False

        from comtypes.gen.UIAutomationClient import (
            ScrollAmount_LargeIncrement,
            ScrollAmount_LargeDecrement,
            ScrollAmount_SmallIncrement,
            ScrollAmount_SmallDecrement,
            ScrollAmount_NoAmount,
        )

        v_inc = ScrollAmount_SmallIncrement if small else ScrollAmount_LargeIncrement
        v_dec = ScrollAmount_SmallDecrement if small else ScrollAmount_LargeDecrement
        v_amount = v_inc if direction == "down" else v_dec

        for _ in range(max(1, repeats)):
            iface.Scroll(ScrollAmount_NoAmount, v_amount)
        return True
    except Exception:
        return False


# ----------------- 6. 마우스 휠 기반 스크롤(폴백) - 2번 로직 이식 -----------------
def _safe_wheel(scroll_container, steps: int):
    if steps == 0:
        return

    ensure_pywinauto()

    try:
        if hasattr(scroll_container, "wheel_scroll"):
            scroll_container.wheel_scroll(steps)
            return
    except Exception:
        pass

    try:
        if hasattr(scroll_container, "wheel_mouse_input"):
            scroll_container.wheel_mouse_input(wheel_dist=steps)
            return
    except Exception:
        pass

    try:
        rect = scroll_container.rectangle()
        center = rect.mid_point()
        try:
            mouse.scroll(coords=(center.x, center.y), wheel_dist=steps)
            return
        except Exception:
            pass
        try:
            mouse.wheel(coords=(center.x, center.y), wheel_dist=steps)
            return
        except Exception:
            pass
    except Exception:
        pass

    try:
        scroll_container.set_focus()
        if steps > 0:
            keyboard.send_keys("{UP %d}" % steps)
        else:
            keyboard.send_keys("{DOWN %d}" % abs(steps))
    except Exception:
        pass


# ----------------- 7. 선택 항목을 가장 빠르게 얻기 - 2번 로직 이식 -----------------
def get_selected_tree_item_fast(tree_control):
    ensure_pywinauto()
    if not _pwa_ready:
        return None

    try:
        if hasattr(tree_control, "get_selection"):
            sel = tree_control.get_selection()
            if sel:
                return sel[0]
    except Exception:
        pass

    try:
        iface_sel = getattr(tree_control, "iface_selection", None)
        if iface_sel:
            arr = iface_sel.GetSelection()
            length = getattr(arr, "Length", 0)
            if length and length > 0:
                el = arr.GetElement(0)
                return UIAWrapper(UIAElementInfo(el))
    except Exception:
        pass

    try:
        for item in tree_control.children():
            try:
                if item.is_selected():
                    return item
            except Exception:
                pass
    except Exception:
        pass

    try:
        for item in tree_control.descendants(control_type="TreeItem"):
            try:
                if item.is_selected():
                    return item
            except Exception:
                pass
    except Exception:
        pass

    return None


# ----------------- 8. 페이지/섹션 컨테이너(Tree/List) 찾기 - ensure 호출 -----------------
def _find_tree_or_list(onenote_window):
    ensure_pywinauto()
    if not _pwa_ready:
        return None
    for ctype in ("Tree", "List"):
        try:
            return onenote_window.child_window(
                control_type=ctype, found_index=0
            ).wrapper_object()
        except Exception:
            continue
    return None


# ----------------- 8.1 지정 텍스트 섹션 찾기/선택 - 2번 로직 이식 -----------------
def _normalize_text(s: Optional[str]) -> str:
    return " ".join(((s or "").strip().split())).lower()


def select_section_by_text(
    onenote_window, text: str, tree_control: Optional[object] = None
) -> bool:
    ensure_pywinauto()
    if not _pwa_ready:
        return False
    try:
        tree_control = tree_control or _find_tree_or_list(onenote_window)
        if not tree_control:
            return False

        target_norm = _normalize_text(text)

        def _scan(types: List[str]):
            for t in types:
                try:
                    for itm in tree_control.descendants(control_type=t):
                        try:
                            if _normalize_text(itm.window_text()) == target_norm:
                                try:
                                    itm.select()
                                except Exception:
                                    try:
                                        itm.click_input()
                                    except Exception:
                                        return False
                                return True
                        except Exception:
                            continue
                except Exception:
                    continue
            return False

        if _scan(["TreeItem"]) or _scan(["ListItem"]):
            _center_element_in_view(
                get_selected_tree_item_fast(tree_control), tree_control
            )
            return True
        return False
    except Exception:
        return False


# ----------------- 9. 요소를 중앙으로 위치시키는 함수(최적화) - ensure 호출 -----------------
def _center_element_in_view(element_to_center, scroll_container):
    ensure_pywinauto()
    if not _pwa_ready:
        return
    try:
        try:
            element_to_center.iface_scroll_item.ScrollIntoView()
        except AttributeError:
            # ScrollItem 미지원 컨트롤
            return

        _wait_rect_settle(
            lambda: element_to_center.rectangle(), timeout=0.3, interval=0.03
        )

        rect_container = scroll_container.rectangle()
        rect_item = element_to_center.rectangle()
        item_center_y = (rect_item.top + rect_item.bottom) / 2
        container_center_y = (rect_container.top + rect_container.bottom) / 2
        offset = item_center_y - container_center_y

        if abs(offset) <= 10:
            return

        def step_for(dy):
            return max(1, min(5, int(abs(dy) / 150)))

        for _ in range(3):
            if abs(offset) <= 10:
                break

            direction = "down" if offset > 0 else "up"
            repeats = step_for(offset)

            used_pattern = _scroll_vertical_via_pattern(
                scroll_container, direction=direction, small=True, repeats=repeats
            )
            if not used_pattern:
                wheel_steps = -repeats if offset > 0 else repeats
                _safe_wheel(scroll_container, wheel_steps)

            time.sleep(0.03)

            rect_container = scroll_container.rectangle()
            rect_item = element_to_center.rectangle()
            item_center_y = (rect_item.top + rect_item.bottom) / 2
            container_center_y = (rect_container.top + rect_container.bottom) / 2
            offset = item_center_y - container_center_y

    except Exception as e:
        print(f"[WARN] 중앙 정렬 중 오류: {e}")


# ----------------- 10. 선택된 항목을 중앙으로 스크롤 - 2번 로직 이식 -----------------
def scroll_selected_item_to_center(
    onenote_window, tree_control: Optional[object] = None
):
    ensure_pywinauto()
    if not _pwa_ready:
        return False, None
    try:
        tree_control = tree_control or _find_tree_or_list(onenote_window)
        if not tree_control:
            return False, None

        selected_item = get_selected_tree_item_fast(tree_control)
        if not selected_item:
            return False, None

        item_name = selected_item.window_text()
        _center_element_in_view(selected_item, tree_control)
        return True, item_name
    except (ElementNotFoundError, TimeoutError):
        return False, None
    except Exception:
        return False, None


# ----------------- 11. 연결 시그니처 저장/스코어 기반 재획득 -----------------
def build_window_signature(win) -> dict:
    try:
        pid = win.process_id()
    except Exception:
        pid = None
    exe_path = get_process_image_path(pid) if pid else None
    exe_name = os.path.basename(exe_path).lower() if exe_path else None
    try:
        handle = win.handle
    except Exception:
        handle = None
    try:
        title = win.window_text()
    except Exception:
        title = None
    try:
        cls_name = win.class_name()
    except Exception:
        cls_name = None

    return {
        "handle": handle,
        "pid": pid,
        "class_name": cls_name,
        "title": title,
        "exe_path": exe_path,
        "exe_name": exe_name,
    }


def save_connection_info(window_element):
    try:
        info = build_window_signature(window_element)
        current_settings = load_settings()
        current_settings["connection_signature"] = info
        save_settings(current_settings)
    except Exception as e:
        print(f"[ERROR] 연결 정보 저장 실패: {e}")


def _score_candidate_dict(c, sig) -> int:
    try:
        title = (c.get("title") or "").lower()
        cls = c.get("class_name") or ""
        pid = c.get("pid")
        exe_path = get_process_image_path(pid) or ""
        exe_name = os.path.basename(exe_path).lower() if exe_path else ""

        score = 0
        if sig.get("handle") and c.get("handle") == sig["handle"]:
            score += 100
        if sig.get("exe_name") and exe_name == sig["exe_name"]:
            score += 50
        if "onenote.exe" in exe_name:
            score += 50
        if "onenote" in title or "원노트" in title:
            score += 25
        if sig.get("class_name") and cls == sig["class_name"]:
            score += 10
        if sig.get("pid") and pid == sig["pid"]:
            score += 8
        prev_title = (sig.get("title") or "").lower()
        if prev_title:
            if prev_title in title:
                score += 6
            else:
                if "onenote" in prev_title and "onenote" in title:
                    score += 4
                if "원노트" in prev_title and "원노트" in title:
                    score += 4
        if cls == ONENOTE_CLASS_NAME:
            score += 5
        return score
    except Exception:
        return -1


def reacquire_window_by_signature(sig) -> Optional[object]:
    ensure_pywinauto()
    if not _pwa_ready:
        return None
    h = sig.get("handle")
    if h:
        try:
            w = Desktop(backend="uia").window(handle=h)
            if w.is_visible():
                return w
        except Exception:
            pass

    candidates = enum_windows_fast(filter_title_substr=None)
    best, best_score = None, -1
    for c in candidates:
        s = _score_candidate_dict(c, sig)
        if s > best_score:
            best, best_score = c, s

    if best and best_score >= 30:
        try:
            w = Desktop(backend="uia").window(handle=best["handle"])
            if w.is_visible():
                return w
        except Exception:
            return None
    return None


# ----------------- 12. 저장된 정보로 재연결 -----------------
def load_connection_info_and_reconnect():
    ensure_pywinauto()
    settings = load_settings()
    sig = settings.get("connection_signature")
    if not sig:
        return None, "연결되지 않음"
    try:
        win = reacquire_window_by_signature(sig)
        if win and win.is_visible():
            window_title = win.window_text()
            try:
                save_connection_info(win)
            except Exception:
                pass
            return win, f"(자동 재연결) '{window_title}'"

        return None, "(재연결 실패) 이전 앱을 찾을 수 없습니다."
    except Exception:
        return None, "연결되지 않음"


# ----------------- 13. 백그라운드 자동 재연결 워커 -----------------
class ReconnectWorker(QThread):
    finished = pyqtSignal(object)

    def run(self):
        try:
            ensure_pywinauto()
            win, status = load_connection_info_and_reconnect()
            if win:
                payload = {
                    "ok": True,
                    "status": status,
                    "sig": build_window_signature(win),
                }
            else:
                payload = {"ok": False, "status": status}
        except Exception as e:
            payload = {"ok": False, "status": f"연결되지 않음 (오류: {e})"}
        self.finished.emit(payload)


# ----------------- 3-A. OneNote 창 목록 스캔 워커 -----------------
class OneNoteWindowScanner(QThread):
    done = pyqtSignal(list)

    def __init__(self, my_pid: int, parent=None):
        super().__init__(parent)
        self.my_pid = my_pid

    def run(self):
        results = []
        try:
            wins = enum_windows_fast(filter_title_substr=None)
            for w in wins:
                try:
                    if is_strict_onenote_window(w, self.my_pid):
                        results.append(w)
                except Exception:
                    continue

            results.sort(
                key=lambda r: (
                    r.get("class_name", "") != ONENOTE_CLASS_NAME,
                    r.get("title", ""),
                )
            )
        except Exception as e:
            print(f"[ERROR] OneNote 창 스캔 중 오류: {e}")
        finally:
            self.done.emit(results)


# ----------------- 3-B/C. 기타 창 스캔 및 선택 다이얼로그 -----------------
class WindowListWorker(QThread):
    done = pyqtSignal(list)

    def run(self):
        try:
            results = enum_windows_fast(filter_title_substr=None)
            self.done.emit(results)
        except Exception:
            self.done.emit([])


class OtherWindowSelectionDialog(QDialog):
    def __init__(self, my_pid: int, parent=None):
        super().__init__(parent)
        self.my_pid = my_pid
        self.setWindowTitle("연결할 창을 더블클릭하세요.")
        self.setGeometry(400, 400, 500, 420)

        self.layout = QVBoxLayout(self)
        self.tip_label = QLabel("창 목록을 검색 중입니다...")
        self.layout.addWidget(self.tip_label)

        self.other_list_widget = QListWidget()
        self.layout.addWidget(self.other_list_widget)
        self.other_list_widget.hide()

        self.windows_info = []
        self.selected_info = None

        self.other_list_widget.itemDoubleClicked.connect(self.on_window_selected)

        self.worker = WindowListWorker()
        self.worker.done.connect(self._on_windows_list_ready)
        self.worker.start()

    def _on_windows_list_ready(self, results):
        self.tip_label.hide()
        if not results:
            self.tip_label.setText("표시할 창이 없습니다. 다시 시도해 주세요.")
            self.tip_label.show()
            return

        for r in results:
            pid = r.get("pid")
            if pid == self.my_pid:
                continue
            if not is_strict_onenote_window(r, self.my_pid):
                self.windows_info.append(r)

        self.windows_info.sort(key=lambda r: r.get("title", ""))

        if self.windows_info:
            items = [
                f'{r["title"]}  [{r["class_name"]}] (0x{r["handle"]:X})'
                for r in self.windows_info
            ]
            self.other_list_widget.addItems(items)
            self.other_list_widget.show()
        else:
            self.tip_label.setText("OneNote를 제외한 다른 창이 없습니다.")
            self.tip_label.show()

    def on_window_selected(self, item):
        row = self.other_list_widget.currentRow()
        if 0 <= row < len(self.windows_info):
            self.selected_info = self.windows_info[row]
            self.accept()


# ----------------- 14-A. 즐겨찾기 트리 위젯 -----------------
class FavoritesTree(QTreeWidget):
    structureChanged = pyqtSignal()
    deleteRequested = pyqtSignal()
    renameRequested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setHeaderHidden(True)
        self.setColumnCount(1)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setIndentation(16)
        self.setAnimated(True)
        self.setExpandsOnDoubleClick(True)

    def dropEvent(self, event):
        # 드래그 중인 아이템과 드롭 대상 아이템을 미리 가져옵니다.
        source_item = self.currentItem()
        target_item = self.itemAt(event.position().toPoint())

        # 소스 아이템이 없으면 이벤트를 무시합니다.
        if not source_item:
            event.ignore()
            return

        # 먼저 Qt의 기본 드롭 이벤트를 실행하여 아이템을 이동시킵니다.
        # 이렇게 하면 복잡한 이동 로직을 직접 구현할 필요가 없습니다.
        super().dropEvent(event)

        # 이제, 이동된 결과를 확인하고 규칙에 어긋나면 위치를 수정합니다.
        if target_item and source_item.parent() == target_item:
            source_type = source_item.data(0, ROLE_TYPE)
            target_type = target_item.data(0, ROLE_TYPE)

            # 규칙: '섹션'은 다른 '섹션'의 자식이 될 수 없습니다.
            if source_type == "section" and target_type == "section":
                # 1. 잘못된 위치(target_item)에서 source_item을 다시 떼어냅니다.
                moved_item = target_item.takeChild(
                    target_item.indexOfChild(source_item)
                )

                if moved_item:
                    # 2. target_item의 부모(즉, 한 단계 위 그룹)를 찾습니다.
                    parent_of_target = target_item.parent()
                    if not parent_of_target:
                        parent_of_target = self.invisibleRootItem()

                    # 3. target_item의 인덱스를 찾습니다.
                    target_index = parent_of_target.indexOfChild(target_item)

                    # 4. 떼어냈던 아이템을 target_item 바로 다음에 삽입합니다.
                    parent_of_target.insertChild(target_index + 1, moved_item)

                    # 5. 사용자가 위치를 인지할 수 있도록 이동된 아이템을 선택합니다.
                    self.setCurrentItem(moved_item)

        # 구조가 변경되었음을 알립니다 (저장을 위함).
        self.structureChanged.emit()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Delete:
            self.deleteRequested.emit()
            event.accept()
            return
        if event.key() == Qt.Key.Key_F2:
            self.renameRequested.emit()
            event.accept()
            return
        super().keyPressEvent(event)


# ----------------- 14. PyQt GUI -----------------
class OneNoteScrollRemoconApp(QMainWindow):
    def __init__(self):
        super().__init__()
        ensure_pywinauto()

        # 1. 설정 로드 및 창 위치 설정
        self.settings = load_settings()
        self.onenote_window = None
        self.tree_control = None
        self._reconnect_worker = None
        self._scanner_worker = None
        self.onenote_windows_info: List[Dict] = []
        self.my_pid = os.getpid()
        self._auto_center_after_activate = True

        # 1.1 애플리케이션 아이콘 설정
        icon_path = resource_path(APP_ICON_PATH)
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        # --- [오류 수정] ---
        # 설정 파일에서 geo 값을 가져올 때 딕셔너리인지 확인
        geo = self.settings.get("window_geometry")
        if isinstance(geo, dict):
            # 딕셔너리가 맞으면 해당 값으로 창 위치 설정
            self.setGeometry(
                geo.get("x", 200),
                geo.get("y", 180),
                geo.get("width", 960),
                geo.get("height", 540),
            )
        else:
            # 딕셔너리가 아니거나(리스트 등) 없으면 기본값으로 설정
            self.setGeometry(200, 180, 960, 540)
            # 메모리의 설정도 올바르게 수정하여, 앱 종료 시 정상적으로 저장되도록 함
            self.settings["window_geometry"] = DEFAULT_SETTINGS["window_geometry"]
        # --- [수정 완료] ---

        self.init_ui("준비됨 (자동 재연결 중...)")

        # 2. 즐겨찾기 로드
        self._load_favorites()

        self.fav_tree.deleteRequested.connect(self._delete_favorite_item)
        self.fav_tree.renameRequested.connect(self._rename_favorite_item)

        QTimer.singleShot(0, self.refresh_onenote_list)
        QTimer.singleShot(0, self._start_auto_reconnect)

    def init_ui(self, initial_status):
        self.setWindowTitle("OneNote 전자필기장 스크롤 리모컨")

        # --- 메뉴바 생성 ---
        menubar = self.menuBar()
        file_menu = menubar.addMenu("&파일")

        export_action = QAction("즐겨찾기 내보내기...", self)
        export_action.triggered.connect(self._export_favorites)
        file_menu.addAction(export_action)

        import_action = QAction("즐겨찾기 가져오기...", self)
        import_action.triggered.connect(self._import_favorites)
        file_menu.addAction(import_action)

        # --- 스타일시트 정의 ---
        COLOR_BACKGROUND = "#2E2E2E"
        COLOR_PRIMARY_TEXT = "#E0E0E0"
        COLOR_SECONDARY_TEXT = "#B0B0B0"
        COLOR_GROUPBOX_BG = "#3C3C3C"
        COLOR_ACCENT = "#A6D854"
        COLOR_ACCENT_HOVER = "#B8E966"
        COLOR_ACCENT_PRESSED = "#95C743"
        COLOR_SECONDARY_BUTTON = "#555555"
        COLOR_SECONDARY_BUTTON_HOVER = "#666666"
        COLOR_SECONDARY_BUTTON_PRESSED = "#444444"
        COLOR_LIST_BG = "#252525"
        COLOR_LIST_SELECTED = "#0078D7"
        COLOR_STATUS_BAR = "#252525"

        self.setStyleSheet(
            f"""
            QWidget {{
                background-color: {COLOR_BACKGROUND};
                color: {COLOR_PRIMARY_TEXT};
                font-family: 'Malgun Gothic';
                font-size: 10pt;
            }}
            QGroupBox {{
                background-color: {COLOR_GROUPBOX_BG};
                border: 1px solid {COLOR_BACKGROUND};
                border-radius: 8px;
                margin-top: 10px;
                font-weight: bold;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 10px;
                left: 10px;
            }}
            QLabel {{
                color: {COLOR_SECONDARY_TEXT};
                font-weight: normal;
            }}
            QListWidget {{
                background-color: {COLOR_LIST_BG};
                border: 1px solid {COLOR_GROUPBOX_BG};
                border-radius: 4px;
            }}
            QListWidget::item:selected {{
                background-color: {COLOR_LIST_SELECTED};
                color: white;
            }}
            QTreeWidget {{
                background-color: {COLOR_LIST_BG};
                border: 1px solid {COLOR_GROUPBOX_BG};
                border-radius: 6px;
            }}
            QToolButton {{
                background-color: {COLOR_SECONDARY_BUTTON};
                color: {COLOR_PRIMARY_TEXT};
                border: none;
                border-radius: 4px;
                padding: 4px 6px;
            }}
            QToolButton:hover {{
                background-color: {COLOR_SECONDARY_BUTTON_HOVER};
            }}
            QToolButton:pressed {{
                background-color: {COLOR_SECONDARY_BUTTON_PRESSED};
            }}
            QPushButton {{
                background-color: {COLOR_SECONDARY_BUTTON};
                color: {COLOR_PRIMARY_TEXT};
                border: none;
                border-radius: 4px;
                padding: 6px 12px;
            }}
            QPushButton:hover {{
                background-color: {COLOR_SECONDARY_BUTTON_HOVER};
            }}
            QPushButton:pressed {{
                background-color: {COLOR_SECONDARY_BUTTON_PRESSED};
            }}
            QPushButton:disabled {{
                background-color: #404040;
                color: #808080;
            }}
            QMenuBar {{
                background-color: {COLOR_GROUPBOX_BG};
                color: {COLOR_PRIMARY_TEXT};
            }}
            QMenuBar::item:selected {{
                background-color: {COLOR_SECONDARY_BUTTON_HOVER};
            }}
            QMenu {{
                background-color: {COLOR_GROUPBOX_BG};
                border: 1px solid {COLOR_SECONDARY_BUTTON};
            }}
            QMenu::item:selected {{
                background-color: {COLOR_LIST_SELECTED};
            }}
            #StatusBarLabel {{
                background-color: {COLOR_STATUS_BAR};
                color: {COLOR_PRIMARY_TEXT};
                padding: 5px 12px;
                font-size: 9pt;
                border-top: 1px solid #444444;
            }}
            /* --- [추가] 검색 입력창 스타일 --- */
            QLineEdit {{
                background-color: {COLOR_LIST_BG};
                border: 1px solid {COLOR_SECONDARY_BUTTON};
                border-radius: 4px;
                padding: 4px 8px;
            }}
            QLineEdit:focus {{
                border: 1px solid {COLOR_LIST_SELECTED};
            }}
            /* --- [추가 완료] --- */
        """
        )

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(10)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        main_layout.addWidget(splitter, stretch=1)

        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        fav_group = QGroupBox("즐겨찾기")
        fav_layout = QVBoxLayout(fav_group)

        # 툴바 - 1행
        tb1_layout = QHBoxLayout()
        self.btn_add_group = QToolButton()
        self.btn_add_group.setText("그룹 추가")
        self.btn_add_group.clicked.connect(self._add_group)

        self.btn_add_section_current = QToolButton()
        self.btn_add_section_current.setText("현재 섹션 추가")
        self.btn_add_section_current.clicked.connect(self._add_section_from_current)

        self.btn_rename = QToolButton()
        self.btn_rename.setText("이름 바꾸기 (F2)")
        self.btn_rename.clicked.connect(self._rename_favorite_item)

        # --- [수정] 버튼 위치 교체 (1) ---
        tb1_layout.addWidget(
            self.btn_add_section_current
        )  # "현재 섹션 추가"를 먼저 배치
        tb1_layout.addWidget(self.btn_rename)
        tb1_layout.addStretch(1)

        # 툴바 - 2행
        tb2_layout = QHBoxLayout()
        self.btn_delete = QToolButton()
        self.btn_delete.setText("삭제 (Del)")
        self.btn_delete.clicked.connect(self._delete_favorite_item)
        self.btn_expand_all = QToolButton()
        self.btn_expand_all.setText("그룹 펼치기")
        self.btn_collapse_all = QToolButton()
        self.btn_collapse_all.setText("그룹 접기")

        # --- [수정] 버튼 위치 교체 (2) ---
        tb2_layout.addWidget(self.btn_add_group)  # "그룹 추가"를 두 번째 줄로 배치
        tb2_layout.addStretch(1)
        tb2_layout.addWidget(self.btn_expand_all)
        tb2_layout.addWidget(self.btn_collapse_all)

        fav_layout.addLayout(tb1_layout)
        fav_layout.addLayout(tb2_layout)

        self.fav_tree = FavoritesTree()
        self.btn_expand_all.clicked.connect(self.fav_tree.expandAll)
        self.btn_collapse_all.clicked.connect(self.fav_tree.collapseAll)
        self.fav_tree.itemDoubleClicked.connect(self._on_fav_item_double_clicked)
        self.fav_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.fav_tree.customContextMenuRequested.connect(self._on_fav_context_menu)
        self.fav_tree.structureChanged.connect(self._save_favorites)
        self.fav_tree.itemChanged.connect(lambda *_: self._save_favorites())

        fav_layout.addWidget(self.fav_tree)

        # --- [수정] 하단 버튼 레이아웃 재구성 ---
        move_buttons_layout = QHBoxLayout()
        move_buttons_layout.addWidget(self.btn_delete)  # 왼쪽에 "삭제" 버튼 추가
        move_buttons_layout.addStretch(1)

        self.btn_move_up = QToolButton()
        self.btn_move_up.setText("▲ 위로")
        self.btn_move_up.clicked.connect(self._move_item_up)

        self.btn_move_down = QToolButton()
        self.btn_move_down.setText("▼ 아래로")
        self.btn_move_down.clicked.connect(self._move_item_down)

        move_buttons_layout.addWidget(self.btn_move_up)
        move_buttons_layout.addWidget(self.btn_move_down)

        fav_layout.addLayout(move_buttons_layout)
        # --- [수정 완료] ---

        self.fav_tree.itemSelectionChanged.connect(self._update_move_button_state)

        left_layout.addWidget(fav_group, stretch=1)
        splitter.addWidget(left_panel)
        splitter.setStretchFactor(0, 0)

        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(10)

        connection_group = QGroupBox("OneNote 창 목록")
        connection_layout = QVBoxLayout(connection_group)

        list_header_layout = QHBoxLayout()
        list_header_layout.addWidget(
            QLabel("더블클릭하여 연결 및 중앙 정렬"),
            alignment=Qt.AlignmentFlag.AlignLeft,
        )
        list_header_layout.addStretch()

        self.refresh_button = QPushButton(" 새로고침")
        refresh_icon = self.style().standardIcon(
            QApplication.style().StandardPixmap.SP_BrowserReload
        )
        self.refresh_button.setIcon(QIcon(refresh_icon))
        self.refresh_button.clicked.connect(self.refresh_onenote_list)
        list_header_layout.addWidget(self.refresh_button)

        connection_layout.addLayout(list_header_layout)

        self.onenote_list_widget = QListWidget()
        self.onenote_list_widget.addItem("자동 재연결 시도 중...")
        self.onenote_list_widget.itemDoubleClicked.connect(
            self.connect_and_center_from_list_item
        )
        connection_layout.addWidget(self.onenote_list_widget)
        right_layout.addWidget(connection_group)

        actions_group = QGroupBox("자동화 기능")
        actions_layout = QVBoxLayout(actions_group)

        self.center_button = QPushButton("현재 선택된 항목 중앙으로 정렬")
        center_icon = self.style().standardIcon(
            QApplication.style().StandardPixmap.SP_ArrowRight
        )
        self.center_button.setIcon(QIcon(center_icon))
        self.center_button.setStyleSheet(
            f"""
            QPushButton {{
                background-color: {COLOR_ACCENT};
                color: #111;
                font-weight: bold;
                padding: 8px 16px;
            }}
            QPushButton:hover {{ background-color: {COLOR_ACCENT_HOVER}; }}
            QPushButton:pressed {{ background-color: {COLOR_ACCENT_PRESSED}; }}
            QPushButton:disabled {{
                background-color: #555555;
                color: #999999;
                border: 1px solid #444444;
            }}
        """
        )
        self.center_button.clicked.connect(self.center_selected_item_action)
        self.center_button.setEnabled(False)
        actions_layout.addWidget(self.center_button)

        other_buttons_layout = QHBoxLayout()
        connect_other_button = QPushButton("다른 앱에 연결...")
        connect_other_button.clicked.connect(self.select_other_window)
        other_buttons_layout.addWidget(connect_other_button)

        disconnect_button = QPushButton("연결 해제")
        disconnect_button.clicked.connect(self.disconnect_and_clear_info)
        other_buttons_layout.addWidget(disconnect_button)
        actions_layout.addLayout(other_buttons_layout)

        right_layout.addWidget(actions_group)

        # --- [수정] '전자필기장 검색' 그룹을 별도로 생성 ---
        # QLineEdit를 추가하기 위해 PyQt6.QtWidgets 임포트 목록에 추가해야 합니다.
        try:
            from PyQt6.QtWidgets import QLineEdit
        except ImportError:  # 혹시 모를 상황 대비
            QLineEdit = lambda: QWidget()

        search_group = QGroupBox("전자필기장 검색")  # 새 그룹박스 생성
        search_group_layout = QVBoxLayout(search_group)  # 그룹박스를 위한 레이아웃

        search_widget_layout = QHBoxLayout()  # 입력창, 버튼을 담을 가로 레이아웃

        # '전자필기장 검색:' 라벨 생성 코드 삭제

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(
            "검색할 섹션 이름 입력..."
        )  # 안내 문구 수정
        self.search_input.returnPressed.connect(self._search_and_select_section)
        self.search_input.setEnabled(False)
        search_widget_layout.addWidget(self.search_input, stretch=1)

        self.search_button = QPushButton("전자필기장 위치")
        # '검색' 버튼 스타일 변경
        self.search_button.setStyleSheet(
            """
            QPushButton {
                background-color: #F39C12; 
                color: #000000; 
                font-weight: bold;
                border: none;
                padding: 6px 12px;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: #F5B041; }
            QPushButton:pressed { background-color: #D68910; }
            QPushButton:disabled { 
                background-color: #555555;
                color: #999999;
            }
        """
        )
        self.search_button.clicked.connect(self._search_and_select_section)
        self.search_button.setEnabled(False)
        search_widget_layout.addWidget(self.search_button)

        search_group_layout.addLayout(search_widget_layout)
        right_layout.addWidget(search_group)
        # --- [수정 완료] ---

        right_layout.addStretch(1)
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(1, 1)

        self.connection_status_label = QLabel(initial_status)
        self.statusBar().addPermanentWidget(self.connection_status_label)
        self.statusBar().setStyleSheet(f"background-color: {COLOR_STATUS_BAR};")

        splitter.setSizes([320, 640])
        self._update_move_button_state()

    # ----------------- 14.1 창 닫기 이벤트 핸들러 (Geometry/Favorites 저장) -----------------
    def closeEvent(self, event):
        self._save_window_geometry()
        self._save_favorites()
        super().closeEvent(event)

    def _save_window_geometry(self):
        geom = self.geometry()
        current_settings = load_settings()
        current_settings["window_geometry"] = {
            "x": geom.x(),
            "y": geom.y(),
            "width": geom.width(),
            "height": geom.height(),
        }
        save_settings(current_settings)

    def update_status_and_ui(self, status_text: str, is_connected: bool):
        self.connection_status_label.setText(status_text)
        self.center_button.setEnabled(is_connected)
        # --- [추가] 검색 관련 UI 활성화/비활성화 ---
        self.search_input.setEnabled(is_connected)
        self.search_button.setEnabled(is_connected)
        # --- [추가 완료] ---

    def _start_auto_reconnect(self):
        self.refresh_button.setEnabled(False)
        self._reconnect_worker = ReconnectWorker()
        self._reconnect_worker.finished.connect(self._on_reconnect_done)
        self._reconnect_worker.start()

    def _on_reconnect_done(self, payload):
        self._reconnect_worker = None
        status = payload.get("status", "연결되지 않음")
        if payload.get("ok"):
            ensure_pywinauto()
            sig = payload.get("sig", {})
            target = None
            try:
                h = sig.get("handle")
                if h:
                    target = Desktop(backend="uia").window(handle=h)
                if not target or not target.is_visible():
                    target = reacquire_window_by_signature(sig)
            except Exception:
                target = None

            if target:
                self.onenote_window = target
                try:
                    save_connection_info(self.onenote_window)
                except Exception:
                    pass
                self.update_status_and_ui(f"연결됨: {status}", True)
                QTimer.singleShot(0, self._cache_tree_control)
                self.refresh_onenote_list()
                return

        self.onenote_window = None
        self.tree_control = None
        self.update_status_and_ui(f"상태: {status}", False)
        self.refresh_onenote_list()

    def refresh_onenote_list(self):
        if self._scanner_worker and self._scanner_worker.isRunning():
            return

        self.onenote_list_widget.clear()
        self.onenote_list_widget.addItem("OneNote 창을 검색 중입니다...")
        self.onenote_list_widget.setEnabled(False)
        self.refresh_button.setEnabled(False)

        self._scanner_worker = OneNoteWindowScanner(self.my_pid)
        self._scanner_worker.done.connect(self._on_onenote_list_ready)
        self._scanner_worker.start()

    def _on_onenote_list_ready(self, results: List[Dict]):
        self.onenote_windows_info = results
        self.onenote_list_widget.clear()

        if not results:
            self.onenote_list_widget.addItem("실행 중인 OneNote 창을 찾지 못했습니다.")
        else:
            items = [f'{r["title"]}  [{r["class_name"]}]' for r in results]
            self.onenote_list_widget.addItems(items)

        self.onenote_list_widget.setEnabled(True)
        self.refresh_button.setEnabled(True)

    def _cache_tree_control(self):
        self.tree_control = _find_tree_or_list(self.onenote_window)
        if self.tree_control:
            try:
                _ = self.tree_control.children()
            except Exception:
                pass

    def _perform_connection(self, info: Dict) -> bool:
        ensure_pywinauto()
        if not _pwa_ready:
            self.update_status_and_ui("pywinauto가 준비되지 않았습니다.", False)
            return False
        try:
            self.onenote_window = Desktop(backend="uia").window(handle=info["handle"])
            if not self.onenote_window.is_visible():
                raise ElementNotFoundError

            window_title = self.onenote_window.window_text()
            save_connection_info(self.onenote_window)

            status_text = f"연결됨: '{window_title}'"
            self.update_status_and_ui(status_text, True)
            QTimer.singleShot(0, self._cache_tree_control)
            return True

        except ElementNotFoundError:
            self.update_status_and_ui("연결 실패: 선택한 창이 보이지 않습니다.", False)
            self.refresh_onenote_list()
            return False
        except Exception as e:
            self.update_status_and_ui(f"연결 실패: {e}", False)
            return False

    def connect_and_center_from_list_item(self, item):
        row = self.onenote_list_widget.currentRow()
        if 0 <= row < len(self.onenote_windows_info):
            info = self.onenote_windows_info[row]
            if self._perform_connection(info):
                QTimer.singleShot(50, self.center_selected_item_action)

    def select_other_window(self):
        dialog = OtherWindowSelectionDialog(self.my_pid, self)
        if dialog.exec():
            info = dialog.selected_info
            if info:
                self._perform_connection(info)

    def disconnect_and_clear_info(self):
        self.onenote_window = None
        self.tree_control = None
        self.update_status_and_ui("연결 해제됨.", False)

        # 연결 시그니처만 지우기
        current_settings = load_settings()
        current_settings["connection_signature"] = None
        save_settings(current_settings)

    def _pre_action_check(self) -> bool:
        ensure_pywinauto()
        if not self.onenote_window:
            self.update_status_and_ui("오류: 앱에 연결되어 있지 않습니다.", False)
            return False
        try:
            if not self.onenote_window.is_visible():
                raise ElementNotFoundError
        except (ElementNotFoundError, AttributeError):
            self.update_status_and_ui(
                "오류: 연결된 창을 찾을 수 없습니다. 연결을 해제합니다.", False
            )
            self.disconnect_and_clear_info()
            return False
        return True

    def center_selected_item_action(self):
        if not self._pre_action_check():
            return

        if not self.tree_control:
            self.tree_control = _find_tree_or_list(self.onenote_window)

        success, item_name = scroll_selected_item_to_center(
            self.onenote_window, self.tree_control
        )

        if success:
            self.update_status_and_ui(f"성공: '{item_name}' 중앙 정렬 완료.", True)
        else:
            # 한번 더 컨테이너 재탐색 후 재시도
            self.tree_control = _find_tree_or_list(self.onenote_window)
            success, item_name = scroll_selected_item_to_center(
                self.onenote_window, self.tree_control
            )
            if success:
                self.update_status_and_ui(f"성공: '{item_name}' 중앙 정렬 완료.", True)
            else:
                self.update_status_and_ui(
                    "실패: 선택 항목을 찾거나 정렬하지 못했습니다.", True
                )

    def _search_and_select_section(self):
        """입력창의 텍스트로 섹션을 검색하고 선택 및 중앙 정렬합니다."""
        if not self._pre_action_check():
            return

        search_text = self.search_input.text().strip()
        if not search_text:
            self.update_status_and_ui("검색할 내용을 입력하세요.", True)
            return

        # tree_control이 캐시되지 않았을 경우를 대비해 다시 한 번 탐색
        if not self.tree_control:
            self.tree_control = _find_tree_or_list(self.onenote_window)

        self.update_status_and_ui(f"'{search_text}' 섹션을 검색 중...", True)

        # 텍스트로 섹션 검색 및 선택
        success = select_section_by_text(
            self.onenote_window, search_text, self.tree_control
        )

        if success:
            # 선택 성공 시, 잠시 후 중앙 정렬 실행 (UI 반영 시간 고려)
            QTimer.singleShot(100, self.center_selected_item_action)
            # 상태바는 즉시 업데이트
            self.update_status_and_ui(f"검색 성공: '{search_text}' 선택 완료.", True)
        else:
            self.update_status_and_ui(
                f"검색 실패: '{search_text}' 섹션을 찾을 수 없습니다.", True
            )

    # ----------------- 15. 즐겨찾기 로드/세이브 -----------------
    def _load_favorites(self):
        self.fav_tree.clear()

        data = self.settings.get("favorites", [])

        if isinstance(data, list):
            for node in data:
                self._append_fav_node(self.fav_tree.invisibleRootItem(), node)
        self.fav_tree.expandAll()

    def _export_favorites(self):
        # 1. 현재 즐겨찾기 데이터를 직렬화합니다.
        self._save_favorites()
        favorites_data = self.settings.get("favorites", [])

        if not favorites_data:
            QMessageBox.information(
                self, "내보내기", "내보낼 즐겨찾기 항목이 없습니다."
            )
            return

        # --- [추가] 현재 날짜와 시간으로 기본 파일 이름 생성 ---
        # 형식: OneNote_Remocon_Favorites_YYYY-MM-DD_HH-MM-SS.json
        timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")
        default_filename = f"OneNote_Remocon_Favorites_{timestamp}.json"
        # --- [추가 완료] ---

        # 2. 파일 저장 대화상자를 엽니다.
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "즐겨찾기 내보내기",
            default_filename,  # --- [수정] 동적으로 생성된 이름 사용 ---
            "JSON Files (*.json);;All Files (*)",
        )

        # 3. 사용자가 파일을 선택한 경우 데이터를 저장합니다.
        if file_path:
            try:
                with open(file_path, "w", encoding="utf-8") as f:
                    json.dump(favorites_data, f, ensure_ascii=False, indent=2)
                QMessageBox.information(
                    self,
                    "성공",
                    f"즐겨찾기를 성공적으로 내보냈습니다.\n\n경로: {file_path}",
                )
            except Exception as e:
                QMessageBox.critical(
                    self, "오류", f"파일을 저장하는 중 오류가 발생했습니다:\n{e}"
                )

    def _import_favorites(self):
        # 1. 파일 열기 대화상자를 바로 엽니다. (확인 창 제거)
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "즐겨찾기 가져오기",
            "",  # 기본 경로
            "JSON Files (*.json);;All Files (*)",
        )

        # 2. 사용자가 파일을 선택한 경우 데이터를 불러옵니다.
        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    imported_data = json.load(f)

                # 3. 데이터 형식이 리스트인지 기본적인 검사를 수행합니다.
                if not isinstance(imported_data, list):
                    raise ValueError("올바른 즐겨찾기 파일 형식이 아닙니다.")

                # 4. 불러온 데이터로 설정과 UI를 업데이트합니다.
                self.settings["favorites"] = imported_data
                self._load_favorites()  # 트리 UI 새로고침
                self._save_favorites()  # 앱의 기본 설정 파일에도 저장
                QMessageBox.information(
                    self, "성공", "즐겨찾기를 성공적으로 가져왔습니다."
                )

            except Exception as e:
                QMessageBox.critical(
                    self, "오류", f"파일을 불러오는 중 오류가 발생했습니다:\n{e}"
                )

    def _save_favorites(self):
        try:
            data = []
            root = self.fav_tree.invisibleRootItem()
            for i in range(root.childCount()):
                data.append(self._serialize_fav_item(root.child(i)))

            current_settings = load_settings()
            current_settings["favorites"] = data
            save_settings(current_settings)

            self.settings["favorites"] = data
        except Exception as e:
            print(f"[ERROR] 즐겨찾기 저장 실패: {e}")

    def _serialize_fav_item(self, item: QTreeWidgetItem) -> Dict[str, Any]:
        node_type = item.data(0, ROLE_TYPE)
        payload = item.data(0, ROLE_DATA) or {}
        node = {
            "type": node_type,
            "id": payload.get("id") or str(uuid.uuid4()),
            "name": item.text(0),
        }
        if node_type == "section":
            node["target"] = payload.get("target", {})
        children = []
        for i in range(item.childCount()):
            children.append(self._serialize_fav_item(item.child(i)))
        if children:
            node["children"] = children
        return node

    def _append_fav_node(
        self, parent: QTreeWidgetItem, node: Dict[str, Any]
    ) -> QTreeWidgetItem:
        item = QTreeWidgetItem(parent)
        node_type = node.get("type", "group")
        name = node.get("name", "이름 없음")
        item.setText(0, name)
        item.setData(0, ROLE_TYPE, node_type)
        payload = {"id": node.get("id", str(uuid.uuid4()))}
        if node_type == "section":
            payload["target"] = node.get("target", {})
            item.setIcon(
                0,
                self.style().standardIcon(
                    QApplication.style().StandardPixmap.SP_FileIcon
                ),
            )
        else:
            item.setIcon(
                0,
                self.style().standardIcon(
                    QApplication.style().StandardPixmap.SP_DirIcon
                ),
            )
        item.setData(0, ROLE_DATA, payload)
        item.setFlags(
            item.flags()
            | Qt.ItemFlag.ItemIsEditable
            | Qt.ItemFlag.ItemIsDragEnabled
            | Qt.ItemFlag.ItemIsDropEnabled
            | Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsSelectable
        )
        for ch in node.get("children", []):
            self._append_fav_node(item, ch)
        return item

    # ----------------- 16. 즐겨찾기 조작 -----------------
    def _current_fav_item(self) -> Optional[QTreeWidgetItem]:
        items = self.fav_tree.selectedItems()
        return items[0] if items else None

    def _move_item_up(self):
        item = self._current_fav_item()
        if not item:
            return

        parent = item.parent() or self.fav_tree.invisibleRootItem()
        index = parent.indexOfChild(item)

        if index > 0:
            # --- [추가] ---
            # 이동 전, 아이템의 펼침 상태를 저장합니다.
            is_expanded = item.isExpanded()
            # --- [추가 완료] ---

            # 아이템을 잠시 떼어낸 후, 한 칸 위 인덱스에 다시 삽입합니다.
            taken_item = parent.takeChild(index)
            parent.insertChild(index - 1, taken_item)

            # --- [추가] ---
            # 이동 후, 저장해둔 펼침 상태를 복원합니다.
            taken_item.setExpanded(is_expanded)
            # --- [추가 완료] ---

            # 이동된 아이템을 다시 선택하여 포커스를 유지합니다.
            self.fav_tree.setCurrentItem(taken_item)
            self._save_favorites()
            self._update_move_button_state()

    def _move_item_down(self):
        item = self._current_fav_item()
        if not item:
            return

        parent = item.parent() or self.fav_tree.invisibleRootItem()
        index = parent.indexOfChild(item)

        # 마지막 아이템이 아닌 경우에만 이동합니다.
        if index < parent.childCount() - 1:
            # --- [추가] ---
            # 이동 전, 아이템의 펼침 상태를 저장합니다.
            is_expanded = item.isExpanded()
            # --- [추가 완료] ---

            taken_item = parent.takeChild(index)
            parent.insertChild(index + 1, taken_item)

            # --- [추가] ---
            # 이동 후, 저장해둔 펼침 상태를 복원합니다.
            taken_item.setExpanded(is_expanded)
            # --- [추가 완료] ---

            self.fav_tree.setCurrentItem(taken_item)
            self._save_favorites()
            self._update_move_button_state()

    def _update_move_button_state(self):
        item = self._current_fav_item()

        # 선택된 아이템이 없으면 두 버튼 모두 비활성화합니다.
        if not item:
            self.btn_move_up.setEnabled(False)
            self.btn_move_down.setEnabled(False)
            return

        parent = item.parent() or self.fav_tree.invisibleRootItem()
        index = parent.indexOfChild(item)

        # 첫 번째 아이템이면 '위로' 버튼을 비활성화합니다.
        self.btn_move_up.setEnabled(index > 0)
        # 마지막 아이템이면 '아래로' 버튼을 비활성화합니다.
        self.btn_move_down.setEnabled(index < parent.childCount() - 1)

    def _add_group(self):
        parent = self._current_fav_item()
        if parent and parent.data(0, ROLE_TYPE) == "section":
            parent = parent.parent()
        parent = parent or self.fav_tree.invisibleRootItem()
        node = {"type": "group", "name": "새 그룹", "children": []}
        item = self._append_fav_node(parent, node)
        self.fav_tree.editItem(item, 0)
        self._save_favorites()

    def _add_section_from_current(self):
        if not self.onenote_window:
            QMessageBox.information(self, "안내", "먼저 연결된 창이 있어야 합니다.")
            return

        title = ""
        try:
            title = self.onenote_window.window_text()
        except Exception:
            pass

        section_text = None
        try:
            tc = self.tree_control or _find_tree_or_list(self.onenote_window)
            if tc:
                sel = get_selected_tree_item_fast(tc)
                if sel:
                    section_text = sel.window_text()
        except Exception:
            pass

        default_name = section_text or title or "새 섹션"
        name, ok = QInputDialog.getText(
            self, "섹션 즐겨찾기 추가", "표시 이름:", text=default_name
        )
        if not ok or not name.strip():
            return

        try:
            sig = build_window_signature(self.onenote_window)
        except Exception:
            sig = {}

        target = {"sig": sig, "section_text": section_text}
        node = {"type": "section", "name": name.strip(), "target": target}

        parent = self._current_fav_item()
        if parent and parent.data(0, ROLE_TYPE) == "section":
            parent = parent.parent()
        parent = parent or self.fav_tree.invisibleRootItem()
        self._append_fav_node(parent, node)
        self._save_favorites()

    def _add_section_from_other_window(self):
        dialog = OtherWindowSelectionDialog(self.my_pid, self)
        if not dialog.exec():
            return
        info = dialog.selected_info
        if not info:
            return

        default_name = (info.get("title") or "새 섹션").strip() or "새 섹션"
        name, ok = QInputDialog.getText(
            self, "섹션 즐겨찾기 추가", "표시 이름:", text=default_name
        )
        if not ok or not name.strip():
            return

        try:
            ensure_pywinauto()
            win = Desktop(backend="uia").window(handle=info["handle"])
            sig = build_window_signature(win)
        except Exception:
            sig = {
                "handle": info.get("handle"),
                "pid": info.get("pid"),
                "class_name": info.get("class_name"),
                "title": info.get("title"),
            }
        target = {"sig": sig, "section_text": None}
        node = {"type": "section", "name": name.strip(), "target": target}

        parent = self._current_fav_item()
        if parent and parent.data(0, ROLE_TYPE) == "section":
            parent = parent.parent()
        parent = parent or self.fav_tree.invisibleRootItem()
        self._append_fav_node(parent, node)
        self._save_favorites()

    def _rename_favorite_item(self):
        item = self._current_fav_item()
        if not item:
            return
        self.fav_tree.editItem(item, 0)

    def _delete_favorite_item(self):
        item = self._current_fav_item()
        if not item:
            return
        node_type = item.data(0, ROLE_TYPE)
        name = item.text(0)

        if node_type == "group" and item.childCount() > 0:
            ret = QMessageBox.question(
                self,
                "삭제 확인",
                f"그룹 '{name}'과(와) 모든 하위 항목을 삭제할까요?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if ret != QMessageBox.StandardButton.Yes:
                return
        else:
            ret = QMessageBox.question(
                self,
                "삭제 확인",
                f"'{name}'을(를) 삭제할까요?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if ret != QMessageBox.StandardButton.Yes:
                return

        parent = item.parent() or self.fav_tree.invisibleRootItem()
        parent.removeChild(item)
        self._save_favorites()

    def _on_fav_context_menu(self, pos):
        item = self._current_fav_item()
        menu = QMenu(self)

        act_add_group = QAction("그룹 추가", self)
        act_add_group.triggered.connect(self._add_group)
        menu.addAction(act_add_group)

        act_add_curr = QAction("현재 섹션 추가", self)
        act_add_curr.triggered.connect(self._add_section_from_current)
        menu.addAction(act_add_curr)

        act_add_other = QAction("다른 창 추가", self)
        act_add_other.triggered.connect(self._add_section_from_other_window)
        menu.addAction(act_add_other)

        if item:
            menu.addSeparator()
            act_rename = QAction("이름 바꾸기 (F2)", self)
            act_rename.triggered.connect(self._rename_favorite_item)
            menu.addAction(act_rename)

            act_delete = QAction("삭제 (Del)", self)
            act_delete.triggered.connect(self._delete_favorite_item)
            menu.addAction(act_delete)

        menu.exec(self.fav_tree.viewport().mapToGlobal(pos))

    def _on_fav_item_double_clicked(self, item: QTreeWidgetItem, col: int):
        if not item:
            return
        node_type = item.data(0, ROLE_TYPE)
        if node_type != "section":
            return
        self._activate_favorite_section(item)

    def _activate_favorite_section(self, item: QTreeWidgetItem):
        ensure_pywinauto()
        if not _pwa_ready:
            self.update_status_and_ui(
                "오류: 자동화 모듈이 로드되지 않았습니다.",
                self.center_button.isEnabled(),
            )
            return

        payload = item.data(0, ROLE_DATA) or {}
        target = payload.get("target") or {}
        display_name = item.text(0)

        sig = target.get("sig") or {}
        if not sig:
            self.update_status_and_ui(
                "오류: 즐겨찾기에 대상 창 정보가 없습니다.",
                self.center_button.isEnabled(),
            )
            return

        win = reacquire_window_by_signature(sig)
        if not win:
            self.update_status_and_ui(
                f"실패: 대상 창 '{display_name}'을(를) 찾을 수 없습니다.",
                self.center_button.isEnabled(),
            )
            return

        try:
            win.set_focus()
        except Exception:
            pass

        try:
            info = {
                "handle": win.handle,
                "title": win.window_text(),
                "class_name": win.class_name(),
                "pid": win.process_id(),
            }
            connected = self._perform_connection(info)
        except Exception:
            connected = False

        if connected and self._auto_center_after_activate:
            exe_name = (sig.get("exe_name") or "").lower()
            if "onenote" in exe_name or "onenote" in (sig.get("title") or "").lower():
                section_text = target.get("section_text")
                if section_text:
                    ok = select_section_by_text(
                        self.onenote_window, section_text, self.tree_control
                    )
                    if ok:
                        QTimer.singleShot(
                            500,
                            lambda: scroll_selected_item_to_center(
                                self.onenote_window, self.tree_control
                            ),
                        )
                        self.update_status_and_ui(f"활성화: '{display_name}'", True)
                    # --- [핵심 수정 로직] ---
                    else:
                        # 실패 시: 이름 변경 및 상태바 업데이트 (알림창 제거)
                        current_name = item.text(0)

                        # 아직 (구) 접미사가 없는 경우에만 이름 변경 수행
                        if not current_name.startswith("(구) "):
                            new_name = f"(구) {current_name}"
                            item.setText(0, new_name)
                            # 변경된 이름을 설정 파일에 즉시 저장
                            self._save_favorites()

                            # 상태바에 실패 및 이름 변경 사실을 알림
                            status_message = (
                                f"섹션 찾기 실패: '{new_name}'(으)로 변경됨"
                            )
                            self.update_status_and_ui(status_message, True)
                        else:
                            # 이미 (구) 접미사가 있는 경우, 실패 사실만 알림
                            status_message = (
                                f"섹션 찾기 실패: '{current_name}' 섹션을 찾을 수 없음"
                            )
                            self.update_status_and_ui(status_message, True)
                    # --- [수정 로직 끝] ---
                    return

        self.update_status_and_ui(f"활성화: '{display_name}'", True)


# ----------------- 17. 엔트리 포인트 -----------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = OneNoteScrollRemoconApp()
    ex.show()

    # 이중 연결 방지 및 그룹 토글 기능 복원
    try:
        ex.fav_tree.itemDoubleClicked.disconnect()
    except TypeError:
        pass

    def _toggle_group_and_activate_section(item, col):
        node_type = item.data(0, ROLE_TYPE)
        if node_type != "section":
            item.setExpanded(not item.isExpanded())
        else:
            ex._on_fav_item_double_clicked(item, col)

    ex.fav_tree.itemDoubleClicked.connect(_toggle_group_and_activate_section)

    sys.exit(app.exec())
