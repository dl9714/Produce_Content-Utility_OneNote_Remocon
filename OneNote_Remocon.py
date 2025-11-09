# -*- coding: utf-8 -*-
import sys
import json
import os
import time
import uuid
import ctypes
from ctypes import wintypes
from typing import Optional, List, Dict, Any
import base64

from PyQt6.QtWidgets import (
    QApplication,
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
    QMainWindow,
    QFileDialog,
    QWidget,
    QLineEdit,
)
from PyQt6.QtCore import (
    QThread,
    pyqtSignal,
    QTimer,
    Qt,
    QSettings,
    QEvent,
    QRect,
    QByteArray,
)
from PyQt6.QtGui import QIcon, QAction

# ----------------- 0. 전역 상수 -----------------
SETTINGS_FILE = "OneNote_Remocon_Setting.json"
APP_ICON_PATH = "app_icon.ico"

ONENOTE_CLASS_NAME = "ApplicationFrameWindow"
SCROLL_STEP_SENSITIVITY = 40

ROLE_TYPE = Qt.ItemDataRole.UserRole + 1
ROLE_DATA = Qt.ItemDataRole.UserRole + 2


# ----------------- 0.0 설정 파일 경로 헬퍼 -----------------
def _get_settings_file_path() -> str:
    """
    설정 파일(쓰기 가능)의 경로를 반환합니다.
    - PyInstaller로 패키징된 경우: 실행 파일(.exe)이 위치한 디렉토리
    - 스크립트 실행인 경우: 현재 작업 디렉토리
    """
    # sys.frozen은 PyInstaller에 의해 생성된 실행 파일인지 확인하는 일반적인 방법입니다.
    if getattr(sys, "frozen", False):
        # 실행 파일(.exe)이 있는 디렉토리
        base_path = os.path.dirname(sys.executable)
    else:
        # 스크립트 실행 환경 (현재 작업 디렉토리)
        base_path = os.path.abspath(".")

    return os.path.join(base_path, SETTINGS_FILE)


# ----------------- 0.0 설정 파일 로드/저장 유틸리티 (즐겨찾기 버퍼 구조 추가) -----------------
DEFAULT_SETTINGS = {
    "window_geometry": {"x": 200, "y": 180, "width": 960, "height": 540},
    "splitter_states": None,  # 새 설정 항목 추가
    "connection_signature": None,
    "favorites_buffers": {"기본 즐겨찾기 버퍼": []},
    "active_buffer": "기본 즐겨찾기 버퍼",
}


def load_settings() -> Dict[str, Any]:
    # 설정 파일 경로를 실행 파일 위치 기준으로 가져옴
    settings_path = _get_settings_file_path()

    if not os.path.exists(settings_path):
        return DEFAULT_SETTINGS.copy()
    try:
        with open(settings_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        # 하위 호환성을 위한 마이그레이션 로직
        if "favorites" in data and "favorites_buffers" not in data:
            print(
                "[INFO] 구 버전 설정 감지. 새 즐겨찾기 버퍼 구조로 마이그레이션합니다."
            )
            data["favorites_buffers"] = {"기본 즐겨찾기 버퍼": data["favorites"]}
            data["active_buffer"] = "기본 즐겨찾기 버퍼"
            del data["favorites"]

        settings = DEFAULT_SETTINGS.copy()
        settings.update(data)
        return settings
    except Exception as e:
        print(f"[ERROR] 설정 파일 로드 실패: {e}")
        return DEFAULT_SETTINGS.copy()


def save_settings(data: Dict[str, Any]):
    # 설정 파일 경로를 실행 파일 위치 기준으로 가져옴
    settings_path = _get_settings_file_path()

    try:
        if "favorites" in data:
            del data["favorites"]

        with open(settings_path, "w", encoding="utf-8") as f:
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
    # NameError 수정: _ppa_ready -> _pwa_ready
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
        pass


# ----------------- 0.2 Win32 빠른 창 열거 -----------------
_user32 = ctypes.windll.user32


def _win_get_window_text(hwnd):
    length = _user32.GetWindowTextLengthW(hwnd)
    buf = ctypes.create_unicode_buffer(length + 1 if length > 0 else 1)
    _user32.GetWindowTextW(hwnd, buf, len(buf))
    return buf.value


def _win_get_class_name(hwnd):
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
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# ----------------- 1. 프로세스 실행 파일 경로 얻기 -----------------
def get_process_image_path(pid: int) -> Optional[str]:
    if not pid:
        return None

    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

    # 64비트 안전: use_last_error로 WinAPI 에러 사용 가능
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    OpenProcess = kernel32.OpenProcess
    OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    OpenProcess.restype = wintypes.HANDLE

    QueryFullProcessImageNameW = kernel32.QueryFullProcessImageNameW
    QueryFullProcessImageNameW.argtypes = [
        wintypes.HANDLE,
        wintypes.DWORD,
        wintypes.LPWSTR,
        ctypes.POINTER(wintypes.DWORD),
    ]
    QueryFullProcessImageNameW.restype = wintypes.BOOL

    CloseHandle = kernel32.CloseHandle
    CloseHandle.argtypes = [wintypes.HANDLE]
    CloseHandle.restype = wintypes.BOOL

    hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not hProcess:
        return None
    try:
        # 1차 버퍼
        size = 512
        while True:
            buf_len = wintypes.DWORD(size)
            buf = ctypes.create_unicode_buffer(buf_len.value)
            ok = QueryFullProcessImageNameW(hProcess, 0, buf, ctypes.byref(buf_len))
            if ok:
                return buf.value
            # 버퍼 부족 시 한 번 정도 키워 봄
            err = ctypes.get_last_error()
            # ERROR_INSUFFICIENT_BUFFER = 122
            if err == 122 and size < 4096:
                size *= 2
                continue
            return None
    finally:
        CloseHandle(hProcess)


# ----------------- 1.1 엄격한 OneNote 창 검증 헬퍼 -----------------
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


# ----------------- 6. 마우스 휠 기반 스크롤(폴백) -----------------
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


# ----------------- 7. 선택 항목을 가장 빠르게 얻기 -----------------
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


# ----------------- 8.1 지정 텍스트 섹션 찾기/선택 -----------------
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


# ----------------- 10. 선택된 항목을 중앙으로 스크롤 -----------------
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
    copyRequested = pyqtSignal()
    pasteRequested = pyqtSignal()

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
        source_item = self.currentItem()
        target_item = self.itemAt(event.position().toPoint())

        if not source_item:
            event.ignore()
            return

        super().dropEvent(event)

        if target_item and source_item.parent() == target_item:
            source_type = source_item.data(0, ROLE_TYPE)
            target_type = target_item.data(0, ROLE_TYPE)

            if source_type == "section" and target_type == "section":
                moved_item = target_item.takeChild(
                    target_item.indexOfChild(source_item)
                )

                if moved_item:
                    parent_of_target = target_item.parent()
                    if not parent_of_target:
                        parent_of_target = self.invisibleRootItem()
                    target_index = parent_of_target.indexOfChild(target_item)
                    parent_of_target.insertChild(target_index + 1, moved_item)
                    self.setCurrentItem(moved_item)

        self.structureChanged.emit()

    def keyPressEvent(self, event):
        # Delete, F2 처리
        if event.key() == Qt.Key.Key_Delete:
            self.deleteRequested.emit()
            event.accept()
            return
        if event.key() == Qt.Key.Key_F2:
            self.renameRequested.emit()
            event.accept()
            return

        # Ctrl+C / Ctrl+V 처리
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            if event.key() == Qt.Key.Key_C:
                self.copyRequested.emit()
                event.accept()
                return
            if event.key() == Qt.Key.Key_V:
                self.pasteRequested.emit()
                event.accept()
                return

        super().keyPressEvent(event)


# ----------------- 14. PyQt GUI -----------------
class OneNoteScrollRemoconApp(QMainWindow):
    def __init__(self):
        super().__init__()
        ensure_pywinauto()

        # 1. 설정 로드 및 창 위치/상태 복원
        self.settings = load_settings()
        self.onenote_window = None
        self.tree_control = None
        self._reconnect_worker = None
        self._scanner_worker = None
        self.onenote_windows_info: List[Dict] = []
        self.my_pid = os.getpid()
        self._auto_center_after_activate = True
        self.active_buffer_name = None

        # --- [START] 창 위치 복원 및 유효성 검사 로직 (수정됨) ---
        geo_settings = self.settings.get(
            "window_geometry", DEFAULT_SETTINGS["window_geometry"]
        )

        # 주 모니터의 사용 가능한 영역 가져오기 (작업 표시줄 제외)
        primary_screen = QApplication.primaryScreen()
        if not primary_screen:  # 헤드리스 환경 등 예외 처리
            # 기본 가상 화면 크기 설정
            screen_rect = QRect(0, 0, 1920, 1080)
        else:
            screen_rect = primary_screen.availableGeometry()

        # 저장된 창 위치 QRect 객체로 생성
        window_rect = QRect(
            geo_settings.get("x", 200),
            geo_settings.get("y", 180),
            geo_settings.get("width", 960),
            geo_settings.get("height", 540),
        )

        # 창이 화면에 보이는지 확인 (최소 100x50 픽셀이 보여야 함)
        intersection = screen_rect.intersected(window_rect)
        is_visible = intersection.width() >= 100 and intersection.height() >= 50

        if not is_visible:
            # 창이 화면 밖에 있으면 화면 중앙으로 이동
            # 창 크기는 유지하되, 화면 크기보다 크지 않도록 조정
            window_rect.setWidth(min(window_rect.width(), screen_rect.width()))
            window_rect.setHeight(min(window_rect.height(), screen_rect.height()))
            # 중앙 정렬
            window_rect.moveCenter(screen_rect.center())

        self.setGeometry(window_rect)
        # --- [END] 창 위치 복원 및 유효성 검사 로직 ---

        # 즐겨찾기 복사 데이터 임시 저장소 (클립보드 역할)
        self.clipboard_data: Optional[Dict] = None

        # 즐겨찾기 버퍼 복사 데이터 임시 저장소
        self.buffer_clipboard_data: Optional[Dict] = None

        # 1.1 애플리케이션 아이콘 설정
        icon_path = resource_path(APP_ICON_PATH)
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.init_ui("준비됨 (자동 재연결 중...)")

        # 2. 즐겨찾기 버퍼 및 즐겨찾기 로드
        self._load_buffers_and_favorites()

        self.fav_tree.deleteRequested.connect(self._delete_favorite_item)
        self.fav_tree.renameRequested.connect(self._rename_favorite_item)

        # 복사/붙여넣기 시그널 연결
        self.fav_tree.copyRequested.connect(self._copy_favorite_item)
        self.fav_tree.pasteRequested.connect(self._paste_favorite_item)

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

        # --- 스타일시트 정의 (생략) ---
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
            QLineEdit {{
                background-color: {COLOR_LIST_BG};
                border: 1px solid {COLOR_SECONDARY_BUTTON};
                border-radius: 4px;
                padding: 4px 8px;
            }}
            QLineEdit:focus {{
                border: 1px solid {COLOR_LIST_SELECTED};
            }}
        """
        )

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(10)

        self.main_splitter = QSplitter(Qt.Orientation.Horizontal)  # self로 저장
        self.main_splitter.setChildrenCollapsible(False)
        main_layout.addWidget(self.main_splitter, stretch=1)

        self.left_splitter = QSplitter(Qt.Orientation.Horizontal)  # self로 저장
        self.left_splitter.setChildrenCollapsible(False)

        # 1. 즐겨찾기 버퍼 관리 패널 (가장 왼쪽)
        buffer_panel = QWidget()
        buffer_layout = QVBoxLayout(buffer_panel)
        buffer_layout.setContentsMargins(0, 0, 0, 0)
        buffer_layout.setSpacing(8)

        buffer_group = QGroupBox("즐겨찾기 버퍼")
        buffer_group_layout = QVBoxLayout(buffer_group)

        # 즐겨찾기 버퍼 상단 툴바: 추가, 이름변경
        buffer_toolbar_top_layout = QHBoxLayout()
        self.btn_add_buffer = QToolButton()
        self.btn_add_buffer.setText("추가")
        self.btn_add_buffer.clicked.connect(self._add_buffer)

        self.btn_rename_buffer = QToolButton()
        self.btn_rename_buffer.setText("이름변경")
        self.btn_rename_buffer.clicked.connect(self._rename_buffer)

        buffer_toolbar_top_layout.addWidget(self.btn_add_buffer)
        buffer_toolbar_top_layout.addWidget(self.btn_rename_buffer)
        buffer_toolbar_top_layout.addStretch(1)
        buffer_group_layout.addLayout(buffer_toolbar_top_layout)

        self.buffer_list_widget = QListWidget()
        self.buffer_list_widget.currentItemChanged.connect(self._on_buffer_changed)
        self.buffer_list_widget.itemSelectionChanged.connect(
            self._update_buffer_move_button_state
        )
        self.buffer_list_widget.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu
        )
        self.buffer_list_widget.customContextMenuRequested.connect(
            self._on_buffer_context_menu
        )

        buffer_group_layout.addWidget(self.buffer_list_widget)

        # 즐겨찾기 버퍼 하단 툴바: 삭제, 위로, 아래로
        buffer_toolbar_bottom_layout = QHBoxLayout()
        self.btn_delete_buffer = QToolButton()
        self.btn_delete_buffer.setText("삭제")
        self.btn_delete_buffer.clicked.connect(self._delete_buffer)

        self.btn_buffer_move_up = QToolButton()
        self.btn_buffer_move_up.setText("▲ 위로")
        self.btn_buffer_move_up.clicked.connect(self._move_buffer_up)

        self.btn_buffer_move_down = QToolButton()
        self.btn_buffer_move_down.setText("▼ 아래로")
        self.btn_buffer_move_down.clicked.connect(self._move_buffer_down)

        buffer_toolbar_bottom_layout.addWidget(self.btn_delete_buffer)
        buffer_toolbar_bottom_layout.addStretch(1)
        buffer_toolbar_bottom_layout.addWidget(self.btn_buffer_move_up)
        buffer_toolbar_bottom_layout.addWidget(self.btn_buffer_move_down)
        buffer_group_layout.addLayout(buffer_toolbar_bottom_layout)

        buffer_layout.addWidget(buffer_group)
        self.left_splitter.addWidget(buffer_panel)

        # 2. 즐겨찾기 관리 패널 (중앙)
        favorites_panel = QWidget()
        left_layout = QVBoxLayout(favorites_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        fav_group = QGroupBox("즐겨찾기")
        fav_layout = QVBoxLayout(fav_group)

        # 툴바 - 1행: 그룹추가, 현재 섹션 추가, 이름 바꾸기
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
        tb1_layout.addWidget(self.btn_add_section_current)
        tb1_layout.addWidget(self.btn_rename)
        tb1_layout.addStretch(1)

        # 툴바 - 2행: 그룹 펼치기, 접기 (삭제 버튼은 하단으로 이동)
        tb2_layout = QHBoxLayout()
        self.btn_expand_all = QToolButton()
        self.btn_expand_all.setText("그룹 펼치기")
        self.btn_collapse_all = QToolButton()
        self.btn_collapse_all.setText("그룹 접기")
        tb2_layout.addWidget(self.btn_add_group)
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

        # 즐겨찾기 하단 툴바: 삭제, 위로, 아래로 (삭제 버튼 재배치)
        move_buttons_layout = QHBoxLayout()

        # 삭제 버튼 (tb2에서 이동)
        self.btn_delete = QToolButton()
        self.btn_delete.setText("삭제 (Del)")
        self.btn_delete.clicked.connect(self._delete_favorite_item)
        move_buttons_layout.addWidget(self.btn_delete)

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

        self.fav_tree.itemSelectionChanged.connect(self._update_move_button_state)
        left_layout.addWidget(fav_group, stretch=1)

        self.left_splitter.addWidget(favorites_panel)
        self.main_splitter.addWidget(self.left_splitter)

        # 3. 오른쪽 패널 (변경 없음)
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

        search_group = QGroupBox("전자필기장 검색")
        search_group_layout = QVBoxLayout(search_group)

        search_widget_layout = QHBoxLayout()

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("검색할 섹션 이름 입력...")
        self.search_input.returnPressed.connect(self._search_and_select_section)
        self.search_input.setEnabled(False)
        search_widget_layout.addWidget(self.search_input, stretch=1)

        self.search_button = QPushButton("전자필기장 위치")
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

        right_layout.addStretch(1)
        self.main_splitter.addWidget(right_panel)

        self.connection_status_label = QLabel(initial_status)
        self.statusBar().addPermanentWidget(self.connection_status_label)
        self.statusBar().setStyleSheet(f"background-color: {COLOR_STATUS_BAR};")

        # --- [START] 스플리터 상태 복원 로직 (수정됨) ---
        # 저장된 스플리터 상태 불러오기
        splitter_states = self.settings.get("splitter_states")
        restored = False
        if isinstance(splitter_states, dict):
            try:
                main_state_b64 = splitter_states.get("main")
                if main_state_b64:
                    main_state = QByteArray.fromBase64(main_state_b64.encode("ascii"))
                    if not main_state.isEmpty():
                        self.main_splitter.restoreState(main_state)

                left_state_b64 = splitter_states.get("left")
                if left_state_b64:
                    left_state = QByteArray.fromBase64(left_state_b64.encode("ascii"))
                    if not left_state.isEmpty():
                        self.left_splitter.restoreState(left_state)

                restored = True
            except Exception as e:
                print(f"[WARN] 스플리터 상태 복원 실패: {e}")
                restored = False

        # 복원에 실패했거나 저장된 상태가 없으면 기본값으로 설정
        if not restored:
            self.left_splitter.setSizes([150, 250])
            self.main_splitter.setSizes([400, 560])
        # --- [END] 스플리터 상태 복원 로직 ---

        self._update_move_button_state()
        self._update_buffer_move_button_state()

    # ----------------- 14.1 창 닫기 이벤트 핸들러 (Geometry/Favorites 저장) -----------------
    def _save_window_state(self):
        """창 지오메트리와 스플리터 상태를 self.settings에 업데이트하고 파일에 즉시 저장합니다."""
        # self.settings (메모리)를 직접 수정합니다. load_settings()를 호출하지 않습니다.
        # 이렇게 함으로써 다른 세션 변경사항이 유지됩니다.
        if not self.isMinimized() and not self.isMaximized():
            geom = self.geometry()
            self.settings["window_geometry"] = {
                "x": geom.x(),
                "y": geom.y(),
                "width": geom.width(),
                "height": geom.height(),
            }

        try:
            self.settings["splitter_states"] = {
                "main": self.main_splitter.saveState()
                .toBase64()
                .data()
                .decode("ascii"),
                "left": self.left_splitter.saveState()
                .toBase64()
                .data()
                .decode("ascii"),
            }
        except Exception as e:
            print(f"[WARN] 스플리터 상태 저장 실패: {e}")

        # 수정된 self.settings 객체 전체를 파일에 저장합니다.
        # 즐겨찾기 등 다른 모든 변경사항도 함께 저장됩니다.
        save_settings(self.settings)

    def closeEvent(self, event):
        self._save_window_state()  # 변경된 함수 호출
        self._save_favorites()
        super().closeEvent(event)

    def update_status_and_ui(self, status_text: str, is_connected: bool):
        self.connection_status_label.setText(status_text)
        self.center_button.setEnabled(is_connected)
        self.search_input.setEnabled(is_connected)
        self.search_button.setEnabled(is_connected)

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

        if not self.tree_control:
            self.tree_control = _find_tree_or_list(self.onenote_window)

        self.update_status_and_ui(f"'{search_text}' 섹션을 검색 중...", True)

        success = select_section_by_text(
            self.onenote_window, search_text, self.tree_control
        )

        if success:
            QTimer.singleShot(100, self.center_selected_item_action)
            self.update_status_and_ui(f"검색 성공: '{search_text}' 선택 완료.", True)
        else:
            self.update_status_and_ui(
                f"검색 실패: '{search_text}' 섹션을 찾을 수 없습니다.", True
            )

    # ----------------- 15. 즐겨찾기 로드/세이브 (즐겨찾기 버퍼 시스템 적용) -----------------
    def _load_buffers_and_favorites(self, select_row: int = -1):
        """설정에서 즐겨찾기 버퍼 목록을 불러와 UI에 표시하고, 활성화된 즐겨찾기 버퍼의 즐겨찾기를 로드합니다."""
        self.buffer_list_widget.blockSignals(True)
        self.buffer_list_widget.clear()

        buffers = self.settings.get("favorites_buffers", {})
        if not buffers:
            buffers = DEFAULT_SETTINGS["favorites_buffers"]
            self.settings["favorites_buffers"] = buffers

        buffer_names = list(buffers.keys())
        self.buffer_list_widget.addItems(buffer_names)

        active_buffer = self.settings.get("active_buffer")
        target_row = -1

        if select_row != -1:
            if 0 <= select_row < len(buffer_names):
                target_row = select_row
                active_buffer = buffer_names[select_row]
        elif active_buffer in buffer_names:
            target_row = buffer_names.index(active_buffer)
        elif self.buffer_list_widget.count() > 0:
            target_row = 0
            active_buffer = buffer_names[0]

        if target_row != -1:
            self.buffer_list_widget.setCurrentRow(target_row)
            self.active_buffer_name = active_buffer

        self.buffer_list_widget.blockSignals(False)
        self._update_buffer_move_button_state()

        if self.active_buffer_name:
            self._load_tree_from_buffer(self.active_buffer_name)

    def _load_tree_from_buffer(self, buffer_name: str):
        """지정된 즐겨찾기 버퍼 이름에 해당하는 즐겨찾기 데이터를 트리에 로드합니다."""
        self.fav_tree.clear()
        self.active_buffer_name = buffer_name

        favorites_data = self.settings.get("favorites_buffers", {}).get(buffer_name, [])

        if isinstance(favorites_data, list):
            for node in favorites_data:
                self._append_fav_node(self.fav_tree.invisibleRootItem(), node)

        self.fav_tree.expandAll()
        self.settings["active_buffer"] = buffer_name
        self._save_settings_to_file()

    def _save_favorites(self):
        """현재 활성화된 즐겨찾기 버퍼의 즐겨찾기 트리 상태를 설정에 저장합니다."""
        if not self.active_buffer_name:
            return

        try:
            data = []
            root = self.fav_tree.invisibleRootItem()
            for i in range(root.childCount()):
                data.append(self._serialize_fav_item(root.child(i)))

            if "favorites_buffers" not in self.settings:
                self.settings["favorites_buffers"] = {}
            self.settings["favorites_buffers"][self.active_buffer_name] = data

            self._save_settings_to_file()

        except Exception as e:
            print(f"[ERROR] 즐겨찾기 저장 실패: {e}")

    def _save_settings_to_file(self):
        """현재 self.settings 객체를 파일에 저장합니다."""
        save_settings(self.settings)

    def _export_favorites(self):
        self._save_favorites()
        if not self.active_buffer_name:
            QMessageBox.warning(self, "내보내기", "활성화된 즐겨찾기 버퍼가 없습니다.")
            return

        favorites_data = self.settings["favorites_buffers"].get(
            self.active_buffer_name, []
        )

        if not favorites_data:
            QMessageBox.information(
                self, "내보내기", "내보낼 즐겨찾기 항목이 없습니다."
            )
            return

        timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")
        default_filename = (
            f"OneNote_Favorites_{self.active_buffer_name}_{timestamp}.json"
        )

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "현재 즐겨찾기 버퍼 즐겨찾기 내보내기",
            default_filename,
            "JSON Files (*.json);;All Files (*)",
        )

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
        reply = QMessageBox.question(
            self,
            "즐겨찾기 가져오기",
            "새 즐겨찾기 버퍼를 만들어 가져오시겠습니까?\n\n"
            "(아니오를 선택하면 현재 활성 즐겨찾기 버퍼에 덮어씁니다)",
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No
            | QMessageBox.StandardButton.Cancel,
        )

        if reply == QMessageBox.StandardButton.Cancel:
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self, "즐겨찾기 가져오기", "", "JSON Files (*.json);;All Files (*)"
        )

        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    imported_data = json.load(f)

                if not isinstance(imported_data, list):
                    raise ValueError("올바른 즐겨찾기 파일 형식이 아닙니다.")

                if reply == QMessageBox.StandardButton.Yes:
                    base_name = os.path.splitext(os.path.basename(file_path))[0]
                    new_buffer_name, ok = QInputDialog.getText(
                        self,
                        "새 즐겨찾기 버퍼 이름",
                        "가져올 즐겨찾기 버퍼의 이름을 입력하세요:",
                        text=base_name,
                    )
                    if not ok or not new_buffer_name:
                        return

                    self.settings["favorites_buffers"][new_buffer_name] = imported_data
                    self._load_buffers_and_favorites()
                    items = self.buffer_list_widget.findItems(
                        new_buffer_name, Qt.MatchFlag.MatchExactly
                    )
                    if items:
                        self.buffer_list_widget.setCurrentItem(items[0])

                else:
                    if not self.active_buffer_name:
                        QMessageBox.critical(
                            self,
                            "오류",
                            "활성화된 즐겨찾기 버퍼가 없어 가져올 수 없습니다.",
                        )
                        return
                    self.settings["favorites_buffers"][
                        self.active_buffer_name
                    ] = imported_data
                    self._load_tree_from_buffer(self.active_buffer_name)

                self._save_settings_to_file()
                QMessageBox.information(
                    self, "성공", "즐겨찾기를 성공적으로 가져왔습니다."
                )

            except Exception as e:
                QMessageBox.critical(
                    self, "오류", f"파일을 불러오는 중 오류가 발생했습니다:\n{e}"
                )

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

    # ----------------- 15-2. 즐겨찾기 버퍼 순서 변경 로직 -----------------
    def _update_buffer_move_button_state(self):
        """즐겨찾기 버퍼 목록에서 선택된 항목에 따라 이동 버튼 활성화 상태를 업데이트합니다."""
        row = self.buffer_list_widget.currentRow()
        count = self.buffer_list_widget.count()

        is_selected = row != -1

        if hasattr(self, "btn_buffer_move_up"):
            self.btn_buffer_move_up.setEnabled(is_selected and row > 0)
            self.btn_buffer_move_down.setEnabled(is_selected and row < count - 1)

    def _move_buffer_up(self):
        """선택된 즐겨찾기 버퍼를 위로 이동시킵니다."""
        current_row = self.buffer_list_widget.currentRow()
        if current_row > 0:
            self._swap_buffer_list(current_row, current_row - 1)

    def _move_buffer_down(self):
        """선택된 즐겨찾기 버퍼를 아래로 이동시킵니다."""
        current_row = self.buffer_list_widget.currentRow()
        count = self.buffer_list_widget.count()
        if current_row < count - 1:
            self._swap_buffer_list(current_row, current_row + 1)

    def _swap_buffer_list(self, index1, index2):
        """실제 즐겨찾기 버퍼 순서 변경 로직 (설정 데이터 및 UI 업데이트)"""

        buffer_names = [
            self.buffer_list_widget.item(i).text()
            for i in range(self.buffer_list_widget.count())
        ]

        if 0 <= index1 < len(buffer_names) and 0 <= index2 < len(buffer_names):
            buffer_names[index1], buffer_names[index2] = (
                buffer_names[index2],
                buffer_names[index1],
            )
        else:
            return

        old_buffers = self.settings["favorites_buffers"]
        new_buffers = {}
        for name in buffer_names:
            if name in old_buffers:
                new_buffers[name] = old_buffers[name]

        self.settings["favorites_buffers"] = new_buffers
        self._save_settings_to_file()

        self._load_buffers_and_favorites(select_row=index2)

    # ----------------- 15-3. 즐겨찾기 버퍼 관리 및 이벤트 핸들러 -----------------
    def _on_buffer_changed(self, current, previous):
        """즐겨찾기 버퍼 리스트에서 선택 항목이 변경될 때 호출됩니다."""
        self._update_buffer_move_button_state()

        if current is not None:
            new_buffer_name = current.text()
            if new_buffer_name != self.active_buffer_name:
                self._load_tree_from_buffer(new_buffer_name)

    def _add_buffer(self):
        """새 즐겨찾기 버퍼를 추가합니다."""
        buffer_name, ok = QInputDialog.getText(
            self, "새 즐겨찾기 버퍼", "새 즐겨찾기 버퍼의 이름을 입력하세요:"
        )
        if ok and buffer_name:
            buffer_name = buffer_name.strip()
            if not buffer_name:
                return

            if buffer_name in self.settings["favorites_buffers"]:
                QMessageBox.warning(self, "오류", "이미 존재하는 이름입니다.")
                return

            self.settings["favorites_buffers"][buffer_name] = []
            self._save_settings_to_file()

            self._load_buffers_and_favorites()
            items = self.buffer_list_widget.findItems(
                buffer_name, Qt.MatchFlag.MatchExactly
            )
            if items:
                self.buffer_list_widget.setCurrentItem(items[0])

    def _rename_buffer(self):
        """선택된 즐겨찾기 버퍼의 이름을 변경합니다."""
        current_item = self.buffer_list_widget.currentItem()
        if not current_item:
            return

        old_name = current_item.text()
        new_name, ok = QInputDialog.getText(
            self, "즐겨찾기 버퍼 이름 변경", "새 이름을 입력하세요:", text=old_name
        )

        if ok and new_name and new_name.strip() != old_name:
            new_name = new_name.strip()
            if not new_name:
                return

            if new_name in self.settings["favorites_buffers"]:
                QMessageBox.warning(self, "오류", "이미 존재하는 이름입니다.")
                return

            buffer_names = [
                self.buffer_list_widget.item(i).text()
                for i in range(self.buffer_list_widget.count())
            ]
            old_buffers = self.settings["favorites_buffers"]
            new_buffers = {}
            for name in buffer_names:
                if name == old_name:
                    new_buffers[new_name] = old_buffers[old_name]
                elif name in old_buffers:
                    new_buffers[name] = old_buffers[name]

            self.settings["favorites_buffers"] = new_buffers

            if self.settings["active_buffer"] == old_name:
                self.settings["active_buffer"] = new_name

            self._save_settings_to_file()

            self._load_buffers_and_favorites()
            items = self.buffer_list_widget.findItems(
                new_name, Qt.MatchFlag.MatchExactly
            )
            if items:
                self.buffer_list_widget.setCurrentItem(items[0])

    def _delete_buffer(self):
        """선택된 즐겨찾기 버퍼를 삭제합니다."""
        current_item = self.buffer_list_widget.currentItem()
        if not current_item:
            return

        if self.buffer_list_widget.count() <= 1:
            QMessageBox.warning(
                self, "삭제 불가", "최소 하나 이상의 즐겨찾기 버퍼가 필요합니다."
            )
            return

        buffer_name = current_item.text()
        reply = QMessageBox.question(
            self,
            "즐겨찾기 버퍼 삭제",
            f"'{buffer_name}' 즐겨찾기 버퍼를 정말 삭제하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )

        if reply == QMessageBox.StandardButton.Yes:
            deleted_name = current_item.text()

            if deleted_name in self.settings["favorites_buffers"]:
                del self.settings["favorites_buffers"][deleted_name]

            if self.settings["active_buffer"] == deleted_name:
                remaining_buffers = list(self.settings["favorites_buffers"].keys())
                self.settings["active_buffer"] = (
                    remaining_buffers[0] if remaining_buffers else None
                )

            self._save_settings_to_file()
            self._load_buffers_and_favorites()

    # ----------------- 15-4. 즐겨찾기 버퍼 복사/붙여넣기 및 메뉴 처리 -----------------
    def _copy_buffer(self):
        """선택된 즐겨찾기 버퍼의 내용을 복사합니다."""
        current_item = self.buffer_list_widget.currentItem()
        if not current_item:
            return

        buffer_name = current_item.text()
        buffer_data = self.settings["favorites_buffers"].get(buffer_name)

        if buffer_data is not None:
            self.buffer_clipboard_data = {"name": buffer_name, "data": buffer_data}
            self.connection_status_label.setText(
                f"즐겨찾기 버퍼 '{buffer_name}' 복사됨."
            )

    def _paste_buffer(self):
        """클립보드의 즐겨찾기 버퍼 내용을 새 즐겨찾기 버퍼로 붙여넣습니다."""
        if not self.buffer_clipboard_data:
            QMessageBox.warning(
                self, "붙여넣기 오류", "복사된 즐겨찾기 버퍼가 없습니다."
            )
            return

        base_name = self.buffer_clipboard_data["name"]

        new_buffer_name, ok = QInputDialog.getText(
            self,
            "즐겨찾기 버퍼 붙여넣기",
            "새 즐겨찾기 버퍼 이름:",
            text=f"복사본 - {base_name}",
        )

        if ok and new_buffer_name:
            new_buffer_name = new_buffer_name.strip()
            if not new_buffer_name:
                return

            if new_buffer_name in self.settings["favorites_buffers"]:
                QMessageBox.warning(self, "오류", "이미 존재하는 이름입니다.")
                return

            self.settings["favorites_buffers"][new_buffer_name] = (
                self.buffer_clipboard_data["data"]
            )
            self._save_settings_to_file()

            self._load_buffers_and_favorites()
            items = self.buffer_list_widget.findItems(
                new_buffer_name, Qt.MatchFlag.MatchExactly
            )
            if items:
                self.buffer_list_widget.setCurrentItem(items[0])

            self.connection_status_label.setText(
                f"즐겨찾기 버퍼 '{new_buffer_name}' 붙여넣기 완료."
            )

    def _on_buffer_context_menu(self, pos):
        """즐겨찾기 버퍼 목록 우클릭 메뉴를 표시합니다."""
        current_item = self.buffer_list_widget.currentItem()
        menu = QMenu(self)

        # 기본 기능
        act_add = QAction("추가", self)
        act_add.triggered.connect(self._add_buffer)
        menu.addAction(act_add)

        if current_item:
            act_rename = QAction("이름 변경", self)
            act_rename.triggered.connect(self._rename_buffer)
            menu.addAction(act_rename)

            act_delete = QAction("삭제", self)
            act_delete.triggered.connect(self._delete_buffer)
            menu.addAction(act_delete)

            # 즐겨찾기 버퍼 이동 (UI 버튼과 동일 기능)
            menu.addSeparator()
            act_move_up = QAction("위로 이동", self)
            act_move_up.setEnabled(self.btn_buffer_move_up.isEnabled())
            act_move_up.triggered.connect(self._move_buffer_up)
            menu.addAction(act_move_up)

            act_move_down = QAction("아래로 이동", self)
            act_move_down.setEnabled(self.btn_buffer_move_down.isEnabled())
            act_move_down.triggered.connect(self._move_buffer_down)
            menu.addAction(act_move_down)

            # 복사
            menu.addSeparator()
            act_copy = QAction("복사", self)
            act_copy.triggered.connect(self._copy_buffer)
            menu.addAction(act_copy)

        # 붙여넣기 (클립보드 데이터 존재 여부에 따라 활성화)
        act_paste = QAction("붙여넣기", self)
        act_paste.triggered.connect(self._paste_buffer)
        act_paste.setEnabled(self.buffer_clipboard_data is not None)
        menu.addAction(act_paste)

        menu.exec(self.buffer_list_widget.viewport().mapToGlobal(pos))

    # ----------------- 16-1. 즐겨찾기 복사/붙여넣기 로직 -----------------
    def _copy_favorite_item(self):
        """선택된 즐겨찾기 항목을 복사합니다."""
        item = self._current_fav_item()
        if not item:
            return

        self.clipboard_data = self._serialize_fav_item(item)
        self.connection_status_label.setText(
            f"'{item.text(0)}' 항목 복사됨."
        )  # 상태바 알림 사용

    def _paste_favorite_item(self):
        """클립보드에 있는 즐겨찾기 항목을 붙여넣습니다."""
        if not self.clipboard_data:
            QMessageBox.warning(
                self, "붙여넣기 오류", "클립보드에 복사된 항목이 없습니다."
            )
            return

        parent = self._current_fav_item()

        if parent and parent.data(0, ROLE_TYPE) == "section":
            parent = parent.parent()

        parent = parent or self.fav_tree.invisibleRootItem()

        def _deep_copy_node(node: Dict[str, Any]) -> Dict[str, Any]:
            new_node = node.copy()
            new_node["id"] = str(uuid.uuid4())
            # new_node["name"] = f"복사본 - {new_node['name']}" # 이 줄을 제거하거나 주석 처리
            if "children" in new_node:
                new_node["children"] = [
                    _deep_copy_node(child) for child in new_node["children"]
                ]
            return new_node

        try:
            copied_node = _deep_copy_node(self.clipboard_data)

            new_item = self._append_fav_node(parent, copied_node)

            self.fav_tree.setCurrentItem(new_item)

            self._save_favorites()
            self.connection_status_label.setText(
                f"'{new_item.text(0)}' 항목 붙여넣기 완료."
            )  # 상태바 알림 사용

        except Exception as e:
            QMessageBox.critical(
                self, "붙여넣기 오류", f"항목을 붙여넣는 중 오류가 발생했습니다: {e}"
            )

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
            is_expanded = item.isExpanded()
            taken_item = parent.takeChild(index)
            parent.insertChild(index - 1, taken_item)
            taken_item.setExpanded(is_expanded)
            self.fav_tree.setCurrentItem(taken_item)
            self._save_favorites()
            self._update_move_button_state()

    def _move_item_down(self):
        item = self._current_fav_item()
        if not item:
            return

        parent = item.parent() or self.fav_tree.invisibleRootItem()
        index = parent.indexOfChild(item)

        if index < parent.childCount() - 1:
            is_expanded = item.isExpanded()
            taken_item = parent.takeChild(index)
            parent.insertChild(index + 1, taken_item)
            taken_item.setExpanded(is_expanded)
            self.fav_tree.setCurrentItem(taken_item)
            self._save_favorites()
            self._update_move_button_state()

    def _update_move_button_state(self):
        item = self._current_fav_item()

        if not item:
            self.btn_move_up.setEnabled(False)
            self.btn_move_down.setEnabled(False)
            return

        parent = item.parent() or self.fav_tree.invisibleRootItem()
        index = parent.indexOfChild(item)

        self.btn_move_up.setEnabled(index > 0)
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

        # 복사/붙여넣기 메뉴
        menu.addSeparator()

        act_copy = QAction("복사 (Ctrl+C)", self)
        act_copy.triggered.connect(self._copy_favorite_item)
        act_copy.setEnabled(item is not None)
        menu.addAction(act_copy)

        act_paste = QAction("붙여넣기 (Ctrl+V)", self)
        act_paste.triggered.connect(self._paste_favorite_item)
        act_paste.setEnabled(self.clipboard_data is not None)
        menu.addAction(act_paste)

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
                        # --- [START] 이름 복원 로직 추가 ---
                        is_name_restored = False
                        current_name = item.text(0)
                        restored_name = current_name
                        if current_name.startswith("(구) "):
                            restored_name = current_name[4:]  # "(구) " 제거
                            item.setText(0, restored_name)
                            self._save_favorites()
                            is_name_restored = True
                        # --- [END] 이름 복원 로직 추가 ---

                        QTimer.singleShot(
                            500,
                            lambda: scroll_selected_item_to_center(
                                self.onenote_window, self.tree_control
                            ),
                        )

                        if is_name_restored:
                            self.update_status_and_ui(
                                f"활성화: '{restored_name}' (이름 복원)", True
                            )
                        else:
                            self.update_status_and_ui(f"활성화: '{display_name}'", True)

                        return
                    else:
                        # --- 실패 시 로직 (기존과 동일) ---
                        current_name = item.text(0)

                        if not current_name.startswith("(구) "):
                            new_name = f"(구) {current_name}"
                            item.setText(0, new_name)
                            self._save_favorites()

                            status_message = (
                                f"섹션 찾기 실패: '{new_name}'(으)로 변경됨"
                            )
                            self.update_status_and_ui(status_message, True)
                        else:
                            status_message = (
                                f"섹션 찾기 실패: '{current_name}' 섹션을 찾을 수 없음"
                            )
                            self.update_status_and_ui(status_message, True)
                    return

        self.update_status_and_ui(f"활성화: '{display_name}'", True)


# ----------------- 17. 엔트리 포인트 -----------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = OneNoteScrollRemoconApp()
    ex.show()

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
