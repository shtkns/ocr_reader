import io, re, os, sys, time, json, ctypes, subprocess, asyncio, difflib, threading, urllib.parse, urllib.request
from datetime import datetime
import tkinter as tk
from tkinter import scrolledtext, messagebox
import keyboard
from PIL import ImageGrab
import winsdk.windows.media.ocr as ocr
import winsdk.windows.globalization as globalization
import winsdk.windows.graphics.imaging as imaging
import winsdk.windows.storage.streams as streams
import pygetwindow as gw


# --- パス取得ユーティリティ ---
def get_base_path():
    """exe化された場合とスクリプト実行時でベースとなるパスを正しく取得する"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# --- 設定読み込み ---
def load_config():
    base_path = get_base_path()
    json_path = os.path.join(base_path, "settings.json")
    default_cfg = {"STABLE_THRESHOLD": 2, "HISTORY_SIZE": 5, "SIMILARITY_THRESHOLD": 0.8, "SLEEP_INTERVAL": 0.15, "BOUYOMI_PORT": 50080}
    try:
        with open(json_path, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
            return (data.get("CHAR_NAMES", []), data.get("ORG_NAMES", []), data.get("REPLACEMENTS", {}), data.get("CONFIG", default_cfg))
    except:
        return [], [], {}, default_cfg


CHAR_NAMES, ORG_NAMES, REPLACEMENTS, CONFIG = load_config()

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    ctypes.windll.user32.SetProcessDPIAware()
exit_event = threading.Event()


# --- 連携ユーティリティ ---
def send_to_bouyomi(text):
    try:
        port = CONFIG.get("BOUYOMI_PORT", 50080)
        encoded_text = urllib.parse.quote(text)
        url = f"http://localhost:{port}/talk?text={encoded_text}"
        with urllib.request.urlopen(url, timeout=1) as response:
            return True
    except:
        return False


class LogManager:
    def __init__(self):
        self.base_dir = os.path.join(get_base_path(), "logs")
        if not os.path.exists(self.base_dir):
            os.makedirs(self.base_dir)

    def write_log(self, text):
        filename = f"log_{datetime.now().strftime('%Y%m%d')}.txt"
        filepath = os.path.join(self.base_dir, filename)
        timestamp = datetime.now().strftime("[%H:%M:%S]")
        with open(filepath, "a", encoding="utf-8") as f:
            f.write(f"{timestamp} {text}\n")


# --- GUIクラス ---
class NovelReaderGUI:
    def __init__(self):
        self.root = None
        self.selection = None
        self.logger = LogManager()
        self.log_area = None

    def start_menu(self):
        self.root = tk.Tk()
        self.root.title("Mode Selection")
        self.root.geometry("300x150")
        self.root.attributes("-topmost", True)
        tk.Label(self.root, text="キャプチャモードを選択してください").pack(pady=10)
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(expand=True)
        tk.Button(btn_frame, text="アプリ選択(非推奨)", width=15, command=self.mode_app).pack(pady=5)
        tk.Button(btn_frame, text="範囲選択", width=15, command=self.mode_range).pack(pady=5)
        self.root.mainloop()

    def mode_app(self):
        self.root.destroy()
        self.root = tk.Tk()
        self.root.title("Select App Window")
        self.root.geometry("400x300")
        tk.Label(self.root, text="対象のウィンドウを選択してください").pack(pady=5)
        windows = [w for w in gw.getAllWindows() if w.title]
        titles = [w.title for w in windows]
        listbox = tk.Listbox(self.root)
        for title in titles:
            listbox.insert(tk.END, title)
        listbox.pack(fill="both", expand=True, padx=10, pady=5)

        def on_select():
            idx = listbox.curselection()
            if not idx:
                return
            target_title = listbox.get(idx)
            target_win = gw.getWindowsWithTitle(target_title)[0]
            self.selection = (target_win.left, target_win.top, target_win.right, target_win.bottom)
            self.root.destroy()

        tk.Button(self.root, text="選択確定", command=on_select).pack(pady=10)
        self.root.mainloop()

    def mode_range(self):
        self.root.destroy()
        self.get_selection()

    def get_selection(self):
        self.root = tk.Tk()
        self.root.attributes("-alpha", 0.3, "-topmost", True, "-toolwindow", True)
        self.root.overrideredirect(True)
        m = [ctypes.windll.user32.GetSystemMetrics(i) for i in [76, 77, 78, 79]]
        self.root.geometry(f"{m[2]}x{m[3]}+{m[0]}+{m[1]}")
        canvas = tk.Canvas(self.root, cursor="cross", bg="grey", highlightthickness=0)
        canvas.pack(fill="both", expand=True)
        rect_data = [None, 0, 0]

        def on_press(e):
            rect_data[1], rect_data[2] = e.x, e.y
            rect_data[0] = canvas.create_rectangle(e.x, e.y, e.x, e.y, outline="red", width=2)

        def on_move(e):
            canvas.coords(rect_data[0], rect_data[1], rect_data[2], e.x, e.y)

        def on_release(e):
            win_x, win_y = self.root.winfo_x(), self.root.winfo_y()
            self.selection = (
                min(rect_data[1], e.x) + win_x,
                min(rect_data[2], e.y) + win_y,
                max(rect_data[1], e.x) + win_x,
                max(rect_data[2], e.y) + win_y,
            )
            self.root.destroy()

        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_move)
        canvas.bind("<ButtonRelease-1>", on_release)
        self.root.mainloop()

    def setup_log_window(self):
        self.root = tk.Tk()
        self.root.title("Novel Reader Console")
        self.root.geometry("600x500")
        self.root.attributes("-topmost", False)

        def on_closing():
            exit_event.set()
            self.root.destroy()
            os._exit(0)

        self.root.protocol("WM_DELETE_WINDOW", on_closing)

        panel = tk.Frame(self.root, bg="#333")
        panel.pack(fill="x")
        tk.Button(panel, text="ログフォルダ", command=lambda: os.startfile(self.logger.base_dir), bg="#555", fg="white").pack(
            side="left", padx=5, pady=5
        )

        def open_editor():
            base_path = get_base_path()
            editor_exe = os.path.join(base_path, "dict_editor.exe")
            if os.path.exists(editor_exe):
                subprocess.Popen([editor_exe])
            else:
                editor_py = os.path.join(base_path, "dict_editor.py")
                subprocess.Popen(["python", editor_py])

        tk.Button(panel, text="辞書・リスト編集", command=open_editor, bg="#007acc", fg="white").pack(side="left", padx=5, pady=5)
        tk.Label(panel, text="Esc:終了", bg="#333", fg="#aaa").pack(side="right", padx=10)
        self.log_area = scrolledtext.ScrolledText(self.root, wrap=tk.WORD, bg="black", fg="white", font=("MS Gothic", 12))
        self.log_area.pack(fill="both", expand=True)

        def watch_esc():
            keyboard.wait("esc")
            exit_event.set()
            self.root.quit()

        threading.Thread(target=watch_esc, daemon=True).start()

    def add_log(self, text):
        if self.log_area:
            self.log_area.insert(tk.END, f"【確定】: {text}\n\n")
            self.log_area.see(tk.END)
            self.logger.write_log(text)


# --- OCRロジック ---
def format_text(result):
    if not result.lines:
        return ""
    lines = sorted(result.lines, key=lambda l: l.words[0].bounding_rect.y if l.words else 0)
    text = "".join([l.text for l in lines])
    jp_char = r"[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FFF\uFF01-\uFF5E=ー]"
    for _ in range(3):
        text = re.sub(f"({jp_char})\s+({jp_char})", r"\1\2", text)
    text = text.replace(" ", "")
    syms = r"[、。！？：:,.!?」」「]"
    text = re.sub(f"\s*({syms})\s*", r"\1", text)
    for old, new in REPLACEMENTS.items():
        text = text.replace(old, new)
    return re.sub(r"\s{2,}", " ", text).replace("\n", "").strip()


def extract_content(raw_text):
    found_speaker, body_text = "", raw_text
    for name in CHAR_NAMES:
        if raw_text.startswith(name):
            found_speaker, body_text = name, raw_text[len(name) :]
            org_pattern = f"^({'|'.join(map(re.escape, ORG_NAMES))})"
            body_text = re.sub(org_pattern, "", body_text)
            break
    if not found_speaker:
        org_pattern = f"^({'|'.join(map(re.escape, ORG_NAMES))})"
        body_text = re.sub(org_pattern, "", body_text)
    return found_speaker, body_text.lstrip("!！.。:： ")


def is_duplicate(current_body, history):
    clean_body = current_body.strip(" 　\t\n\r。、.4")
    if not clean_body:
        return True
    for old_body in history:
        if difflib.SequenceMatcher(None, clean_body, old_body).ratio() > CONFIG["SIMILARITY_THRESHOLD"]:
            return True
    return False


async def process_ocr(engine, selection):
    img = ImageGrab.grab(bbox=selection, all_screens=True)
    byte_io = io.BytesIO()
    img.save(byte_io, format="PNG")
    stream = streams.InMemoryRandomAccessStream()
    writer = streams.DataWriter(stream)
    writer.write_bytes(byte_io.getvalue())
    await writer.store_async()
    writer.detach_stream()
    stream.seek(0)
    decoder = await imaging.BitmapDecoder.create_async(stream)
    bitmap = await decoder.get_software_bitmap_async()
    result = await engine.recognize_async(bitmap)
    stream.close()
    return result


async def run_monitoring(selection, gui):
    # exe化を考慮した正しいJSONパス指定
    json_path = os.path.join(get_base_path(), "settings.json")
    last_mtime = os.path.getmtime(json_path) if os.path.exists(json_path) else 0
    engine = ocr.OcrEngine.try_create_from_language(globalization.Language("ja-JP"))
    last_speaker, last_raw_text, sent_history = "", "", []
    current_stable, stable_count = "", 0

    while not exit_event.is_set():
        try:
            # 1. 設定のリアルタイム更新チェック
            try:
                current_mtime = os.path.getmtime(json_path)
                if current_mtime > last_mtime:
                    global CHAR_NAMES, ORG_NAMES, REPLACEMENTS, CONFIG
                    CHAR_NAMES, ORG_NAMES, REPLACEMENTS, CONFIG = load_config()
                    last_mtime = current_mtime
                    gui.add_log("--- 設定の更新を反映しました ---")
            except:
                pass

            result = await process_ocr(engine, selection)
            raw_text = format_text(result)
            if not raw_text:
                current_stable, stable_count = "", 0
                await asyncio.sleep(CONFIG["SLEEP_INTERVAL"])
                continue

            if raw_text == current_stable:
                stable_count += 1
            else:
                current_stable, stable_count = raw_text, 1

            if stable_count < CONFIG["STABLE_THRESHOLD"] or raw_text == last_raw_text:
                await asyncio.sleep(CONFIG["SLEEP_INTERVAL"])
                continue

            speaker, body = extract_content(raw_text)
            clean_body = body.strip(" 　\t\n\r。、.4")

            if is_duplicate(clean_body, sent_history):
                last_raw_text = raw_text
                await asyncio.sleep(CONFIG["SLEEP_INTERVAL"])
                continue

            final_text = body if speaker == last_speaker else f"{speaker}: {body}" if speaker else body
            content = final_text.strip()

            if content:
                # 可変速ロジック
                base_speed = 150
                char_count = len(content)
                final_speed = min(base_speed + (char_count * 2), 350)
                speed_tag = f"速度({final_speed})"

                if send_to_bouyomi(speed_tag + content):
                    gui.add_log(f"[{final_speed}%速] {content}")
                    last_raw_text, last_speaker = raw_text, speaker
                    sent_history.append(clean_body)
                    if len(sent_history) > CONFIG["HISTORY_SIZE"]:
                        sent_history.pop(0)
                    await asyncio.sleep(0.2)

            await asyncio.sleep(CONFIG["SLEEP_INTERVAL"])
        except Exception:
            await asyncio.sleep(0.5)


if __name__ == "__main__":
    app = NovelReaderGUI()
    app.start_menu()
    if app.selection:
        app.setup_log_window()
        loop = asyncio.new_event_loop()
        threading.Thread(target=lambda: loop.run_until_complete(run_monitoring(app.selection, app)), daemon=True).start()
        app.root.mainloop()
