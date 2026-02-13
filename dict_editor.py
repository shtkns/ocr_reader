import tkinter as tk
from tkinter import messagebox, ttk
import json, os, sys, re


class DictEditor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("ブルアカOCR 設定エディタ")
        self.root.geometry("600x700")
        base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
        self.json_path = os.path.join(base_path, "settings.json")
        self.data = self.load_json()
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        self.setup_replacement_tab()
        self.setup_list_tab("CHAR_NAMES", "キャラクター名")
        self.setup_list_tab("ORG_NAMES", "所属名")
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(fill="x", padx=10, pady=5)
        tk.Button(btn_frame, text="設定を保存", command=self.save_json, bg="#4caf50", fg="white", height=2).pack(fill="x")
        self.root.mainloop()

    def load_json(self):
        try:
            with open(self.json_path, "r", encoding="utf-8-sig") as f:
                return json.load(f)
        except:
            return {"REPLACEMENTS": {}, "CHAR_NAMES": [], "ORG_NAMES": [], "CONFIG": {}}

    def save_json(self):
        try:
            with open(self.json_path, "w", encoding="utf-8", newline="") as f:
                json.dump(self.data, f, indent=4, ensure_ascii=False)
            messagebox.showinfo("成功", "保存しました")
        except:
            messagebox.showerror("失敗", "保存できませんでした")

    def setup_replacement_tab(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="置換辞書")
        input_f = tk.Frame(frame, pady=10)
        input_f.pack(fill="x")
        tk.Label(input_f, text="誤字:").grid(row=0, column=0)
        ent_b = tk.Entry(input_f)
        ent_b.grid(row=0, column=1, sticky="ew")
        tk.Label(input_f, text="正解:").grid(row=0, column=2)
        ent_a = tk.Entry(input_f)
        ent_a.grid(row=0, column=3, sticky="ew")
        tree = ttk.Treeview(frame, columns=("B", "A"), show="headings")
        tree.pack(fill="both", expand=True)
        tree.heading("B", text="誤字")
        tree.heading("A", text="正解")

        def update_list():
            tree.delete(*tree.get_children())
            for b, a in self.data["REPLACEMENTS"].items():
                tree.insert("", "end", values=(b, a))

        def add():
            b = ent_b.get().strip()
            if b:
                self.data["REPLACEMENTS"][b] = ent_a.get().strip()
                update_list()
                ent_b.delete(0, tk.END)
                ent_a.delete(0, tk.END)

        def add_bulk():
            top = tk.Toplevel(self.root)
            top.title("一括追加")
            top.geometry("300x400")
            txt = tk.Text(top)
            txt.pack(fill="both", expand=True)

            def commit():
                for line in txt.get("1.0", tk.END).splitlines():
                    parts = re.split(r"[,，\s\t]+", line.strip())
                    if len(parts) >= 2:
                        self.data["REPLACEMENTS"][parts[0]] = parts[1]
                update_list()
                top.destroy()

            tk.Button(top, text="実行", command=commit).pack(fill="x")

        tk.Button(input_f, text="追加", command=add).grid(row=0, column=4)
        ctrl_f = tk.Frame(frame)
        ctrl_f.pack(fill="x")
        tk.Button(ctrl_f, text="一括追加...", command=add_bulk).pack(side="left")

        def del_sel():
            for i in tree.selection():
                val = tree.item(i)["values"][0]
                if str(val) in self.data["REPLACEMENTS"]:
                    del self.data["REPLACEMENTS"][str(val)]
            update_list()

        tk.Button(ctrl_f, text="削除", command=del_sel).pack(side="right")
        update_list()

    def setup_list_tab(self, key, label):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text=label)
        input_f = tk.Frame(frame, pady=10)
        input_f.pack(fill="x")
        ent = tk.Entry(input_f)
        ent.pack(side="left", fill="x", expand=True)
        lb = tk.Listbox(frame)
        lb.pack(fill="both", expand=True)

        def update_list():
            lb.delete(0, tk.END)
            for i in self.data[key]:
                lb.insert(tk.END, i)

        def add_single():
            val = ent.get().strip()
            if val and val not in self.data[key]:
                self.data[key].append(val)
                update_list()
                ent.delete(0, tk.END)

        def add_bulk():
            top = tk.Toplevel(self.root)
            top.geometry("300x400")
            txt = tk.Text(top)
            txt.pack(fill="both", expand=True)

            def commit():
                for line in txt.get("1.0", tk.END).splitlines():
                    v = line.strip()
                    if v and v not in self.data[key]:
                        self.data[key].append(v)
                update_list()
                top.destroy()

            tk.Button(top, text="実行", command=commit).pack(fill="x")

        tk.Button(input_f, text="追加", command=add_single).pack(side="right")
        tk.Button(frame, text="一括追加...", command=add_bulk).pack(side="left")
        tk.Button(
            frame, text="削除", command=lambda: [self.data[key].remove(lb.get(i)) for i in reversed(lb.curselection()) if update_list() or True]
        ).pack(side="right")
        update_list()


if __name__ == "__main__":
    DictEditor()
