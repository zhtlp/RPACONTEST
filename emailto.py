# -*- coding: utf-8 -*-
"""
邮件模板管理 + 发送助手 GUI
涛哥专属 v4（新增：导入导出模板）
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import webbrowser
import urllib.parse

TEMPLATE_FILE = "email_templates.json"

# --------------------------------------------------
# 模板文件加载与保存
# --------------------------------------------------
def load_templates():
    if not os.path.exists(TEMPLATE_FILE):
        return {}
    with open(TEMPLATE_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_templates(data):
    with open(TEMPLATE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

# --------------------------------------------------
# 主窗口
# --------------------------------------------------
class EmailGUI:
    def __init__(self, root):
        self.root = root
        root.title("邮件模板助手（高端定制 v4）")
        root.geometry("900x700")
        root.resizable(False, False)

        self.templates = load_templates()

        # ==================== 左侧模板区 ====================
        frame_left = tk.Frame(root)
        frame_left.pack(side=tk.LEFT, fill=tk.Y, padx=30, pady=10)

        tk.Label(frame_left, text="模板列表", font=("Arial", 12, "bold")).pack(anchor="w")

        # Listbox + 滚动条
        list_frame = tk.Frame(frame_left)
        list_frame.pack(fill=tk.Y, expand=True)

        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        self.template_list = tk.Listbox(
            list_frame,
            width=25,
            height=10,
            yscrollcommand=scrollbar.set
        )
        scrollbar.config(command=self.template_list.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.template_list.pack(side=tk.LEFT, fill=tk.Y)
        self.template_list.bind("<<ListboxSelect>>", self.load_selected_template)



        # ---------------- 原有三个按钮 ----------------
        btn_new = tk.Button(frame_left, text="新建模板", command=self.new_template)
        btn_new.pack(fill="x", pady=5)

        btn_save = tk.Button(frame_left, text="保存模板", command=self.save_template, bg="#cde1ff")
        btn_save.pack(fill="x", pady=5)

        btn_delete = tk.Button(frame_left, text="删除模板", command=self.delete_template, bg="#ffcccc")
        btn_delete.pack(fill="x", pady=(5,35))
        # ---------------- 新增：导入导出按钮 ----------------
        btn_export = tk.Button(frame_left, text="  导出全部模板  ", command=self.export_templates, bg="#fff3cd")
        btn_export.pack(fill="x", pady=4)

        btn_import = tk.Button(frame_left, text="  导入模板文件  ", command=self.import_templates, bg="#d1f0ff")
        btn_import.pack(fill="x", pady=4)
        # ==================== 右侧编辑区 ====================
        frame_right = tk.Frame(root)
        frame_right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=30, pady=10)

        tk.Label(frame_right, text="收件人（TO，用;分隔）").grid(row=0, column=0, sticky="w", pady=4)
        self.entry_to = tk.Entry(frame_right, width=70)
        self.entry_to.grid(row=0, column=1, sticky="w", pady=4)

        tk.Label(frame_right, text="抄送（CC，用;分隔）").grid(row=1, column=0, sticky="w", pady=4)
        self.entry_cc = tk.Entry(frame_right, width=70)
        self.entry_cc.grid(row=1, column=1, sticky="w", pady=4)

        tk.Label(frame_right, text="主题").grid(row=2, column=0, sticky="w", pady=4)
        self.entry_subject = tk.Entry(frame_right, width=70)
        self.entry_subject.grid(row=2, column=1, sticky="w", pady=4)

        tk.Label(frame_right, text="邮件正文（纯文本）").grid(row=3, column=0, sticky="nw", pady=4)
        self.text_body = tk.Text(frame_right, width=70, height=28, wrap=tk.WORD)
        self.text_body.grid(row=3, column=1, sticky="w", pady=4)

        btn_send = tk.Button(frame_right, text="生成邮件窗口", command=self.generate_mailto, bg="#d4ffc4", height=3,width=30)
        btn_send.grid(row=4, column=1, pady=20, sticky="w")

        self.refresh_template_list()

    # --------------------------------------------------
    # 新增：导入导出功能
    # --------------------------------------------------
    def export_templates(self):
        if not self.templates:
            messagebox.showinfo("提示", "当前没有模板可导出！")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON 文件", "*.json")],
            title="导出模板 - 请选择保存位置"
        )
        if file_path:
            try:
                with open(file_path, "w", encoding="utf-8") as f:
                    json.dump(self.templates, f, ensure_ascii=False, indent=4)
                messagebox.showinfo("成功", f"模板已导出到：\n{file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败：{e}")

    def import_templates(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON 文件", "*.json")],
            title="导入模板 - 请选择文件"
        )
        if not file_path:
            return
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                imported = json.load(f)
            if not isinstance(imported, dict):
                messagebox.showerror("错误", "文件格式不正确，必须是模板字典格式！")
                return

            count = len(imported)
            self.templates.update(imported)  # 同名模板会被覆盖
            save_templates(self.templates)
            self.refresh_template_list()
            messagebox.showinfo("成功", f"成功导入 {count} 个模板！\n同名模板已覆盖。")
        except Exception as e:
            messagebox.showerror("错误", f"导入失败：{e}")

    # --------------------------------------------------
    # 其余功能完全不变（略）
    # --------------------------------------------------
    def refresh_template_list(self):
        self.template_list.delete(0, tk.END)
        for key in sorted(self.templates.keys()):
            self.template_list.insert(tk.END, key)

    def new_template(self):
        self.entry_to.delete(0, tk.END)
        self.entry_cc.delete(0, tk.END)
        self.entry_subject.delete(0, tk.END)
        self.text_body.delete("1.0", tk.END)
        self.template_list.selection_clear(0, tk.END)

    def load_selected_template(self, event):
        idx = self.template_list.curselection()
        if not idx:
            return
        name = self.template_list.get(idx)
        data = self.templates.get(name, {})

        self.entry_to.delete(0, tk.END)
        self.entry_to.insert(0, data.get("to", ""))

        self.entry_cc.delete(0, tk.END)
        self.entry_cc.insert(0, data.get("cc", ""))

        self.entry_subject.delete(0, tk.END)
        self.entry_subject.insert(0, data.get("subject", ""))

        self.text_body.delete("1.0", tk.END)
        self.text_body.insert("1.0", data.get("body", ""))

    def save_template(self):
        name = self.entry_subject.get().strip()
        if not name:
            messagebox.showerror("错误", "模板名称使用邮件主题，请输入主题！")
            return

        self.templates[name] = {
            "to": self.entry_to.get().strip(),
            "cc": self.entry_cc.get().strip(),
            "subject": self.entry_subject.get().strip(),
            "body": self.text_body.get("1.0", tk.END).strip(),
        }
        save_templates(self.templates)
        self.refresh_template_list()
        messagebox.showinfo("成功", "模板已保存！")

    def delete_template(self):
        idx = self.template_list.curselection()
        if not idx:
            messagebox.showerror("错误", "请选择要删除的模板！")
            return

        name = self.template_list.get(idx)
        if name in self.templates:
            del self.templates[name]
            save_templates(self.templates)
            self.refresh_template_list()
            self.new_template()
            messagebox.showinfo("成功", "模板已删除！")

    def generate_mailto(self):
        to = self.entry_to.get().strip()
        cc = self.entry_cc.get().strip()

        subject = urllib.parse.quote(self.entry_subject.get().strip())
        body_raw = self.text_body.get("1.0", tk.END)
        body = urllib.parse.quote(body_raw, safe="<>/\"'=")

        mailto = f"mailto:{to}?subject={subject}&body={body}"
        if cc:
            mailto += f"&cc={urllib.parse.quote(cc)}"

        webbrowser.open(mailto)
        messagebox.showinfo("提示", "邮件窗口已打开，如需附件请在客户端中手动添加。")

# --------------------------------------------------
# main
# --------------------------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = EmailGUI(root)
    root.mainloop()