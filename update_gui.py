import json
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import warnings
warnings.filterwarnings("ignore", category=UserWarning)

try:
    from openpyxl import load_workbook
    from openpyxl.drawing.image import Image as XLImage
except ImportError:
    messagebox.showerror("缺少库", "请先运行: pip install openpyxl")
    sys.exit(1)

CONFIG_FILE = "update_config.json"

def save_config(data):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def load_config():
    if not os.path.exists(CONFIG_FILE):
        return {"data_source_path": "", "template_path": "", "output_suffix": "_已更新",
                "image_global_width_cm": 3.5, "image_global_height_cm": 2.8,
                "data_mappings": [], "image_mappings": []}
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel报告一键更新器")
        self.root.geometry("820x650")
        self.config = load_config()

        # 路径配置区
        frm_path = ttk.LabelFrame(root, text="基础路径")
        frm_path.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(frm_path, text="数据源路径:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.src_path_var = tk.StringVar(value=self.config["data_source_path"])
        ttk.Entry(frm_path, textvariable=self.src_path_var, width=80).grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(frm_path, text="浏览", command=self.browse_src).grid(row=0, column=2, padx=5)

        ttk.Label(frm_path, text="模板路径:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.tpl_path_var = tk.StringVar(value=self.config["template_path"])
        ttk.Entry(frm_path, textvariable=self.tpl_path_var, width=80).grid(row=1, column=1, padx=5, pady=2)
        ttk.Button(frm_path, text="浏览", command=self.browse_tpl).grid(row=1, column=2, padx=5)

        # 图片尺寸
        frm_img_size = ttk.LabelFrame(root, text="图片统一尺寸(厘米)")
        frm_img_size.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(frm_img_size, text="宽度:").grid(row=0, column=0, padx=5)
        self.img_w_var = tk.DoubleVar(value=self.config["image_global_width_cm"])
        ttk.Entry(frm_img_size, textvariable=self.img_w_var, width=10).grid(row=0, column=1, padx=5)
        ttk.Label(frm_img_size, text="高度:").grid(row=0, column=2, padx=5)
        self.img_h_var = tk.DoubleVar(value=self.config["image_global_height_cm"])
        ttk.Entry(frm_img_size, textvariable=self.img_h_var, width=10).grid(row=0, column=3, padx=5)

        # 映射管理区（Notebook）
        nb = ttk.Notebook(root)
        nb.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 数据映射页
        frm_data = ttk.Frame(nb)
        nb.add(frm_data, text="数据映射")
        self.data_tree = ttk.Treeview(frm_data, columns=("source", "target"), show="headings", height=10)
        self.data_tree.heading("source", text="数据源单元格 (如 Sheet1!B2)")
        self.data_tree.heading("target", text="模板目标单元格 (如 Sheet1!D5)")
        self.data_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        btn_frm_data = ttk.Frame(frm_data)
        btn_frm_data.pack(fill=tk.X, padx=5, pady=2)
        ttk.Button(btn_frm_data, text="添加", command=self.add_data_mapping).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frm_data, text="删除选中", command=lambda: self.delete_selected(self.data_tree)).pack(side=tk.LEFT, padx=5)

        # 图片映射页
        frm_img = ttk.Frame(nb)
        nb.add(frm_img, text="图片映射")
        self.img_tree = ttk.Treeview(frm_img, columns=("match_cell", "folder", "target_cell"), show="headings", height=10)
        self.img_tree.heading("match_cell", text="匹配编号来源单元格")
        self.img_tree.heading("folder", text="图片文件夹")
        self.img_tree.heading("target_cell", text="插入目标单元格")
        self.img_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        btn_frm_img = ttk.Frame(frm_img)
        btn_frm_img.pack(fill=tk.X, padx=5, pady=2)
        ttk.Button(btn_frm_img, text="添加", command=self.add_image_mapping).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frm_img, text="删除选中", command=lambda: self.delete_selected(self.img_tree)).pack(side=tk.LEFT, padx=5)

        # 执行按钮
        ttk.Button(root, text="一键更新报告", command=self.run_update).pack(pady=10)

        # 加载已有映射到界面
        self.refresh_data_tree()
        self.refresh_image_tree()

    def browse_src(self):
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")])
        if path:
            self.src_path_var.set(path)

    def browse_tpl(self):
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")])
        if path:
            self.tpl_path_var.set(path)

    def add_data_mapping(self):
        popup = tk.Toplevel(self.root)
        popup.title("添加数据映射")
        ttk.Label(popup, text="源单元格 (如 Sheet1!B2):").grid(row=0, column=0, padx=5, pady=5)
        src_entry = ttk.Entry(popup, width=30)
        src_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(popup, text="目标单元格 (如 Sheet1!D5):").grid(row=1, column=0, padx=5, pady=5)
        tgt_entry = ttk.Entry(popup, width=30)
        tgt_entry.grid(row=1, column=1, padx=5, pady=5)

        def save():
            src = src_entry.get().strip()
            tgt = tgt_entry.get().strip()
            if src and tgt:
                self.config["data_mappings"].append({"source_cell": src, "target_cell": tgt})
                self.refresh_data_tree()
                popup.destroy()
            else:
                messagebox.showwarning("输入不完整", "两个单元格都不能为空")
        ttk.Button(popup, text="确定", command=save).grid(row=2, column=0, columnspan=2, pady=10)

    def add_image_mapping(self):
        popup = tk.Toplevel(self.root)
        popup.title("添加图片映射")
        ttk.Label(popup, text="匹配编号来源单元格:").grid(row=0, column=0, padx=5, pady=5)
        match_entry = ttk.Entry(popup, width=30)
        match_entry.grid(row=0, column=1, padx=5)
        ttk.Label(popup, text="图片文件夹:").grid(row=1, column=0, padx=5, pady=5)
        folder_var = tk.StringVar()
        folder_entry = ttk.Entry(popup, textvariable=folder_var, width=30)
        folder_entry.grid(row=1, column=1, padx=5)
        ttk.Button(popup, text="浏览", command=lambda: folder_var.set(filedialog.askdirectory())).grid(row=1, column=2, padx=5)
        ttk.Label(popup, text="插入目标单元格:").grid(row=2, column=0, padx=5, pady=5)
        tgt_entry = ttk.Entry(popup, width=30)
        tgt_entry.grid(row=2, column=1, padx=5)

        def save():
            match = match_entry.get().strip()
            folder = folder_var.get().strip()
            tgt = tgt_entry.get().strip()
            if match and folder and tgt:
                self.config["image_mappings"].append({
                    "match_value_source_cell": match,
                    "image_folder": folder,
                    "target_cell": tgt
                })
                self.refresh_image_tree()
                popup.destroy()
            else:
                messagebox.showwarning("输入不完整", "所有字段必填")
        ttk.Button(popup, text="确定", command=save).grid(row=3, column=0, columnspan=3, pady=10)

    def delete_selected(self, tree):
        selected = tree.selection()
        if not selected:
            return
        idx = tree.index(selected[0])
        if tree == self.data_tree:
            del self.config["data_mappings"][idx]
            self.refresh_data_tree()
        else:
            del self.config["image_mappings"][idx]
            self.refresh_image_tree()

    def refresh_data_tree(self):
        for i in self.data_tree.get_children():
            self.data_tree.delete(i)
        for m in self.config.get("data_mappings", []):
            self.data_tree.insert("", tk.END, values=(m["source_cell"], m["target_cell"]))

    def refresh_image_tree(self):
        for i in self.img_tree.get_children():
            self.img_tree.delete(i)
        for m in self.config.get("image_mappings", []):
            self.img_tree.insert("", tk.END, values=(m["match_value_source_cell"], m["image_folder"], m["target_cell"]))

    def save_current_config(self):
        self.config["data_source_path"] = self.src_path_var.get()
        self.config["template_path"] = self.tpl_path_var.get()
        self.config["image_global_width_cm"] = self.img_w_var.get()
        self.config["image_global_height_cm"] = self.img_h_var.get()
        save_config(self.config)

    def run_update(self):
        self.save_current_config()
        # 验证路径
        cfg = self.config
        if not cfg["data_source_path"] or not os.path.exists(cfg["data_source_path"]):
            messagebox.showerror("错误", "数据源文件不存在")
            return
        if not cfg["template_path"] or not os.path.exists(cfg["template_path"]):
            messagebox.showerror("错误", "模板文件不存在")
            return

        try:
            wb_src = load_workbook(cfg["data_source_path"], data_only=True)
            wb = load_workbook(cfg["template_path"])
        except Exception as e:
            messagebox.showerror("打开文件失败", str(e))
            return

        # 数据写入
        for m in cfg.get("data_mappings", []):
            try:
                src_sh, src_cell = m["source_cell"].split("!")
                tgt_sh, tgt_cell = m["target_cell"].split("!")
                ws_src = wb_src[src_sh]
                ws_tgt = wb[tgt_sh]
                ws_tgt[tgt_cell].value = ws_src[src_cell].value
            except Exception as e:
                wb_src.close()
                wb.close()
                messagebox.showerror("数据写入出错", f"{m}\n{e}")
                return

        # 图片插入
        img_w = cfg["image_global_width_cm"]
        img_h = cfg["image_global_height_cm"]
        for m in cfg.get("image_mappings", []):
            try:
                src_sh, src_cell = m["match_value_source_cell"].split("!")
                match_val = wb_src[src_sh][src_cell].value
                if match_val is None:
                    continue
                match_val = str(match_val).strip()
                folder = Path(m["image_folder"])
                img_path = None
                for ext in [".jpg", ".jpeg", ".png", ".bmp", ".gif"]:
                    candidate = folder / f"{match_val}{ext}"
                    if candidate.exists():
                        img_path = str(candidate)
                        break
                if img_path is None:
                    print(f"图片未找到: {match_val} in {folder}")
                    continue
                tgt_sh, tgt_cell = m["target_cell"].split("!")
                ws_tgt = wb[tgt_sh]
                img = XLImage(img_path)
                img.width = img_w
                img.height = img_h
                ws_tgt.add_image(img, tgt_cell)
            except Exception as e:
                wb_src.close()
                wb.close()
                messagebox.showerror("图片插入出错", f"{m}\n{e}")
                return

        # 保存新文件
        tpl_path = Path(cfg["template_path"])
        suffix = cfg.get("output_suffix", "_已更新")
        out_path = tpl_path.parent / f"{tpl_path.stem}{suffix}.xlsx"
        try:
            wb.save(str(out_path))
        except Exception as e:
            wb_src.close()
            wb.close()
            messagebox.showerror("保存失败", str(e))
            return

        wb_src.close()
        wb.close()
        # 自动打开
        try:
            os.startfile(out_path)
        except:
            pass
        messagebox.showinfo("完成", f"报告已生成:\n{out_path}")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
