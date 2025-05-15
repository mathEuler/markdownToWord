import tkinter as tk
from tkinter import filedialog, ttk, messagebox, scrolledtext
import os
import subprocess
import tempfile
import sys
import platform
import datetime
import glob


class MarkdownToWordConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Markdown 转 Word 工具 - 数学公式增强版")
        self.root.geometry("1000x600")
        # 初始化转换计数器
        self.conversion_count = 0
        # 设置默认保存文件名前缀
        self.filename_prefix = "markdownToWord"
        # 默认保存位置
        self.default_save_dir = os.path.join(os.path.expanduser("~"), "Desktop")
        self.setup_ui()
        self.check_pandoc()

    def check_pandoc(self):
        """检查 Pandoc 是否已安装"""
        try:
            result = subprocess.run(
                ["pandoc", "--version"], capture_output=True, text=True, check=True
            )
            self.status_label.config(
                text=f"Pandoc 已安装: {result.stdout.splitlines()[0]}"
            )
            return True
        except (subprocess.CalledProcessError, FileNotFoundError):
            self.status_label.config(
                text="错误: 未检测到 Pandoc! 请确保已安装 Pandoc。"
            )
            return False

    def get_default_output_path(self):
        """生成默认输出路径，格式为桌面/前缀_YYYY-MM-DD_n.docx"""
        # 使用当前设置的保存目录
        save_dir = self.default_save_dir
        if not os.path.exists(save_dir):
            # 如果目录不存在，回退到桌面
            save_dir = os.path.join(os.path.expanduser("~"), "Desktop")
            
        # 使用当前设置的文件名前缀
        prefix = self.filename_prefix
        
        # 当前日期
        today = datetime.datetime.now().strftime("%Y-%m-%d")

        # 查找今天已存在的文件数量
        pattern = os.path.join(save_dir, f"{prefix}_{today}_*.docx")
        existing_files = glob.glob(pattern)

        # 确定新文件的序号
        if not existing_files:
            count = 1
        else:
            # 提取现有文件的最大序号并加1
            max_count = 0
            for file in existing_files:
                try:
                    count_str = os.path.basename(file).split("_")[-1].split(".")[0]
                    count = int(count_str)
                    max_count = max(max_count, count)
                except (ValueError, IndexError):
                    pass
            count = max_count + 1

        # 构建新文件名
        filename = f"{prefix}_{today}_{count}.docx"
        return os.path.join(save_dir, filename)

    def setup_ui(self):
        # 创建左右分割面板
        self.paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 左侧面板 - Markdown 输入
        left_frame = ttk.LabelFrame(self.paned_window, text="Markdown 输入")
        self.paned_window.add(left_frame, weight=3)

        # 右侧面板 - 设置和输出
        right_frame = ttk.LabelFrame(self.paned_window, text="设置和输出")
        self.paned_window.add(right_frame, weight=1)

        # Markdown 输入区域
        self.text_input = scrolledtext.ScrolledText(
            left_frame, wrap=tk.WORD, font=("Consolas", 11)
        )
        self.text_input.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 右侧控制区域
        control_frame = ttk.Frame(right_frame)
        control_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 数学公式设置
        math_frame = ttk.LabelFrame(control_frame, text="数学公式设置")
        math_frame.pack(fill=tk.X, padx=5, pady=5)

        self.math_option = tk.StringVar(value="mathjax")
        ttk.Radiobutton(
            math_frame, text="MathJax", variable=self.math_option, value="mathjax"
        ).pack(anchor=tk.W, padx=5, pady=2)
        ttk.Radiobutton(
            math_frame, text="MathML", variable=self.math_option, value="mathml"
        ).pack(anchor=tk.W, padx=5, pady=2)
        ttk.Radiobutton(
            math_frame, text="标准（默认）", variable=self.math_option, value="none"
        ).pack(anchor=tk.W, padx=5, pady=2)

        # 其他设置
        options_frame = ttk.LabelFrame(control_frame, text="其他设置")
        options_frame.pack(fill=tk.X, padx=5, pady=5)

        # 添加自动打开选项
        self.auto_open = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame, text="转换后自动打开文件", variable=self.auto_open
        ).pack(anchor=tk.W, padx=5, pady=2)

        # 添加默认路径选项
        self.use_default_path = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="使用默认保存路径",
            variable=self.use_default_path,
        ).pack(anchor=tk.W, padx=5, pady=2)
        
        # 添加默认保存目录设置
        save_dir_frame = ttk.Frame(options_frame)
        save_dir_frame.pack(fill=tk.X, padx=5, pady=2)
        ttk.Label(save_dir_frame, text="默认保存位置:").pack(side=tk.LEFT)
        self.save_dir_var = tk.StringVar(value=self.default_save_dir)
        ttk.Entry(save_dir_frame, textvariable=self.save_dir_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True
        )
        ttk.Button(
            save_dir_frame, text="浏览...", command=self.browse_save_dir
        ).pack(side=tk.RIGHT)
        
        # 添加文件名前缀设置
        prefix_frame = ttk.Frame(options_frame)
        prefix_frame.pack(fill=tk.X, padx=5, pady=2)
        ttk.Label(prefix_frame, text="文件名前缀:").pack(side=tk.LEFT)
        self.prefix_var = tk.StringVar(value=self.filename_prefix)
        ttk.Entry(prefix_frame, textvariable=self.prefix_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True
        )
        ttk.Button(
            prefix_frame, text="应用", command=self.update_prefix_settings
        ).pack(side=tk.RIGHT)

        self.use_reference_docx = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            options_frame, text="使用参考文档样式", variable=self.use_reference_docx
        ).pack(anchor=tk.W, padx=5, pady=2)

        ref_docx_frame = ttk.Frame(options_frame)
        ref_docx_frame.pack(fill=tk.X, padx=5, pady=2)

        self.reference_docx_path = tk.StringVar()
        ttk.Entry(ref_docx_frame, textvariable=self.reference_docx_path).pack(
            side=tk.LEFT, fill=tk.X, expand=True
        )
        ttk.Button(
            ref_docx_frame, text="浏览...", command=self.browse_reference_docx
        ).pack(side=tk.RIGHT)

        # 操作按钮
        buttons_frame = ttk.Frame(control_frame)
        buttons_frame.pack(fill=tk.X, padx=5, pady=10)

        ttk.Button(
            buttons_frame, text="从文件加载 Markdown", command=self.load_markdown
        ).pack(fill=tk.X, pady=2)

        # 修改转换按钮，添加快速转换选项
        ttk.Button(
            buttons_frame, text="转换为 Word（选择路径）", command=self.convert_to_word
        ).pack(fill=tk.X, pady=2)

        ttk.Button(
            buttons_frame, text="快速转换（使用默认路径）", command=self.quick_convert
        ).pack(fill=tk.X, pady=2)

        ttk.Button(
            buttons_frame,
            text="清空",
            command=lambda: self.text_input.delete(1.0, tk.END),
        ).pack(fill=tk.X, pady=2)

        # 状态信息
        self.status_label = ttk.Label(control_frame, text="就绪", anchor=tk.W)
        self.status_label.pack(fill=tk.X, padx=5, pady=5)
        
    def browse_save_dir(self):
        """浏览选择默认保存目录"""
        dir_path = filedialog.askdirectory(initialdir=self.default_save_dir)
        if dir_path:
            self.save_dir_var.set(dir_path)
            
    def update_prefix_settings(self):
        """更新文件名前缀和保存目录设置"""
        # 更新文件名前缀
        new_prefix = self.prefix_var.get()
        if new_prefix:
            self.filename_prefix = new_prefix
        
        # 更新保存目录
        new_dir = self.save_dir_var.get()
        if new_dir and os.path.exists(new_dir):
            self.default_save_dir = new_dir
        elif new_dir:
            try:
                os.makedirs(new_dir, exist_ok=True)
                self.default_save_dir = new_dir
            except:
                messagebox.showerror("错误", f"无法创建目录: {new_dir}")
                self.save_dir_var.set(self.default_save_dir)
        
        # 更新状态
        default_path = self.get_default_output_path()
        self.status_label.config(text=f"设置已更新。下一个文件将保存为: {default_path}")

    def browse_reference_docx(self):
        filepath = filedialog.askopenfilename(filetypes=[("Word 文档", "*.docx")])
        if filepath:
            self.reference_docx_path.set(filepath)

    def load_markdown(self):
        filepath = filedialog.askopenfilename(
            filetypes=[
                ("Markdown 文件", "*.md"),
                ("文本文件", "*.txt"),
                ("所有文件", "*.*"),
            ]
        )
        if filepath:
            try:
                with open(filepath, "r", encoding="utf-8") as file:
                    content = file.read()
                    self.text_input.delete(1.0, tk.END)
                    self.text_input.insert(tk.END, content)
                self.status_label.config(text=f"已加载: {filepath}")
            except Exception as e:
                messagebox.showerror("加载错误", f"无法加载文件: {str(e)}")

    def quick_convert(self):
        """使用默认路径快速转换"""
        markdown_text = self.text_input.get("1.0", tk.END).strip()
        if not markdown_text:
            self.status_label.config(text="请输入 Markdown 内容！")
            return

        # 获取默认输出路径
        output_path = self.get_default_output_path()

        # 执行转换
        self.do_conversion(markdown_text, output_path)

    def convert_to_word(self):
        """转换为 Word，让用户选择保存路径"""
        markdown_text = self.text_input.get("1.0", tk.END).strip()
        if not markdown_text:
            self.status_label.config(text="请输入 Markdown 内容！")
            return

        # 获取默认路径作为对话框的初始路径
        default_path = self.get_default_output_path()
        default_dir = os.path.dirname(default_path)
        default_file = os.path.basename(default_path)

        # 用户选择保存路径
        output_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word文件", "*.docx")],
            initialdir=default_dir,
            initialfile=default_file,
        )
        if not output_path:
            return

        # 执行转换
        self.do_conversion(markdown_text, output_path)

    def do_conversion(self, markdown_text, output_path):
        """执行实际的转换工作"""
        # 创建临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix=".md") as temp_file:
            temp_markdown_file = temp_file.name
        with open(temp_markdown_file, "w", encoding="utf-8") as f:
            f.write(markdown_text)

        try:
            # 构建命令参数
            cmd = ["pandoc", temp_markdown_file, "-o", output_path]

            # 添加数学公式处理选项
            if self.math_option.get() != "none":
                cmd.append(f"--{self.math_option.get()}")

            # 添加参考文档
            if self.use_reference_docx.get() and self.reference_docx_path.get():
                cmd.extend(["--reference-doc", self.reference_docx_path.get()])

            # 添加其他转换增强选项
            cmd.extend(["--standalone", "--wrap=none"])

            # 执行转换
            self.status_label.config(text="正在转换...")
            result = subprocess.run(cmd, capture_output=True, text=True, check=True)

            self.status_label.config(text=f"转换成功! 已保存到: {output_path}")

            # 转换计数器加1
            self.conversion_count += 1

            # 根据设置自动打开或询问是否打开
            if self.auto_open.get():
                self.open_file(output_path)
            else:
                if messagebox.askyesno(
                    "转换完成", f"Word 文件已成功保存到: {output_path}\n是否立即打开?"
                ):
                    self.open_file(output_path)

        except subprocess.CalledProcessError as e:
            error_msg = f"转换失败: {e.stderr}"
            self.status_label.config(text=error_msg)
            messagebox.showerror("转换错误", error_msg)
        finally:
            if os.path.exists(temp_markdown_file):
                os.remove(temp_markdown_file)

    def open_file(self, filepath):
        """使用系统默认程序打开文件"""
        if platform.system() == "Windows":
            os.startfile(filepath)
        elif platform.system() == "Darwin":  # macOS
            subprocess.run(["open", filepath])
        else:  # Linux
            subprocess.run(["xdg-open", filepath])


# 在 main() 函数中添加异常处理
def main():
    try:
        root = tk.Tk()
        app = MarkdownToWordConverter(root)
        root.mainloop()
    except Exception as e:
        # 创建错误日志文件
        with open("error_log.txt", "w") as f:
            import traceback

            f.write(f"发生错误: {str(e)}\n")
            f.write(traceback.format_exc())
        # 显示错误消息框
        import ctypes

        ctypes.windll.user32.MessageBoxW(
            0, f"程序启动出错: {str(e)}\n详细信息已保存到error_log.txt", "错误", 0
        )


if __name__ == "__main__":
    main()