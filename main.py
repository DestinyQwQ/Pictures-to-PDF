import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import threading


class ImageToPdfConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("图片转PDF转换器")
        self.root.geometry("650x450")
        self.root.resizable(True, True)

        # 设置中文字体支持
        self.style = ttk.Style()
        self.style.configure("TLabel", font=("SimHei", 10))
        self.style.configure("TButton", font=("SimHei", 10))
        self.style.configure("TEntry", font=("SimHei", 10))

        # 源文件夹路径
        self.source_dir = tk.StringVar()
        # 输出PDF路径/文件夹
        self.output_path = tk.StringVar(value="output.pdf")
        # 转换模式：合并为一个PDF还是每个图片一个PDF
        self.conversion_mode = tk.StringVar(value="single")  # "single" 或 "multiple"
        # 转换进度
        self.progress_var = tk.DoubleVar()

        self.create_widgets()

    def create_widgets(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 源文件夹选择
        ttk.Label(main_frame, text="图片文件夹:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.source_dir, width=50).grid(row=0, column=1, pady=5)
        ttk.Button(main_frame, text="浏览...", command=self.browse_source).grid(row=0, column=2, padx=5, pady=5)

        # 转换模式选择
        mode_frame = ttk.Frame(main_frame)
        mode_frame.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=5)

        ttk.Label(mode_frame, text="转换模式:").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(
            mode_frame,
            text="所有图片合并为一个PDF",
            variable=self.conversion_mode,
            value="single",
            command=self.update_output_label
        ).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(
            mode_frame,
            text="每个图片单独生成一个PDF",
            variable=self.conversion_mode,
            value="multiple",
            command=self.update_output_label
        ).pack(side=tk.LEFT, padx=5)

        # 输出路径
        self.output_label = ttk.Label(main_frame, text="输出PDF文件:")
        self.output_label.grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_path, width=50).grid(row=2, column=1, pady=5)
        ttk.Button(main_frame, text="选择...", command=self.browse_output).grid(row=2, column=2, padx=5, pady=5)

        # 转换按钮
        ttk.Button(main_frame, text="开始转换", command=self.start_conversion).grid(row=3, column=0, columnspan=3,
                                                                                    pady=10)

        # 进度条
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=4, column=0, columnspan=3, sticky=tk.EW, pady=10)

        # 状态标签
        self.status_label = ttk.Label(main_frame, text="就绪")
        self.status_label.grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=5)

        # 日志文本框
        ttk.Label(main_frame, text="转换日志:").grid(row=6, column=0, sticky=tk.NW, pady=5)
        self.log_text = tk.Text(main_frame, height=10, width=60)
        self.log_text.grid(row=6, column=1, columnspan=2, pady=5, sticky=tk.NSEW)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(main_frame, command=self.log_text.yview)
        scrollbar.grid(row=6, column=3, sticky=tk.NS)
        self.log_text.config(yscrollcommand=scrollbar.set)

        # 设置权重，使控件可以随窗口大小调整
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)

    def update_output_label(self):
        """根据转换模式更新输出标签文本"""
        if self.conversion_mode.get() == "single":
            self.output_label.config(text="输出PDF文件:")
            if not self.output_path.get().endswith(".pdf"):
                self.output_path.set("output.pdf")
        else:
            self.output_label.config(text="输出文件夹:")
            # 如果当前是PDF文件路径，提取其所在文件夹
            current_path = self.output_path.get()
            if current_path.endswith(".pdf"):
                self.output_path.set(os.path.dirname(current_path) or os.getcwd())

    def browse_source(self):
        """选择图片所在的文件夹"""
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.source_dir.set(dir_path)
            self.log(f"已选择图片文件夹: {dir_path}")

    def browse_output(self):
        """根据转换模式选择输出路径"""
        if self.conversion_mode.get() == "single":
            # 单个PDF文件
            file_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
            )
            if file_path:
                self.output_path.set(file_path)
                self.log(f"已选择输出PDF文件: {file_path}")
        else:
            # 多个PDF文件的输出文件夹
            dir_path = filedialog.askdirectory()
            if dir_path:
                self.output_path.set(dir_path)
                self.log(f"已选择输出文件夹: {dir_path}")

    def log(self, message):
        """在日志文本框中添加信息"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)  # 滚动到最后一行

    def update_status(self, message):
        """更新状态标签"""
        self.status_label.config(text=message)

    def update_progress(self, value):
        """更新进度条"""
        self.progress_var.set(value)

    def is_image_file(self, filename):
        """检查文件是否为支持的图片格式"""
        ext = os.path.splitext(filename)[1].lower()
        return ext in ['.jpg', '.jpeg', '.png']

    def convert_single_image_to_pdf(self, image_path, output_path):
        """将单张图片转换为PDF"""
        try:
            with Image.open(image_path) as img:
                # 创建PDF
                c = canvas.Canvas(output_path, pagesize=letter)
                pdf_width, pdf_height = letter

                # 计算缩放比例
                img_width, img_height = img.size
                scale = min(pdf_width / img_width, pdf_height / img_height)
                new_width = img_width * scale
                new_height = img_height * scale

                # 居中放置图片
                x = (pdf_width - new_width) / 2
                y = (pdf_height - new_height) / 2

                # 添加图片到PDF
                c.drawImage(image_path, x, y, width=new_width, height=new_height)
                c.save()
            return True
        except Exception as e:
            self.log(f"处理图片时出错: {str(e)}")
            return False

    def convert_images_to_pdf(self):
        """将图片转换为PDF的核心函数"""
        source_dir = self.source_dir.get()
        output_path = self.output_path.get()
        conversion_mode = self.conversion_mode.get()

        # 验证输入
        if not source_dir:
            messagebox.showerror("错误", "请选择图片文件夹")
            self.update_status("就绪")
            return

        if not os.path.isdir(source_dir):
            messagebox.showerror("错误", f"文件夹不存在: {source_dir}")
            self.update_status("就绪")
            return

        if not output_path:
            messagebox.showerror("错误", "请指定输出路径")
            self.update_status("就绪")
            return

        # 验证输出路径
        if conversion_mode == "single":
            # 确保输出文件夹存在
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                messagebox.showerror("错误", f"输出文件夹不存在: {output_dir}")
                self.update_status("就绪")
                return
        else:
            # 确保输出文件夹存在
            if not os.path.exists(output_path):
                try:
                    os.makedirs(output_path)
                except Exception as e:
                    messagebox.showerror("错误", f"无法创建输出文件夹: {str(e)}")
                    self.update_status("就绪")
                    return

        # 获取所有图片文件
        image_files = [f for f in os.listdir(source_dir)
                       if os.path.isfile(os.path.join(source_dir, f))
                       and self.is_image_file(f)]

        if not image_files:
            messagebox.showinfo("提示", "所选文件夹中没有找到JPG或PNG图片")
            self.update_status("就绪")
            return

        # 按文件名排序
        image_files.sort()
        total_files = len(image_files)

        self.log(f"找到 {total_files} 个图片文件，开始转换...")

        try:
            if conversion_mode == "single":
                # 所有图片合并为一个PDF
                c = canvas.Canvas(output_path, pagesize=letter)

                for i, filename in enumerate(image_files):
                    file_path = os.path.join(source_dir, filename)
                    self.log(f"处理: {filename}")

                    try:
                        with Image.open(file_path) as img:
                            # 计算图片在PDF中的尺寸和位置（保持比例）
                            img_width, img_height = img.size
                            pdf_width, pdf_height = letter

                            # 计算缩放比例
                            scale = min(pdf_width / img_width, pdf_height / img_height)
                            new_width = img_width * scale
                            new_height = img_height * scale

                            # 居中放置图片
                            x = (pdf_width - new_width) / 2
                            y = (pdf_height - new_height) / 2

                            # 添加图片到PDF
                            c.drawImage(file_path, x, y, width=new_width, height=new_height)
                            c.showPage()  # 新建一页

                    except Exception as e:
                        self.log(f"处理 {filename} 时出错: {str(e)}")

                    # 更新进度
                    progress = (i + 1) / total_files * 100
                    self.update_progress(progress)
                    self.update_status(f"正在转换: {i + 1}/{total_files}")

                # 保存PDF
                c.save()
                self.update_progress(100)
                self.update_status("转换完成")
                self.log(f"PDF文件已生成: {output_path}")
                messagebox.showinfo("成功", f"已成功将 {total_files} 张图片转换为一个PDF\n文件保存至: {output_path}")

            else:
                # 每个图片单独生成一个PDF
                success_count = 0

                for i, filename in enumerate(image_files):
                    file_path = os.path.join(source_dir, filename)
                    # 生成输出PDF文件名（与图片同名，替换扩展名为.pdf）
                    base_name = os.path.splitext(filename)[0]
                    pdf_filename = f"{base_name}.pdf"
                    pdf_path = os.path.join(output_path, pdf_filename)

                    self.log(f"处理: {filename} -> {pdf_filename}")

                    if self.convert_single_image_to_pdf(file_path, pdf_path):
                        success_count += 1

                    # 更新进度
                    progress = (i + 1) / total_files * 100
                    self.update_progress(progress)
                    self.update_status(f"正在转换: {i + 1}/{total_files}")

                self.update_progress(100)
                self.update_status("转换完成")
                self.log(f"转换完成，共生成 {success_count} 个PDF文件")
                messagebox.showinfo("成功",
                                    f"已成功将 {success_count}/{total_files} 张图片转换为PDF\n文件保存至: {output_path}")

        except Exception as e:
            self.log(f"转换失败: {str(e)}")
            messagebox.showerror("错误", f"转换过程中发生错误: {str(e)}")
            self.update_status("转换失败")

    def start_conversion(self):
        """开始转换过程（在新线程中运行以避免界面冻结）"""
        # 清空日志
        self.log_text.delete(1.0, tk.END)
        self.update_progress(0)
        self.update_status("准备转换...")

        # 在新线程中执行转换，避免界面冻结
        conversion_thread = threading.Thread(target=self.convert_images_to_pdf)
        conversion_thread.daemon = True
        conversion_thread.start()


if __name__ == "__main__":
    root = tk.Tk()
    app = ImageToPdfConverter(root)
    root.mainloop()