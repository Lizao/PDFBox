import sys
import os
import traceback

def resource_path(relative_path):
    """获取资源的绝对路径，用于PyInstaller打包"""
    try:
        # PyInstaller创建临时文件夹，将路径存储在_MEIPASS中
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def main():


    import sys
    import os
    from io import BytesIO
    from PyQt5.QtWidgets import (
        QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
        QFileDialog, QMessageBox, QListWidget, QInputDialog, QLabel, QDialog,
        QScrollArea, QFrame, QGridLayout, QStyle
    )
    from PyQt5.QtGui import QPixmap, QImage, QFont, QIcon, QPainter
    from PyQt5.QtCore import Qt, QSize
    from PyPDF2 import PdfReader, PdfWriter
    import fitz  # PyMuPDF

    # ------------------ 压缩图标（Unicode 📦） ------------------
    def get_compress_icon():
        icon = QIcon()
        pixmap = QPixmap(36,36)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        painter.setFont(QFont("Segoe UI Emoji", 28))
        painter.drawText(pixmap.rect(), Qt.AlignCenter, "📦")
        painter.end()
        icon.addPixmap(pixmap)
        return icon

    # ---------------- 预览窗口 ----------------
    class PreviewWindow(QDialog):
        def __init__(self, pages, page_index=0):
            super().__init__()
            self.setWindowTitle("PDF 页面预览")
            self.setGeometry(100, 50, 1000, 800)
            self.pages = pages
            self.page_index = page_index
            self.scale = 1.0
            self.cache = {}

            self.scroll = QScrollArea()
            self.scroll.setWidgetResizable(True)
            self.label = QLabel()
            self.label.setAlignment(Qt.AlignCenter)
            self.scroll.setWidget(self.label)

            btn_prev = QPushButton("⬅ 上一页")
            btn_next = QPushButton("下一页 ➡")
            btn_zoom_in = QPushButton("放大 +")
            btn_zoom_out = QPushButton("缩小 -")
            btn_reset = QPushButton("重置")

            btn_prev.clicked.connect(self.prev_page)
            btn_next.clicked.connect(self.next_page)
            btn_zoom_in.clicked.connect(self.zoom_in)
            btn_zoom_out.clicked.connect(self.zoom_out)
            btn_reset.clicked.connect(self.reset_zoom)

            btn_layout = QHBoxLayout()
            for btn in [btn_prev, btn_next, btn_zoom_in, btn_zoom_out, btn_reset]:
                btn.setFixedHeight(50)
                btn.setFont(QFont("Microsoft YaHei", 12, QFont.Bold))
                btn_layout.addWidget(btn)

            layout = QVBoxLayout()
            layout.addWidget(self.scroll)
            layout.addLayout(btn_layout)
            self.setLayout(layout)

            self.show_page()

        def render_page(self):
            key = (self.page_index, self.scale)
            if key in self.cache:
                self.label.setPixmap(self.cache[key])
                return

            pdf_writer = PdfWriter()
            pdf_writer.add_page(self.pages[self.page_index])
            buffer = BytesIO()
            pdf_writer.write(buffer)
            buffer.seek(0)

            doc = fitz.open(stream=buffer.read(), filetype="pdf")
            page = doc[0]
            mat = fitz.Matrix(self.scale, self.scale)
            pix = page.get_pixmap(matrix=mat)

            if pix.alpha:
                img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGBA8888)
            else:
                img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)

            pixmap = QPixmap.fromImage(img)
            self.label.setPixmap(pixmap)
            self.label.adjustSize()
            self.cache[key] = pixmap
            doc.close()

        def show_page(self):
            self.setWindowTitle(f"PDF 预览 - 页 {self.page_index+1}/{len(self.pages)} (缩放 {self.scale*100:.0f}%)")
            self.render_page()

        def prev_page(self):  self.page_index = max(self.page_index-1, 0); self.show_page()
        def next_page(self):  self.page_index = min(self.page_index+1, len(self.pages)-1); self.show_page()
        def zoom_in(self):   self.scale *= 1.2; self.show_page()
        def zoom_out(self):  self.scale /= 1.2; self.show_page()
        def reset_zoom(self): self.scale = 1.0; self.show_page()

    # ---------------- PDF 工具 ----------------
    class PDFTool(QWidget):
        def __init__(self):
            super().__init__()
            self.setWindowTitle("PDF 工具箱")
            self.setGeometry(200, 200, 1400, 750)
            self.pdf_path = ""
            self.pages = []
            self.reader = None
            self.initUI()

        def initUI(self):
            main_layout = QHBoxLayout()

            # ---------- 左侧按钮 ----------
            left_frame = QFrame()
            left_frame.setStyleSheet("""
                QFrame {
                    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                                stop:0 #4a90e2, stop:1 #3a70c1);
                    border-radius: 12px;
                }
            """)
            left_layout = QVBoxLayout()
            left_layout.setContentsMargins(20, 20, 20, 20)
            left_layout.setSpacing(25)

            left_buttons_info = [
                ("剪切 PDF", self.style().standardIcon(QStyle.SP_DesktopIcon), self.cut_pdf),
                ("合并 PDF", self.style().standardIcon(QStyle.SP_DirOpenIcon), self.merge_pdf),
                ("拆分 PDF", self.style().standardIcon(QStyle.SP_FileDialogDetailedView), self.split_pdf),
                ("旋转 PDF", self.style().standardIcon(QStyle.SP_BrowserReload), self.rotate_pdf),
                ("压缩 PDF", get_compress_icon(), self.compress_pdf)
            ]
            for text, icon, slot in left_buttons_info:
                btn = QPushButton(text)
                btn.setFixedSize(220, 70)
                btn.setIcon(icon)
                btn.setIconSize(QSize(36,36))
                btn.setFont(QFont("Microsoft YaHei", 16, QFont.Bold))
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: #ffffff;
                        color: #3a70c1;
                        border-radius: 14px;
                        text-align: left;
                        padding-left: 20px;
                    }
                    QPushButton:hover { background-color: #e6f0ff; }
                    QPushButton:pressed { background-color: #cce0ff; }
                """)
                btn.clicked.connect(slot)
                left_layout.addWidget(btn)
            left_layout.addStretch()
            left_frame.setLayout(left_layout)

            # ---------- 右侧编辑界面 ----------
            right_frame = QFrame()
            right_frame.setStyleSheet("background-color: #6fa8dc; border-radius: 12px;")
            right_layout = QVBoxLayout()
            right_layout.setContentsMargins(15,15,15,15)
            right_layout.setSpacing(15)

            btn_open = QPushButton("📂 打开 PDF 编辑")
            btn_open.setIcon(self.style().standardIcon(QStyle.SP_DirOpenIcon))
            btn_open.setIconSize(QSize(28,28))
            btn_open.setFixedHeight(60)
            btn_open.setFont(QFont("Microsoft YaHei", 17, QFont.Bold))
            btn_open.setStyleSheet("""
                QPushButton {
                    background-color: #ffffff;
                    color: #3a70c1;
                    border-radius:12px;
                    text-align: center;
                }
                QPushButton:hover {background-color: #e6f0ff;}
                QPushButton:pressed {background-color: #cce0ff;}
            """)
            btn_open.clicked.connect(self.open_pdf_edit)
            right_layout.addWidget(btn_open)

            self.page_list = QListWidget()
            self.page_list.setStyleSheet("""
                QListWidget {
                    background-color:#ffffff;
                    border:1px solid #3a70c1;
                    font-size:16px;
                }
                QListWidget::item:selected {
                    background-color:#f0f0ff;
                    color:#000;
                }
            """)
            right_layout.addWidget(self.page_list)

            page_btn_grid = QGridLayout()
            page_btn_grid.setSpacing(15)
            page_buttons_info = [
                ("上移页", self.style().standardIcon(QStyle.SP_ArrowUp), self.move_up),
                ("下移页", self.style().standardIcon(QStyle.SP_ArrowDown), self.move_down),
                ("删除页", self.style().standardIcon(QStyle.SP_TrashIcon), self.delete_page),
                ("旋转页", self.style().standardIcon(QStyle.SP_BrowserReload), self.rotate_page),
                ("插入页", self.style().standardIcon(QStyle.SP_FileDialogNewFolder), self.insert_page),
                ("预览页", self.style().standardIcon(QStyle.SP_FileDialogContentsView), self.open_preview),
                ("保存 PDF", self.style().standardIcon(QStyle.SP_DialogSaveButton), self.save_pdf)
            ]
            for index, (text, icon, slot) in enumerate(page_buttons_info):
                btn = QPushButton(text)
                btn.setFixedSize(200, 55)
                btn.setIcon(icon)
                btn.setIconSize(QSize(28,28))
                btn.setFont(QFont("Microsoft YaHei", 14, QFont.Bold))
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: #ffffff;
                        color: #3a70c1;
                        border-radius:12px;
                        text-align: center;
                    }
                    QPushButton:hover {background-color: #e6f0ff;}
                    QPushButton:pressed {background-color: #cce0ff;}
                """)
                btn.clicked.connect(slot)
                row = index // 2
                col = index % 2
                page_btn_grid.addWidget(btn,row,col)
            right_layout.addLayout(page_btn_grid)
            right_frame.setLayout(right_layout)

            main_layout.addWidget(left_frame)
            main_layout.addSpacing(20)
            main_layout.addWidget(right_frame, stretch=3)
            self.setLayout(main_layout)

        # ---------------- PDF 功能 ----------------
        def cut_pdf(self):
            file, _ = QFileDialog.getOpenFileName(self, "选择 PDF", "", "PDF Files (*.pdf)")
            if not file: return
            start, ok1 = QInputDialog.getInt(self, "起始页", "起始页码：", 1, 1)
            if not ok1: return
            end, ok2 = QInputDialog.getInt(self, "终止页", "终止页码：", start, start)
            if not ok2: return
            reader = PdfReader(file)
            writer = PdfWriter()
            for i in range(start - 1, end):
                if i < len(reader.pages):
                    writer.add_page(reader.pages[i])
            save_path = os.path.join(os.path.dirname(file),
                                     f"{os.path.splitext(os.path.basename(file))[0]}({start}-{end}).pdf")
            with open(save_path, "wb") as f:
                writer.write(f)
            QMessageBox.information(self, "完成", f"剪切完成\n保存为 {save_path}")

        def merge_pdf(self):
            dlg = MergeDialog(self)
            dlg.exec_()

        def split_pdf(self):
            file, _ = QFileDialog.getOpenFileName(self, "选择 PDF 文件", "", "PDF Files (*.pdf)")
            if not file: return
            base_name = os.path.splitext(os.path.basename(file))[0]
            dir_name = os.path.join(os.path.dirname(file), base_name)
            os.makedirs(dir_name, exist_ok=True)
            reader = PdfReader(file)
            for i, page in enumerate(reader.pages):
                writer = PdfWriter()
                writer.add_page(page)
                save_path = os.path.join(dir_name, f"{base_name}({i + 1}).pdf")
                with open(save_path, "wb") as f: writer.write(f)
            QMessageBox.information(self, "完成", f"拆分完成\n保存至 {dir_name}")

        def rotate_pdf(self):
            file, _ = QFileDialog.getOpenFileName(self, "选择 PDF 文件", "", "PDF Files (*.pdf)")
            if not file: return
            angle, ok = QInputDialog.getInt(self, "旋转角度", "旋转角度（90/180/270）：", 90)
            if not ok or angle % 90 != 0:
                QMessageBox.warning(self, "错误", "请输入有效角度")
                return
            reader = PdfReader(file)
            writer = PdfWriter()
            for page in reader.pages:
                page.rotate(angle)
                writer.add_page(page)
            save_file, _ = QFileDialog.getSaveFileName(self, "保存旋转后的 PDF", "rotated.pdf", "PDF Files (*.pdf)")
            if save_file:
                with open(save_file, "wb") as f: writer.write(f)
                QMessageBox.information(self, "完成", f"旋转完成\n保存为 {save_file}")

        def compress_pdf(self):
            from PyQt5.QtCore import QThread, pyqtSignal
            from PyQt5.QtWidgets import QProgressDialog

            class CompressThread(QThread):
                progress = pyqtSignal(int)
                finished = pyqtSignal(str, bool)  # 参数: 保存路径, 是否成功
                error = pyqtSignal(str)

                def __init__(self, input_path, output_path, quality_level):
                    super().__init__()
                    self.input_path = input_path
                    self.output_path = output_path
                    self.quality_level = quality_level

                def run(self):
                    try:
                        doc = fitz.open(self.input_path)
                        total_pages = len(doc)

                        # 根据压缩等级设置参数
                        if self.quality_level == "高质量 (大文件)":
                            # 高质量：使用最少的压缩，保持高质量
                            compress_params = {
                                "garbage": 1,  # 最小程度的垃圾回收
                                "deflate": False,  # 不使用压缩
                                "clean": True,
                                "linear": False  # 移除linear参数
                            }
                        elif self.quality_level == "中等质量":
                            # 中等质量：平衡压缩和质量
                            compress_params = {
                                "garbage": 3,  # 中等垃圾回收
                                "deflate": True,  # 使用压缩
                                "clean": True,
                                "linear": False  # 移除linear参数
                            }
                        else:  # 小文件 (低质量)
                            # 低质量：最大压缩
                            compress_params = {
                                "garbage": 4,  # 最大垃圾回收
                                "deflate": True,  # 使用压缩
                                "clean": True,
                                "linear": False  # 移除linear参数
                            }

                        # 计算实际压缩步骤
                        total_steps = 50  # 总共50步，更精细的控制

                        # 步骤1: 读取和准备文档
                        for i in range(10):
                            self.progress.emit(i * 2)
                            self.msleep(10)

                        # 步骤2: 执行压缩
                        self.progress.emit(20)

                        # 保存PDF，移除linear和ascii参数
                        doc.save(self.output_path,
                                 garbage=compress_params["garbage"],
                                 deflate=compress_params["deflate"],
                                 clean=compress_params["clean"])

                        # 步骤3: 完成压缩
                        for i in range(20, 101):
                            self.progress.emit(i)
                            self.msleep(5)

                        doc.close()

                        self.finished.emit(self.output_path, True)

                    except Exception as e:
                        self.error.emit(str(e))

            # 选择PDF文件
            file, _ = QFileDialog.getOpenFileName(self, "选择 PDF 文件", "", "PDF Files (*.pdf)")
            if not file:
                return

            # 获取原文件名和路径
            original_name = os.path.splitext(os.path.basename(file))[0]
            original_dir = os.path.dirname(file)

            # 选择压缩等级并确定文件名后缀
            levels = {
                "高质量 (大文件)": "高质量",
                "中等质量": "中质量",
                "小文件 (低质量)": "低质量"
            }

            level, ok = QInputDialog.getItem(
                self,
                "选择压缩等级",
                "压缩等级:",
                list(levels.keys()),
                1,
                False
            )

            if not ok:
                return

            # 生成新的文件名（在原文件夹中）
            suffix = levels[level]
            default_name = f"{original_name}_{suffix}.pdf"
            save_path = os.path.join(original_dir, default_name)

            # 弹出保存对话框，默认位置为原文件夹
            save_file, _ = QFileDialog.getSaveFileName(
                self,
                "保存压缩后的 PDF",
                save_path,  # 默认路径
                "PDF Files (*.pdf)"
            )

            if not save_file:
                return

            # 创建进度对话框
            progress_dialog = QProgressDialog("正在压缩PDF...", "取消", 0, 100, self)
            progress_dialog.setWindowTitle("PDF压缩")
            progress_dialog.setWindowModality(Qt.WindowModal)
            progress_dialog.setMinimumDuration(0)
            progress_dialog.setAutoClose(False)
            progress_dialog.setAutoReset(False)
            progress_dialog.setMinimumWidth(300)

            # 创建并启动压缩线程
            compress_thread = CompressThread(file, save_file, level)

            # 连接信号
            def update_progress(value):
                progress_dialog.setValue(value)

            def on_finished(output_path, success):
                # 关闭进度对话框
                progress_dialog.close()
                if success:
                    # 检查文件是否存在
                    if os.path.exists(output_path):
                        # 显示文件大小对比
                        original_size = os.path.getsize(file) / 1024  # KB
                        compressed_size = os.path.getsize(output_path) / 1024  # KB
                        reduction = (1 - compressed_size / original_size) * 100 if original_size > 0 else 0

                        QMessageBox.information(
                            self,
                            "压缩完成",
                            f"压缩完成！\n\n"
                            f"原文件: {os.path.basename(file)} ({original_size:.1f} KB)\n"
                            f"新文件: {os.path.basename(output_path)} ({compressed_size:.1f} KB)\n"
                            f"压缩率: {reduction:.1f}%\n"
                            f"保存位置: {output_path}"
                        )
                    else:
                        QMessageBox.critical(self, "错误", f"压缩失败：输出文件不存在\n{output_path}")
                else:
                    QMessageBox.critical(self, "错误", "压缩过程中发生错误")

            def on_error(error_msg):
                progress_dialog.close()
                QMessageBox.critical(self, "压缩错误", f"压缩失败:\n{error_msg}")

            compress_thread.progress.connect(update_progress)
            compress_thread.finished.connect(on_finished)
            compress_thread.error.connect(on_error)

            # 取消按钮的处理
            def cancel_compress():
                if compress_thread.isRunning():
                    compress_thread.terminate()
                    compress_thread.wait()
                progress_dialog.close()

            progress_dialog.canceled.connect(cancel_compress)

            # 启动线程
            compress_thread.start()

        # ---------------- 页面操作函数 ----------------
        def open_pdf_edit(self):
            file, _ = QFileDialog.getOpenFileName(self, "选择 PDF 编辑", "", "PDF Files (*.pdf)")
            if not file: return
            self.pdf_path = file
            self.reader = PdfReader(file)
            self.pages = [page for page in self.reader.pages]
            self.refresh_page_list()

        def refresh_page_list(self):
            self.page_list.clear()
            for i in range(len(self.pages)):
                self.page_list.addItem(f"页 {i + 1}")

        def move_up(self):
            row = self.page_list.currentRow()
            if row > 0:
                self.pages[row - 1], self.pages[row] = self.pages[row], self.pages[row - 1]
                self.refresh_page_list()
                self.page_list.setCurrentRow(row - 1)

        def move_down(self):
            row = self.page_list.currentRow()
            if row < len(self.pages) - 1 and row >= 0:
                self.pages[row + 1], self.pages[row] = self.pages[row], self.pages[row + 1]
                self.refresh_page_list()
                self.page_list.setCurrentRow(row + 1)

        def delete_page(self):
            row = self.page_list.currentRow()
            if row >= 0:
                del self.pages[row]
                self.refresh_page_list()

        def rotate_page(self):
            row = self.page_list.currentRow()
            if row >= 0:
                angle, ok = QInputDialog.getInt(self, "旋转角度", "输入旋转角度(90/180/270):", 90)
                if ok and angle % 90 == 0:
                    try:
                        self.pages[row].rotate(angle)
                    except Exception as e:
                        QMessageBox.warning(self, "错误", f"旋转失败: {str(e)}")
                else:
                    QMessageBox.warning(self, "错误", "请输入有效角度(90/180/270)")

        def insert_page(self):
            file, _ = QFileDialog.getOpenFileName(self, "选择 PDF 文件插入", "", "PDF Files (*.pdf)")
            if not file: return
            reader = PdfReader(file)
            row = self.page_list.currentRow()
            if row < 0: row = len(self.pages) - 1
            for i, page in enumerate(reader.pages):
                self.pages.insert(row + i + 1, page)
            self.refresh_page_list()

        def open_preview(self):
            if not self.pages:
                QMessageBox.warning(self, "提示", "请先加载 PDF")
                return
            row = self.page_list.currentRow()
            if row < 0:
                row = 0
            self.preview_window = PreviewWindow(self.pages, row)
            self.preview_window.show()

        def save_pdf(self):
            if not self.pages:
                QMessageBox.warning(self, "错误", "没有可保存的 PDF 页面")
                return
            save_path = os.path.splitext(self.pdf_path)[0] + "_edited.pdf"
            writer = PdfWriter()
            for page in self.pages:
                writer.add_page(page)
            with open(save_path, "wb") as f:
                writer.write(f)
            QMessageBox.information(self, "完成", f"PDF 保存成功\n路径: {save_path}")

    from PyQt5.QtWidgets import QDialog, QListWidget, QVBoxLayout, QPushButton, QHBoxLayout

    class MergeDialog(QDialog):
        """用于选择和拖动排序合并 PDF 文件"""
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setWindowTitle("合并 PDF 文件")
            self.setGeometry(200, 200, 500, 600)
            self.pdf_files = []

            layout = QVBoxLayout()

            self.list_widget = QListWidget()
            self.list_widget.setAcceptDrops(True)
            self.list_widget.setDragEnabled(True)
            self.list_widget.setDragDropMode(QListWidget.InternalMove)
            layout.addWidget(self.list_widget)

            btn_layout = QHBoxLayout()
            btn_add = QPushButton("添加文件")
            btn_remove = QPushButton("删除文件")
            btn_ok = QPushButton("合并并保存")
            btn_cancel = QPushButton("取消")
            btn_layout.addWidget(btn_add)
            btn_layout.addWidget(btn_remove)
            btn_layout.addWidget(btn_ok)
            btn_layout.addWidget(btn_cancel)

            layout.addLayout(btn_layout)
            self.setLayout(layout)

            btn_add.clicked.connect(self.add_files)
            btn_remove.clicked.connect(self.remove_file)
            btn_ok.clicked.connect(self.merge_files)
            btn_cancel.clicked.connect(self.reject)

        def add_files(self):
            files, _ = QFileDialog.getOpenFileNames(self, "选择 PDF 文件", "", "PDF Files (*.pdf)")
            if files:
                for f in files:
                    if f not in self.pdf_files:
                        self.pdf_files.append(f)
                        self.list_widget.addItem(f)

        def remove_file(self):
            row = self.list_widget.currentRow()
            if row >= 0:
                self.pdf_files.pop(row)
                self.list_widget.takeItem(row)

        def merge_files(self):
            if len(self.pdf_files) < 2:
                QMessageBox.warning(self, "提示", "请选择至少两个 PDF 文件")
                return
            save_file, _ = QFileDialog.getSaveFileName(self, "保存合并 PDF", "merged.pdf", "PDF Files (*.pdf)")
            if not save_file:
                return
            writer = PdfWriter()
            for i in range(self.list_widget.count()):
                pdf_path = self.list_widget.item(i).text()
                reader = PdfReader(pdf_path)
                for page in reader.pages:
                    writer.add_page(page)
            with open(save_file, "wb") as f:
                writer.write(f)
            QMessageBox.information(self, "完成", f"合并完成\n保存为 {save_file}")
            self.accept()

    # ---------------- 主程序 ----------------
    if __name__ == "__main__":
        app = QApplication(sys.argv)
        app.setFont(QFont("Microsoft YaHei", 11))
        window = PDFTool()
        window.show()
        sys.exit(app.exec_())

if __name__ == "__main__":
    try:
        main()
    except Exception:
        # write log
        with open(os.path.join(base_path, "error_log.txt"), "w", encoding="utf-8") as f:
            traceback.print_exc(file=f)
        print("程序异常，请查看 error_log.txt")
        input("按回车退出")
        sys.exit(1)