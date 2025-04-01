import os
import sys
import shutil
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QListWidget, QPushButton, QFileDialog, 
                           QLabel, QProgressBar, QMessageBox, QTableWidget, 
                           QTableWidgetItem, QHeaderView, QStackedWidget)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QColor, QIcon
from pdf_merger import PDFMerger

def resource_path(relative_path):
    """获取资源的绝对路径，兼容开发环境和PyInstaller打包后的环境"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller打包后的路径
        base_path = sys._MEIPASS
    else:
        # 开发环境下的路径
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
    return os.path.join(base_path, relative_path)

class PDFProcessThread(QThread):
    progress = pyqtSignal(str, int)  # status message, folder count
    finished = pyqtSignal(str)  # output file path
    error = pyqtSignal(str)
    status_update = pyqtSignal(int, int)  # success_count, ignored_count

    def __init__(self, merger, folder_path, output_path=''):
        super().__init__()
        self.merger = merger
        self.folder_path = folder_path
        self.output_path = output_path
        
    def run(self):
        try:
            # 首先获取要处理的文件夹列表
            subfolders = [os.path.join(self.folder_path, subfolder) 
                         for subfolder in os.listdir(self.folder_path) 
                         if os.path.isdir(os.path.join(self.folder_path, subfolder))]
            
            doc = None
            try:
                doc = self.merger.create_document()
                total = len(subfolders)
                success_count = 0
                
                for i, subfolder_path in enumerate(subfolders, 1):
                    folder_name = os.path.basename(subfolder_path)
                    self.progress.emit(f"正在处理: {folder_name} ({i}/{total})", int((i / total) * 100))
                    if self.merger.merge_invoice_and_images_to_total_pdf(subfolder_path, doc):
                        success_count += 1
                    self.status_update.emit(success_count, len(self.merger.ignored_folders))
                
                if self.merger.folder_count > 0:
                    # 保存文件
                    timestamp = self.merger.get_timestamp()
                    parent_folder_name = os.path.basename(os.path.abspath(self.folder_path))
                    output_filename = f'{parent_folder_name}_报销单_自动生成_{self.merger.folder_count}张发票_{timestamp}.pdf'
                    output_path = os.path.join(self.folder_path, output_filename)
                    
                    doc.save(output_path)
                    self.finished.emit(output_path)
                else:
                    self.error.emit("没有成功处理任何文件夹")
            finally:
                if doc:
                    doc.close()
                    
        except Exception as e:
            self.error.emit(str(e))

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.merger = PDFMerger(debug_mode=True)
        self.current_path = os.getcwd()
        self.selected_folder = None
        self.output_file = None
        
        # 设置窗口图标，使用兼容打包环境的路径
        icon_path = resource_path(os.path.join('icon', 'icon.ico'))
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        else:
            print(f"图标文件未找到: {icon_path}")
            
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('发票处理助手')
        self.setGeometry(100, 100, 600, 600)

        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 创建堆叠式窗口部件
        self.stack = QStackedWidget()
        layout.addWidget(self.stack)

        # 创建各个页面
        self.createFolderSelectPage()
        self.createReportPage()
        self.createProcessPage()

        # 创建导航按钮布局
        nav_layout = QHBoxLayout()
        self.prev_button = QPushButton("上一步")
        self.next_button = QPushButton("下一步")
        nav_layout.addWidget(self.prev_button)
        nav_layout.addWidget(self.next_button)
        self.nav_layout = nav_layout  # 保存引用以便后续控制显示/隐藏
        layout.addLayout(nav_layout)

        # 绑定导航按钮事件
        self.prev_button.clicked.connect(self.prevPage)
        self.next_button.clicked.connect(self.nextPage)
        
        # 初始化显示第一个页面
        self.stack.setCurrentIndex(0)
        self.updateNavButtons()
        self.scanFolders()

    def createFolderSelectPage(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        
        # 创建和设置文件夹列表
        self.folder_list = QListWidget()
        layout.addWidget(QLabel("请选择要处理的文件夹:"))
        layout.addWidget(self.folder_list)
        
        # 添加双击处理事件
        self.folder_list.itemDoubleClicked.connect(lambda: self.nextPage())

        # 选择其他文件夹按钮
        select_button = QPushButton("选择其他文件夹")
        select_button.clicked.connect(self.selectFolder)
        layout.addWidget(select_button)

        self.stack.addWidget(page)

    def createReportPage(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        
        # 创建表格
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["文件夹", "PDF文件", "图片文件", "问题说明"])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        layout.addWidget(self.table)

        # 添加刷新按钮
        button_layout = QHBoxLayout()
        self.report_refresh_button = QPushButton("刷新")
        self.report_refresh_button.clicked.connect(self.refreshFolder)
        button_layout.addStretch()
        button_layout.addWidget(self.report_refresh_button)
        layout.addLayout(button_layout)

        self.stack.addWidget(page)

    def createProcessPage(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        
        # 状态和进度部分
        status_group = QWidget()
        status_layout = QVBoxLayout(status_group)
        status_layout.setContentsMargins(0, 0, 0, 20)  # 添加底部间距
        
        self.status_label = QLabel("准备处理...")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        status_layout.addWidget(self.status_label)
        status_layout.addWidget(self.progress_bar)
        layout.addWidget(status_group)
        
        # 统计信息部分
        self.stats_label = QLabel("")
        layout.addWidget(self.stats_label)
        
        # 忽略文件夹列表部分
        list_group = QWidget()
        list_layout = QVBoxLayout(list_group)
        list_layout.setContentsMargins(0, 10, 0, 10)  # 添加上下间距
        
        list_layout.addWidget(QLabel("处理失败的文件夹:"))
        self.ignored_list = QTableWidget()
        self.ignored_list.setColumnCount(4)
        self.ignored_list.setHorizontalHeaderLabels(["文件夹", "PDF文件", "图片文件", "原因"])
        header = self.ignored_list.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        list_layout.addWidget(self.ignored_list)
        layout.addWidget(list_group)

        # 结果和操作按钮部分
        result_group = QWidget()
        result_layout = QVBoxLayout(result_group)
        result_layout.setContentsMargins(0, 20, 0, 0)  # 添加顶部间距
        
        self.result_label = QLabel()
        result_layout.addWidget(self.result_label)
        
        # 文件操作按钮组
        self.file_ops_widget = QWidget()
        file_ops_layout = QHBoxLayout(self.file_ops_widget)
        file_ops_layout.setContentsMargins(0, 10, 0, 0)  # 添加顶部间距
        
        self.open_button = QPushButton("打开文件")
        self.move_button = QPushButton("移动文件")
        self.delete_button = QPushButton("删除文件")
        self.regenerate_button = QPushButton("重新处理")
        
        file_ops_layout.addWidget(self.open_button)
        file_ops_layout.addWidget(self.move_button)
        file_ops_layout.addWidget(self.delete_button)
        file_ops_layout.addWidget(self.regenerate_button)
        
        self.open_button.clicked.connect(self.openFile)
        self.move_button.clicked.connect(self.moveFile)
        self.delete_button.clicked.connect(self.deleteFile)
        self.regenerate_button.clicked.connect(self.regenerateFile)
        
        result_layout.addWidget(self.file_ops_widget)
        layout.addWidget(result_group)
        
        # 初始隐藏文件操作按钮
        self.file_ops_widget.hide()
        
        self.stack.addWidget(page)

    def updateNavButtons(self):
        current = self.stack.currentIndex()
        self.prev_button.setEnabled(current > 0)
        
        if current == 0:  # 文件夹选择页面
            self.next_button.setEnabled(bool(self.folder_list.currentItem()))
        elif current == 1:  # 报告页面
            if hasattr(self, 'folder_stats'):
                valid_count, total_count = self.folder_stats
                self.next_button.setText(f"开始处理 ({valid_count}/{total_count})")
                # 只有在有符合条件的文件夹时才启用开始处理按钮
                self.next_button.setEnabled(valid_count > 0)
            else:
                self.next_button.setText("开始处理")
        elif current == 2:  # 处理页面
            if not self.output_file:  # 正在处理中
                self.prev_button.setEnabled(False)
                self.next_button.setEnabled(False)
            else:  # 处理完成
                self.prev_button.hide()
                self.next_button.hide()

    def prevPage(self):
        current = self.stack.currentIndex()
        if current > 0:
            self.stack.setCurrentIndex(current - 1)
            self.updateNavButtons()

    def nextPage(self):
        current = self.stack.currentIndex()
        if current == 0:  # 从文件夹选择到报告页面
            self.selected_folder = os.path.join(self.current_path, 
                                              self.folder_list.currentItem().text())
            self.analyzeFolder()
        elif current == 1:  # 从报告页面到处理页面
            self.startProcessing()
        elif current == 3:  # 从结果页面完成
            self.close()
            return
            
        self.stack.setCurrentIndex(current + 1)
        self.updateNavButtons()

    def scanFolders(self):
        self.folder_list.clear()
        subfolders = [f for f in os.listdir(self.current_path) 
                     if os.path.isdir(os.path.join(self.current_path, f))]
        self.folder_list.addItems(subfolders)
        self.folder_list.itemClicked.connect(lambda: self.updateNavButtons())

    def selectFolder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹", self.current_path)
        if folder:
            self.current_path = folder
            self.scanFolders()

    def analyzeFolder(self):
        if not self.selected_folder:
            return
            
        subfolders = [f for f in os.listdir(self.selected_folder) 
                     if os.path.isdir(os.path.join(self.selected_folder, f))]
        
        self.table.setRowCount(len(subfolders))
        
        valid_count = 0  # 统计符合条件的文件夹数
        total_count = len(subfolders)  # 总文件夹数
        
        for i, folder in enumerate(subfolders):
            folder_path = os.path.join(self.selected_folder, folder)
            files = os.listdir(folder_path)
            pdf_count = len([f for f in files if f.lower().endswith('.pdf')])
            img_count = len([f for f in files if f.lower().endswith(('.png', '.jpg', '.jpeg'))])
            
            # 检查是否符合条件
            if pdf_count == 1 and img_count >= 2:
                valid_count += 1
            
            # 设置表格项
            folder_item = QTableWidgetItem(folder)
            pdf_item = QTableWidgetItem(str(pdf_count))
            img_item = QTableWidgetItem(str(img_count))
            reason_item = QTableWidgetItem("")
            
            # 设置颜色和原因
            reasons = []
            if pdf_count != 1:
                pdf_item.setBackground(QColor(255, 200, 200))
                reasons.append(f'缺少PDF')
            if img_count < 2:
                img_item.setBackground(QColor(255, 200, 200))
                reasons.append(f'缺少图片')
            
            if reasons:
                reason_item.setText("、".join(reasons))
                
            self.table.setItem(i, 0, folder_item)
            self.table.setItem(i, 1, pdf_item)
            self.table.setItem(i, 2, img_item)
            self.table.setItem(i, 3, reason_item)
        
        # 保存统计信息供updateNavButtons使用
        self.folder_stats = (valid_count, total_count)
        self.updateNavButtons()

    def startProcessing(self):
        self.merger = PDFMerger(debug_mode=True)
        self.thread = PDFProcessThread(self.merger, self.selected_folder)
        self.thread.progress.connect(self.updateProgress)
        self.thread.status_update.connect(self.updateStats)
        self.thread.finished.connect(self.processingFinished)
        self.thread.error.connect(self.processingError)
        self.thread.start()
        
        self.prev_button.setEnabled(False)
        self.next_button.setEnabled(False)
        
        # 清空忽略列表
        self.ignored_list.setRowCount(0)
        self.stats_label.setText("")

    def updateProgress(self, status, progress):
        self.status_label.setText(status)
        self.progress_bar.setValue(progress)
        if progress == 100:
            self.status_label.setText("处理完成，正在生成PDF文件...")

    def updateStats(self, success_count, ignored_count):
        self.stats_label.setText(f"成功: {success_count} 个文件夹，忽略: {ignored_count} 个文件夹")
        
        # 更新忽略文件夹列表
        self.ignored_list.setRowCount(len(self.merger.ignored_folders))
        self._updateIgnoredList(self.ignored_list)

    def _updateIgnoredList(self, table_widget):
        """更新忽略文件夹列表到指定的表格部件"""
        table_widget.setRowCount(len(self.merger.ignored_folders))
        for i, folder_data in enumerate(self.merger.ignored_folders):
            if len(folder_data) == 4:  # 包含PDF和图片计数的情况
                folder_path, pdf_count, img_count, reason = folder_data
                folder_name = os.path.basename(folder_path)
                
                folder_item = QTableWidgetItem(folder_name)
                pdf_item = QTableWidgetItem(str(pdf_count))
                img_item = QTableWidgetItem(str(img_count))
                reason_item = QTableWidgetItem("")  # 创建原因列
                
                # 设置颜色并生成原因文本
                reasons = []
                if pdf_count != 1:
                    pdf_item.setBackground(QColor(255, 200, 200))
                    reasons.append(f"缺少PDF")
                if img_count < 2:
                    img_item.setBackground(QColor(255, 200, 200))
                    reasons.append(f"缺少图片")
                
                reason_item.setText("、".join(reasons))
                    
                table_widget.setItem(i, 0, folder_item)
                table_widget.setItem(i, 1, pdf_item)
                table_widget.setItem(i, 2, img_item)
                table_widget.setItem(i, 3, reason_item)  # 添加原因列

    def processingFinished(self, output_file):
        self.output_file = output_file
        self.status_label.setText("处理完成！")
        self.progress_bar.setValue(100)
        self.result_label.setText(f"文件已保存到：{output_file}")
        
        # 更新忽略的文件夹列表
        self._updateIgnoredList(self.ignored_list)
        
        # 显示文件操作按钮，隐藏导航按钮
        self.file_ops_widget.show()
        self.prev_button.hide()
        self.next_button.hide()
        
        # 确保所有按钮都是启用的
        self.open_button.setEnabled(True)
        self.move_button.setEnabled(True)
        self.delete_button.setEnabled(True)
        self.regenerate_button.setEnabled(True)

    def processingError(self, error_message):
        self.status_label.setText("处理出错！")
        QMessageBox.critical(self, "错误", f"处理过程中出现错误：{error_message}")
        self.stack.setCurrentIndex(1)  # 返回到报告页面
        self.prev_button.show()
        self.next_button.show()
        self.file_ops_widget.hide()
        self.updateNavButtons()

    def moveFile(self):
        if not self.output_file or not os.path.exists(self.output_file):
            QMessageBox.warning(self, "警告", "找不到输出文件!")
            return
            
        target_dir = QFileDialog.getExistingDirectory(self, "选择目标文件夹")
        if target_dir:
            try:
                new_path = os.path.join(target_dir, os.path.basename(self.output_file))
                shutil.move(self.output_file, new_path)
                self.output_file = new_path
                self.result_label.setText(f"文件已移动到：{new_path}")
                QMessageBox.information(self, "成功", f"文件已移动到：{new_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"移动文件时出错：{str(e)}")

    def deleteFile(self):
        if not self.output_file or not os.path.exists(self.output_file):
            QMessageBox.warning(self, "警告", "找不到输出文件!")
            return
            
        reply = QMessageBox.question(self, "确认删除", 
                                   "确定要删除生成的PDF文件吗？",
                                   QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            try:
                os.remove(self.output_file)
                QMessageBox.information(self, "成功", "文件已删除")
                self.output_file = None
                self.result_label.setText("文件已删除")
                
                # 更新按钮状态
                self.open_button.setEnabled(False)
                self.move_button.setEnabled(False)
                self.delete_button.setEnabled(False)
                self.regenerate_button.setEnabled(True)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"删除文件时出错：{str(e)}")

    def openFile(self):
        """打开生成的PDF文件"""
        if not self.output_file or not os.path.exists(self.output_file):
            QMessageBox.warning(self, "警告", "找不到输出文件!")
            return
        
        # 使用系统默认的PDF查看器打开文件
        import subprocess
        try:
            if os.name == 'nt':  # Windows
                os.startfile(self.output_file)
            else:  # Linux/Mac
                subprocess.run(['xdg-open', self.output_file])
        except Exception as e:
            QMessageBox.critical(self, "错误", f"打开文件时出错：{str(e)}")

    def refreshFolder(self):
        """刷新当前文件夹的分析结果"""
        if not self.selected_folder:
            return
            
        # 清空标签
        self.stats_label.clear()
        self.result_label.clear()
            
        # 重新分析文件夹
        subfolders = [f for f in os.listdir(self.selected_folder) 
                     if os.path.isdir(os.path.join(self.selected_folder, f))]
        
        has_invalid = False
        for folder in subfolders:
            folder_path = os.path.join(self.selected_folder, folder)
            files = os.listdir(folder_path)
            pdf_count = len([f for f in files if f.lower().endswith('.pdf')])
            img_count = len([f for f in files if f.lower().endswith(('.png', '.jpg', '.jpeg'))])
            
            if pdf_count != 1 or img_count < 2:
                has_invalid = True
                break
        
        if not has_invalid:
            reply = QMessageBox.question(self, "确认", 
                                     "所有发票都已符合条件，是否重新处理？",
                                     QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.regenerateFile()
        else:
            # 更新显示
            self.analyzeFolder()

    def regenerateFile(self):
        """重新处理PDF文件"""
        if self.output_file and os.path.exists(self.output_file):
            try:
                os.remove(self.output_file)
            except Exception as e:
                QMessageBox.warning(self, "警告", f"删除原文件时出错：{str(e)}")
                return
        
        # 清空所有状态
        self.output_file = None
        self.result_label.clear()
        self.status_label.setText("准备处理...")
        self.progress_bar.setValue(0)
        self.stats_label.clear()
        self.ignored_list.setRowCount(0)
        
        # 返回到报告页面并更新界面状态
        self.stack.setCurrentIndex(1)
        self.prev_button.show()
        self.next_button.show()
        self.next_button.setEnabled(True)  # 确保"开始处理"按钮是启用的
        self.next_button.setText("开始处理")
        self.file_ops_widget.hide()
        
        # 刷新文件夹分析
        self.analyzeFolder()

if __name__ == '__main__':
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec_()