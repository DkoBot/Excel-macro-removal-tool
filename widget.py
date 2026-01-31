# This Python file uses the following encoding: utf-8
import os
import sys
import Core
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QFileDialog,
    QMenu,
    QMessageBox,
    QTreeWidgetItem,
    QWidget,
)

# Important:
# You need to run the following command to generate the ui_form.py file
#     pyside6-uic form.ui -o ui_form.py, or
#     pyside2-uic form.ui -o ui_form.py
import rc_Ico  # 确保图标资源已注册
from ui_form import Ui_Widget

# 支持的 Office 文件扩展名（与选择器一致）
OFFICE_EXTENSIONS = (
    ".docm", ".dotm", ".xlsm", ".xltm", ".pptm", ".potm",
    ".doc", ".xls", ".ppt", ".xlsx",
)
FOLDER_ICON_PATH = ":/icopng/ico/文件夹.png"
DELETE_ICON_PATH = ":/icopng/ico/删除.png"


class ScanFolderWorker(QThread):
    """在子线程中递归扫描文件夹，通过信号汇报进度和结果。"""
    progress = Signal(int, str)  # files_found_count, current_dirpath
    finished_result = Signal(str, str, list)  # folder_name, root, [(full_path, rel_path), ...]

    def __init__(self, dir_path: str):
        super().__init__()
        self._dir_path = dir_path

    def run(self) -> None:
        root = os.path.abspath(self._dir_path)
        folder_name = os.path.basename(root.rstrip(os.sep)) or root
        to_add: list[tuple[str, str]] = []
        try:
            for dirpath, _dirnames, filenames in os.walk(root):
                for name in filenames:
                    ext = os.path.splitext(name)[1].lower()
                    if ext not in OFFICE_EXTENSIONS:
                        continue
                    full = os.path.normpath(os.path.join(dirpath, name))
                    rel = os.path.relpath(full, root)
                    to_add.append((full, rel))
                self.progress.emit(len(to_add), dirpath)
        except OSError:
            pass
        self.finished_result.emit(folder_name, root, to_add)


class CleanMacroWorker(QThread):
    """在子线程中执行宏清理，通过信号汇报进度和结果。"""
    progress = Signal(int, int)   # current_index, total
    finished_result = Signal(int, int)  # success_count, total_vba_size

    def __init__(self, file_paths: list[str], replace_original: bool, generate_report: bool, is_english: bool = False):
        super().__init__()
        self._file_paths = file_paths
        self._replace_original = replace_original
        self._generate_report = generate_report
        self._is_english = is_english

    def run(self) -> None:
        total = len(self._file_paths)
        sun_count = 0
        total_vba_size = 0
        for i, path in enumerate(self._file_paths):
            sun_count += Core.clean_vba_macro(path, self._replace_original, self._generate_report, self._is_english)
            total_vba_size += Core.get_last_vba_size()
            self.progress.emit(i + 1, total)
        self.finished_result.emit(sun_count, total_vba_size)


def _delete_icon() -> QIcon:
    """获取删除菜单图标（资源优先，失败时从文件加载）。"""
    icon = QIcon(DELETE_ICON_PATH)
    if icon.isNull():
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ico", "删除.png")
        icon = QIcon(path)
    return icon


class Widget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ui = Ui_Widget()
        self.ui.setupUi(self)
        # 字典：文件完整路径 -> (所在目录, 文件名)，支持递归与同名文件
        self._file_path_map: dict[str, tuple[str, str]] = {}
        self.ui.pushButton_2.clicked.connect(self._on_select_path_clicked)
        self.ui.pushButton.clicked.connect(self._on_push_button_clicked)
        self._next_file_id = 1
        self._is_english = False
        self._setup_tree_context_menu()
        self.ui.comboBox.currentIndexChanged.connect(self._on_language_changed)
        # 初始语言与 comboBox 一致：0=中文，1=英文
        self.EditLanguage(self.ui.comboBox.currentIndex() == 1)

    def _on_language_changed(self) -> None:
        """comboBox 切换时更新界面语言。"""
        self.EditLanguage(self.ui.comboBox.currentIndex() == 1)

    def EditLanguage(self, is_english: bool = False) -> None:
        """根据 is_english 设置所有控件为中文或英文。"""
        if is_english:
            self.setWindowTitle("Excel Macro Cleaner By: CNFsToT")
            self.ui.pushButton_2.setText("Select file or folder")
            self.ui.treeWidget.headerItem().setText(0, "ID")
            self.ui.treeWidget.headerItem().setText(1, "File name")
            self.ui.checkBox.setText("Extract macro for forensics and generate report")
            self.ui.label_2.setText("Total macro size:")
            self.ui.checkBox_2.setText("Keep original file")
            self.ui.comboBox.setItemText(0, "Chinese")
            self.ui.comboBox.setItemText(1, "English")
            self.ui.label.setText("Total processed:")
            self.ui.pushButton.setText("Start cleaning")
        else:
            self.setWindowTitle("Excel 宏清理工具 By: CNFsToT")
            self.ui.pushButton_2.setText("选中文件或文件夹")
            self.ui.treeWidget.headerItem().setText(0, "ID")
            self.ui.treeWidget.headerItem().setText(1, "文件名")
            self.ui.checkBox.setText("提取宏脚本用于取证生成报告")
            self.ui.label_2.setText("宏总大小:")
            self.ui.checkBox_2.setText("保留原文件")
            self.ui.comboBox.setItemText(0, "中文")
            self.ui.comboBox.setItemText(1, "English")
            self.ui.label.setText("共计处理:")
            self.ui.pushButton.setText("开始清理")
        self._is_english = is_english



    def _on_push_button_clicked(self) -> None:
        """pushButton（开始清理）点击事件：在子线程执行，progressBar 显示进度。"""
        file_paths = list(self._file_path_map.keys())
        if not file_paths:
            return
        # 不保留原文件（直接覆盖）时弹出红色警示确认框
        if not self.ui.checkBox_2.isChecked():
            msg = QMessageBox(self)
            msg.setWindowTitle("确认" if not self._is_english else "Confirm")
            msg.setTextFormat(Qt.TextFormat.RichText)
            msg.setText(
                '<span style="color: red;">是否确认不保留原文件直接覆盖? 确认后原文件将无法恢复!</span>'
                if not self._is_english
                else '<span style="color: red;">Confirm overwrite without keeping the original file?</span>'
            )
            msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            msg.setDefaultButton(QMessageBox.StandardButton.No)
            if msg.exec() != QMessageBox.StandardButton.Yes:
                return
        total = len(file_paths)
        self.ui.pushButton.setEnabled(False)
        self.ui.progressBar.setMaximum(total)
        self.ui.progressBar.setValue(0)
        self._clean_worker = CleanMacroWorker(
            file_paths,
            self.ui.checkBox_2.isChecked(),
            self.ui.checkBox.isChecked(),
            self._is_english,
        )
        self._clean_worker.progress.connect(self._on_clean_progress)
        self._clean_worker.finished_result.connect(self._on_clean_finished)
        self._clean_thread = QThread()
        self._clean_worker.moveToThread(self._clean_thread)
        self._clean_thread.started.connect(self._clean_worker.run)
        self._clean_worker.finished_result.connect(self._clean_thread.quit)
        self._clean_thread.start()

    def _on_clean_progress(self, current: int, total: int) -> None:
        self.ui.progressBar.setValue(current)

    def _on_clean_finished(self, success_count: int, total_vba_size: int) -> None:
        self.ui.progressBar.setValue(self.ui.progressBar.maximum())
        if self._is_english:
            self.ui.textEdit.setText(f"Successfully processed files with macro: {success_count}")
        else:
            self.ui.textEdit.setText(f"成功处理含宏文件:{success_count}个")
        self.ui.textEdit_3.setText(self._format_size(total_vba_size))
        self.ui.pushButton.setEnabled(True)

    def _format_size(self, size_bytes: int) -> str:
        """将字节数格式化为可读大小。"""
        if size_bytes >= 1024 * 1024:
            return f"{size_bytes / (1024 * 1024):.2f} MB"
        if size_bytes >= 1024:
            return f"{size_bytes / 1024:.2f} KB"
        return f"{size_bytes} B"



    def _setup_tree_context_menu(self) -> None:
        """为 treeWidget 设置右键菜单（删除），支持多选后删除。"""
        tree = self.ui.treeWidget
        tree.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        tree.customContextMenuRequested.connect(self._on_tree_context_menu)

    def _on_tree_context_menu(self, pos) -> None:
        tree = self.ui.treeWidget
        item_at = tree.itemAt(pos)
        if not item_at:
            return
        menu = QMenu(self)
        delete_text = "Delete" if self._is_english else "删除"
        delete_action = menu.addAction(_delete_icon(), delete_text)
        action = menu.exec(tree.viewport().mapToGlobal(pos))
        if action == delete_action:
            selected = tree.selectedItems()
            if selected:
                self._remove_items_from_tree_and_dict(selected)
            else:
                self._remove_item_from_tree_and_dict(item_at)
    def _is_ancestor_of(self, ancestor: QTreeWidgetItem, item: QTreeWidgetItem) -> bool:
        """判断 ancestor 是否为 item 的祖先（沿 parent 向上走）。"""
        p = item.parent()
        while p is not None:
            if p is ancestor:
                return True
            p = p.parent()
        return False

    def _remove_items_from_tree_and_dict(self, items: list[QTreeWidgetItem]) -> None:
        """批量从树和字典中移除多项（多选删除）；只移除“顶层”选中项避免重复。"""
        if not items:
            return
        # 只保留选中项中不被其它选中项包含的项，避免删父时子已无效
        top_items = [
            it for it in items
            if not any(self._is_ancestor_of(other, it) for other in items if other is not it)
        ]
        for item in top_items:
            self._remove_item_from_tree_and_dict(item, reorganize=False)
        self._reorganize_file_ids()

    def _remove_item_from_tree_and_dict(
        self, item: QTreeWidgetItem, reorganize: bool = True
    ) -> None:
        """从树和字典中移除该项；若为文件夹则一并移除其下所有文件的记录；删除后可选重组 ID。"""
        tree = self.ui.treeWidget
        # 收集该项及其所有子项中的文件完整路径，从字典中移除
        def collect_full_paths(it: QTreeWidgetItem) -> list:
            paths = []
            if it.childCount() == 0 and it.text(0).strip():
                fp = it.data(1, Qt.ItemDataRole.UserRole)
                if fp is not None:
                    paths.append(fp)
            for i in range(it.childCount()):
                paths.extend(collect_full_paths(it.child(i)))
            return paths
        for full_path in collect_full_paths(item):
            self._file_path_map.pop(full_path, None)
        # 从树中移除
        parent = item.parent()
        if parent is None:
            idx = tree.indexOfTopLevelItem(item)
            if idx >= 0:
                tree.takeTopLevelItem(idx)
        else:
            parent.removeChild(item)
        if reorganize:
            self._reorganize_file_ids()
    def _reorganize_file_ids(self) -> None:
        """按树中顺序重新为所有文件节点分配连续 ID（1, 2, 3, ...），并更新 _next_file_id。"""
        tree = self.ui.treeWidget
        next_id = 1
        for i in range(tree.topLevelItemCount()):
            top = tree.topLevelItem(i)
            if top.text(0).strip():  # 顶层文件项
                top.setText(0, str(next_id))
                next_id += 1
            for j in range(top.childCount()):
                child = top.child(j)
                child.setText(0, str(next_id))
                next_id += 1
        self._next_file_id = next_id
    def _apply_folder_scan_result(self, folder_name: str, root: str, to_add: list[tuple[str, str]]) -> None:
        """将扫描结果应用到树（过滤已存在项）。在主线程调用。"""
        to_add_filtered = [(f, r) for f, r in to_add if f not in self._file_path_map]
        if not to_add_filtered:
            return
        folder_item = QTreeWidgetItem(self.ui.treeWidget)
        folder_item.setText(1, folder_name)
        folder_item.setText(0, "")
        folder_item.setIcon(0, QIcon(FOLDER_ICON_PATH))
        for full_path, rel_path in sorted(to_add_filtered, key=lambda x: x[1].lower()):
            self._file_path_map[full_path] = (os.path.dirname(full_path), os.path.basename(full_path))
            child = QTreeWidgetItem(folder_item)
            child.setText(0, str(self._next_file_id))
            child.setText(1, rel_path)
            child.setData(1, Qt.ItemDataRole.UserRole, full_path)
            self._next_file_id += 1

    def _on_scan_folder_progress(self, files_count: int, current_dir: str) -> None:
        """扫描文件夹进度：更新进度条与状态文字。"""
        self.ui.progressBar.setMaximum(0)  # 不定进度
        if self._is_english:
            self.ui.textEdit.setPlainText(f"Scanning... {files_count} file(s) found")
        else:
            self.ui.textEdit.setPlainText(f"扫描中... 已找到 {files_count} 个文件")

    def _on_scan_folder_finished(self, folder_name: str, root: str, to_add: list[tuple[str, str]]) -> None:
        """扫描完成：写入树、恢复进度条与按钮。"""
        self._apply_folder_scan_result(folder_name, root, to_add)
        self.ui.progressBar.setMaximum(100)
        self.ui.progressBar.setValue(0)
        self.ui.pushButton_2.setEnabled(True)
        self.ui.textEdit.clear()
    def _add_file_to_tree(self, file_path: str) -> None:
        full_path = os.path.normpath(os.path.abspath(file_path))
        if full_path in self._file_path_map:
            return
        name = os.path.basename(full_path)
        dir_path = os.path.dirname(full_path)
        self._file_path_map[full_path] = (dir_path, name)
        file_item = QTreeWidgetItem(self.ui.treeWidget)
        file_item.setText(0, str(self._next_file_id))
        file_item.setText(1, name)
        file_item.setData(1, Qt.ItemDataRole.UserRole, full_path)
        self._next_file_id += 1
    def _on_select_path_clicked(self):
        """弹出选择：文件夹 或 Office 文件，加入 treeWidget 并写入 textEdit_2。"""
        en = self._is_english
        msg = QMessageBox(self)
        msg.setWindowTitle("Select type" if en else "选择类型")
        msg.setText("Please choose what to add:" if en else "请选择要添加的内容：")
        btn_folder = msg.addButton("Select folder" if en else "选择文件夹", QMessageBox.ButtonRole.ActionRole)
        btn_file = msg.addButton("Select Office file" if en else "选择 Office 文件", QMessageBox.ButtonRole.ActionRole)
        msg.addButton("Cancel" if en else "取消", QMessageBox.ButtonRole.RejectRole)
        msg.exec()

        if msg.clickedButton() == btn_folder:
            path = QFileDialog.getExistingDirectory(
                self,
                "Select folder" if en else "选择文件夹",
                "",
                QFileDialog.Option.ShowDirsOnly | QFileDialog.Option.DontResolveSymlinks,
            )
            if path:
                self.ui.textEdit_2.setPlainText(path)
                self.ui.pushButton_2.setEnabled(False)
                self._scan_worker = ScanFolderWorker(path)
                self._scan_worker.progress.connect(self._on_scan_folder_progress)
                self._scan_worker.finished_result.connect(self._on_scan_folder_finished)
                self._scan_worker.start()
        elif msg.clickedButton() == btn_file:
            path, _ = QFileDialog.getOpenFileName(
                self,
                "Select Office file" if en else "选择 Office 文件",
                "",
                "Office files (*.docm *.dotm *.xlsm *.xltm *.pptm *.potm *.doc *.xls *.ppt *.xlsx);;All files (*)" if en
                else "Office 文件 (*.docm *.dotm *.xlsm *.xltm *.pptm *.potm *.doc *.xls *.ppt *.xlsx);;所有文件 (*)",
            )
            if path:
                self.ui.textEdit_2.setPlainText(path)
                self._add_file_to_tree(path)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    widget = Widget()
    widget.show()
    sys.exit(app.exec())
