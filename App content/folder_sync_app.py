import sys
import os
import shutil
from pathlib import Path
from datetime import datetime

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QTreeWidget, QTreeWidgetItem,
    QLabel, QMessageBox, QHeaderView, QTreeWidgetItemIterator,
    QLineEdit
)
from PyQt6.QtGui import QFont, QColor, QBrush
from PyQt6.QtCore import Qt, QTimer

# Define colors for file status
COLOR_MISSING = QColor("#dbeafe")
COLOR_OLDER = QColor("#fee2e2")
COLOR_NORMAL = QColor("white") # For the target tree

# Define a set of folders to exclude from scanning
EXCLUDED_DIRS = {"python-embed", ".git", "__pycache__", ".vscode"}


class FileSyncApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.source_path = ""
        self.target_path = ""
        
        self.diff_files = []

        self.setWindowTitle("Folder Sync and Compare Tool")
        self.setGeometry(100, 100, 1200, 800)

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)

        controls_layout = QHBoxLayout()
        self.btn_analyze = QPushButton("▶️ Analyze and Show Differences")
        self.btn_analyze.setFont(QFont("Arial", 10, QFont.Weight.Bold))
        self.btn_analyze.clicked.connect(self.run_full_analysis)
        self.btn_analyze.setEnabled(False)

        self.btn_copy = QPushButton("⬇️ Copy Checked to Target")
        self.btn_copy.setFont(QFont("Arial", 10, QFont.Weight.Bold))
        self.btn_copy.clicked.connect(self.copy_checked_files)
        self.btn_copy.setEnabled(False)

        controls_layout.addWidget(self.btn_analyze)
        controls_layout.addWidget(self.btn_copy)
        controls_layout.addStretch()

        trees_layout = QHBoxLayout()
        source_box = self.create_source_layout()
        target_box = self.create_target_layout()
        trees_layout.addLayout(source_box)
        trees_layout.addLayout(target_box)

        main_layout.addLayout(controls_layout)
        main_layout.addLayout(trees_layout)

    def create_source_layout(self) -> QVBoxLayout:
        source_box = QVBoxLayout()
        self.source_label = QLabel("Source Folder: Not Selected")
        self.btn_select_source = QPushButton("Select Source Folder")
        self.btn_select_source.clicked.connect(self.select_source_folder)
        
        filter_layout = QHBoxLayout()
        filter_label = QLabel("Filter extensions (e.g., .txt):")
        self.filter_input = QLineEdit()
        self.filter_input.setPlaceholderText("Apply after analysis")
        self.filter_timer = QTimer(self)
        self.filter_timer.setSingleShot(True)
        self.filter_timer.setInterval(500)
        self.filter_timer.timeout.connect(self.apply_filter)
        self.filter_input.textChanged.connect(self.filter_timer.start)
        filter_layout.addWidget(filter_label)
        filter_layout.addWidget(self.filter_input)

        self.source_tree = self.create_tree_widget()
        self.source_tree.itemChanged.connect(self.handle_item_changed)
        
        source_box.addWidget(self.source_label)
        source_box.addWidget(self.btn_select_source)
        source_box.addLayout(filter_layout)
        source_box.addWidget(self.source_tree)
        return source_box

    def create_target_layout(self) -> QVBoxLayout:
        target_box = QVBoxLayout()
        self.target_label = QLabel("Target Folder: Not Selected")
        self.btn_select_target = QPushButton("Select Target Folder")
        self.btn_select_target.clicked.connect(self.select_target_folder)
        self.target_tree = self.create_tree_widget(add_checkboxes=False)
        
        target_box.addWidget(self.target_label)
        target_box.addWidget(self.btn_select_target)
        target_box.addSpacing(self.filter_input.sizeHint().height() + 15)
        target_box.addWidget(self.target_tree)
        return target_box

    def create_tree_widget(self, add_checkboxes: bool = True) -> QTreeWidget:
        tree = QTreeWidget()
        tree.setHeaderLabels(["Name", "Size (KB)", "Date Modified"])
        tree.header().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        tree.setSortingEnabled(True)
        return tree
    
    def select_source_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Source Folder")
        if folder:
            self.source_path = Path(folder)
            self.source_label.setText(f"Source Folder: {self.source_path}")
            self.check_if_ready_to_analyze()

    def select_target_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Target Folder")
        if folder:
            self.target_path = Path(folder)
            self.target_label.setText(f"Target Folder: {self.target_path}")
            self.check_if_ready_to_analyze()
    
    def check_if_ready_to_analyze(self):
        ready = bool(self.source_path and self.target_path)
        self.btn_analyze.setEnabled(ready)
        if not ready:
            self.btn_copy.setEnabled(False)

    # ### CORRECTION IS HERE ###
    def run_full_analysis(self):
        # Directly check if paths are set, instead of calling the other function
        if not (self.source_path and self.target_path):
            QMessageBox.warning(self, "Warning", "Please select both source and target folders first.")
            return
        
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        
        source_files = self._scan_directory(self.source_path)
        target_files = self._scan_directory(self.target_path)
        
        self.diff_files.clear()
        for rel_path, src_stat in source_files.items():
            if rel_path not in target_files:
                self.diff_files.append({'path': src_stat['path'], 'status': 'missing'})
            else:
                tgt_stat = target_files[rel_path]
                if src_stat['mtime'] > tgt_stat['mtime']:
                    self.diff_files.append({'path': src_stat['path'], 'status': 'older'})
        
        self.populate_source_tree_with_diffs()
        self.populate_target_tree()
        self.apply_filter()

        QApplication.restoreOverrideCursor()

        self.btn_copy.setEnabled(bool(self.diff_files))
        
        if self.diff_files:
            QMessageBox.information(self, "Analysis Complete", f"Found {len(self.diff_files)} differences.")
        else:
            QMessageBox.information(self, "Analysis Complete", "Folders are in sync. No differences found.")
    # ### END OF CORRECTION ###

    def _scan_directory(self, root_path: Path) -> dict:
        file_map = {}
        for dirpath, dirnames, filenames in os.walk(root_path):
            dirnames[:] = [d for d in dirnames if d not in EXCLUDED_DIRS]
            for filename in filenames:
                full_path = Path(dirpath) / filename
                try:
                    stat = full_path.stat()
                    rel_path = full_path.relative_to(root_path)
                    file_map[rel_path] = {'path': full_path, 'mtime': stat.st_mtime}
                except (FileNotFoundError, PermissionError):
                    continue
        return file_map

    def populate_source_tree_with_diffs(self):
        self.source_tree.clear()
        self.source_tree.setSortingEnabled(False)
        nodes = {self.source_path: self.source_tree.invisibleRootItem()}

        for file_info in sorted(self.diff_files, key=lambda x: x['path']):
            full_path = file_info['path']
            parent_path = full_path.parent
            
            if parent_path != self.source_path:
                relative_parts = parent_path.relative_to(self.source_path).parts
                path_tracker = self.source_path
                for part in relative_parts:
                    path_tracker = path_tracker / part
                    if path_tracker not in nodes:
                        dir_item = QTreeWidgetItem(nodes[path_tracker.parent], [part])
                        dir_item.setData(0, Qt.ItemDataRole.UserRole, path_tracker)
                        dir_item.setFlags(dir_item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                        dir_item.setCheckState(0, Qt.CheckState.Unchecked)
                        nodes[path_tracker] = dir_item
            
            parent_item = nodes[parent_path]
            try:
                stat = full_path.stat()
                size_kb = f"{stat.st_size / 1024:.2f}"
                mod_time = datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')

                file_item = QTreeWidgetItem(parent_item, [full_path.name, size_kb, mod_time])
                file_item.setData(0, Qt.ItemDataRole.UserRole, full_path)
                file_item.setFlags(file_item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                file_item.setCheckState(0, Qt.CheckState.Unchecked)

                color = COLOR_MISSING if file_info['status'] == 'missing' else COLOR_OLDER
                for i in range(file_item.columnCount()):
                    file_item.setBackground(i, QBrush(color))
            except (FileNotFoundError, PermissionError) as e:
                print(f"Could not access {full_path}: {e}")

        self.source_tree.setSortingEnabled(True)
        self.source_tree.expandToDepth(0)

    def populate_target_tree(self):
        tree_widget = self.target_tree
        root_path = self.target_path
        
        tree_widget.clear()
        tree_widget.setSortingEnabled(False)
        nodes = {root_path: tree_widget.invisibleRootItem()}

        for dirpath, dirnames, filenames in os.walk(root_path, topdown=True):
            dirnames[:] = [d for d in dirnames if d not in EXCLUDED_DIRS]
            current_dir_path = Path(dirpath)
            parent_item = nodes.get(current_dir_path.parent, tree_widget.invisibleRootItem())
            
            if current_dir_path != root_path:
                dir_item = QTreeWidgetItem(parent_item, [current_dir_path.name])
                nodes[current_dir_path] = dir_item

            for filename in filenames:
                file_path = current_dir_path / filename
                try:
                    stat = file_path.stat()
                    size_kb = f"{stat.st_size / 1024:.2f}"
                    mod_time = datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                    QTreeWidgetItem(nodes[current_dir_path], [filename, size_kb, mod_time])
                except (FileNotFoundError, PermissionError):
                    continue
        
        tree_widget.setSortingEnabled(True)
        tree_widget.expandToDepth(0)

    def handle_item_changed(self, item: QTreeWidgetItem, column: int):
        if column == 0:
            self.source_tree.blockSignals(True)
            state = item.checkState(0)
            if item.childCount() > 0:
                self.set_child_check_state(item, state)
            self.source_tree.blockSignals(False)

    def set_child_check_state(self, parent_item: QTreeWidgetItem, state: Qt.CheckState):
        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            child.setCheckState(0, state)
            if child.childCount() > 0:
                self.set_child_check_state(child, state)

    def apply_filter(self):
        filter_text = self.filter_input.text().strip().lower()
        extensions = {f".{ext.strip().lstrip('.')}" for ext in filter_text.split(',') if ext.strip()} if filter_text else None
        
        iterator = QTreeWidgetItemIterator(self.source_tree)
        while iterator.value():
            iterator.value().setHidden(False)
            iterator += 1

        if extensions:
            iterator = QTreeWidgetItemIterator(self.source_tree)
            while iterator.value():
                item = iterator.value()
                item_path = item.data(0, Qt.ItemDataRole.UserRole)
                if item_path and item_path.is_file() and item_path.suffix.lower() not in extensions:
                    item.setHidden(True)
                iterator += 1
        
        self.hide_empty_folders(self.source_tree.invisibleRootItem())

    def hide_empty_folders(self, parent_item: QTreeWidgetItem):
        is_empty = True
        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            if child.childCount() > 0:
                if not self.hide_empty_folders(child):
                    is_empty = False
            elif not child.isHidden():
                is_empty = False
        
        if parent_item != self.source_tree.invisibleRootItem():
             parent_item.setHidden(is_empty)
        return is_empty

    def copy_checked_files(self):
        items_to_copy = []
        iterator = QTreeWidgetItemIterator(self.source_tree)
        while iterator.value():
            item = iterator.value()
            if not item.isHidden() and item.checkState(0) == Qt.CheckState.Checked:
                source_file_path = item.data(0, Qt.ItemDataRole.UserRole)
                if source_file_path and source_file_path.is_file():
                    items_to_copy.append(item)
            iterator += 1
        
        if not items_to_copy:
            QMessageBox.warning(self, "No Files Checked", "Please check one or more files from the source list to copy.")
            return

        copied_count = 0
        errors = []
        for item in items_to_copy:
            source_file_path = item.data(0, Qt.ItemDataRole.UserRole)
            relative_path = source_file_path.relative_to(self.source_path)
            target_file_path = self.target_path / relative_path
            
            try:
                target_file_path.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(source_file_path, target_file_path)
                copied_count += 1
            except Exception as e:
                errors.append(f"Failed to copy {source_file_path}: {e}")
        
        message = f"Successfully copied {copied_count} file(s)."
        if errors:
            message += f"\n\nEncountered {len(errors)} error(s):\n" + "\n".join(errors)
            QMessageBox.critical(self, "Copy Operation Finished with Errors", message)
        else:
            QMessageBox.information(self, "Copy Operation Successful", message)
        
        self.run_full_analysis()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = FileSyncApp()
    window.show()
    sys.exit(app.exec())