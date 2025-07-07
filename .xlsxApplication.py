import sys
from typing import List

import pandas as pd
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
    QMainWindow,
)


class ExcelProcessorApp(QMainWindow):
    """Minimal invasive rewrite of the original widget‑based Excel helper.

    ‑ Keeps the overall flow (file‑select page → table page → save) intact
    ‑ Removes style‑sheet properties unsupported by Qt (e.g., box‑shadow)
    ‑ Fixes crash reason: duplicate *cellChanged* connections & recursive signals
    ‑ Implements K‑column multiplication + L‑column stock check with “#SİPARİŞ VER”.
    """

    # --- Constants ------------------------------------------------------- #
    SHEET1_COLS = {"A": 0, "C": 2, "G": 6, "E": 4}
    SHEET2_COLS = {"B": 1, "J": 9}
    SHEET3_COLS = {"B": 1, "J": 9, "K": 10}
    COMMON_MATCH_COL = {"G": 6}

    HEADER_LABELS = [
        "A", "B", "C", "D",  # Sheet‑1
        "E", "F",              # Sheet‑2
        "H", "I", "J",         # Sheet‑3
        "K", "L",              # New columns
    ]

    # --- Init / UI ------------------------------------------------------- #
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Veri Yönetimi Uygulaması")
        self.setGeometry(100, 100, 1200, 800)
        self.setWindowIcon(QIcon("icon.png"))  # put your icon next to script

        self._updating = False  # recursion‑guard for cellChanged
        self._cell_connected = False  # track signal connection

        self.excel_data = {}
        self.selected_file_path = ""
        self.sheet_names: List[str] = []

        self._build_style()
        self._build_pages()

    # -------------------------------------------------------------------- #
    #                           UI Construction
    # -------------------------------------------------------------------- #
    def _build_style(self):
        self.setStyleSheet(
            """
            QMainWindow { background: #f0f2f5; }
            QWidget     { font-family: 'Segoe UI', sans-serif; font-size: 14px; }
            QLabel#titleLabel { font-size: 28px; font-weight: bold; color: #2c3e50; margin-bottom: 20px; }
            QPushButton { background: #3498db; color: white; border-radius: 8px; padding: 12px 25px; font-size: 15px; font-weight: bold; border: none; }
            QPushButton:hover    { background: #2980b9; }
            QPushButton:disabled { background: #cccccc; color: #666666; }
            QFrame#card { background: white; border-radius: 10px; padding: 30px; }
            QTableWidget         { background: white; border: 1px solid #dcdcdc; gridline-color: #f0f0f0; selection-background-color: #aed6f1; font-size: 13px; }
            QHeaderView::section { background: #e9ecef; color: #495057; padding: 8px; border: 1px solid #dcdcdc; font-weight: bold; }
            QLabel#filePathLabel { font-style: italic; color: #555; font-size: 13px; margin-top: 10px; }
            """
        )

    def _build_pages(self):
        self.stacked_widget = QStackedWidget(self)
        self.setCentralWidget(self.stacked_widget)

        # 1) File‑select page ------------------------------------------------
        self.file_page = QWidget()
        main_v = QVBoxLayout(self.file_page)

        card = QFrame(objectName="card")
        card.setFixedSize(500, 350)
        card_v = QVBoxLayout(card)

        ttl = QLabel("Excel Dosyası Seçin", objectName="titleLabel", alignment=Qt.AlignCenter)
        card_v.addWidget(ttl)
        card_v.addSpacing(30)

        self.btn_select = QPushButton("Dosya Seç", clicked=self._select_file)
        card_v.addWidget(self.btn_select)

        self.lbl_file = QLabel("Seçilen Dosya: Yüklenmedi", objectName="filePathLabel", alignment=Qt.AlignCenter)
        card_v.addWidget(self.lbl_file)
        card_v.addSpacing(20)

        self.btn_open = QPushButton("Uyarlanmış Dosyayı Görüntüle", enabled=False, clicked=self._open_table_page)
        card_v.addWidget(self.btn_open)

        main_v.addStretch(1)
        main_v.addWidget(card, alignment=Qt.AlignCenter)
        main_v.addStretch(1)

        self.stacked_widget.addWidget(self.file_page)

        # 2) Table page -----------------------------------------------------
        self.table_page = QWidget()
        tv = QVBoxLayout(self.table_page)

        lbl2 = QLabel("Uyarlanmış Excel Verileri", objectName="titleLabel", alignment=Qt.AlignCenter)
        tv.addWidget(lbl2)
        tv.addSpacing(15)

        self.table = QTableWidget(editTriggers=QTableWidget.DoubleClicked | QTableWidget.AnyKeyPressed, alternatingRowColors=True)
        tv.addWidget(self.table)

        hbox = QHBoxLayout()
        self.btn_save = QPushButton("Değişiklikleri Kaydet", clicked=self._save_excel)
        self.btn_back = QPushButton("Geri Dön", clicked=lambda: self.stacked_widget.setCurrentWidget(self.file_page))
        hbox.addStretch(1)
        hbox.addWidget(self.btn_save)
        hbox.addWidget(self.btn_back)
        hbox.addStretch(1)
        tv.addLayout(hbox)
        tv.addSpacing(20)

        self.stacked_widget.addWidget(self.table_page)

    # -------------------------------------------------------------------- #
    #                           File selection
    # -------------------------------------------------------------------- #
    def _select_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Excel Dosyası Seç", "", "Excel Dosyaları (*.xlsx)")
        if not path:
            return
        self.selected_file_path = path
        self.lbl_file.setText(f"Seçilen Dosya: {path.split('/')[-1]}")
        self._load_excel()

    def _load_excel(self):
        try:
            xls = pd.ExcelFile(self.selected_file_path)
            self.sheet_names = xls.sheet_names
            if len(self.sheet_names) < 3:
                raise ValueError("Seçilen Excel dosyasında en az 3 sayfa bulunmalıdır.")
            self.excel_data = {
                "s1": pd.read_excel(xls, sheet_name=self.sheet_names[0], header=None),
                "s2": pd.read_excel(xls, sheet_name=self.sheet_names[1], header=None),
                "s3": pd.read_excel(xls, sheet_name=self.sheet_names[2], header=None),
            }
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Excel dosyası yüklenirken bir hata oluştu:\n{e}")
            self.btn_open.setEnabled(False)
            return

        QMessageBox.information(self, "Başarılı", "Excel dosyası başarıyla yüklendi.")
        self.btn_open.setEnabled(True)

    # -------------------------------------------------------------------- #
    #                           Table population
    # -------------------------------------------------------------------- #
    def _open_table_page(self):
        if not self.excel_data:
            return
        self._populate_table()
        self.stacked_widget.setCurrentWidget(self.table_page)

    def _populate_table(self):
        df1 = self.excel_data["s1"]
        df2 = self.excel_data["s2"]
        df3 = self.excel_data["s3"]

        # 1) Filter + deduplicate rows in Sheet‑1 --------------------------
        df1 = df1[
            df1[self.SHEET1_COLS["C"]].notna() & df1[self.SHEET1_COLS["A"]].notna()
        ].copy()
        df1.drop_duplicates(subset=self.SHEET1_COLS["C"], keep="first", inplace=True)

        # 2) Prepare table dimensions -------------------------------------
        self.table.setColumnCount(len(self.HEADER_LABELS))
        self.table.setHorizontalHeaderLabels(self.HEADER_LABELS)
        self.table.setRowCount(len(df1))

        # 3) Fill rows -----------------------------------------------------
        for r, row in enumerate(df1.itertuples(index=False)):
            # Sheet‑1 cols
            self.table.setItem(r, 0, QTableWidgetItem(str(row[self.SHEET1_COLS["A"]])))
            self.table.setItem(r, 1, QTableWidgetItem(str(row[self.SHEET1_COLS["C"]])))
            self.table.setItem(r, 2, QTableWidgetItem(str(row[self.SHEET1_COLS["G"]])))
            self.table.setItem(r, 3, QTableWidgetItem(str(row[self.SHEET1_COLS["E"]])))

            match_val = row[self.SHEET1_COLS["C"]]

            # Sheet‑2 match
            s2_match = df2[df2[self.COMMON_MATCH_COL["G"]] == match_val]
            if not s2_match.empty:
                m = s2_match.iloc[0]
                self.table.setItem(r, 4, QTableWidgetItem(str(m[self.SHEET2_COLS["B"]])))
                self.table.setItem(r, 5, QTableWidgetItem(str(m[self.SHEET2_COLS["J"]])))
            else:
                self.table.setItem(r, 4, QTableWidgetItem(""))
                self.table.setItem(r, 5, QTableWidgetItem(""))

            # Sheet‑3 match
            s3_match = df3[df3[self.COMMON_MATCH_COL["G"]] == match_val]
            if not s3_match.empty:
                m3 = s3_match.iloc[0]
                self.table.setItem(r, 6, QTableWidgetItem(str(m3[self.SHEET3_COLS["B"]])))
                self.table.setItem(r, 7, QTableWidgetItem(str(m3[self.SHEET3_COLS["J"]])))
                self.table.setItem(r, 8, QTableWidgetItem(str(m3[self.SHEET3_COLS["K"]])))
            else:
                self.table.setItem(r, 6, QTableWidgetItem(""))
                self.table.setItem(r, 7, QTableWidgetItem(""))
                self.table.setItem(r, 8, QTableWidgetItem(""))

            # K (need) initially empty
            self.table.setItem(r, 9, QTableWidgetItem(""))

            # L initial calc (K assumed 0)
            self._update_l_column(r)

        # 4) Resize + connect once ----------------------------------------
        self.table.resizeColumnsToContents()
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.table.verticalHeader().setDefaultSectionSize(30)

        if self._cell_connected:
            try:
                self.table.cellChanged.disconnect(self._cell_changed)
            except TypeError:
                pass
        self.table.cellChanged.connect(self._cell_changed)
        self._cell_connected = True

    # -------------------------------------------------------------------- #
    #                        Cell change handlers
    # -------------------------------------------------------------------- #
    def _cell_changed(self, row: int, col: int):
        if self._updating or col != 9:
            return

        try:
            k_raw = self.table.item(row, col).text().replace(",", ".")
            k_coef = float(k_raw)
        except (ValueError, AttributeError):
            return  # invalid input

        try:
            d_val = float(self.table.item(row, 3).text().replace(",", "."))
        except (ValueError, AttributeError):
            d_val = 0.0

        need = d_val * k_coef

        # propagate to selection if multiple K cells are selected ----------
        sel_ranges = self.table.selectedRanges()
        k_rows = {row}
        for rng in sel_ranges:
            if rng.leftColumn() <= 9 <= rng.rightColumn():
                k_rows.update(range(rng.topRow(), rng.bottomRow() + 1))

        self._updating = True
        for r in k_rows:
            # recalculate per row (different D)
            try:
                d_cell = float(self.table.item(r, 3).text().replace(",", "."))
            except (ValueError, AttributeError):
                d_cell = 0.0
            self.table.setItem(r, 9, QTableWidgetItem(str(d_cell * k_coef)))
            self._update_l_column(r)
        self._updating = False

    def _to_float(self, item: QTableWidgetItem) -> float:
        try:
            return float(item.text().replace(",", "."))
        except (ValueError, AttributeError):
            return 0.0

    def _update_l_column(self, row: int):
        f_val = self._to_float(self.table.item(row, 5))
        i_val = self._to_float(self.table.item(row, 7))
        j_val = self._to_float(self.table.item(row, 8))
        k_val = self._to_float(self.table.item(row, 9))

        result = f_val + i_val + j_val - k_val
        text = f"{result} #SİPARİŞ VER" if result < 0 else str(result)

        item = self.table.item(row, 10)
        if item is None:
            item = QTableWidgetItem()
            item.setFlags(item.flags() ^ Qt.ItemIsEditable)
            self.table.setItem(row, 10, item)
        item.setText(text)

    # -------------------------------------------------------------------- #
    #                           Save to Excel
    # -------------------------------------------------------------------- #
    def _save_excel(self):
        path, _ = QFileDialog.getSaveFileName(self, "Uyarlanmış Excel Dosyasını Kaydet", "uyarlanmis_veri.xlsx", "Excel Dosyaları (*.xlsx)")
        if not path:
            return

        rows, cols = self.table.rowCount(), self.table.columnCount()
        data = [[self.table.item(r, c).text() if self.table.item(r, c) else "" for c in range(cols)] for r in range(rows)]

        def col_name(n):
            name = ""
            while n >= 0:
                name = chr(n % 26 + ord("A")) + name
                n = n // 26 - 1
            return name

        headers = [col_name(i) for i in range(cols)]
        pd.DataFrame(data, columns=headers).to_excel(path, index=False)
        QMessageBox.information(self, "Başarılı", f"Dosya kaydedildi: {path.split('/')[-1]}")


# ----------------------------------------------------------------------- #
#                                   main
# ----------------------------------------------------------------------- #
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelProcessorApp()
    window.show()
    sys.exit(app.exec_())
