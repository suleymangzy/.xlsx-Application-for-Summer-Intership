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
    # Column mappings for original sheets (0-indexed)
    SHEET1_COLS = {"A": 0, "C": 2, "G": 6, "E": 4}
    SHEET2_COLS = {"B": 1, "J": 9}  # J is the column to be summed
    SHEET3_COLS = {"B": 1, "J": 9, "K": 10}  # J and K are the columns to be summed
    COMMON_MATCH_COL = {"G": 6}  # Column G (index 6) is used for matching across sheets
    # New: Column mappings for the 4th sheet
    # Assuming 2nd index (column C) for matching material, 8th index (column I) for ordered quantity
    SHEET4_COLS = {"C": 2, "I": 8}

    # Header labels for the displayed QTableWidget
    HEADER_LABELS = [
        "Ü.Ağacı Sev", "Malzeme", "Açıklama", "Miktar",  # Corresponds to Sheet-1 columns A, C, G, E
        "Depo 100", "Kullanılabilir Stok",  # Corresponds to Sheet-2 columns B, J (summed)
        "Depo 110", "Kullanılabilir Stok", "Kalite Stoğu",  # Corresponds to Sheet-3 columns B, J (summed), K (summed)
        "İhtiyaç", "Durum",  # New columns K and L
        "Verilen Sipariş Miktarı", "Verilmesi Gereken Sipariş Miktarı"  # New columns for order quantities
    ]

    # --- Init / UI ------------------------------------------------------- #
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Veri Yönetimi Uygulaması")
        self.setGeometry(100, 100, 1200, 800)
        # Set window icon (ensure 'icon.png' is in the same directory as the script)
        self.setWindowIcon(QIcon("icon.png"))

        self._updating = False  # Guard to prevent recursive calls during cell updates
        self._cell_connected = False  # Flag to track if cellChanged signal is connected

        self.excel_data = {}  # Stores pandas DataFrames for each sheet
        self.selected_file_path = ""  # Path of the currently selected Excel file
        self.sheet_names: List[str] = []  # Names of sheets in the loaded Excel file

        self._build_style()  # Apply custom CSS styling
        self._build_pages()  # Construct the UI pages

    # -------------------------------------------------------------------- #
    #                           UI Construction
    # -------------------------------------------------------------------- #
    def _build_style(self):
        """Applies custom CSS styling to the application widgets."""
        self.setStyleSheet(
            """
            QMainWindow { background: #f0f2f5; } /* Light grey background for main window */
            QWidget     { font-family: 'Segoe UI', sans-serif; font-size: 14px; } /* Default font for widgets */
            QLabel#titleLabel { font-size: 28px; font-weight: bold; color: #2c3e50; margin-bottom: 20px; } /* Styling for title labels */
            QPushButton { /* Styling for buttons */
                background: #3498db;
                color: white;
                border-radius: 8px;
                padding: 12px 25px;
                font-size: 15px;
                font-weight: bold;
                border: none;
            }
            QPushButton:hover    { background: #2980b9; } /* Darker blue on hover */
            QPushButton:disabled { background: #cccccc; color: #666666; } /* Greyed out for disabled buttons */
            QFrame#card { background: white; border-radius: 10px; padding: 30px; } /* Styling for card-like frames */
            QTableWidget         { /* Styling for the table widget */
                background: white;
                border: 1px solid #dcdcdc;
                gridline-color: #f0f0f0;
                selection-background-color: #aed6f1;
                font-size: 13px;
            }
            QHeaderView::section { /* Styling for table headers */
                background: #e9ecef;
                color: #495057;
                padding: 8px;
                border: 1px solid #dcdcdc;
                font-weight: bold;
            }
            QLabel#filePathLabel { font-style: italic; color: #555; font-size: 13px; margin-top: 10px; } /* Styling for file path label */
            """
        )

    def _build_pages(self):
        """Constructs the two main pages of the application: file selection and table view."""
        self.stacked_widget = QStackedWidget(self)
        self.setCentralWidget(self.stacked_widget)

        # 1) File‑select page ------------------------------------------------
        self.file_page = QWidget()
        main_v = QVBoxLayout(self.file_page)  # Main layout for the file selection page

        card = QFrame(objectName="card")  # Card frame for buttons and labels
        card.setFixedSize(500, 350)  # Fixed size for the card
        card_v = QVBoxLayout(card)  # Layout for the card content

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

        main_v.addStretch(1)  # Push content to center
        main_v.addWidget(card, alignment=Qt.AlignCenter)
        main_v.addStretch(1)

        self.stacked_widget.addWidget(self.file_page)  # Add file page to stacked widget

        # 2) Table page -----------------------------------------------------
        self.table_page = QWidget()
        tv = QVBoxLayout(self.table_page)  # Main layout for the table page

        lbl2 = QLabel("Uyarlanmış Excel Verileri", objectName="titleLabel", alignment=Qt.AlignCenter)
        tv.addWidget(lbl2)
        tv.addSpacing(15)

        self.table = QTableWidget(
            editTriggers=QTableWidget.DoubleClicked | QTableWidget.AnyKeyPressed,
            alternatingRowColors=True  # Zebra striping for rows
        )
        tv.addWidget(self.table)

        hbox = QHBoxLayout()  # Layout for save and back buttons
        self.btn_save = QPushButton("Değişiklikleri Kaydet", clicked=self._save_excel)
        self.btn_back = QPushButton("Geri Dön", clicked=lambda: self.stacked_widget.setCurrentWidget(self.file_page))
        hbox.addStretch(1)
        hbox.addWidget(self.btn_save)
        hbox.addWidget(self.btn_back)
        hbox.addStretch(1)
        tv.addLayout(hbox)
        tv.addSpacing(20)

        self.stacked_widget.addWidget(self.table_page)  # Add table page to stacked widget

    # -------------------------------------------------------------------- #
    #                           File selection
    # -------------------------------------------------------------------- #
    def _select_file(self):
        """Opens a file dialog for the user to select an Excel file."""
        # GetOpenFileName returns (filePath, filter), we only need filePath
        path, _ = QFileDialog.getOpenFileName(self, "Excel Dosyası Seç", "", "Excel Dosyaları (*.xlsx)")
        if not path:  # If no file is selected, return
            return
        self.selected_file_path = path
        self.lbl_file.setText(f"Seçilen Dosya: {path.split('/')[-1]}")  # Display selected file name
        self._load_excel()  # Attempt to load the selected Excel file

    def _load_excel(self):
        """Loads data from the selected Excel file into pandas DataFrames."""
        try:
            xls = pd.ExcelFile(self.selected_file_path)  # Create an ExcelFile object
            self.sheet_names = xls.sheet_names  # Get all sheet names
            # New: Check if at least 4 sheets are present
            if len(self.sheet_names) < 4:
                raise ValueError("Seçilen Excel dosyasında en az 4 sayfa bulunmalıdır.")
            # Load the first four sheets into DataFrames, skipping the first row (index 0)
            # This ensures that the original Excel file's first row is not processed.
            self.excel_data = {
                "s1": pd.read_excel(xls, sheet_name=self.sheet_names[0], header=None, skiprows=[0]),
                "s2": pd.read_excel(xls, sheet_name=self.sheet_names[1], header=None, skiprows=[0]),
                "s3": pd.read_excel(xls, sheet_name=self.sheet_names[2], header=None, skiprows=[0]),
                "s4": pd.read_excel(xls, sheet_name=self.sheet_names[3], header=None, skiprows=[0]),
                # New: Load 4th sheet
            }
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Excel dosyası yüklenirken bir hata oluştu:\n{e}")
            self.btn_open.setEnabled(False)  # Disable open button on error
            return

        QMessageBox.information(self, "Başarılı", "Excel dosyası başarıyla yüklendi.")
        self.btn_open.setEnabled(True)  # Enable open button on successful load

    # -------------------------------------------------------------------- #
    #                           Table population
    # -------------------------------------------------------------------- #
    def _open_table_page(self):
        """Switches to the table view page and populates the table."""
        if not self.excel_data:  # Ensure data is loaded
            return
        self._populate_table()  # Fill the QTableWidget with processed data
        self._process_fsnkp_rows()  # Process FSNKP rows after initial population
        self.stacked_widget.setCurrentWidget(self.table_page)  # Switch to table page

    def _populate_table(self):
        """Populates the QTableWidget with data from the loaded Excel sheets,
        performing filtering, deduplication, and aggregation."""
        df1 = self.excel_data["s1"]
        df2 = self.excel_data["s2"]
        df3 = self.excel_data["s3"]
        df4 = self.excel_data["s4"]  # New: Get 4th DataFrame

        # 1) Filter + deduplicate rows in Sheet‑1 --------------------------
        # Filter rows where column C and A are not empty
        df1 = df1[
            df1[self.SHEET1_COLS["C"]].notna() & df1[self.SHEET1_COLS["A"]].notna()
            ].copy()
        # Remove duplicate rows based on column C, keeping the first occurrence
        df1.drop_duplicates(subset=self.SHEET1_COLS["C"], keep="first", inplace=True)

        # Filter rows where 'Malzeme' column (original index 2) has 3 or more hyphens
        # Convert to string to use .str methods, then count occurrences of '-'
        df1 = df1[df1[self.SHEET1_COLS["C"]].astype(str).str.count('-') < 3].copy()

        # 2) Prepare table dimensions -------------------------------------
        self.table.setColumnCount(len(self.HEADER_LABELS))  # Set number of columns based on headers
        self.table.setHorizontalHeaderLabels(self.HEADER_LABELS)  # Set column headers
        self.table.setRowCount(len(df1))  # Set number of rows based on filtered Sheet 1 data

        # 3) Fill rows -----------------------------------------------------
        for r, row in enumerate(df1.itertuples(index=False)):  # Iterate over rows of filtered Sheet 1
            # Sheet‑1 cols (Populate columns A, B, C, D in the table)
            self.table.setItem(r, 0, QTableWidgetItem(str(row[self.SHEET1_COLS["A"]])))  # Col A
            self.table.setItem(r, 1, QTableWidgetItem(str(row[self.SHEET1_COLS["C"]])))  # Col B
            self.table.setItem(r, 2, QTableWidgetItem(str(row[self.SHEET1_COLS["G"]])))  # Col C
            self.table.setItem(r, 3, QTableWidgetItem(str(row[self.SHEET1_COLS["E"]])))  # Col D

            match_val = row[self.SHEET1_COLS["C"]]  # Value from Sheet 1 column C for matching

            # Sheet‑2 match and aggregation (Populate columns E, F in the table)
            s2_matches = df2[df2[self.COMMON_MATCH_COL["G"]] == match_val]
            if not s2_matches.empty:
                # Column E (table index 4): Take the value from Sheet 2's column B of the first match
                self.table.setItem(r, 4, QTableWidgetItem(str(s2_matches.iloc[0][self.SHEET2_COLS["B"]])))
                # Column F (table index 5): Sum all matching values from Sheet 2's column J
                sum_j_s2 = s2_matches[self.SHEET2_COLS["J"]].apply(self._to_float_series).sum()
                self.table.setItem(r, 5, QTableWidgetItem(str(sum_j_s2)))
            else:
                self.table.setItem(r, 4, QTableWidgetItem(""))
                self.table.setItem(r, 5, QTableWidgetItem(""))

            # Sheet‑3 match and aggregation (Populate columns H, I, J in the table)
            s3_matches = df3[df3[self.COMMON_MATCH_COL["G"]] == match_val]
            if not s3_matches.empty:
                # Column H (table index 6): Take the value from Sheet 3's column B of the first match
                self.table.setItem(r, 6, QTableWidgetItem(str(s3_matches.iloc[0][self.SHEET3_COLS["B"]])))
                # Column I (table index 7): Sum all matching values from Sheet 3's column J
                sum_j_s3 = s3_matches[self.SHEET3_COLS["J"]].apply(self._to_float_series).sum()
                self.table.setItem(r, 7, QTableWidgetItem(str(sum_j_s3)))
                # Column J (table index 8): Sum all matching values from Sheet 3's column K
                sum_k_s3 = s3_matches[self.SHEET3_COLS["K"]].apply(self._to_float_series).sum()
                self.table.setItem(r, 8, QTableWidgetItem(str(sum_k_s3)))
            else:
                self.table.setItem(r, 6, QTableWidgetItem(""))
                self.table.setItem(r, 7, QTableWidgetItem(""))
                self.table.setItem(r, 8, QTableWidgetItem(""))

            # K (İhtiyaç) column (table index 9) initially empty
            self.table.setItem(r, 9, QTableWidgetItem(""))

            # L (Durum) column (table index 10) initial calculation (K assumed 0)
            self._update_l_column(r)

            # New: Update "Verilen Sipariş Miktarı" and "Verilmesi Gereken Sipariş Miktarı" columns
            self._update_order_quantities(r, df4)

        # 4) Resize + connect once ----------------------------------------
        self.table.resizeColumnsToContents()  # Adjust column widths to fit content
        # Allow users to resize columns manually
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        # Fix row height
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.table.verticalHeader().setDefaultSectionSize(30)

        # Disconnect existing cellChanged signal to prevent duplicate connections
        if self._cell_connected:
            try:
                self.table.cellChanged.disconnect(self._cell_changed)
            except TypeError:  # Handle case where signal might not be connected yet
                pass
        self.table.cellChanged.connect(self._cell_changed)  # Connect the signal
        self._cell_connected = True  # Set flag that signal is connected

    # -------------------------------------------------------------------- #
    #                        Cell change handlers
    # -------------------------------------------------------------------- #
    def _cell_changed(self, row: int, col: int):
        """Handles changes in table cells, specifically for the 'İhtiyaç' (K) column.
        Propagates the entered value multiplied by column D to all cells below in the same column."""
        # If an update is already in progress or the changed column is not 'K' (index 9), return
        if self._updating or col != 9:
            return

        try:
            # Get the text from the changed cell, replace comma with dot for float conversion
            k_raw = self.table.item(row, col).text().replace(",", ".")
            k_input_value = float(k_raw)  # Convert to float
        except (ValueError, AttributeError):
            # If the input is not a valid number, clear the cell and recalculate L
            self._updating = True  # Set updating flag to prevent recursion
            self.table.setItem(row, col, QTableWidgetItem(""))  # Clear the invalid input
            self._update_l_column(row)  # Recalculate L for the current row with K=0
            # New: Also update order quantities if K changes
            self._update_order_quantities(row, self.excel_data["s4"])
            self._updating = False  # Reset updating flag
            return

        self._updating = True  # Set updating flag to prevent recursion
        # Apply the new K value (multiplied by D column) to the changed cell and all cells below it in column K
        for r_idx in range(row, self.table.rowCount()):
            # Get the value from column D (index 3) for the current row
            d_val = self._to_float(self.table.item(r_idx, 3))

            # Calculate the new 'İhtiyaç' value by multiplying D column value with the user's input
            calculated_k_value = d_val * k_input_value

            # Set the K column item for the current row to the calculated value
            self.table.setItem(r_idx, 9, QTableWidgetItem(str(calculated_k_value)))
            # Recalculate and update the L column for the current row based on the new K value
            self._update_l_column(r_idx)
            # New: Also update order quantities if K changes
            self._update_order_quantities(r_idx, self.excel_data["s4"])
        self._updating = False  # Reset updating flag

    def _to_float(self, item: QTableWidgetItem) -> float:
        """Converts a QTableWidgetItem's text to a float, handling commas and empty strings."""
        try:
            if item is None or item.text() == "":
                return 0.0
            return float(item.text().replace(",", "."))
        except (ValueError, AttributeError):
            return 0.0

    def _to_float_series(self, value) -> float:
        """Converts a value from a pandas Series to float, handling non-numeric values
        and comma decimal separators."""
        try:
            if isinstance(value, str):
                # Remove thousands separators (if any) and replace comma decimal with dot
                value = value.replace(".", "").replace(",", ".")
            return float(value)
        except (ValueError, TypeError):
            return 0.0

    def _update_l_column(self, row: int):
        """Calculates and updates the 'Durum' (L) column based on values in
        columns F, I, J, and K for a given row."""
        # Get values from relevant columns, converting them to float
        f_val = self._to_float(self.table.item(row, 5))  # Column F (Sheet 2 J sum)
        i_val = self._to_float(self.table.item(row, 7))  # Column I (Sheet 3 J sum)
        j_val = self._to_float(self.table.item(row, 8))  # Column J (Sheet 3 K sum)
        k_val = self._to_float(self.table.item(row, 9))  # Column K (İhtiyaç)

        # Calculate the result for 'Durum' (L)
        result = f_val + i_val + j_val - k_val
        # Format the text: if result is negative, add "#SİPARİŞ VER"
        text = f"{result} #SİPARİŞ VER" if result < 0 else str(result)

        # Get or create the QTableWidgetItem for the L column
        item = self.table.item(row, 10)
        if item is None:
            item = QTableWidgetItem()
            # Make the L column non-editable as it's a calculated field
            item.setFlags(item.flags() ^ Qt.ItemIsEditable)
            self.table.setItem(row, 10, item)
        item.setText(text)  # Set the calculated text

    def _update_order_quantities(self, row: int, df4: pd.DataFrame):
        """
        Calculates and updates 'Verilen Sipariş Miktarı' and 'Verilmesi Gereken Sipariş Miktarı'
        for a given row based on 'Durum' column and 4th Excel sheet.
        """
        durum_item = self.table.item(row, 10)  # 'Durum' column
        malzeme_item = self.table.item(row, 1)  # 'Malzeme' column

        verilen_siparis_miktari = 0.0
        verilmesi_gereken_siparis_miktari = 0.0

        if durum_item and "#SİPARİŞ VER" in durum_item.text():
            try:
                # Extract the numeric part of the 'Durum' value
                durum_numeric_str = durum_item.text().split(" #SİPARİŞ VER")[0].replace(",", ".")
                durum_numeric_val = float(durum_numeric_str)
            except (ValueError, AttributeError):
                durum_numeric_val = 0.0

            if malzeme_item:
                malzeme_val = malzeme_item.text()
                # Find matches in the 4th sheet based on 'Malzeme' (column C, index 2)
                s4_matches = df4[df4[self.SHEET4_COLS["C"]] == malzeme_val]

                if not s4_matches.empty:
                    # Sum values from the 8th index column (column I) in the 4th sheet
                    verilen_siparis_miktari = s4_matches[self.SHEET4_COLS["I"]].apply(self._to_float_series).sum()

            # Calculate 'Verilmesi Gereken Sipariş Miktarı'
            remaining_needed = durum_numeric_val - verilen_siparis_miktari
            verilmesi_gereken_siparis_miktari = max(0.0, remaining_needed)  # Ensure it's not negative

        # Set items for "Verilen Sipariş Miktarı" (index 11)
        item_verilen = self.table.item(row, 11)
        if item_verilen is None:
            item_verilen = QTableWidgetItem()
            item_verilen.setFlags(item_verilen.flags() ^ Qt.ItemIsEditable)  # Make non-editable
            self.table.setItem(row, 11, item_verilen)
        item_verilen.setText(str(verilen_siparis_miktari))

        # Set items for "Verilmesi Gereken Sipariş Miktarı" (index 12)
        item_gereken = self.table.item(row, 12)
        if item_gereken is None:
            item_gereken = QTableWidgetItem()
            item_gereken.setFlags(item_gereken.flags() ^ Qt.ItemIsEditable)  # Make non-editable
            self.table.setItem(row, 12, item_gereken)
        item_gereken.setText(str(verilmesi_gereken_siparis_miktari))

    def _process_fsnkp_rows(self):
        """
        Processes rows to remove 'FSNKP' entries and update the 'Durum' column of the preceding row.
        Iterates backwards to handle row deletion correctly.
        """
        self._updating = True  # Prevent cellChanged signal from firing during this process

        rows_to_remove = []
        # Iterate backwards from the second-to-last row down to the first data row (index 1)
        # We need to check against the previous row, so we stop at index 1.
        for r_idx in range(self.table.rowCount() - 1, 0, -1):
            current_malzeme = self.table.item(r_idx, 1).text() if self.table.item(r_idx, 1) else ""
            prev_malzeme = self.table.item(r_idx - 1, 1).text() if self.table.item(r_idx - 1, 1) else ""
            current_aciklama = self.table.item(r_idx, 2).text() if self.table.item(r_idx, 2) else ""

            # Check if current row's 'Malzeme' (column 1) matches previous row's 'Malzeme'
            # and current row's 'Açıklama' (column 2) contains "FSNKP"
            if current_malzeme == prev_malzeme and "FSNKP" in current_aciklama:
                # Add "#FSNKP" to the 'Durum' column (column 10) of the previous row
                prev_durum_item = self.table.item(r_idx - 1, 10)
                if prev_durum_item:
                    current_durum_text = prev_durum_item.text()
                    if "#FSNKP" not in current_durum_text:  # Avoid adding duplicate "#FSNKP"
                        prev_durum_item.setText(current_durum_text + " #FSNKP")

                # Mark the current row for removal
                rows_to_remove.append(r_idx)

        # Remove rows marked for removal (from highest index to lowest to avoid index shifting issues)
        for r_idx in sorted(rows_to_remove, reverse=True):
            self.table.removeRow(r_idx)

        self._updating = False  # Re-enable cellChanged signal

    # -------------------------------------------------------------------- #
    #                           Save to Excel
    # -------------------------------------------------------------------- #
    def _save_excel(self):
        """Saves the current data from the QTableWidget to a new Excel file."""
        # Open a save file dialog
        path, _ = QFileDialog.getSaveFileName(self, "Uyarlanmış Excel Dosyasını Kaydet", "uyarlanmis_veri.xlsx",
                                              "Excel Dosyaları (*.xlsx)")
        if not path:  # If no path is selected, return
            return

        rows, cols = self.table.rowCount(), self.table.columnCount()
        # Extract all data from the QTableWidget into a list of lists,
        # starting from the second row (index 1) to exclude the first data row from the saved file.
        data = [[self.table.item(r, c).text() if self.table.item(r, c) else "" for c in range(cols)] for r in
                range(1, rows)]

        # Use the predefined HEADER_LABELS directly as column headers for the DataFrame
        pd.DataFrame(data, columns=self.HEADER_LABELS).to_excel(path, index=False)
        QMessageBox.information(self, "Başarılı", f"Dosya kaydedildi: {path.split('/')[-1]}")


# ----------------------------------------------------------------------- #
#                                   main
# ----------------------------------------------------------------------- #
if __name__ == "__main__":
    app = QApplication(sys.argv)  # Create the QApplication instance
    window = ExcelProcessorApp()  # Create an instance of the main application window
    window.show()  # Display the window
    sys.exit(app.exec_())  # Start the application event loop
