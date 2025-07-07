import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QStackedWidget, QTableWidget,
    QTableWidgetItem, QMessageBox, QHeaderView, QFrame
)
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QIcon, QFont

class ExcelProcessorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Veri Yönetimi Uygulaması")
        self.setGeometry(100, 100, 1200, 800)
        self.setWindowIcon(QIcon("icon.png")) # Uygulama ikonu (isteğe bağlı, kendi ikonunuzu ekleyin)

        # Ana stil ayarları
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f2f5;
            }
            QWidget {
                font-family: 'Segoe UI', sans-serif;
                font-size: 14px;
            }
            QLabel#titleLabel {
                font-size: 28px;
                font-weight: bold;
                color: #2c3e50;
                margin-bottom: 20px;
            }
            QPushButton {
                background-color: #3498db;
                color: white;
                border-radius: 8px;
                padding: 12px 25px;
                font-size: 15px;
                font-weight: bold;
                border: none;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
            QFrame#cardFrame {
                background-color: white;
                border-radius: 10px;
                padding: 30px;
                box-shadow: 0px 4px 15px rgba(0, 0, 0, 0.1);
            }
            QTableWidget {
                background-color: white;
                border: 1px solid #dcdcdc;
                gridline-color: #f0f0f0;
                selection-background-color: #aed6f1;
                selection-color: black;
                font-size: 13px;
            }
            QHeaderView::section {
                background-color: #e9ecef;
                color: #495057;
                padding: 8px;
                border: 1px solid #dcdcdc;
                font-weight: bold;
            }
            QLabel#filePathLabel {
                font-style: italic;
                color: #555;
                font-size: 13px;
                margin-top: 10px;
            }
        """)

        self.stacked_widget = QStackedWidget()
        self.setCentralWidget(self.stacked_widget)

        self.setup_ui()

        self.excel_data = {} # Tüm excel verisini tutacak (DataFrame'ler)
        self.selected_file_path = None
        self.sheet_names = [] # Sayfa isimlerini tutacak

        # Sütun indeksleri tanımlamaları (Excel'deki 0-tabanlı indekslere göre)
        self.sheet1_cols_map = {
            'A': 0,
            'C': 2, # Eşleştirme için kullanılacak C indeksi
            'G': 6,
            'E': 4
        }
        # Sheet 2'den alınacak indeksler: B, J (K hariç)
        # Sheet 3'ten alınacak indeksler: B, J, K
        self.sheet2_specific_cols_map = {
            'B': 1,
            'J': 9,
        }
        self.sheet3_specific_cols_map = {
            'B': 1,
            'J': 9,
            'K': 10
        }
        # Ortak eşleşme indeksi
        self.common_match_col_map = {
            'G': 6 # Eşleştirme için 2. ve 3. sayfalardaki G indeksi
        }

        # Eşleşme için kullanılacak sütunların indeksleri
        self.matching_col_sheet1_idx = self.sheet1_cols_map['C'] # 1. sayfadaki eşleşme indeksi (C)
        self.matching_col_sheet23_idx = self.common_match_col_map['G'] # 2. ve 3. sayfalardaki eşleşme indeksi (G)

    def setup_ui(self):
        # --- Dosya Seç Sayfası ---
        self.file_selection_page = QWidget()
        main_layout_fs = QVBoxLayout()
        self.file_selection_page.setLayout(main_layout_fs)

        card_frame_fs = QFrame()
        card_frame_fs.setObjectName("cardFrame")
        card_layout_fs = QVBoxLayout()
        card_frame_fs.setLayout(card_layout_fs)
        card_frame_fs.setFixedWidth(500)
        card_frame_fs.setFixedHeight(350)

        title_label_fs = QLabel("Excel Dosyası Seçin")
        title_label_fs.setObjectName("titleLabel")
        title_label_fs.setAlignment(Qt.AlignCenter)
        card_layout_fs.addWidget(title_label_fs)

        card_layout_fs.addSpacing(30)

        self.select_file_button = QPushButton("Dosya Seç")
        self.select_file_button.clicked.connect(self.select_excel_file)
        card_layout_fs.addWidget(self.select_file_button)

        self.selected_file_label = QLabel("Seçilen Dosya: Yüklenmedi")
        self.selected_file_label.setObjectName("filePathLabel")
        self.selected_file_label.setAlignment(Qt.AlignCenter)
        card_layout_fs.addWidget(self.selected_file_label)

        card_layout_fs.addSpacing(20)

        self.go_to_custom_button = QPushButton("Uyarlanmış Dosyayı Görüntüle")
        self.go_to_custom_button.setEnabled(False)
        self.go_to_custom_button.clicked.connect(self.go_to_custom_file_page)
        card_layout_fs.addWidget(self.go_to_custom_button)

        main_layout_fs.addStretch()
        main_layout_fs.addWidget(card_frame_fs, alignment=Qt.AlignCenter)
        main_layout_fs.addStretch()

        self.stacked_widget.addWidget(self.file_selection_page)

        # --- Uyarlanmış Dosya Sayfası ---
        self.custom_file_page = QWidget()
        custom_file_layout = QVBoxLayout()
        self.custom_file_page.setLayout(custom_file_layout)

        custom_title_label = QLabel("Uyarlanmış Excel Verileri")
        custom_title_label.setObjectName("titleLabel")
        custom_title_label.setAlignment(Qt.AlignCenter)
        custom_file_layout.addWidget(custom_title_label)

        custom_file_layout.addSpacing(15)

        self.table_widget = QTableWidget()
        self.table_widget.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.AnyKeyPressed)
        self.table_widget.setAlternatingRowColors(True)
        custom_file_layout.addWidget(self.table_widget)

        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.save_button = QPushButton("Değişiklikleri Kaydet")
        self.save_button.clicked.connect(self.save_custom_excel)
        button_layout.addWidget(self.save_button)

        self.back_button = QPushButton("Geri Dön")
        self.back_button.clicked.connect(lambda: self.stacked_widget.setCurrentWidget(self.file_selection_page))
        button_layout.addWidget(self.back_button)
        button_layout.addStretch()

        custom_file_layout.addLayout(button_layout)
        custom_file_layout.addSpacing(20)

        self.stacked_widget.addWidget(self.custom_file_page)

    def select_excel_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self,
                                                   "Excel Dosyası Seç",
                                                   "",
                                                   "Excel Dosyaları (*.xlsx)",
                                                   options=options)
        if file_path:
            self.selected_file_path = file_path
            self.selected_file_label.setText(f"Seçilen Dosya: {file_path.split('/')[-1]}")
            self.go_to_custom_button.setEnabled(False)
            self.load_excel_data()
        else:
            self.selected_file_path = None
            self.selected_file_label.setText("Seçilen Dosya: Yüklenmedi")
            self.go_to_custom_button.setEnabled(False)

    def load_excel_data(self):
        if not self.selected_file_path:
            QMessageBox.warning(self, "Hata", "Lütfen önce bir Excel dosyası seçin.")
            return

        try:
            xls = pd.ExcelFile(self.selected_file_path)
            self.sheet_names = xls.sheet_names

            if len(self.sheet_names) < 3:
                QMessageBox.critical(self, "Hata", "Seçilen Excel dosyasında en az 3 sayfa bulunmalıdır.")
                self.go_to_custom_button.setEnabled(False)
                return

            self.excel_data['Sheet1_Indexed'] = pd.read_excel(xls, sheet_name=self.sheet_names[0], header=None)
            self.excel_data['Sheet2_Indexed'] = pd.read_excel(xls, sheet_name=self.sheet_names[1], header=None)
            self.excel_data['Sheet3_Indexed'] = pd.read_excel(xls, sheet_name=self.sheet_names[2], header=None)

            QMessageBox.information(self, "Başarılı", "Excel dosyası başarıyla yüklendi.")
            self.go_to_custom_button.setEnabled(True)

        except FileNotFoundError:
            QMessageBox.critical(self, "Hata", "Dosya bulunamadı.")
            self.go_to_custom_button.setEnabled(False)
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Excel dosyası yüklenirken bir hata oluştu: {e}")
            self.go_to_custom_button.setEnabled(False)

    def go_to_custom_file_page(self):
        if not all(k in self.excel_data for k in ['Sheet1_Indexed', 'Sheet2_Indexed', 'Sheet3_Indexed']):
            QMessageBox.warning(self, "Uyarı", "Excel verileri tam olarak yüklenemedi. Lütfen tekrar deneyin.")
            return

        self.populate_custom_table()
        self.stacked_widget.setCurrentWidget(self.custom_file_page)

    def populate_custom_table(self):
        sheet1_df_original = self.excel_data.get('Sheet1_Indexed')
        sheet2_df = self.excel_data.get('Sheet2_Indexed')
        sheet3_df = self.excel_data.get('Sheet3_Indexed')

        # Yeni gereksinim: Sadece 1. sayfada hem 2. (C) hem de 0. (A) indekslerde değeri olan satırları filtrele
        sheet1_df = sheet1_df_original[
            sheet1_df_original[self.sheet1_cols_map['C']].notna() &
            sheet1_df_original[self.sheet1_cols_map['A']].notna()
        ].copy() # Filtreleme sonrası kopya oluşturarak SettingWithCopyWarning'i önle

        sheet1_df.drop_duplicates(
            subset=self.sheet1_cols_map['C'],  # C sütunu
            keep='first',
            inplace=True
        )
        # Sütun indeksleri listesi
        sheet1_display_indices = [
            self.sheet1_cols_map['A'],
            self.sheet1_cols_map['C'],
            self.sheet1_cols_map['G'],
            self.sheet1_cols_map['E']
        ]
        sheet2_display_indices = [
            self.sheet2_specific_cols_map['B'],
            self.sheet2_specific_cols_map['J']
        ]
        sheet3_display_indices = [
            self.sheet3_specific_cols_map['B'],
            self.sheet3_specific_cols_map['J'],
            self.sheet3_specific_cols_map['K']
        ]

        # Sütun indekslerinin varlığını kontrol et
        max_col_sheet1 = sheet1_df_original.shape[1] - 1 if not sheet1_df_original.empty else -1
        max_col_sheet2 = sheet2_df.shape[1] - 1 if not sheet2_df.empty else -1
        max_col_sheet3 = sheet3_df.shape[1] - 1 if not sheet3_df.empty else -1

        # Sheet1'den çekilecek tüm indeksler ve eşleşme indeksi
        all_required_sheet1_indices = list(set(sheet1_display_indices + [self.matching_col_sheet1_idx]))
        # Sheet2'den çekilecek tüm indeksler ve eşleşme indeksi
        all_required_sheet2_indices = list(set(sheet2_display_indices + [self.matching_col_sheet23_idx]))
        # Sheet3'ten çekilecek tüm indeksler ve eşleşme indeksi
        all_required_sheet3_indices = list(set(sheet3_display_indices + [self.matching_col_sheet23_idx]))

        if not all(idx <= max_col_sheet1 for idx in all_required_sheet1_indices):
            missing_indices = [idx for idx in all_required_sheet1_indices if idx > max_col_sheet1]
            QMessageBox.critical(self, "Hata", f"İlk sayfada eksik sütun indeksleri var: {missing_indices}. Mevcut max indeks: {max_col_sheet1}")
            return
        if not all(idx <= max_col_sheet2 for idx in all_required_sheet2_indices):
            missing_indices = [idx for idx in all_required_sheet2_indices if idx > max_col_sheet2]
            QMessageBox.critical(self, "Hata", f"İkinci sayfada eksik sütun indeksleri var: {missing_indices}. Mevcut max indeks: {max_col_sheet2}")
            return
        if not all(idx <= max_col_sheet3 for idx in all_required_sheet3_indices):
            missing_indices = [idx for idx in all_required_sheet3_indices if idx > max_col_sheet3]
            QMessageBox.critical(self, "Hata", f"Üçüncü sayfada eksik sütun indeksleri var: {missing_indices}. Mevcut max indeks: {max_col_sheet3}")
            return

        # Toplam sütun sayısı: Sheet1 (4) + Sheet2 (2) + Sheet3 (3) + K & L (2) = 11
        self.table_widget.setColumnCount(
            len(sheet1_display_indices) +
            len(sheet2_display_indices) +
            len(sheet3_display_indices) + 2  # +K, +L
        )

        headers = [
            'A', 'B', 'C', 'D',  # Sheet‑1
            'E', 'F',  # Sheet‑2
            'H', 'I', 'J',  # Sheet‑3
            'K', 'L'  # Yeni sütunlar
        ]
        self.table_widget.setHorizontalHeaderLabels(headers)

        # Tabloya veri doldurma
        # Filtrelenmiş sheet1_df'in satır sayısını kullan
        self.table_widget.setRowCount(len(sheet1_df))

        for row_idx_new, row_data in enumerate(sheet1_df.itertuples(index=False)): # itertuples daha performanslı
            # 1. sayfadan belirtilen indeksleri al
            self.table_widget.setItem(row_idx_new, 0, QTableWidgetItem(str(row_data[self.sheet1_cols_map['A']])))
            self.table_widget.setItem(row_idx_new, 1, QTableWidgetItem(str(row_data[self.sheet1_cols_map['C']])))
            self.table_widget.setItem(row_idx_new, 2, QTableWidgetItem(str(row_data[self.sheet1_cols_map['G']])))
            self.table_widget.setItem(row_idx_new, 3, QTableWidgetItem(str(row_data[self.sheet1_cols_map['E']])))

            # Eşleştirme için 1. sayfadan C indeksi değerini al
            matching_value_sheet1_c = row_data[self.matching_col_sheet1_idx]

            # 2. sayfadan G indeksi ile eşleşen verileri çek (B ve J'yi al, K'yı alma)
            matched_sheet2_rows = sheet2_df[sheet2_df[self.matching_col_sheet23_idx] == matching_value_sheet1_c]
            if not matched_sheet2_rows.empty:
                s2_row = matched_sheet2_rows.iloc[0]
                self.table_widget.setItem(row_idx_new, 4, QTableWidgetItem(str(s2_row[self.sheet2_specific_cols_map['B']])))
                self.table_widget.setItem(row_idx_new, 5, QTableWidgetItem(str(s2_row[self.sheet2_specific_cols_map['J']])))
            else:
                for col_offset in range(len(sheet2_display_indices)): # B, J için boş bırak (2 sütun)
                    self.table_widget.setItem(row_idx_new, 4 + col_offset, QTableWidgetItem(""))

            # 3. sayfadan G indeksi ile eşleşen verileri çek (B, J, K'yı al)
            matched_sheet3_rows = sheet3_df[sheet3_df[self.matching_col_sheet23_idx] == matching_value_sheet1_c]
            if not matched_sheet3_rows.empty:
                s3_row = matched_sheet3_rows.iloc[0]
                self.table_widget.setItem(row_idx_new, 6, QTableWidgetItem(str(s3_row[self.sheet3_specific_cols_map['B']])))
                self.table_widget.setItem(row_idx_new, 7, QTableWidgetItem(str(s3_row[self.sheet3_specific_cols_map['J']])))
                self.table_widget.setItem(row_idx_new, 8, QTableWidgetItem(str(s3_row[self.sheet3_specific_cols_map['K']])))
            else:
                for col_offset in range(len(sheet3_display_indices)): # B, J, K için boş bırak (3 sütun)
                    self.table_widget.setItem(row_idx_new, 6 + col_offset, QTableWidgetItem(""))

            # --- K sütunu (ihtiyaç) başlangıçta boş; kullanıcı girince çarpım yapılacak
            self.table_widget.setItem(row_idx_new, 9, QTableWidgetItem(""))

            # --- L sütunu: (F + I + J) - K  → ilk etapta K=0 kabul edilir
            f_val = float(self.table_widget.item(row_idx_new, 5).text() or 0)
            i_val = float(self.table_widget.item(row_idx_new, 7).text() or 0)
            j_val = float(self.table_widget.item(row_idx_new, 8).text() or 0)
            l_result = f_val + i_val + j_val  # K şimdilik 0

            l_item = QTableWidgetItem(str(l_result))
            l_item.setFlags(l_item.flags() ^ Qt.ItemIsEditable)  # Kullanıcı düzenleyemesin
            self.table_widget.setItem(row_idx_new, 10, l_item)

        self.table_widget.resizeColumnsToContents()
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_widget.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.table_widget.verticalHeader().setDefaultSectionSize(30)
        self.table_widget.cellChanged.connect(self.handle_cell_changed)

    def handle_cell_changed(self, row: int, col: int):
            # Yalnızca K sütunu (index 9) izlenir
            if col != 9:
                return

            try:
                k_input = float(self.table_widget.item(row, col).text().replace(',', '.'))
            except (ValueError, AttributeError):
                return  # Geçersiz giriş

            # D sütunundaki miktar (index 3)
            try:
                d_val = float(self.table_widget.item(row, 3).text().replace(',', '.'))
            except (ValueError, AttributeError):
                d_val = 0.0

            product = d_val * k_input

            # Sonsuz döngüyü önlemek için sinyalleri kapat
            self.table_widget.blockSignals(True)
            self.table_widget.item(row, col).setText(str(product))
            self.update_l_column(row)
            self.table_widget.blockSignals(False)

    def update_l_column(self, row: int):
            def to_float(item):
                try:
                    return float(item.text().replace(',', '.'))
                except (ValueError, AttributeError):
                    return 0.0

            f_val = to_float(self.table_widget.item(row, 5))
            i_val = to_float(self.table_widget.item(row, 7))
            j_val = to_float(self.table_widget.item(row, 8))
            k_val = to_float(self.table_widget.item(row, 9))

            result = f_val + i_val + j_val - k_val
            text = f"{result} #SİPARİŞ VER" if result < 0 else str(result)

            l_item = self.table_widget.item(row, 10)
            if l_item is None:
                l_item = QTableWidgetItem(text)
                l_item.setFlags(l_item.flags() ^ Qt.ItemIsEditable)
                self.table_widget.setItem(row, 10, l_item)
            else:
                l_item.setText(text)

    def save_custom_excel(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self,
                                                   "Uyarlanmış Excel Dosyasını Kaydet",
                                                   "uyarlanmis_veri.xlsx",
                                                   "Excel Dosyaları (*.xlsx)",
                                                   options=options)
        if file_path:
            try:
                row_count = self.table_widget.rowCount()
                column_count = self.table_widget.columnCount()
                data = []

                for row in range(row_count):
                    row_data = []
                    for col in range(column_count):
                        item = self.table_widget.item(row, col)
                        row_data.append(item.text() if item else "")
                    data.append(row_data)

                # Yeni oluşturulan dosya için A, B, C... şeklinde sütun isimleri oluştur
                def excel_column_name(n):
                    name = ''
                    while n >= 0:
                        name = chr(n % 26 + ord('A')) + name
                        n = n // 26 - 1
                    return name

                new_excel_headers = [excel_column_name(i) for i in range(column_count)]

                df_to_save = pd.DataFrame(data, columns=new_excel_headers)

                df_to_save.to_excel(file_path, index=False)
                QMessageBox.information(self, "Başarılı", f"Dosya başarıyla kaydedildi:\n{file_path.split('/')[-1]}")
            except Exception as e:
                QMessageBox.critical(self, "Hata", f"Dosya kaydedilirken bir hata oluştu: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = ExcelProcessorApp()
    ex.show()
    sys.exit(app.exec_())