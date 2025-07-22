import sys
from typing import List
import datetime

import pandas as pd
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QColor, QBrush  # QColor and QBrush added for highlighting
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

# Matplotlib imports for charting
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure


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
    SHEET3_COLS = {"B": 1, "J": 9, "K": 10, "L": 11}  # J, K and L are the columns to be summed
    COMMON_MATCH_COL = {"G": 6}  # Column G (index 6) is used for matching across sheets
    # New: Column mappings for the 4th sheet
    # Assuming 2nd index (column C) for matching material, 8th index (column I) for ordered quantity
    # New: 18th index (column S) for delivery date
    SHEET4_COLS = {"C": 2, "I": 8, "S": 18}

    # Header labels for the displayed QTableWidget
    # These are the headers for the *data* columns, and will be used for the inserted rows.
    HEADER_LABELS = [
        "Ü.Ağacı Sev", "Malzeme", "Açıklama", "Miktar",
        "Depo 100", "Kullanılabilir Stok",
        "Depo 110", "Kullanılabilir Stok", "Kalite Stoğu",
        "İhtiyaç", "Durum",
        "Verilen Sipariş Miktarı", "Verilmesi Gereken Sipariş Miktarı",
        "Teslim Tarihi"
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
        self.chart_figure = None  # To store the matplotlib figure for saving
        self.highlighted_rows = []  # Store indices of rows to be highlighted in Excel

        self._build_style()  # Apply custom CSS styling
        self._build_pages()  # Construct the UI pages

    # -------------------------------------------------------------------- #
    #                           UI Construction
    # -------------------------------------------------------------------- #
    def _build_style(self):
        """Uygulama widget'larına özel CSS stilini uygular."""
        self.setStyleSheet(
            """
            QMainWindow { background: #f0f2f5; } /* Ana pencere için açık gri arka plan */
            QWidget     { font-family: 'Segoe UI', sans-serif; font-size: 14px; } /* Widget'lar için varsayılan yazı tipi */
            QLabel#titleLabel { font-size: 28px; font-weight: bold; color: #2c3e50; margin-bottom: 20px; } /* Başlık etiketleri için stil */
            QPushButton { /* Butonlar için stil */
                background: #3498db;
                color: white;
                border-radius: 8px;
                padding: 12px 25px;
                font-size: 15px;
                font-weight: bold;
                border: none;
            }
            QPushButton:hover    { background: #2980b9; } /* Üzerine gelindiğinde daha koyu mavi */
            QPushButton:disabled { background: #cccccc; color: #666666; } /* Devre dışı bırakılan butonlar için gri tonları */
            QFrame#card { background: white; border-radius: 10px; padding: 30px; } /* Kart benzeri çerçeveler için stil */
            QTableWidget         { /* Tablo widget'ı için stil */
                background: white;
                border: 1px solid #dcdcdc;
                gridline-color: #f0f0f0;
                selection-background-color: #aed6f1;
                font-size: 13px;
            }
            QHeaderView::section { /* Tablo başlıkları için stil */
                background: #e9ecef;
                color: #495057;
                padding: 8px;
                border: 1px solid #dcdcdc;
                font-weight: bold;
            }
            QLabel#filePathLabel { font-style: italic; color: #555; font-size: 13px; margin-top: 10px; } /* Dosya yolu etiketi için stil */
            QFrame#chartContainer {
                background: #e0e0e0; /* Grafik kapsayıcısı için açık gri */
                border: 1px solid #ccc; /* Grafik kapsayıcısı için kenarlık */
                border-radius: 10px;
                padding: 10px;
            }
            """
        )

    def _build_pages(self):
        """Uygulamanın iki ana sayfasını (dosya seçimi ve tablo görünümü) oluşturur."""
        self.stacked_widget = QStackedWidget(self)
        self.setCentralWidget(self.stacked_widget)

        # 1) Dosya seçimi sayfası ------------------------------------------------
        self.file_page = QWidget()
        main_v = QVBoxLayout(self.file_page)  # Dosya seçimi sayfası için ana düzen

        card = QFrame(objectName="card")  # Butonlar ve etiketler için kart çerçevesi
        card.setFixedSize(500, 350)  # Kart için sabit boyut
        card_v = QVBoxLayout(card)  # Kart içeriği için düzen

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

        main_v.addStretch(1)  # İçeriği merkeze it
        main_v.addWidget(card, alignment=Qt.AlignCenter)
        main_v.addStretch(1)

        self.stacked_widget.addWidget(self.file_page)  # Dosya sayfasını yığılmış widget'a ekle

        # 2) Tablo sayfası -----------------------------------------------------
        self.table_page = QWidget()
        tv = QVBoxLayout(self.table_page)  # Tablo sayfası için ana düzen

        lbl2 = QLabel("Uyarlanmış Excel Verileri", objectName="titleLabel", alignment=Qt.AlignCenter)
        tv.addWidget(lbl2)
        tv.addSpacing(15)

        self.table = QTableWidget(
            editTriggers=QTableWidget.DoubleClicked | QTableWidget.AnyKeyPressed,
            alternatingRowColors=True  # Satırlar için zebra şeritleri
        )
        tv.addWidget(self.table)

        hbox = QHBoxLayout()  # Kaydet ve geri butonları için düzen
        self.btn_save = QPushButton("Değişiklikleri Kaydet", clicked=self._save_excel)
        self.btn_back = QPushButton("Geri Dön", clicked=lambda: self.stacked_widget.setCurrentWidget(self.file_page))
        # Yeni: Grafik sayfasını açma butonu
        self.btn_show_chart = QPushButton("İş Tamamlanma Grafiği", clicked=self._open_chart_page)
        hbox.addStretch(1)
        hbox.addWidget(self.btn_save)
        hbox.addWidget(self.btn_back)
        hbox.addWidget(self.btn_show_chart)  # Grafik butonunu ekle
        hbox.addStretch(1)
        tv.addLayout(hbox)
        tv.addSpacing(20)

        self.stacked_widget.addWidget(self.table_page)  # Tablo sayfasını yığılmış widget'a ekle

        # 3) Grafik sayfası -----------------------------------------------------
        self.chart_page = QWidget()
        chart_v_layout = QVBoxLayout(self.chart_page)
        chart_v_layout.addStretch(1)  # Dikey ortalama için üst boşluk

        chart_page_title = QLabel("İş Tamamlanma Grafiği", objectName="titleLabel", alignment=Qt.AlignCenter)
        chart_v_layout.addWidget(chart_page_title)
        chart_v_layout.addSpacing(15)

        self.chart_container = QFrame(objectName="chartContainer")
        self.chart_container.setMinimumHeight(460)  # Grafik için minimum yükseklik ayarla
        chart_layout = QVBoxLayout(self.chart_container)
        chart_layout.setAlignment(Qt.AlignCenter)  # Grafik içeriğini kapsayıcısında ortala
        chart_v_layout.addWidget(self.chart_container)

        chart_hbox = QHBoxLayout()  # Grafik sayfası butonları için yeni düzen
        self.btn_chart_back = QPushButton("Geri Dön",
                                          clicked=lambda: self.stacked_widget.setCurrentWidget(self.table_page))
        self.btn_save_chart = QPushButton("Grafiği Kaydet", clicked=self._save_chart_as_image)
        chart_hbox.addStretch(1)
        chart_hbox.addWidget(self.btn_save_chart)
        chart_hbox.addWidget(self.btn_chart_back)
        chart_hbox.addStretch(1)
        chart_v_layout.addLayout(chart_hbox)
        chart_v_layout.addSpacing(20)
        chart_v_layout.addStretch(1)  # Dikey ortalama için alt boşluk

        self.stacked_widget.addWidget(self.chart_page)  # Grafik sayfasını yığılmış widget'a ekle

    # -------------------------------------------------------------------- #
    #                           Dosya Seçimi
    # -------------------------------------------------------------------- #
    def _select_file(self):
        """Kullanıcının bir Excel dosyası seçmesi için dosya iletişim kutusunu açar."""
        # GetOpenFileName (filePath, filter) döndürür, sadece filePath'a ihtiyacımız var
        path, _ = QFileDialog.getOpenFileName(self, "Excel Dosyası Seç", "", "Excel Dosyaları (*.xlsx)")
        if not path:  # Dosya seçilmezse geri dön
            return
        self.selected_file_path = path
        self.lbl_file.setText(f"Seçilen Dosya: {path.split('/')[-1]}")  # Seçilen dosya adını göster
        self._load_excel()  # Seçilen Excel dosyasını yüklemeyi dene

    def _load_excel(self):
        """Seçilen Excel dosyasından verileri pandas DataFrame'lerine yükler."""
        try:
            xls = pd.ExcelFile(self.selected_file_path)  # Bir ExcelFile nesnesi oluştur
            self.sheet_names = xls.sheet_names  # Tüm sayfa adlarını al
            # Yeni: En az 4 sayfa olup olmadığını kontrol et
            if len(self.sheet_names) < 4:
                raise ValueError("Seçilen Excel dosyasında en az 4 sayfa bulunmalıdır.")
            # İlk dört sayfayı DataFrame'lere yükle, ilk satırı (indeks 0) atla
            # Bu, orijinal Excel dosyasının ilk satırının işlenmemesini sağlar.
            self.excel_data = {
                "s1": pd.read_excel(xls, sheet_name=self.sheet_names[0], header=None, skiprows=[0]),
                "s2": pd.read_excel(xls, sheet_name=self.sheet_names[1], header=None, skiprows=[0]),
                "s3": pd.read_excel(xls, sheet_name=self.sheet_names[2], header=None, skiprows=[0]),
                "s4": pd.read_excel(xls, sheet_name=self.sheet_names[3], header=None, skiprows=[0]),
            }
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Excel dosyası yüklenirken bir hata oluştu:\n{e}")
            self.btn_open.setEnabled(False)  # Hata durumunda aç butonunu devre dışı bırak
            return

        QMessageBox.information(self, "Başarılı", "Excel dosyası başarıyla yüklendi.")
        self.btn_open.setEnabled(True)  # Başarılı yüklemede aç butonunu etkinleştir

    # -------------------------------------------------------------------- #
    #                           Tablo Doldurma
    # -------------------------------------------------------------------- #
    def _open_table_page(self):
        """Tablo görünümü sayfasına geçer ve tabloyu doldurur."""
        if not self.excel_data:  # Verilerin yüklendiğinden emin ol
            return
        self._populate_table()  # QTableWidget'ı işlenmiş verilerle doldur
        self._process_fsnkp_rows()  # İlk doldurmadan sonra FSNKP satırlarını işle
        self.stacked_widget.setCurrentWidget(self.table_page)  # Tablo sayfasına geç

    def _open_chart_page(self):
        """Grafik görünümü sayfasına geçer ve grafiği günceller."""
        if not self.excel_data:
            QMessageBox.warning(self, "Uyarı", "Lütfen önce bir Excel dosyası yükleyin ve tabloyu görüntüleyin.")
            return
        self._update_completion_chart()  # Sayfayı göstermeden önce grafiği güncelle
        self.stacked_widget.setCurrentWidget(self.chart_page)

    def _populate_table(self):
        """Yüklenen Excel sayfalarındaki verileri QTableWidget'a doldurur,
        tüm verileri herhangi bir koşul gözetmeksizin dahil eder, dinamik blok başlıkları da dahil."""
        df1 = self.excel_data["s1"]
        df2 = self.excel_data["s2"]
        df3 = self.excel_data["s3"]
        df4 = self.excel_data["s4"]

        # Dinamik tablo içeriği için hazırla
        final_table_content = []
        self.highlighted_rows = []  # Mevcut doldurma için vurgulanan satırları sıfırla

        # Mevcut blok için "####-####-####" kit kodunu tutar
        active_kit_code_for_block = None
        # Başlık satırından sonraki ilk veri satırını atlamak için bayrak
        skip_next_data_row_after_header = False

        highlight_color_brush = QBrush(QColor("#FFCCCC"))  # Vurgulama için açık kırmızı

        # Eklenen satırlarda görünecek gerçek sütun başlıklarını tanımla.
        # İlk öğe "Ü.Ağacı Sev" değeri olacak, ardından diğer başlıklar gelecek.
        internal_column_headers = self.HEADER_LABELS[1:]  # "Ü.Ağacı Sev" hariç tüm başlıklar

        # Sayfa 1'in tüm satırları üzerinde yinele
        for r_original_idx, row in enumerate(df1.itertuples(index=False)):
            # Sayfa 1'in 'A' sütunundaki ham değeri al.
            raw_val_from_sheet1_A = str(row[self.SHEET1_COLS["A"]])

            # Bu ham değerin bir kit kodu olup olmadığını kontrol et (tire içeriyor mu ve harf içeriyor mu?)
            is_kit_code = "-" in raw_val_from_sheet1_A and any(char.isalpha() for char in raw_val_from_sheet1_A)

            add_new_block_header = False
            if r_original_idx == 0:
                add_new_block_header = True
            elif is_kit_code and (
                    active_kit_code_for_block is None or raw_val_from_sheet1_A != active_kit_code_for_block):
                add_new_block_header = True

            if add_new_block_header:
                active_kit_code_for_block = raw_val_from_sheet1_A
                block_header_row_content = [active_kit_code_for_block] + internal_column_headers
                final_table_content.append(block_header_row_content)
                self.highlighted_rows.append(len(final_table_content) - 1)
                skip_next_data_row_after_header = True  # Bir sonraki veri satırını atlamak için işaretle

            # Bu satırın bir başlığın hemen altında olduğu için atlanıp atlanmayacağını kontrol et
            if skip_next_data_row_after_header:
                skip_next_data_row_after_header = False  # Bir sonraki yineleme için bayrağı sıfırla
                continue  # Bu satırı atla

            # Buraya ulaşırsak, satır nihai tablo içeriğine eklenmelidir
            current_data_row = [""] * len(self.HEADER_LABELS)

            # A sütununa yüklenen excel dosyasında 1. sayfadaki 0. indeksli sütundaki değeri yaz
            current_data_row[0] = str(row[self.SHEET1_COLS["A"]])

            current_data_row[1] = str(row[self.SHEET1_COLS["C"]])  # Malzeme
            current_data_row[2] = str(row[self.SHEET1_COLS["G"]])  # Açıklama
            current_data_row[3] = str(row[self.SHEET1_COLS["E"]])  # Miktar

            match_val = row[self.SHEET1_COLS["C"]]  # Sayfa 1 C sütunundaki eşleşme değeri

            # Sayfa 2 eşleşmesi ve toplama
            s2_matches = df2[df2[self.COMMON_MATCH_COL["G"]] == match_val]
            if not s2_matches.empty:
                current_data_row[4] = str(s2_matches.iloc[0][self.SHEET2_COLS["B"]])  # Depo 100
                # Toplamı al
                val_j_s2_sum = s2_matches[self.SHEET2_COLS["J"]].apply(self._to_float_series).sum()
                current_data_row[5] = str(val_j_s2_sum)  # Kullanılabilir Stok (Depo 100)

            # Sayfa 3 eşleşmesi ve toplama
            s3_matches = df3[df3[self.COMMON_MATCH_COL["G"]] == match_val]
            if not s3_matches.empty:
                current_data_row[6] = str(s3_matches.iloc[0][self.SHEET3_COLS["B"]])  # Depo 110
                # Toplamı al: Sayfa 3 K sütunundaki değerler Kullanılabilir Stok (Depo 110) sütununa
                val_k_s3_sum = s3_matches[self.SHEET3_COLS["K"]].apply(self._to_float_series).sum()
                current_data_row[7] = str(val_k_s3_sum)  # Kullanılabilir Stok (Depo 110)
                # Toplamı al: Sayfa 3 L sütunundaki değerler Kalite Stoğu sütununa
                val_l_s3_sum = s3_matches[self.SHEET3_COLS["L"]].apply(self._to_float_series).sum()
                current_data_row[8] = str(val_l_s3_sum)  # Kalite Stoğu

            current_data_row[9] = ""
            current_data_row[10] = ""
            current_data_row[11] = ""
            current_data_row[12] = ""
            current_data_row[13] = ""

            final_table_content.append(current_data_row)

        # Tablo boyutlarını ayarla
        self.table.setColumnCount(
            len(self.HEADER_LABELS))  # Sütunlar için hala orijinal HEADER_LABELS uzunluğunu kullan
        self.table.setRowCount(len(final_table_content))
        # Başlıkları satır olarak eklediğimiz için varsayılan yatay başlıkları temizle
        self.table.setHorizontalHeaderLabels([""] * len(self.HEADER_LABELS))

        # QTableWidget'ı doldur ve vurgulama uygula
        for r_idx, row_data in enumerate(final_table_content):
            for c_idx, cell_value in enumerate(row_data):
                item = QTableWidgetItem(str(cell_value))
                if r_idx in self.highlighted_rows:
                    item.setBackground(highlight_color_brush)
                    # Başlık satırlarını düzenlenemez yap
                    item.setFlags(item.flags() ^ Qt.ItemIsEditable)
                elif c_idx != 9:  # Veri satırları için, 'İhtiyaç' (indeks 9) hariç tüm hücreleri düzenlenemez yap
                    item.setFlags(item.flags() ^ Qt.ItemIsEditable)
                self.table.setItem(r_idx, c_idx, item)

        # Doldurmadan sonra, 'Durum' ve sipariş miktarlarını hesaplamak için tekrar yinele
        # Bu, 'İhtiyaç' başlangıçta boş olabileceğinden ve hesaplama için D sütununa ihtiyaç duyulduğundan gereklidir
        # Ve sipariş miktarları 'Durum'a bağlıdır
        for r_idx in range(self.table.rowCount()):
            # Bu hesaplamalar için başlık satırlarını atla
            if r_idx in self.highlighted_rows:
                continue
            self._update_l_column(r_idx)
            self._update_order_quantities(r_idx, df4)
            # Veri satırları için "Teslim Tarihi" sütununu (tablo indeks 13) doldur
            malzeme_item = self.table.item(r_idx, 1)  # Eşleşme için Malzeme sütunu
            if malzeme_item:
                malzeme_val = malzeme_item.text()
                teslim_tarihi_val = ""
                s4_delivery_matches = df4[df4[self.SHEET4_COLS["C"]] == malzeme_val]
                if not s4_delivery_matches.empty:
                    raw_date = s4_delivery_matches.iloc[0][self.SHEET4_COLS["S"]]
                    try:
                        formatted_date = pd.to_datetime(raw_date).strftime('%d.%m.%Y')
                        teslim_tarihi_val = formatted_date
                    except (ValueError, TypeError):
                        teslim_tarihi_val = str(raw_date) if pd.notna(raw_date) else ""

                item_teslim_tarihi = self.table.item(r_idx, 13)
                if item_teslim_tarihi is None:
                    item_teslim_tarihi = QTableWidgetItem()
                    item_teslim_tarihi.setFlags(item_teslim_tarihi.flags() ^ Qt.ItemIsEditable)
                    self.table.setItem(r_idx, 13, item_teslim_tarihi)
                item_teslim_tarihi.setText(teslim_tarihi_val)

        # 4) Boyutlandırma + bir kez bağla
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
    #                        Hücre Değişikliği İşleyicileri
    # -------------------------------------------------------------------- #
    def _cell_changed(self, row: int, col: int):
        """Tablo hücrelerindeki değişiklikleri, özellikle 'İhtiyaç' (K) sütunu için işler.
        Girilen değeri D sütunuyla çarparak aynı sütundaki tüm alt hücrelere yayar."""
        # Bir güncelleme zaten devam ediyorsa veya değişen sütun 'K' (indeks 9) değilse geri dön
        # Ayrıca, vurgulanmış bir başlık satırını düzenlemeye çalışmadığımızdan emin ol
        if self._updating or col != 9 or row in self.highlighted_rows:
            return

        try:
            # Değişen hücreden metni al, ondalık dönüşüm için virgülü noktayla değiştir
            k_raw = self.table.item(row, col).text().replace(",", ".")
            k_input_value = float(k_raw)  # Float'a dönüştür
        except (ValueError, AttributeError):
            # Giriş geçerli bir sayı değilse, hücreyi temizle ve L'yi yeniden hesapla
            self._updating = True  # Özyinelemeyi önlemek için güncelleme bayrağını ayarla
            self.table.setItem(row, col, QTableWidgetItem(""))  # Geçersiz girişi temizle
            self._update_l_column(row)  # K=0 ile mevcut satır için L'yi yeniden hesapla
            # K değişirse sipariş miktarlarını da güncelle
            self._update_order_quantities(row, self.excel_data["s4"])
            # Grafik güncelleme artık grafik sayfası açıldığında veya tüm değişiklikler yapıldıktan sonra işlenir
            self._updating = False  # Güncelleme bayrağını sıfırla
            return

        self._updating = True  # Özyinelemeyi önlemek için güncelleme bayrağını ayarla
        # Yeni K değerini (D sütunuyla çarpılarak) değişen hücreye ve aynı sütundaki tüm alt hücrelere uygula
        for r_idx in range(row, self.table.rowCount()):
            # Değişiklikleri yayarken vurgulanmış başlık satırlarını atla
            if r_idx in self.highlighted_rows:
                continue

            # Mevcut satır için D sütunundaki değeri al (indeks 3)
            d_val = self._to_float(self.table.item(r_idx, 3))

            # Kullanıcının girişiyle D sütunu değerini çarparak yeni 'İhtiyaç' değerini hesapla
            calculated_k_value = d_val * k_input_value

            # Mevcut satır için K sütunu öğesini hesaplanan değere ayarla
            self.table.setItem(r_idx, 9, QTableWidgetItem(str(calculated_k_value)))
            # Yeni K değerine göre mevcut satır için L sütununu yeniden hesapla ve güncelle
            self._update_l_column(r_idx)
            # K değişirse sipariş miktarlarını da güncelle
            self._update_order_quantities(r_idx, self.excel_data["s4"])
        # Grafik güncelleme artık grafik sayfası açıldığında veya tüm değişiklikler yapıldıktan sonra işlenir
        self._updating = False  # Güncelleme bayrağını sıfırla

    def _to_float(self, item: QTableWidgetItem) -> float:
        """QTableWidgetItem'ın metnini float'a dönüştürür, virgülleri ve boş dizeleri işler."""
        try:
            if item is None or item.text() == "":
                return 0.0
            return float(item.text().replace(",", "."))
        except (ValueError, AttributeError):
            return 0.0

    def _to_float_series(self, value) -> float:
        """Pandas Serisindeki bir değeri float'a dönüştürür, sayısal olmayan değerleri
        ve virgül ondalık ayırıcılarını işler."""
        try:
            if isinstance(value, str):
                # Binlik ayırıcıları (varsayılıyorsa) kaldırır ve virgül ondalık ayırıcısını nokta ile değiştirir
                value = value.replace(".", "").replace(",", ".")
            return float(value)
        except (ValueError, TypeError):
            return 0.0

    def _update_l_column(self, row: int):
        """Belirli bir satır için F, I, J ve K sütunlarındaki değerlere göre 'Durum' (L) sütununu hesaplar ve günceller."""
        # İlgili sütunlardan değerleri alır, float'a dönüştürür
        f_val = self._to_float(self.table.item(row, 5))  # Sütun F (Sayfa 2 J değeri)
        i_val = self._to_float(self.table.item(row, 7))  # Sütun I (Sayfa 3 J değeri)
        j_val = self._to_float(self.table.item(row, 8))  # Sütun J (Sayfa 3 K değeri)
        k_val = self._to_float(self.table.item(row, 9))  # Sütun K (İhtiyaç)

        # 'Durum' (L) için sonucu hesaplar
        result = f_val + i_val + j_val - k_val
        # Metni biçimlendirir: eğer sonuç negatifse, "#SİPARİŞ VER" ekler
        text = f"{result} #SİPARİŞ VER" if result < 0 else str(result)

        # L sütunu için QTableWidgetItem'ı alır veya oluşturur
        item = self.table.item(row, 10)
        if item is None:
            item = QTableWidgetItem()
            # L sütununu hesaplanmış bir alan olduğu için düzenlenemez yapar
            item.setFlags(item.flags() ^ Qt.ItemIsEditable)
            self.table.setItem(row, 10, item)
        item.setText(text)  # Hesaplanan metni ayarlar

    def _update_order_quantities(self, row: int, df4: pd.DataFrame):
        """
        'Durum' sütunu ve 4. Excel sayfasına göre belirli bir satır için 'Verilen Sipariş Miktarı' ve
        'Verilmesi Gereken Sipariş Miktarı'nı hesaplar ve günceller.
        """
        durum_item = self.table.item(row, 10)  # 'Durum' sütunu
        malzeme_item = self.table.item(row, 1)  # 'Malzeme' sütunu

        # Initialize to 0.0
        verilen_siparis_miktari = 0.0
        verilmesi_gereken_siparis_miktari = 0.0
        durum_numeric_val = 0.0

        # Calculate "Verilen Sipariş Miktarı" based on sheet 4, column I (index 8)
        if malzeme_item:
            malzeme_val = malzeme_item.text()
            # Match current row's 'Malzeme' (column 1) with SHEET4_COLS["C"] (index 2)
            # and sum SHEET4_COLS["I"] (index 8)
            s4_matches = df4[df4[self.SHEET4_COLS["C"]] == malzeme_val]
            if not s4_matches.empty:
                # Sum the values in the 'I' column (index 8) from sheet 4
                verilen_siparis_miktari = s4_matches[self.SHEET4_COLS["I"]].apply(self._to_float_series).sum()

        # Extract numeric value from 'Durum' column
        if durum_item and "#SİPARİŞ VER" in durum_item.text():
            try:
                # Extract the numeric part of the 'Durum' value
                durum_numeric_str = durum_item.text().split(" #SİPARİŞ VER")[0].replace(",", ".")
                durum_numeric_val = float(durum_numeric_str)
            except (ValueError, AttributeError):
                durum_numeric_val = 0.0

        # Calculate "Verilmesi Gereken Sipariş Miktarı"
        # It should be (Value from 'Durum' column) - (Value from 'Verilen Sipariş Miktarı' column)
        # Taking into account that durum_numeric_val is already (F+I+J-K) from _update_l_column
        # And if it has #SİPARİŞ VER, it means it's negative.
        # So, the "remaining_needed" is the absolute value of durum_numeric_val, minus verilen_siparis_miktari
        # If durum_numeric_val is -100 and verilen_siparis_miktari is 20, then remaining needed is 100 - 20 = 80
        # If durum_numeric_val is -100 and verilen_siparis_miktari is 120, then remaining needed is 0 (already fulfilled and more)
        if durum_numeric_val < 0:  # Only if '#SİPARİŞ VER' is present, meaning it's a negative stock
            # The absolute need is -durum_numeric_val
            net_need = abs(durum_numeric_val) - verilen_siparis_miktari
            if net_need > 0:
                verilmesi_gereken_siparis_miktari = net_need
            else:
                verilmesi_gereken_siparis_miktari = 0.0
        else:  # If Durum is not negative (no #SİPARİŞ VER), then no order is needed
            verilmesi_gereken_siparis_miktari = 0.0

        # Set items for "Verilen Sipariş Miktarı" (index 11)
        item_verilen = self.table.item(row, 11)
        if item_verilen is None:
            item_verilen = QTableWidgetItem()
            item_verilen.setFlags(item_verilen.flags() ^ Qt.ItemIsEditable)  # Make uneditable
            self.table.setItem(row, 11, item_verilen)
        item_verilen.setText(str(verilen_siparis_miktari))

        # Set items for "Verilmesi Gereken Sipariş Miktarı" (index 12)
        item_gereken = self.table.item(row, 12)
        if item_gereken is None:
            item_gereken = QTableWidgetItem()
            item_gereken.setFlags(item_gereken.flags() ^ Qt.ItemIsEditable)  # Make uneditable
            self.table.setItem(row, 12, item_gereken)
        item_gereken.setText(str(verilmesi_gereken_siparis_miktari))

    def _process_fsnkp_rows(self):
        """
        'FSNKP' girişlerini kaldırmak ve önceki satırın 'Durum' sütununu güncellemek için satırları işler.
        Satır silme işlemini doğru şekilde ele almak için geriye doğru yineler.
        """
        self._updating = True

        rows_to_remove = []
        # Satır kaldırma işlemini doğru şekilde ele almak için geriye doğru yinele
        for r_idx in range(self.table.rowCount() - 1, 0, -1):  # Sondan ikinci satırdan başla, 1. satıra kadar git
            # Metne erişmeden önce öğelerin mevcut olduğundan emin ol
            current_malzeme_item = self.table.item(r_idx, 1)
            prev_malzeme_item = self.table.item(r_idx - 1, 1)
            current_aciklama_item = self.table.item(r_idx, 2)

            current_malzeme = current_malzeme_item.text() if current_malzeme_item else ""
            prev_malzeme = prev_malzeme_item.text() if prev_malzeme_item else ""
            current_aciklama = current_aciklama_item.text() if current_aciklama_item else ""

            # Mevcut satırın 'Malzeme' (sütun 1) önceki satırın 'Malzeme'siyle eşleşiyor mu kontrol et
            # ve mevcut satırın 'Açıklama'sı (sütun 2) "FSNKP" içeriyor mu kontrol et
            # Ayrıca mevcut satırın kendisinin bir başlık satırı OLMADIĞINDAN emin ol (sütun 1'in "Malzeme" olup olmadığını kontrol et)
            if current_malzeme == prev_malzeme and "FSNKP" in current_aciklama and current_malzeme != "Malzeme":
                # Önceki satırın 'Durum' sütununa (sütun 10) "#FSNKP" ekle
                prev_durum_item = self.table.item(r_idx - 1, 10)
                if prev_durum_item:
                    current_durum_text = prev_durum_item.text()
                    if "#FSNKP" not in current_durum_text:  # Yinelenen "#FSNKP" eklemeyi önle
                        prev_durum_item.setText(current_durum_text + " #FSNKP")

                # Mevcut satırı kaldırmak için işaretle
                rows_to_remove.append(r_idx)

        # Kaldırılacak satırları kaldır (dizin kaydırma sorunlarını önlemek için en yüksek dizinden en düşüğe doğru)
        for r_idx in sorted(rows_to_remove, reverse=True):
            self.table.removeRow(r_idx)

        # Tüm FSNKP işleme ve satır kaldırma işlemlerinden sonra, vurgulanan satırları yeniden belirle
        self.highlighted_rows = []
        for r_idx in range(self.table.rowCount()):
            # Bir satır, ikinci sütunu (indeks 1) "Malzeme" ise bir blok başlığıdır
            malzeme_header_item = self.table.item(r_idx, 1)
            if malzeme_header_item and malzeme_header_item.text() == "Malzeme":
                self.highlighted_rows.append(r_idx)

        self._updating = False

    def _update_completion_chart(self):
        """
        'Durum' sütunundaki hücreleri sayar ve tamamlanma durumunu gösteren bir pasta grafiği oluşturur.
        Ayrıca, en geç teslim tarihini ve Ü.Ağacı Sev değerini grafiğe ekler.
        """
        completed_count = 0
        incomplete_count = 0
        total_rows = self.table.rowCount()
        latest_delivery_date = None
        u_agaci_sev_value = ""

        # Grafik başlığı için en üst blok başlığının A sütunundaki değeri al
        if self.highlighted_rows:
            first_header_row_idx = self.highlighted_rows[0]
            u_agaci_sev_item = self.table.item(first_header_row_idx, 0)
            if u_agaci_sev_item:
                u_agaci_sev_value = u_agaci_sev_item.text()

        for r_idx in range(total_rows):
            # Tamamlanma durumunu hesaplarken başlık satırlarını atla
            if r_idx in self.highlighted_rows:
                continue

            durum_item = self.table.item(r_idx, 10)  # 'Durum' sütunu (indeks 10)
            if durum_item and "#SİPARİŞ VER" in durum_item.text():
                incomplete_count += 1
            else:
                completed_count += 1

            # En geç teslim tarihini bul
            teslim_tarihi_item = self.table.item(r_idx, 13)  # 'Teslim Tarihi' sütunu (indeks 13)
            if teslim_tarihi_item:
                date_str = teslim_tarihi_item.text()
                try:
                    # GG.AA.YYYY formatını ayrıştır
                    current_date = datetime.datetime.strptime(date_str, '%d.%m.%Y').date()
                    if latest_delivery_date is None or current_date > latest_delivery_date:
                        latest_delivery_date = current_date
                except ValueError:
                    pass  # Geçersiz tarih formatlarını yoksay

        # Kapsayıcıdan mevcut grafiği temizle
        for i in reversed(range(self.chart_container.layout().count())):
            widget_to_remove = self.chart_container.layout().itemAt(i).widget()
            if widget_to_remove:
                widget_to_remove.setParent(None)

        # Yüzde hesaplaması için toplam veri satırlarını (başlıklar hariç) hesapla
        total_data_rows = total_rows - len(self.highlighted_rows)
        if total_data_rows == 0:
            no_data_label = QLabel("Grafik için veri yok.", alignment=Qt.AlignCenter)
            self.chart_container.layout().addWidget(no_data_label)
            self.chart_figure = None  # Veri yoksa figürü temizle
            return

        completed_percentage = (completed_count / total_data_rows) * 100
        incomplete_percentage = (incomplete_count / total_data_rows) * 100

        # Sıfır yüzdeleri zarifçe işlemek için özel bir autopct fonksiyonu tanımla
        def autopct_format(pct):
            return ('%1.1f%%' % pct) if pct > 0 else ''

        # Bir Matplotlib figürü ve eksenleri oluştur
        self.chart_figure = Figure(figsize=(7, 4.6), dpi=100)  # Figür boyutunu 700x460 piksel olarak ayarla
        ax = self.chart_figure.add_subplot(111)

        # Güncellenmiş renkler: Tamamlanan için Mavi, Tamamlanmayan için Turuncu
        labels = ['Tamamlandı ({:.1f}%)'.format(completed_percentage),
                  'Tamamlanmadı ({:.1f}%)'.format(incomplete_percentage)]
        sizes = [completed_percentage, incomplete_percentage]
        colors = ['#1f77b4', '#ff7f0e']  # Mavi ve Turuncu
        explode = (0.05, 0)  # 'Tamamlandı' dilimini hafifçe ayır

        ax.pie(sizes, explode=explode, labels=labels, colors=colors,
               autopct=autopct_format, shadow=True, startangle=90,
               textprops={'fontsize': 10, 'color': 'black', 'fontweight': 'bold'})  # Etiketler için metin özellikleri
        ax.axis('equal')  # Eşit en boy oranı, pastanın bir daire olarak çizilmesini sağlar.

        # Grafik başlığını Ü.Ağacı Sev değeriyle ayarla
        chart_title_text = f"{u_agaci_sev_value} İş Tamamlanma Durumu"
        ax.set_title(chart_title_text, fontsize=16, color='#2c3e50', fontweight='bold')

        # En geç teslim tarihini grafiğe ekle - başlığın altında sağ alt köşeye konumlandırıldı
        if latest_delivery_date:
            date_text = f"En Geç Teslim Tarihi: {latest_delivery_date.strftime('%d.%m.%Y')}"
            # Tarih metnini figüre göre sağ alt köşeye konumlandır, eksenlere göre değil
            # Y koordinatı hafifçe yukarı ayarlandı (0.02 yerine 0.05) böylece en alt kenarla çakışma olmaz.
            self.chart_figure.text(0.98, 0.05, date_text, transform=self.chart_figure.transFigure,
                                   fontsize=10, color='#555555', ha='right', va='bottom', fontweight='bold')

        # Etiketlerin/başlığın çakışmasını önlemek için düzeni ayarla
        self.chart_figure.tight_layout()

        # Matplotlib figürünü bir PyQt widget'ına göm
        # Tuvalin düzeni içinde genişlemesini sağlamak için stretch=1 eklendi
        canvas = FigureCanvas(self.chart_figure)
        self.chart_container.layout().addWidget(canvas, stretch=1)
        canvas.draw()

    def _save_chart_as_image(self):
        """Oluşturulan pasta grafiğini bir JPEG/PNG görüntüsü olarak kaydeder."""
        if self.chart_figure is None:
            QMessageBox.warning(self, "Uyarı", "Kaydedilecek bir grafik bulunamadı. Lütfen önce tabloyu görüntüleyin.")
            return

        # Kullanıcıdan kaydetme dosya yolunu al
        path, _ = QFileDialog.getSaveFileName(self, "Grafiği Kaydet", "iş_tamamlanma_grafiği.png",
                                              "Görüntü Dosyaları (*.png *.jpg *.jpeg)")
        if not path:
            return

        try:
            # Figürü belirtilen boyutlarla kaydet
            self.chart_figure.savefig(path, dpi=100)  # dpi=100 ve figsize=(7, 4.6) 700x460 piksel verir
            QMessageBox.information(self, "Başarılı", f"Grafik kaydedildi: {path.split('/')[-1]}")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Grafik kaydedilirken bir hata oluştu:\n{e}")

    # -------------------------------------------------------------------- #
    #                           Excel'e Kaydet
    # -------------------------------------------------------------------- #
    def _save_excel(self):
        """Geçerli verileri QTableWidget'tan yeni bir Excel dosyasına kaydeder ve belirginleştirme uygular."""
        # Bir kaydetme dosyası iletişim kutusu açar
        path, _ = QFileDialog.getSaveFileName(self, "Uyarlanmış Excel Dosyasını Kaydet", "uyarlanmis_veri.xlsx",
                                              "Excel Dosyaları (*.xlsx)")
        if not path:  # Eğer yol seçilmezse, geri döner
            return

        rows, cols = self.table.rowCount(), self.table.columnCount()
        data_to_save = []
        for r in range(rows):  # QTableWidget'taki tüm satırları döngüye al
            row_data = []
            for c in range(cols):
                item = self.table.item(r, c)
                row_data.append(item.text() if item else "")
            data_to_save.append(row_data)

        # Başlıklar artık data_to_save'in bir parçası olduğu için açık başlıklar olmadan DataFrame oluştur
        df_to_save = pd.DataFrame(data_to_save)

        try:
            # ExcelWriter kullanarak belirginleştirme için xlsxwriter motorunu kullan
            writer = pd.ExcelWriter(path, engine='xlsxwriter')
            # DataFrame'i Excel'e yaz, indeksleri ve başlıkları dahil et
            df_to_save.to_excel(writer, sheet_name='Uyarlanmış Veri', index=False, header=False)

            workbook = writer.book
            worksheet = writer.sheets['Uyarlanmış Veri']

            # Belirginleştirme için formatı tanımla (açık kırmızı)
            highlight_format = workbook.add_format({'bg_color': '#FFCCCC'})
            # Başlık satırları için kalın font formatı
            header_font_format = workbook.add_format({'bold': True, 'bg_color': '#FFCCCC'})

            # Belirginleştirilecek satırlara formatı uygula
            # Excel satırları 1-indeksli olduğu için r_idx + 1 kullanılır.
            for r_idx in self.highlighted_rows:
                # set_row, xlsxwriter için 0-indeksli satır numarasını alır
                # Başlık satırlarına başlık yazı tipi formatını uygula
                worksheet.set_row(r_idx, None, header_font_format)

            # Diğer vurgulanan satırlara genel vurgulama formatını uygula (varsa, ancak bu mantıkta sadece başlıklar)
            # Bu döngü, highlighted_rows yalnızca başlık satırlarını içeriyorsa teknik olarak gereksizdir
            # Ancak mantık daha sonra değişirse sağlamlık için tutuldu.
            for r_idx in self.highlighted_rows:
                if r_idx not in self.highlighted_rows:  # Bu koşul her zaman yanlış olacaktır
                    worksheet.set_row(r_idx, None, highlight_format)

            writer.close()
            QMessageBox.information(self, "Başarılı", f"Dosya kaydedildi: {path.split('/')[-1]}")
        except Exception as e:
            QMessageBox.critical(self, "Hata",
                                 f"Dosya kaydedilirken ve belirginleştirme uygulanırken bir hata oluştu:\\n{e}")


# ----------------------------------------------------------------------- #
#                                   main
# ----------------------------------------------------------------------- #
if __name__ == "__main__":
    app = QApplication(sys.argv)  # QApplication örneğini oluşturur
    window = ExcelProcessorApp()  # Ana uygulama penceresinin bir örneğini oluşturur
    window.show()  # Pencereyi gösterir
    sys.exit(app.exec_())  # Uygulama olay döngüsünü başlat
