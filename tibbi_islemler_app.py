# -*- coding: utf-8 -*-
"""
Tıbbi İşlemler Listesi - Gelişmiş Arama Uygulaması
Excel dosyasındaki tıbbi işlemleri arayan, filtreleyen ve analiz eden masaüstü uygulaması.
"""
import os
import sys
import csv
import json
import tkinter as tk
import urllib.request
import urllib.error
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from collections import Counter
import openpyxl

APP_VERSION = "2.3"
GITHUB_REPO = "erdiege01/tibbi-islemler-app"
VERSION_URL = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/version.json"

# Opsiyonel kütüphaneler
try:
    import customtkinter as ctk
    USE_CTK = True
except ImportError:
    USE_CTK = False

try:
    from PIL import Image, ImageTk
    USE_PIL = True
except ImportError:
    USE_PIL = False

try:
    import matplotlib
    matplotlib.use('Agg')
    from matplotlib import pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    USE_MPL = True
except ImportError:
    USE_MPL = False

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    USE_DND = True
except ImportError:
    USE_DND = False


EXCEL_PATH = r"C:\Users\m.atacan\Downloads\4-genel-tibbi-islemler-listesi-excellxlsx.xlsx"
CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".tibbi_islemler_config.json")


def resource_path(relative_path):
    """PyInstaller ile paketlenmiş dosyalara erişim."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def tr_lower(text):
    """Türkçe uyumlu küçük harfe çevirme (büyük/küçük harf duyarsız arama için).

    Python'un varsayılan lower() metodu İ/I harflerini Türkçe kurallara göre
    çevirmez (örn. 'İSTANBUL' -> 'i̇stanbul'). Bu fonksiyon önce İ->i ve
    I->ı dönüşümlerini yapar, sonra lower() uygular.
    """
    return str(text).replace("İ", "i").replace("I", "ı").lower()


class Config:
    """Uygulama ayarlarını yönetir."""

    def __init__(self):
        self.data = {
            "theme": "light",
            "window_size": "1400x800",
            "column_widths": {},
            "search_history": [],
            "favorites": [],
            "last_file": EXCEL_PATH,
        }
        self.load()

    def load(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                    self.data.update(saved)
            except Exception:
                pass

    def save(self):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def get(self, key, default=None):
        return self.data.get(key, default)

    def set(self, key, value):
        self.data[key] = value
        self.save()


class TibbiIslemlerApp:
    def __init__(self, root):
        self.root = root
        self.config = Config()

        # Tema
        self.is_dark = self.config.get("theme") == "dark"
        self.bg_color = "#1e1e1e" if self.is_dark else "#f0f0f0"
        self.fg_color = "#ffffff" if self.is_dark else "#000000"
        self.accent_color = "#0078d4" if not self.is_dark else "#0098ff"

        self.root.title(f"Tıbbi İşlemler Listesi - Gelişmiş Arama v{APP_VERSION}")
        self.root.geometry(self.config.get("window_size", "1400x800"))
        self.root.minsize(1000, 600)
        try:
            self.root.state("zoomed")  # Program her zaman tam ekran açılsın
        except Exception:
            pass

        # Veri
        self.all_rows = []
        self.filtered_rows = []
        self.search_history = self.config.get("search_history", [])
        self.favorites = self.config.get("favorites", [])
        self.sort_column = None
        self.sort_reverse = False

        # Filtre durumu
        self.filter_puan_min = None
        self.filter_puan_max = None
        self.filter_grup = None

        self._build_ui()
        self._load_data()

        # Pencere kapanırken ayarları kaydet
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self):
        """Kullanıcı arayüzünü oluşturur."""
        # Ana container (grid yerleşim: tablo sütunu esner, tüm genişliği kaplar)
        self.main_frame = tk.Frame(self.root, bg=self.bg_color)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=0)
        self.main_frame.grid_columnconfigure(1, weight=1)

        # Üst toolbar
        self._build_toolbar()

        # Sol panel - Filtreler ve istatistikler
        self._build_side_panel()

        # Sağ panel - Tablo
        self._build_table_panel()

        # Alt panel - Durum çubuğu
        self._build_status_bar()

        # Menü çubuğu
        self._build_menu()

        # Sürükle-bırak desteği
        if USE_DND:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self._on_drop)

        # Kısayollar
        self.root.bind("<Control-f>", lambda e: self.search_entry.focus_set())
        self.root.bind("<Control-e>", lambda e: self._export_results())
        self.root.bind("<Control-r>", lambda e: self._load_data())
        self.root.bind("<Control-o>", lambda e: self._browse_file())
        self.root.bind("<Escape>", lambda e: self._clear_search())
        self.root.bind("<F5>", lambda e: self._load_data())

    def _build_toolbar(self):
        """Üst araç çubuğunu oluşturur (sadece arama)."""
        self.file_var = tk.StringVar(value=self.config.get("last_file", EXCEL_PATH))

        toolbar = tk.Frame(self.main_frame, bg=self.bg_color, padx=10, pady=5)
        toolbar.grid(row=0, column=0, columnspan=2, sticky="ew")

        # Tema değiştir (önce yerleştirilir ki dar ekranda kesilmesin)
        theme_text = "Karanlık" if not self.is_dark else "Aydınlık"
        tk.Button(toolbar, text=theme_text, command=self._toggle_theme,
                  bg=self.accent_color, fg="white").pack(side=tk.RIGHT, padx=5)

        # Arama (tam genişlik)
        tk.Label(toolbar, text="Ara:", bg=self.bg_color, fg=self.fg_color).pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_search_change)
        self.search_entry = tk.Entry(toolbar, textvariable=self.search_var, font=("Segoe UI", 11),
                                    bg=self.bg_color, fg=self.fg_color, insertbackground=self.fg_color)
        self.search_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # Arama temizleme butonu
        tk.Button(toolbar, text="Temizle", command=self._clear_search_text,
                  bg=self.accent_color, fg="white").pack(side=tk.LEFT, padx=2)

        # Arama geçmişi dropdown
        self.history_var = tk.StringVar()
        self.history_combo = ttk.Combobox(toolbar, textvariable=self.history_var,
                                          values=self.search_history[:20], width=25)
        self.history_combo.pack(side=tk.LEFT, padx=2)
        self.history_combo.bind("<<ComboboxSelected>>", self._on_history_select)

    def _build_side_panel(self):
        """Sol paneli oluşturur (filtreler, istatistikler, sık kullanılanlar)."""
        self.side_frame = tk.Frame(self.main_frame, bg=self.bg_color, width=300)
        self.side_frame.grid(row=1, column=0, sticky="ns", padx=(10, 0), pady=10)
        self.side_frame.pack_propagate(False)

        # Filtreler
        filter_frame = tk.LabelFrame(self.side_frame, text="Filtreler", bg=self.bg_color,
                                     fg=self.fg_color, padx=10, pady=10)
        filter_frame.pack(fill=tk.X, pady=(0, 10))

        # Puan filtresi
        tk.Label(filter_frame, text="Puan Aralığı:", bg=self.bg_color, fg=self.fg_color).pack(anchor=tk.W)
        puan_frame = tk.Frame(filter_frame, bg=self.bg_color)
        puan_frame.pack(fill=tk.X, pady=2)
        self.puan_min_var = tk.StringVar()
        self.puan_max_var = tk.StringVar()
        tk.Entry(puan_frame, textvariable=self.puan_min_var, width=8,
                 bg=self.bg_color, fg=self.fg_color).pack(side=tk.LEFT)
        tk.Label(puan_frame, text=" - ", bg=self.bg_color, fg=self.fg_color).pack(side=tk.LEFT)
        tk.Entry(puan_frame, textvariable=self.puan_max_var, width=8,
                 bg=self.bg_color, fg=self.fg_color).pack(side=tk.LEFT)
        tk.Button(puan_frame, text="Uygula", command=self._apply_puan_filter,
                  bg=self.accent_color, fg="white").pack(side=tk.LEFT, padx=5)

        # Ameliyat grubu filtresi
        tk.Label(filter_frame, text="Ameliyat Grubu:", bg=self.bg_color, fg=self.fg_color).pack(anchor=tk.W, pady=(10, 0))
        self.grup_var = tk.StringVar(value="Tümü")
        self.grup_combo = ttk.Combobox(filter_frame, textvariable=self.grup_var, state="readonly", width=25)
        self.grup_combo.pack(fill=tk.X, pady=2)
        self.grup_combo.bind("<<ComboboxSelected>>", lambda e: self._apply_filter())

        # Filtre temizle
        tk.Button(filter_frame, text="Filtreleri Temizle", command=self._clear_filters,
                  bg=self.accent_color, fg="white").pack(fill=tk.X, pady=(10, 0))

        # İstatistikler
        stats_frame = tk.LabelFrame(self.side_frame, text="İstatistikler", bg=self.bg_color,
                                    fg=self.fg_color, padx=10, pady=10)
        stats_frame.pack(fill=tk.X, pady=(0, 10))

        self.stats_labels = {}
        stats_items = [
            ("total", "Toplam Kayıt:"),
            ("avg_puan", "Ortalama Puan:"),
            ("max_puan", "En Yüksek Puan:"),
            ("min_puan", "En Düşük Puan:"),
            ("unique_grup", "Ameliyat Grubu Sayısı:"),
        ]
        for key, label in stats_items:
            frame = tk.Frame(stats_frame, bg=self.bg_color)
            frame.pack(fill=tk.X, pady=1)
            tk.Label(frame, text=label, bg=self.bg_color, fg=self.fg_color).pack(side=tk.LEFT)
            self.stats_labels[key] = tk.Label(frame, text="-", bg=self.bg_color, fg=self.accent_color)
            self.stats_labels[key].pack(side=tk.RIGHT)

        # Sık kullanılar
        fav_frame = tk.LabelFrame(self.side_frame, text="Sık Kullanılanlar", bg=self.bg_color,
                                  fg=self.fg_color, padx=10, pady=10)
        fav_frame.pack(fill=tk.BOTH, expand=True)

        self.fav_listbox = tk.Listbox(fav_frame, bg=self.bg_color, fg=self.fg_color,
                                      selectmode=tk.SINGLE, height=10)
        self.fav_listbox.pack(fill=tk.BOTH, expand=True)
        self.fav_listbox.bind("<Double-Button-1>", self._on_fav_select)

        fav_btn_frame = tk.Frame(fav_frame, bg=self.bg_color)
        fav_btn_frame.pack(fill=tk.X, pady=5)
        tk.Button(fav_btn_frame, text="Ekle", command=self._add_favorite,
                  bg=self.accent_color, fg="white").pack(side=tk.LEFT, padx=2)
        tk.Button(fav_btn_frame, text="Sil", command=self._remove_favorite,
                  bg=self.accent_color, fg="white").pack(side=tk.LEFT, padx=2)

        # Grafik butonu
        if USE_MPL:
            tk.Button(self.side_frame, text="Grafikleri Göster", command=self._show_charts,
                      bg=self.accent_color, fg="white").pack(fill=tk.X, pady=(10, 0))

    def _build_table_panel(self):
        """Sağ paneli oluşturur (tablo)."""
        table_frame = tk.Frame(self.main_frame, bg=self.bg_color)
        table_frame.grid(row=1, column=1, sticky="nsew", padx=10, pady=10)

        # Tablo
        columns = ("kod", "ad", "aciklama", "puan", "grup")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="extended")

        self.tree.heading("kod", text="İşlem Kodu", command=lambda: self._sort_by("kod"))
        self.tree.heading("ad", text="İşlem Adı", command=lambda: self._sort_by("ad"))
        self.tree.heading("aciklama", text="Açıklama", command=lambda: self._sort_by("aciklama"))
        self.tree.heading("puan", text="İşlem Puanı", command=lambda: self._sort_by("puan"))
        self.tree.heading("grup", text="Ameliyat Grupları", command=lambda: self._sort_by("grup"))

        # Sütun genişlikleri
        default_widths = {"kod": 100, "ad": 250, "aciklama": 400, "puan": 90, "grup": 150}
        saved_widths = self.config.get("column_widths", {})
        for col in columns:
            width = saved_widths.get(col, default_widths.get(col, 100))
            self.tree.column(col, width=width, minwidth=50, stretch=True)

        # Scrollbar
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # Pencere genişleyince sütunlar otomatik yayılsın (sağda boşluk kalmasın)
        self._fit_guard = False
        self.tree.bind("<Configure>", self._on_tree_configure)

        # Sağ tık menüsü
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Kopyala", command=self._copy_selection)
        self.context_menu.add_command(label="Sık kullanılanlara ekle", command=self._add_selected_to_favorites)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="CSV olarak dışa aktar", command=self._export_results)
        self.tree.bind("<Button-3>", self._show_context_menu)
        self.tree.bind("<ButtonRelease-1>", self._on_tree_click)

        # Sonuç sayısı
        self.count_label = tk.Label(table_frame, text="", bg=self.bg_color, fg=self.fg_color)
        self.count_label.grid(row=2, column=0, sticky=tk.W)

    def _on_tree_click(self, event):
        """Açıklama hücresine tıklanınca detay penceresi açar."""
        if self.tree.identify("region", event.x, event.y) != "cell":
            return
        if self.tree.identify_column(event.x) != "#3":
            return
        item = self.tree.identify_row(event.y)
        if not item:
            return
        self.tree.selection_set(item)
        self._show_aciklama_detail(self.tree.item(item, "values"))

    def _show_aciklama_detail(self, values):
        """Üstte ameliyat kodu+adı, altta açıklamanın tamamını gösterir."""
        # Zaten açık detay penceresi varsa kapat (üst üste binmesin)
        if getattr(self, "_detail_win", None) is not None:
            try:
                if self._detail_win.winfo_exists():
                    self._detail_win.destroy()
            except Exception:
                pass
            self._detail_win = None

        vals = list(values) + [""] * 5
        kod = vals[0] if vals[0] not in (None, "None", "") else "-"
        ad = vals[1] if vals[1] not in (None, "None", "") else "-"
        aciklama = vals[2] if vals[2] not in (None, "None", "") else "-"

        win = tk.Toplevel(self.root)
        self._detail_win = win
        win.title("Açıklama Detayı")
        win.transient(self.root)
        W, H = 600, 400
        win.geometry(f"{W}x{H}")
        win.update_idletasks()
        try:
            rx, ry = self.root.winfo_x(), self.root.winfo_y()
            rw, rh = self.root.winfo_width(), self.root.winfo_height()
            x = max(rx + (rw - W) // 2, 0)
            y = max(ry + (rh - H) // 2, 0)
            win.geometry(f"{W}x{H}+{x}+{y}")
        except Exception:
            pass
        win.configure(bg=self.bg_color)

        header = tk.Label(win, text=f"{kod} - {ad}", bg=self.bg_color, fg=self.fg_color,
                          font=("Segoe UI", 12, "bold"), anchor=tk.W, justify=tk.LEFT)
        header.pack(fill=tk.X, padx=15, pady=(15, 10))
        # Başlık pencere genişliğine göre kaydırsın
        win.bind("<Configure>", lambda e: header.config(wraplength=max(win.winfo_width() - 40, 200)))

        text_frame = tk.Frame(win, bg=self.bg_color)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=15)
        text = tk.Text(text_frame, wrap=tk.WORD, bg=self.bg_color, fg=self.fg_color,
                       insertbackground=self.fg_color, relief=tk.FLAT)
        sb = ttk.Scrollbar(text_frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=sb.set)
        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        text.insert("1.0", str(aciklama))
        text.configure(state=tk.DISABLED)

        tk.Button(win, text="Kapat", command=win.destroy,
                  bg=self.accent_color, fg="white").pack(pady=10)

    def _on_tree_configure(self, event):
        """Tablo genişleyince sütunları orantılı olarak yayar, sağda boşluk bırakmaz."""
        if self._fit_guard:
            return
        cols = ("kod", "ad", "aciklama", "puan", "grup")
        try:
            widths = [self.tree.column(c, "width") for c in cols]
        except Exception:
            return
        total = sum(widths)
        avail = event.width
        if avail <= 1 or total >= avail:
            return
        self._fit_guard = True
        try:
            extra = avail - total
            for c, w in zip(cols, widths):
                self.tree.column(c, width=int(w + extra * w / total))
        finally:
            self._fit_guard = False

    def _build_status_bar(self):
        """Durum çubuğunu oluşturur."""
        status_frame = tk.Frame(self.main_frame, bg=self.bg_color)
        status_frame.grid(row=3, column=0, columnspan=2, sticky="ew")

        self.status_var = tk.StringVar(value="Hazır")
        tk.Label(status_frame, textvariable=self.status_var, bg=self.bg_color,
                 fg=self.fg_color, anchor=tk.W).pack(side=tk.LEFT, padx=10)

        self.time_var = tk.StringVar()
        tk.Label(status_frame, textvariable=self.time_var, bg=self.bg_color,
                 fg=self.fg_color).pack(side=tk.RIGHT, padx=10)
        self._update_time()

        # Aktif dosya bilgisi (ince ikinci satır)
        file_bar = tk.Frame(self.main_frame, bg=self.bg_color)
        file_bar.grid(row=2, column=0, columnspan=2, sticky="ew")
        tk.Label(file_bar, text="Aktif Dosya:", bg=self.bg_color,
                 fg=self.fg_color).pack(side=tk.LEFT, padx=(10, 2))
        self.file_label_var = tk.StringVar(value=self.file_var.get())
        tk.Label(file_bar, textvariable=self.file_label_var, bg=self.bg_color,
                 fg=self.accent_color, anchor=tk.W).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(file_bar, text="Dosya Değiştir", command=self._browse_file,
                  bg=self.accent_color, fg="white").pack(side=tk.RIGHT, padx=10, pady=2)

    def _build_menu(self):
        """Menü çubuğunu oluşturur."""
        menubar = tk.Menu(self.root)

        # Dosya menüsü (dosya ekleme/değiştirme buradan yapılır)
        self.file_menu = tk.Menu(menubar, tearoff=0)
        self.file_menu.add_command(label="Dosya Seç / Değiştir...", command=self._browse_file, accelerator="Ctrl+O")
        self.file_menu.add_command(label="Yenile", command=self._load_data, accelerator="F5")
        self.file_menu.add_command(label="Mevcut Dosya Bilgisi", command=self._show_file_info)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="CSV Olarak Dışa Aktar", command=self._export_results, accelerator="Ctrl+E")
        self.file_menu.add_command(label="Excel Olarak Dışa Aktar", command=self._export_excel)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Güncelle (sürümü denetle)", command=self._check_for_updates)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Çıkış", command=self._on_close)
        menubar.add_cascade(label="Dosya", menu=self.file_menu)

        # Görünüm menüsü
        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label="Tema Değiştir", command=self._toggle_theme)
        view_menu.add_command(label="Sütun Genişliklerini Sıfırla", command=self._reset_column_widths)
        menubar.add_cascade(label="Görünüm", menu=view_menu)

        # Yardım menüsü
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Kısayollar", command=self._show_shortcuts)
        help_menu.add_command(label="Hakkında", command=self._show_about)
        menubar.add_cascade(label="Yardım", menu=help_menu)

        self.root.config(menu=menubar)

    def _update_time(self):
        """Saat güncellemesi."""
        self.time_var.set(datetime.now().strftime("%H:%M:%S"))
        self.root.after(1000, self._update_time)

    def _update_file_label(self):
        """Alt bardaki aktif dosya yolunu günceller."""
        if hasattr(self, "file_label_var"):
            self.file_label_var.set(self.file_var.get())

    def _show_file_info(self):
        """Mevcut dosya bilgisini gösterir."""
        path = self.file_var.get()
        exists = os.path.exists(path) if path else False
        messagebox.showinfo("Mevcut Dosya", f"Dosya:\n{path}\n\nDurum: {'Mevcut' if exists else 'Bulunamadı'}")

    @staticmethod
    def _parse_version(v):
        """'2.2' -> (2, 2) gibi karşılaştırılabilir hale getirir."""
        parts = []
        for p in str(v).strip().lstrip("v").split("."):
            try:
                parts.append(int(p))
            except ValueError:
                parts.append(0)
        return tuple(parts)

    def _check_for_updates(self):
        """GitHub'daki sürümü denetler, yenisi varsa indirip kurar."""
        if "KULLANICI" in GITHUB_REPO:
            messagebox.showinfo("Güncelle", "Güncelleme adresi henüz ayarlanmamış.")
            return
        self.status_var.set("Sürüm denetleniyor...")
        self.root.update()
        try:
            req = urllib.request.Request(VERSION_URL, headers={"User-Agent": "TibbiIslemlerApp"})
            with urllib.request.urlopen(req, timeout=10) as r:
                info = json.loads(r.read().decode("utf-8"))
        except Exception as e:
            self.status_var.set("Hazır")
            messagebox.showerror("Güncelle", f"Sürüm bilgisi alınamadı:\n{e}")
            return

        remote = str(info.get("version", "0"))
        if self._parse_version(remote) <= self._parse_version(APP_VERSION):
            self.status_var.set("Hazır")
            messagebox.showinfo("Güncelle", f"Program güncel (sürüm {APP_VERSION}).")
            return

        notes = info.get("notes", "")
        url = info.get("url", "")
        if not url:
            messagebox.showerror("Güncelle", "Sürüm bilgisinde indirme adresi yok.")
            return
        if not messagebox.askyesno("Güncelle",
                                    f"Yeni sürüm bulundu: {remote} (sizde {APP_VERSION})\n\n{notes}\n\n"
                                    "Şimdi indirip kurulsun mu? Program yeniden başlayacak."):
            self.status_var.set("Hazır")
            return

        if not getattr(sys, "frozen", False):
            messagebox.showinfo("Güncelle", "Kaynak koddan çalışıyorsunuz, 'git pull' ile güncelleyin.")
            self.status_var.set("Hazır")
            return

        try:
            self.status_var.set(f"Sürüm {remote} indiriliyor...")
            self.root.update()
            exe_path = sys.executable
            new_path = exe_path + ".yeni"
            req = urllib.request.Request(url, headers={"User-Agent": "TibbiIslemlerApp"})
            with urllib.request.urlopen(req, timeout=120) as r, open(new_path, "wb") as f:
                while True:
                    chunk = r.read(1024 * 256)
                    if not chunk:
                        break
                    f.write(chunk)
        except Exception as e:
            self.status_var.set("Hazır")
            messagebox.showerror("Güncelle", f"İndirme başarısız:\n{e}")
            return

        # Çalışan exe'yi kapatınca yenisiyle değiştiren küçük script
        bat_path = os.path.join(os.path.dirname(exe_path), "guncelle.bat")
        pid = os.getpid()
        with open(bat_path, "w", encoding="utf-8") as f:
            f.write(f'@echo off\n'
                    f':bekle\n'
                    f'tasklist /FI "PID eq {pid}" 2>nul | find "{pid}" >nul && '
                    f'(timeout /t 1 /nobreak >nul & goto bekle)\n'
                    f'move /Y "{new_path}" "{exe_path}" >nul\n'
                    f'start "" "{exe_path}"\n'
                    f'del "%~f0"\n')
        messagebox.showinfo("Güncelle", "Güncelleme indirildi. Program kapanıp yeni sürüm açılacak.")
        os.startfile(bat_path)
        self._on_close()

    def _browse_file(self):
        """Dosya seçme penceresi (Dosya menüsünden çağrılır)."""
        path = filedialog.askopenfilename(
            title="Excel Dosyası Seç",
            filetypes=[("Excel Dosyaları", "*.xlsx *.xls"), ("Tüm Dosyalar", "*.*")]
        )
        if path:
            self.file_var.set(path)
            self.config.set("last_file", path)
            self._update_file_label()
            self._load_data()

    def _load_data(self):
        """Excel dosyasını yükler."""
        path = self.file_var.get()
        if not path or not os.path.exists(path):
            messagebox.showerror("Hata", f"Dosya bulunamadı:\n{path}")
            return

        self.status_var.set("Dosya yükleniyor...")
        self.root.update()

        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            self.all_rows = []
            for row in ws.iter_rows(min_row=3, values_only=True):
                kod, ad, aciklama, puan, grup = (list(row) + [None] * 5)[:5]
                if kod is None and ad is None:
                    continue
                self.all_rows.append((kod, ad, aciklama, puan, grup))
            wb.close()
        except Exception as e:
            messagebox.showerror("Hata", f"Dosya okunamadı:\n{e}")
            self.status_var.set("Hata oluştu")
            return

        self.status_var.set(f"{len(self.all_rows)} kayıt yüklendi")
        self._update_grup_filter()
        self._update_favorites_list()
        self._apply_filter()

    def _update_grup_filter(self):
        """Ameliyat grubu filtresini günceller."""
        grupler = sorted(set(str(row[4]) for row in self.all_rows if row[4] is not None))
        self.grup_combo['values'] = ["Tümü"] + grupler
        self.grup_combo.set("Tümü")

    def _on_search_change(self, *args):
        """Arama değişikliklerini işler."""
        self._apply_filter()

    def _on_history_select(self, event):
        """Arama geçmişinden seçim yapıldığında."""
        selected = self.history_var.get()
        if selected:
            self.search_var.set(selected)

    def _add_to_history(self, query):
        """Arama sorgusunu geçmişe ekler."""
        if query and query not in self.search_history:
            self.search_history.insert(0, query)
            self.search_history = self.search_history[:50]  # Maksimum 50 kayıt
            self.config.set("search_history", self.search_history)
            self.history_combo.values = self.search_history[:20]

    def _clear_search(self):
        """Aramayı temizler."""
        self.search_var.set("")
        self.puan_min_var.set("")
        self.puan_max_var.set("")
        self.grup_var.set("Tümü")
        self._apply_filter()

    def _clear_search_text(self):
        """Arama kutusundaki tüm kelimeleri siler (filtrelere dokunmaz)."""
        self.search_var.set("")
        self.history_var.set("")
        self.search_entry.focus_set()
        self._apply_filter()

    def _apply_puan_filter(self):
        """Puan filtresini uygular."""
        try:
            min_val = self.puan_min_var.get().strip()
            max_val = self.puan_max_var.get().strip()
            self.filter_puan_min = float(min_val) if min_val else None
            self.filter_puan_max = float(max_val) if max_val else None
            self._apply_filter()
        except ValueError:
            messagebox.showwarning("Uyarı", "Geçerli bir puan aralığı girin.")

    def _clear_filters(self):
        """Tüm filtreleri temizler."""
        self.puan_min_var.set("")
        self.puan_max_var.set("")
        self.grup_var.set("Tümü")
        self.filter_puan_min = None
        self.filter_puan_max = None
        self.filter_grup = None
        self._apply_filter()

    def _apply_filter(self):
        """Filtreleri uygular."""
        query = self.search_var.get()
        self._add_to_history(query)

        self.filtered_rows = []
        for row in self.all_rows:
            # Arama filtresi (büyük/küçük harf duyarsız, Türkçe uyumlu)
            if query:
                query_lower = tr_lower(query)
                found = False
                for val in row:
                    if val is not None and query_lower in tr_lower(val):
                        found = True
                        break
                if not found:
                    continue

            # Puan filtresi
            if self.filter_puan_min is not None or self.filter_puan_max is not None:
                puan = row[3]
                if puan is None:
                    continue
                try:
                    puan_val = float(puan)
                except (ValueError, TypeError):
                    continue
                if self.filter_puan_min is not None and puan_val < self.filter_puan_min:
                    continue
                if self.filter_puan_max is not None and puan_val > self.filter_puan_max:
                    continue

            # Grup filtresi
            if self.grup_var.get() != "Tümü":
                if str(row[4]) != self.grup_var.get():
                    continue

            self.filtered_rows.append(row)

        self._populate_tree()
        self._update_statistics()

    def _populate_tree(self):
        """Tabloyu doldurur."""
        self.tree.delete(*self.tree.get_children())
        for row in self.filtered_rows:
            self.tree.insert("", tk.END, values=row)
        self.count_label.config(text=f"{len(self.filtered_rows)} kayıt gösteriliyor")

    def _update_statistics(self):
        """İstatistikleri günceller."""
        if not self.filtered_rows:
            for label in self.stats_labels.values():
                label.config(text="-")
            return

        puanlar = []
        for row in self.filtered_rows:
            if row[3] is not None:
                try:
                    puanlar.append(float(row[3]))
                except (ValueError, TypeError):
                    pass
        grupler = set(str(row[4]) for row in self.filtered_rows if row[4] is not None)

        self.stats_labels["total"].config(text=str(len(self.filtered_rows)))
        self.stats_labels["avg_puan"].config(text=f"{sum(puanlar)/len(puanlar):.2f}" if puanlar else "-")
        self.stats_labels["max_puan"].config(text=str(max(puanlar)) if puanlar else "-")
        self.stats_labels["min_puan"].config(text=str(min(puanlar)) if puanlar else "-")
        self.stats_labels["unique_grup"].config(text=str(len(grupler)))

    def _sort_by(self, col):
        """Sütuna göre sıralama yapar."""
        col_idx = {"kod": 0, "ad": 1, "aciklama": 2, "puan": 3, "grup": 4}[col]

        if self.sort_column == col:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = col
            self.sort_reverse = False

        def sort_key(row):
            val = row[col_idx]
            if val is None:
                return (1, "")
            if isinstance(val, (int, float)):
                return (0, val)
            return (0, str(val))

        self.filtered_rows.sort(key=sort_key, reverse=self.sort_reverse)
        self._populate_tree()

    def _toggle_theme(self):
        """Tema değiştirir."""
        self.is_dark = not self.is_dark
        self.config.set("theme", "dark" if self.is_dark else "light")
        messagebox.showinfo("Tema", "Tema değişikliği için uygulamayı yeniden başlatın.")

    def _reset_column_widths(self):
        """Sütun genişliklerini sıfırlar."""
        default_widths = {"kod": 100, "ad": 250, "aciklama": 400, "puan": 90, "grup": 150}
        for col, width in default_widths.items():
            self.tree.column(col, width=width)
        self.config.set("column_widths", {})

    def _show_context_menu(self, event):
        """Sağ tık menüsünü gösterir."""
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def _copy_selection(self):
        """Seçili satırları kopyalar."""
        selected = self.tree.selection()
        if not selected:
            return
        lines = []
        for item in selected:
            values = self.tree.item(item, "values")
            lines.append("\t".join(str(v) for v in values))
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(lines))
        self.status_var.set(f"{len(selected)} satır kopyalandı")

    def _add_selected_to_favorites(self):
        """Seçili satırları sık kullanılanlara ekler."""
        selected = self.tree.selection()
        for item in selected:
            values = self.tree.item(item, "values")
            fav_text = f"{values[0]} - {values[1]}"
            if fav_text not in self.favorites:
                self.favorites.append(fav_text)
        self.config.set("favorites", self.favorites)
        self._update_favorites_list()

    def _add_favorite(self):
        """Mevcut aramayı sık kullanılanlara ekler."""
        query = self.search_var.get()
        if query and query not in self.favorites:
            self.favorites.append(query)
            self.config.set("favorites", self.favorites)
            self._update_favorites_list()

    def _remove_favorite(self):
        """Seçili sık kullanılanı siler."""
        selection = self.fav_listbox.curselection()
        if selection:
            idx = selection[0]
            self.favorites.pop(idx)
            self.config.set("favorites", self.favorites)
            self._update_favorites_list()

    def _update_favorites_list(self):
        """Sık kullanılanlar listesini günceller."""
        self.fav_listbox.delete(0, tk.END)
        for fav in self.favorites:
            self.fav_listbox.insert(tk.END, fav)

    def _on_fav_select(self, event):
        """Sık kullanılandan seçim yapıldığında."""
        selection = self.fav_listbox.curselection()
        if selection:
            idx = selection[0]
            fav_text = self.favorites[idx]
            # Eğer "kod - ad" formatındaysa sadece ad kısmını ara
            if " - " in fav_text:
                search_text = fav_text.split(" - ", 1)[1]
            else:
                search_text = fav_text
            self.search_var.set(search_text)
            self._apply_filter()

    def _export_results(self):
        """Sonuçları CSV olarak dışa aktarır."""
        if not self.filtered_rows:
            messagebox.showwarning("Uyarı", "Dışa aktarılacak veri yok.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Dosyası", "*.csv"), ("Tüm Dosyalar", "*.*")]
        )
        if not path:
            return

        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(["İşlem Kodu", "İşlem Adı", "Açıklama", "İşlem Puanı", "Ameliyat Grupları"])
                writer.writerows(self.filtered_rows)
            self.status_var.set(f"{len(self.filtered_rows)} kayıt CSV olarak aktarıldı")
        except Exception as e:
            messagebox.showerror("Hata", f"Dışa aktarma başarısız:\n{e}")

    def _export_excel(self):
        """Sonuçları Excel olarak dışa aktarır."""
        if not self.filtered_rows:
            messagebox.showwarning("Uyarı", "Dışa aktarılacak veri yok.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Dosyası", "*.xlsx"), ("Tüm Dosyalar", "*.*")]
        )
        if not path:
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["İşlem Kodu", "İşlem Adı", "Açıklama", "İşlem Puanı", "Ameliyat Grupları"])
            for row in self.filtered_rows:
                ws.append(row)
            wb.save(path)
            self.status_var.set(f"{len(self.filtered_rows)} kayıt Excel olarak aktarıldı")
        except Exception as e:
            messagebox.showerror("Hata", f"Dışa aktarma başarısız:\n{e}")

    def _show_charts(self):
        """Grafikleri gösterir."""
        if not USE_MPL:
            messagebox.showinfo("Bilgi", "Matplotlib kütüphanesi yüklü değil.")
            return

        if not self.filtered_rows:
            messagebox.showwarning("Uyarı", "Grafik için veri yok.")
            return

        chart_window = tk.Toplevel(self.root)
        chart_window.title("Grafikler")
        chart_window.geometry("800x600")

        # Puan dağılımı
        puanlar = [float(row[3]) for row in self.filtered_rows if row[3] is not None]
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))

        ax1.hist(puanlar, bins=20, color=self.accent_color, edgecolor="black")
        ax1.set_title("Puan Dağılımı")
        ax1.set_xlabel("Puan")
        ax1.set_ylabel("Frekans")

        # Grup dağılımı
        grupler = Counter(str(row[4]) for row in self.filtered_rows if row[4] is not None)
        ax2.barh(list(grupler.keys())[:10], list(grupler.values())[:10], color=self.accent_color)
        ax2.set_title("En Çok İşlem Yapılan Gruplar (Top 10)")
        ax2.set_xlabel("İşlem Sayısı")

        plt.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=chart_window)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def _show_shortcuts(self):
        """Kısayolları gösterir."""
        shortcuts = """
Kısayollar:
  Ctrl+F      - Arama kutusuna odaklan
  Ctrl+E      - CSV olarak dışa aktar
  Ctrl+R      - Verileri yenile
  F5          - Verileri yenile
  Esc         - Aramayı temizle
  Ctrl+O      - Dosya aç
        """
        messagebox.showinfo("Kısayollar", shortcuts)

    def _show_about(self):
        """Hakkında penceresi."""
        messagebox.showinfo("Hakkında",
                          "Tıbbi İşlemler Listesi - Gelişmiş Arama Uygulaması\n"
                          f"Sürüm: {APP_VERSION}\n"
                          "Excel dosyasındaki tıbbi işlemleri arama, filtreleme ve analiz aracı.")

    def _on_drop(self, event):
        """Sürükle-bırak ile dosya yükleme."""
        path = event.data.strip("{}")
        if os.path.exists(path):
            self.file_var.set(path)
            self.config.set("last_file", path)
            self._update_file_label()
            self._load_data()

    def _on_close(self):
        """Uygulama kapanırken ayarları kaydeder."""
        # Sütun genişliklerini kaydet
        widths = {}
        for col in ("kod", "ad", "aciklama", "puan", "grup"):
            widths[col] = self.tree.column(col, "width")
        self.config.set("column_widths", widths)
        self.config.set("window_size", f"{self.root.winfo_width()}x{self.root.winfo_height()}")
        self.root.destroy()


def main():
    if USE_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = TibbiIslemlerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
