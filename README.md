# Tıbbi İşlemler Listesi - Gelişmiş Arama

Excel dosyasındaki tıbbi işlemleri arayan, filtreleyen ve analiz eden masaüstü uygulaması (Python + tkinter).

## Kullanım (son kullanıcı)

1. GitHub **Releases** sayfasından en güncel `TibbiIslemlerApp.exe` dosyasını indirin.
2. Çift tıklayıp çalıştırın (kurulum gerekmez).
3. `Dosya → Dosya Seç / Değiştir...` ile Excel dosyanızı seçin.

## Güncelleme (son kullanıcı)

Program açıkken sol üstte **Dosya → Güncelle (sürümü denetle)** menüsüne tıklayın.
Yeni sürüm varsa otomatik indirilir, program kapanıp yeni sürüm açılır.

## Geliştirici

```bash
# Gerekli kütüphane
C:\Users\kevser.cerci2\AppData\Local\Programs\Python\Python313\python.exe -m pip install openpyxl pyinstaller

# Çalıştır
.../Python313/python.exe tibbi_islemler_app.py

# Exe üret
.../Python313/python.exe -m PyInstaller --clean --noconfirm TibbiIslemlerApp.spec
# Çıktı: dist/TibbiIslemlerApp.exe
```

## Yeni sürüm yayınlama (yönetici)

1. `tibbi_islemler_app.py` içindeki `APP_VERSION` değerini yükseltin.
2. Exe'yi yeniden üretin.
3. Değişiklikleri commit + push yapın.
4. `version.json` içindeki `version`, `url` (yeni exe'nin Release indirme adresi) ve `notes` alanlarını güncelleyip push yapın.
5. GitHub'da yeni bir **Release** açıp exe'yi ekleyin.

Kullanıcılar `Dosya → Güncelle` ile yeni sürüme geçer.
