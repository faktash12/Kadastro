# Kadastro Harç Hesaplama (Masaüstü)

Tapu ve Kadastro döner sermaye tarifelerine göre:
- Teknik Bilgi-Belge hesaplaması
- Kontrollük hesaplaması
- İl/ilçe bazlı yöresel katsayı (YK) yönetimi
- Excel şablonundan katsayı yükleme
- PDF çıktı üretimi

uygulamalarını yapan Python/Tkinter masaüstü uygulaması.

## Özellikler

- Teknik Bilgi-Belge:
  - Parsel köşe noktası, poligon, ölçü krokisi, dönüşüm parametresi ve diğer belgeler
  - Toplam bedel ve PDF çıktısı
- Kontrollük:
  - Değişiklik işlemleri (15-16. Madde), parselasyon, kamulaştırma, mülkiyet raporu vb. işlem türleri
  - İl/ilçe seçimine göre Kadastro YK uygulaması
  - Asgari kontrollük sınırı gereken kalemlerde otomatik uygulanır
- Admin/Ayar:
  - Varsayılan il
  - Excel’den toplu katsayı yükleme
- Tema:
  - Aydınlık / Karanlık

## Kurulum

```bash
pip install -r requirements.txt
```

## Çalıştırma

```bash
python desktop_app.py
```

## EXE Derleme (PyInstaller)

```bash
pyinstaller --noconfirm KadastroDesktopPro.spec
```

Derleme çıktısı:

`dist/KadastroDesktopPro/KadastroDesktopPro.exe`

## Dosyalar

- `desktop_app.py` : Ana masaüstü uygulaması
- `katsayilar.json` : İl/ilçe katsayı verisi
- `KadastroDesktopPro.spec` : EXE derleme yapılandırması

## Not

Bu proje kurum içi kullanım senaryoları için geliştirilmiştir. Tarife değişikliklerinde katsayı/ücret alanları güncellenmelidir.
