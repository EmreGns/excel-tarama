[README.md](https://github.com/user-attachments/files/28836280/README.md)
# 🔍 Excel Gelişmiş Arama Programı

Excel Gelişmiş Arama, Excel dosyalarında yazım hatası, eksik harf veya karakter farkı gibi nedenlerle bulunamayan verileri tespit etmek amacıyla geliştirilmiştir.

Gelişmiş benzerlik algoritması sayesinde, aradığınız kelimeye en yakın sonuçlar **yüzde cinsinden benzerlik oranlarıyla** birlikte listelenir. Arama işlemi, kullanıcı tarafından seçilecek bir klasör içerisinde gerçekleştirilir.

Seçilen klasörde birden fazla Excel (`.xlsx`) dosyası varsa, program bunların tamamını tarayarak ilgili sonuçları sunar.

> **Örnek:** Bir klasör içinde yer alan 8 farklı Excel dosyasında belirli bir etiket veya metni aramak istiyorsanız, bu dosyaları tek bir klasörde toplayıp program üzerinden klasörü seçmeniz yeterlidir.

Ayrıca, arama yapmak istediğiniz Excel dosyasında yazım yanlışları varsa ve Excel'in kendi "Bul" özelliği bu kelimeleri bulamıyorsa, bu program size **en yakın eşleşmeleri** göstererek aradığınızı kolayca bulmanızı sağlar.

---

## 🖥️ Arayüz

<img width="920" height="680" alt="proje_gorseli" src="https://github.com/user-attachments/assets/ac00057a-1423-4a40-b72d-5134704e4df0" />
---

## 📦 Gereksinimler

| Paket | Sürüm |
|-------|-------|
| Python | 3.13.7 |
| pandas | 2.3.2 |
| openpyxl | 3.1.5 |
| rapidfuzz | 3.14.0 |

---

## ⚙️ Kurulum

### 1. Python Kurulumu

Python 3.13.7 sürümünü aşağıdaki adresten indirin:

🔗 https://www.python.org/downloads/

> ⚠️ Kurulum sırasında **"Add Python to PATH"** seçeneğini işaretlemeyi unutmayın. Bu, Python'un komut satırından çalışabilmesi için gereklidir.

### 2. Kütüphane Kurulumu

Aşağıdaki komutu CMD'ye yapıştırarak gerekli kütüphaneleri yükleyin:

```
pip install pandas==2.3.2 openpyxl==3.1.5 rapidfuzz==3.14.0
```

---

## 🗂️ Dosya Yapısı

Program iki ana dosyadan oluşur:

| Dosya | Açıklama |
|-------|----------|
| `excel_arama_backend.py` | Arka plan işlemleri (algoritma ve veri işleme) |
| `Excel Gelişmiş Arama Programı.py` | Kullanıcı arayüzü |

Kurulumlar tamamlandıktan sonra **`Excel Gelişmiş Arama Programı.py`** dosyasına çift tıklayarak programı başlatabilirsiniz.

---

## ⚠️ Dikkat Edilmesi Gerekenler

- Program çalışırken bir **terminal penceresi** açılır. Bu pencereyi kapatmayın — kodlar bu terminal üzerinden çalışır. Terminali kapatırsanız program da kapanır. İsterseniz sağ üstteki `_` butonuna basarak terminali küçültüp arka plana alabilirsiniz.
- Program **klasör halinde** çalışacak şekilde tasarlanmıştır. İçerisindeki dosyaları taşırsanız veya isim değişikliği yaparsanız program düzgün çalışmayabilir.

---

## 🛠️ Özelleştirme

`excel_arama_backend.py` dosyasının **15. satırında** bir kısaltma sözlüğü bulunmaktadır.

Bu sözlük sayesinde örneğin `"P-"` araması yapıldığında algoritma bunu `"pressure"` olarak algılayabilir. Benzer özelleştirmeleri bu sözlüğe ekleyerek programı kendi ihtiyaçlarınıza göre uyarlayabilirsiniz.

---

## 💡 Tavsiyeler

- Masaüstünüzde bir klasör *(örn. `Excel_taranacak_klasörler`)* oluşturup taramak istediğiniz Excel dosyalarını buraya eklerseniz, programı ilk açtığınızda bu klasörü **bir defaya mahsus** seçmeniz yeterlidir. Seçtiğiniz klasör uygulama tarafından otomatik olarak kaydedilir ve bir sonraki çalıştırmada tekrar seçmenize gerek kalmaz.
- Program bazı bilgisayarlarda yavaş çalışabiliyor. İsteğe göre **birden fazla arayüz** açıp her birinde farklı aramalar yapabilirsiniz.

---

## ℹ️ Ek Bilgi

`__pycache__` klasörü, programı ilk kez çalıştırdığınızda otomatik olarak oluşur. Bu klasör, programın sonraki çalıştırmalarda daha hızlı açılmasını sağlar — silmenize gerek yoktur.
