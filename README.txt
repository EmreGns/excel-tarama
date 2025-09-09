Excell Gelişmiş Arama Programı Nedir?

Excell Gelişmiş Arama, Excel dosyalarında yazım hatası, eksik harf veya karakter farkı gibi nedenlerle bulunamayan verileri
tespit etmek amacıyla geliştirilmiştir. Gelişmiş benzerlik algoritması sayesinde, aradığınız kelimeye en yakın sonuçlar yüzde
cinsinden benzerlik oranlarıyla birlikte listelenir. Arama işlemi, kullanıcı tarafından seçilecek bir klasör içerisinde
gerçekleştirilir. Seçilen klasörde birden fazla Excel (.xlsx) dosyası varsa, program bunların tamamını tarayarak ilgili
sonuçları sunar. Örneğin, bir klasör içinde yer alan 8 farklı Excel dosyasında belirli bir tag’i veya metni aramak
istiyorsanız, bu dosyaları tek bir klasörde toplayıp program üzerinden klasörü seçmeniz yeterlidir. Ayrıca, arama yapmak
istediğiniz Excel dosyasında yazım hataları varsa ve Excel’in kendi "Bul" özelliği bu kelimeleri bulamıyorsa, bu program
size en yakın eşleşmeleri göstererek aradığınızı kolayca bulmanızı sağlar.
--------------------------------------------------------------------------------------------------------------------------------
Programın çalışabilmesi için aşağıdaki Python kütüphaneleri gereklidir:
Python 3.13.7
pandas 2.3.2
openpyxl 3.1.5
rapidfuzz 3.14.0

Kurulumları yapmak için:

Python 3.13.7 sürümünü aşağıdaki adresten indirebilirsiniz:
https://www.python.org/downloads/

Kurulum sırasında "Add Python to PATH" seçeneğini işaretlemeyi unutmayın. Bu, Python'un komut satırından çalışabilmesi için 
gereklidir.

Cmd'ye aşağıdaki komutu yapıştırdığınızda gerekli kütüphaneler de kurulmuş olacaktır.
pip install pandas==2.3.2 openpyxl==3.1.5 rapidfuzz==3.14.0
--------------------------------------------------------------------------------------------------------------------------------
Excell Gelişmiş Arama programı iki ana dosyadan oluşur:
excell_arama_backend.py dosyası kodun arka plan kısmıdır.
Excell Gelişmiş Arama Programı ise arayüz kısmıdır.
 
Kurulumlar tamamlandıktan sonra Excell Gelişmiş Arama Programı.py dosyasına çift tıklayarak programı başlatabilirsiniz.
--------------------------------------------------------------------------------------------------------------------------------
Dikkat Edilmesi Gerekenler:
Program çalışırken bir terminal penceresi açılır. Bu pencereyi kapatmayın. Kodlar bu terminal üzerinden çalışır. Terminali 
kapatırsanız program da kapanır. İsterseniz sağ üstteki _ butonuna basarak terminali küçültüp arka plana alabilirsiniz.

Program klasör halinde çalışacak şekilde tasarlanmıştır. İçerisindeki dosyaları taşırsanız, isim değişikliği yaparsanız program 
düzgün çalışmayabilir. 
--------------------------------------------------------------------------------------------------------------------------------
excell_arama_backend.py dosyasında 15.satırda kısaltma sözlüğü bulunmakta. Bu sözlük "P-" araması yapıldığında algoritmanın "pressure" 
olarak algılamasını sağlıyor. Bu tarz düzenlemeleri kod içerisine ekleyebilirsiniz. 
--------------------------------------------------------------------------------------------------------------------------------
Tavsiyeler:
Masaüstünüzde bir klasör (ör. Excell_taranacak_klasörler) oluşturup taramak istediğiniz Excel dosyalarını buraya eklerseniz,
programı ilk açtığınızda bu klasörü bir defaya mahsus seçmeniz yeterlidir. Seçtiğiniz klasör, uygulama tarafından otomatik olarak 
kaydedilir ve bir sonraki çalıştırmada tekrar klasör seçmenize gerek kalmaz.

Program bazı bilgisayarlarda yavaş çalışabiliyor. İsteğe göre birden fazla arayüz açıp hepsinde farklı aramalar yapabilirsiniz.
--------------------------------------------------------------------------------------------------------------------------------
Ek bilgi:
__pycache__ klasörü, programı ilk kez çalıştırdığınızda otomatik olarak oluşur. Bu klasör, programın sonraki çalıştırmalarda 
daha hızlı açılmasını sağlar. Silmenize gerek yoktur.

