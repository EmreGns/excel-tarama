Excell Gelişmiş Aramanın Amacı:
Excell Gelişmiş Arama, Excel dosyalarında yazım farkı, eksik harf veya karakter nedeniyle bulunamayan verileri tespit eder.  
Gelişmiş arama algoritması sayesinde aradığınız kelimeye en yakın sonuçları benzerlik oranlarıyla birlikte listeler.
Arama, seçilecek olan klasör içerisinde yapılmaktadır. Klasör içerisinde birden bazla excell dosyası varsa hepsini tarayacaktır. 
Yani birkaç excell dosyası içerisinde bir tag'in hangi excell içerisinde olduğunu bulmak için de kullanabilirsiniz.
--------------------------------------------------------------------------------------------------------------------------------
Programın çalışabilmesi için aşağıdaki Python kütüphaneleri gereklidir:
Python 3.13.7
pandas 2.3.2
openpyxl 3.1.5
rapidfuzz 3.14.0

Kurulumları yapmak için:

Python 3.13.7 için aşağıdaki adresten yükleme yapabilirsiniz.
https://www.python.org/downloads/

Cmd'ye aşağıdaki komutu yapıştırdığınızda gerekli kütüphaneler de kurulmuş olacaktır.
pip install pandas==2.3.2 openpyxl==3.1.5 rapidfuzz==3.14.0
--------------------------------------------------------------------------------------------------------------------------------
Excell Gelişmiş Arama programı iki ana dosyadan oluşur:
excell_arama.py dosyası kodun arka plan kısmıdır. 
Kurulumlar tamamlandıktan sonra Excell Gelişmiş Arama.py dosyasına çift tıklayarak programı başlatabilirsiniz.
--------------------------------------------------------------------------------------------------------------------------------
Dikkat Edilmesi Gerekenler:
Program çalışırken bir terminal penceresi açılır. Bu pencereyi kapatmayın. Kodlar bu terminal üzerinden çalışır. Terminali 
kapatırsanız program da kapanır. İsterseniz sağ üstteki _ butonuna basarak terminali küçültüp arka plana alabilirsiniz.

Program klasör halinde çalışacak şekilde tasarlanmıştır. İçerisindeki dosyaları taşırsanız, isim değişikliği yaparsanız program 
düzgün çalışmayabilir. 
--------------------------------------------------------------------------------------------------------------------------------
excell_arama.py dosyasında 15.satırda kısaltma sözlüğü bulunmakta. Bu sözlük "P-" araması yapıldığında algoritmanın "pressure" 
olarak algılamasını sağlıyor. Bu tarz düzenlemeleri kod içerisine ekleyebilirsiniz. 
--------------------------------------------------------------------------------------------------------------------------------
Tavsiyeler:
Bu program belli bir klasör içerisini taradığı için Desktop'unuzda bir klasör açıp (ör. Excell_bul) arayacağınız .xlsx dosyalarını 
o klasöre ekleyip çıkartarak programı kullanabilirsiniz.

Program bazı bilgisayarlarda yavaş çalışabiliyor. İsteğe göre birden fazla arayüz açıp aynı anda hepsinde aramalar yapabilirsiniz.
--------------------------------------------------------------------------------------------------------------------------------
Ek bilgi:
__pycache__ klasörü, programı ilk kez çalıştırdığınızda otomatik olarak oluşur. Bu klasör, programın sonraki çalıştırmalarda 
daha hızlı açılmasını sağlar. Silmenize gerek yoktur.
