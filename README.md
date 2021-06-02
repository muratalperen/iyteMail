# İyte Mail Otomasyonu

Zimbra tabanlı Iyte mail kutusu hızla ve çoğu ilgilendirmeyen maillerle doluyor. Bu reponun amacı maillerle yapılacak işlemleri otomatiğe bağlamaktır.

Python Selenium modülü ile yazılmıştır. İleride kullanılabilmesi için veriler csv olarak kaydedilebilir.

Kullanmak için config dosyasını kendinize göre düzenleyin:

+ Kullanıcı adı ve şifrenizi belirleyin
+ Gönderen ve kelime tabanlı black listinizi kendinize göre düzenleyin

*zimbraMail.py* dosyasına girip selenium driver konumunu belirtin. `python3 filter.py` ile çalıştırın.

### Bilinen hatalar

+ eski dosyaya ekleme yapıldığında son birkaç satırı kontrol edin, zaten var olan mailleri ekleyebiliyor
