NEBULA-S WPF Uygulaması - Teknofest 9. Model Uydu Yarışması

Proje Amacı
Bu uygulamanın temel amacı, uydumuzdan gelen verileri etkili bir biçimde görselleştirmek ve kullanıcı arayüzü aracılığıyla komutlar göndermektir. Uygulamanın sunduğu başlıca özellikler şunlardır:
Uydunun GPS Konumunu Belirleme: Gelen verilere göre uydunun GPS konumunu belirleyip harita üzerinde göstermek.
Uydunun Havadaki Duruşunu Gösterme: Uydunun roll, yaw ve pitch değerlerini anlık olarak görselleştirerek, havadaki duruşunu kullanıcıya sunmak.
Veri Görselleştirme: Sıcaklık, nem, hız, basınç, yükseklik, pil gerilimi gibi önemli telemetri verilerini LiveChart kütüphanesi kullanarak görselleştirmek.
Gerçek Zamanlı Kamera Görüntüsü: Uydudan gelen canlı kamera görüntüsünü arayüzde göstermek ve kayıt tuşuna basarak görüntüyü kaydetmek.
Alarm Sistemi: Gelen verilere göre kritik durumları belirleyip, alarm sistemi üzerinden kullanıcıyı anında uyarmak.

Kullanılan Teknolojiler

.NET Framework: Uygulama geliştirme için temel platform.
C#: Uygulamanın arka plan kodlaması için kullanılan programlama dili.
WPF (Windows Presentation Foundation): Kullanıcı arayüzü geliştirmek için kullanılan framework.
MVVM (Model-View-ViewModel): Uygulamanın modüler, test edilebilir ve ölçeklenebilir bir yapıya sahip olmasını sağlamak için kullanılan tasarım deseni.
INotifyPropertyChanged: Gerçek zamanlı veri güncellemeleri için kullanılan arayüz.
Asenkron Programlama: Performans ve verimliliği artırmak için kullanılan yapı.

Uygulamanın Teknik Özellikleri
1. Tasarım ve Modüler Yapı
Uygulama, MVVM (Model-View-ViewModel) tasarım deseni kullanılarak geliştirildi. Bu desen, uygulamanın modüler, esnek ve test edilebilir olmasını sağladı. Kullanıcı dostu ve yüksek performanslı bir arayüz uygulaması oluşturuldu.

2. Gerçek Zamanlı Güncellemeler
INotifyPropertyChanged arayüzü ile uygulamanın veri güncellemelerini anlık olarak arayüze yansıtmasını sağladım. Bu mekanizma sayesinde, verilerdeki değişiklikler anında kontrol edilerek en güncel bilgiler arayüzde görüntülendi.

3. Verimlilik ve Performans Optimizasyonu
Uygulamanın performansını artırmak için asenkron yapılar kullanıldı. Bu, veri işleme süreçlerini hızlandırarak gerçek zamanlı verilere kesintisiz erişim sağladı. Performans ve kullanıcı deneyimi ön planda tutularak en uygun araçlar seçildi.

4. Esneklik ve Ölçeklenebilirlik
Uygulamanın modüler yapısı, gelecekteki gereksinimlere uyum sağlamayı kolaylaştırdı. Bu sayede, yeni veri kaynakları entegre edilebilir ve ek fonksiyonlar kolaylıkla eklenebilir hale getirildi.

Sonuç
Bu WPF uygulaması, NEBULA-S takımının uydu operasyonlarını verimli ve etkili bir şekilde yönetmesine olanak tanıyan kapsamlı bir çözüm sundu. Hem geliştirici hem de kullanıcı perspektifinden bakıldığında, bu uygulama modern yazılım geliştirme ilkelerinin ve teknolojik yeniliklerin bir uyum içinde kullanıldığını göstermektedir.
