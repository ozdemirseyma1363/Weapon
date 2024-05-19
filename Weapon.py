import face_recognition#yüz tanıma kütüphanesi
import dlib#yapay zeka kütüphanesi
import cv2#görüntü işleme kütüphanesi
import imutils#yeniden boyutlandırma için gerekli kütüphane
import keyboard#klavye kullanımı için gerekli kütüphane
import datetime#zaman kullanımı için gerekli kütüphane
import openpyxl #excel işlemleri için kurulan kütüphane
wb = openpyxl.load_workbook('mert.xlsx')
sayfa=wb.active#dosyayı etkinleştirme
tarih = datetime.datetime.now()#bize içindeki bulunduğumuz andaki tarih, saat ve zaman bilgilerini verir.
a="BEYZA"#kişi adı
silah_cascade = cv2.CascadeClassifier("C:\\Users\\User\\Desktop\\cascade.xml")#cascade sınıflandırıcısının dosya yolu
bicak_cascade=cv2.CascadeClassifier("C:\\Users\\User\\Desktop\\bicak1\\classifier\\cascade.xml")#cascade sınıflandırıcısının dosya yolu
detector=dlib.get_frontal_face_detector()#yüzleri tespit etmek için kullanabileceğimiz sınıfın ir nesnesini döndürecektir.
ben=face_recognition.load_image_file("49.PNG")#yüz tespiti yapılacak kişinin resmi
ben1=face_recognition.face_encodings(ben)[0]
cap=cv2.VideoCapture(0)#kamera kullanımı
firstFrame = None
gun_exist = False
while True:#sonsuz döngü
    ret,frame=cap.read()#kamera okuma
    if not ret:#kameraya ulaşamıyorsa
       break#çık
    face_log=[]#yüzlerin boş dizisini oluşturma
    faces=detector(frame)#nesnelerinin bir dizisini döndürür.
    for face in faces:#tüm yüzlerin
        x=face.left()#koordinatlar
        y=face.top()#koordinatlar
        w=face.right()#koordinatlar
        h=face.bottom()#koordinatlar
        face_log.append((y,w,h,x))#yüzlerin kordinatına ekle
    face2=face_recognition.face_encodings(frame,face_log)#tüm yüzler
    i=0#for döngüsü için değişken ataması
    frame = imutils.resize(frame, width=500)#yeniden boyutlandırmsa
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)#griye dönüştürme
    gray = cv2.GaussianBlur(gray, (21, 21), 0)#filtreleme
    silah = silah_cascade.detectMultiScale(gray, 1.3, 12, minSize=(100, 100))#silahların kordinat değerleri bulunur
    bicak= bicak_cascade.detectMultiScale(gray,1.3, 300)#bıçakların  kordinat değerleri bulunur
    for (x, y, w, h) in silah:
        frame = cv2.rectangle(frame, (x, y), (x + w, y + h), (255, 0, 0), 2)#silahları dikdörtgen içine al
        cv2.putText(frame, "silah", (x, y), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 0), cv2.LINE_4)
        gray1= gray[y:y + h, x:x + w]
        color1 = frame[y:y + h, x:x + w]
        #ctypes.windll.user32.MessageBoxW(0, a+"  KİŞİSİNDE SİLAH TESPİT EDİLMİŞTİR", "UYARI", 0)#message box ile uyarı ver
        sayfa.append([a +" KİSİSİ "+str(tarih)+" TARİHİNDE SİLAH İLE GİRİŞ YAPTI "])  # excel dosyasına yazısını ekler
        wb.save("mert.xlsx")  # kaydeder
        #ctypes.windll.user32.MessageBoxW(0, str(tarih)+" SİLAH TESPİT EDİLMİŞTİR", "UYARI", 0)  # messagebox ile uyarı ver
    for (x, y, w, h) in bicak:
        frame = cv2.rectangle(frame, (x, y), (x + w, y + h), (255, 0, 0), 2)#bicakları dikdörtgen içine a
        cv2.putText(frame, "bicak", (x, y), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 0), cv2.LINE_4)
        gray2 = gray[y:y + h, x:x + w]
        color2 = frame[y:y + h, x:x + w]
        #ctypes.windll.user32.MessageBoxW(0,a+" KİŞİSİNDE BIÇAK TESPİT EDİLMİŞTİR", "UYARI", 0)#messagebox ile uyarı ver
        #ctypes.windll.user32.MessageBoxW(0, str(tarih)+" BIÇAK TESPİT EDİLMİŞTİR", "UYARI", 0)#messagebox ile uyarı ver
        sayfa.append([a +" KİSİSİ "+str(tarih)+" TARİHİNDE BICAK İLE GİRİŞ YAPTI "])  # excel dosyasına  yazısını ekler
        wb.save("mert.xlsx")  # kaydeder
    if firstFrame is None:#eğer kare yoksa
        firstFrame = gray#gray ilk karedir
        continue#devam et
    for face in face2:#tüm yüzlerin döngüsü
        y,w,h,x=face_log[i]#kordinatları belirler
        sonuc=face_recognition.compare_faces([ben1],face)#yüzleri karşılaştırır
        if sonuc[0]==True:#yüz tespit edilmişse
            cv2.rectangle(frame,(x-40,y-20),(w-80,h-20),(220,220,220),2)#dikdörtgen içine al
            cv2.putText(frame,a,(x-40,h+10),cv2.FONT_HERSHEY_COMPLEX, 1, (220,220,220), 1)#yazı yazdır
    out = cv2.resize(frame, (1600, 880), 2, 2, interpolation=cv2.INTER_CUBIC)#görüntü ekranının boyutlandırılması
    cv2.namedWindow('image', cv2.WND_PROP_FULLSCREEN)#yeni pencerenin full ekran olması
    cv2.setWindowProperty('image', cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)
    cv2.imshow("image", out)#görüntüleme
    cv2.waitKey(1)
    if keyboard.is_pressed("f"):#f tuşuna basıldığında
        cv2.destroyAllWindows()#pencereyi kapat
        cap.release()#kamerayı durdur
        wb.close()
        break#çık
