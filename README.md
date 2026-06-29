# 📊 Excel Web Editor

Ücretsiz, tarayıcı tabanlı çoklu Excel düzenleyici. Mobil + masaüstü uyumlu, çok cihazlı senkronizasyon.

## Özellikler

- 📂 **Çoklu Sekme**: Birden fazla Excel dosyasını aynı anda aç, sekmeler arası geçiş yap
- ✏️ **Mobil Düzenleme**: Dokunmatik uyumlu hücre seçimi ve alt panel editör
- 💾 **Otomatik Kaydetme**: Değişiklikler localStorage'a otomatik kaydedilir
- ☁️ **Firebase Sync**: Çok cihazlı gerçek zamanlı senkronizasyon (anonim, ücretsiz)
- 📤 **Dışa Aktar**: Düzenlenmiş veriyi .xlsx olarak indir (Ctrl+E)
- 🌓 **Koyu/Açık Tema**: Sistem tercihine göre otomatik, manuel toggle
- 🔍 **Arama**: Tabloda anlık filtreleme
- 📊 **Sıralama**: Sütun başlığına tıkla → artan/azalan
- ⌨️ **Klavye Kısayolları**: Ctrl+S tümünü kaydet, Ctrl+E dışa aktar, Ctrl+N yeni sekme, oklar + Enter + Delete
- 📱 **Mobil Alt Bar**: Düzenle, Aç, Kaydet, Yeni butonları

## Kullanım

1. Siteyi aç
2. Excel dosyası sürükle veya 📂 Aç'a tıkla → yeni sekmede açılır
3. Hücreye dokun → seç, çift tıkla veya ✏️ Düzenle → alt panelde değiştir
4. Diğer sekmeye geç → önceki otomatik kaydedilir
5. 💾 Kaydet veya Ctrl+S ile dışa aktar

## Mobil Kullanım

- Hücre seç: tek dokunuş
- Düzenle: çift dokunuş veya alt bardaki ✏️
- Sekme değiştir: üstteki sekmelere dokun
- Dosya aç: alt bardaki 📂

## Firebase Sync Kurulumu

1. [Firebase Console](https://console.firebase.google.com)'da yeni proje oluştur
2. Firestore Database'i etkinleştir (test modu)
3. Authentication → Anonim girişi etkinleştir
4. Proje ayarlarından Firebase config'i kopyala
5. `js/sync.js` içindeki `firebaseConfig` değişkenini güncelle
6. Siteyi deploy et → tüm cihazlardan aynı veriye eriş

## Teknolojiler

- **SheetJS** — Excel okuma/yazma
- **Firebase Firestore** — Çok cihazlı sync (ücretsiz tier)
- **Vanilla JS** — Sıfır bağımlılık
- **CSS Custom Properties** — Tema desteği
- **localStorage** — Çevrimdışı cache
- **GitHub Pages** — Ücretsiz hosting

## Lisans

MIT
