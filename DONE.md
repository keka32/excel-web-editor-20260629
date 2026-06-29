# excel-web-editor-20260629 — v2 Tamamlandı

## Yeni Özellikler (v2)
- 📂 Çoklu sekme sistemi (TabManager) — birden fazla Excel aynı anda
- 📱 Mobil düzenleme — dokunmatik uyumlu hücre seçimi + alt panel editör
- ⌨️ Alt gezinme çubuğu (mobil) — Düzenle, Aç, Kaydet, Yeni
- ☁️ Firebase Firestore sync — çok cihazlı gerçek zamanlı veri (opsiyonel)
- 🔄 Sekmeler arası otomatik kaydetme + localStorage cache
- 🎨 Gelişmiş sekme stilleri (aktif/pasif/dirty/kapat)

## Mimari
-  — TabManager: sekme ekle/çıkar/değiştir, state yönetimi
-  — CloudSync: Firebase Firestore anonim sync
-  — Ana uygulama: TabManager + mobil UX + klavye + export
-  — AutoSaveStorage: localStorage debounce + quota
-  — Dark/light tema toggle
-  — Sürükle-bırak dosya yükleme

## Dosyalar
- DONE.md
- QUALITY_REPORT.md
- README.md
- css/base.css
- css/table-editor.css
- index.html
- js/app.js
- js/sort-filter.js
- js/storage.js
- js/sync.js
- js/table-editor.js
- js/tabs.js
- js/theme.js
- js/uploader.js
- uploader-component.html