# Temizlik Sunum Sihirbazı — Vega Taurus AVM

Temizlik biriminin aylık faaliyet sunumunu adım adım soru sorarak hazırlayan masaüstü programı.

- `uygulama/` — programın kendisi. Temizlik bilgisayarındaki program her açılışta bu klasörün
  `surum.json` dosyasına bakar; sürüm numarası büyükse `dosyalar` listesindekileri indirip günceller.
- `baslatici.py` — `Temizlik Sunum.exe`'nin kaynağı (sadece güncelleme + başlatma).
