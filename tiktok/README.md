# TikTok-план Amber Galbin · жовтень 2026

- `Amber_Galbin_TikTok_October_2026.xlsx` — готовий файл (імпортується в Google Таблиці).
- `content_october.py` — весь контент: рубрики, свята жовтня–грудня, референси, 16 сценаріїв.
  Критерій свят: церковні (новий календар ПЦУ/УГКЦ), загальнодержавні та професійні дні платоспроможних професій.
- `build_xlsx.py` — збирає таблицю (основний формат): календар жовтня з відео (клік веде на сценарій) і календарі листопада й грудня зі святами; далі сценарії, рубрики, свята, референси, заміри. Свята в календарі підтягуються з вкладки «Свята» формулою.
- `build_doc.py` — запасний варіант: Google Документ (HTML для імпорту) з тим самим контентом.
- `repack.py` — щільно перепаковує .xlsx (zopfli), щоб файл легше завантажувався на Drive.

Перезібрати:

```
pip install openpyxl zopfli
cd tiktok
python3 build_xlsx.py Amber_Galbin_TikTok_October_2026.xlsx
python3 repack.py Amber_Galbin_TikTok_October_2026.xlsx Amber_Galbin_TikTok_October_2026.xlsx
```

Документ:

```
cd tiktok
python3 build_doc.py Amber_Galbin_TikTok_October_2026.html
```

Шрифти таблиці: Montserrat (заголовки) і Roboto (текст). Lexend не використовувати: у ньому немає кирилиці. У Google Docs шрифт при імпорті стає Arial.
