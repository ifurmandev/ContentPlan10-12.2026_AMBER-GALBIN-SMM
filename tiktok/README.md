# TikTok-план Amber Galbin · жовтень 2026

- `Amber_Galbin_TikTok_October_2026.xlsx` — готовий файл (імпортується в Google Таблиці).
- `content_october.py` — весь контент: рубрики, свята, референси, 16 сценаріїв.
- `build_doc.py` — збирає Google Документ (HTML для імпорту): календар, з якого кожне відео веде на свою картку сценарію, рубрики, свята, референси, заміри. Основний формат.
- `build_xlsx.py` — збирає таблицю: календар, сценарії, рубрики, свята, референси, заміри.
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

У Google Docs шрифт при імпорті стає Arial. Lexend не використовувати: у ньому немає кирилиці.
