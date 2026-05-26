[app]
# 1. ОСНОВНЫЕ ДАННЫЕ ПРИЛОЖЕНИЯ
title = OIL MATE
package.name = oilmate
package.domain = ru.urtushkov
version = 1.0.2
orientation = portrait
source.dir = .

# КОРРЕКТНЫЙ ЗНАЧКА ПРИЛОЖЕНИЯ НА ЭКРАНЕ ТЕЛЕФОНА
icon.filename = %(source.dir)s/icon.png

# 2. ВКЛЮЧЕНИЕ ВСЕХ ФАЙЛОВ И ТАБЛИЦ EXCEL
source.include_exts = py,png,jpg,kv,atlas,xlsx,ttf,json
source.include_patterns = data/*.xlsx, *.png, *.json, *.ttf

# 3. ВЧЕРАШНИЙ РАБОЧИЙ СТЕКRequirements (Без ручного указания версий питона)
requirements = python3, kivy, kivymd, pillow, openpyxl, jdcal, et_xmlfile, sqlite3


# ХАКИ ДЛЯ БЕЗОПАСНОЙ СБОРКИ В КОЛЛАБЕ
p4a.skip_pip_dependencies_check = True
p4a.extra_args = 



android.add_compile_options = true

# 4. РАЗРЕШЕНИЯ (Для корректного доступа к файлам на Android 14)
android.permissions = INTERNET, READ_EXTERNAL_STORAGE, WRITE_EXTERNAL_STORAGE, MANAGE_EXTERNAL_STORAGE

# 5. СБОРКА ПОД СВЕЖИЕ ТРЕБОВАНИЯ GOOGLE PLAY (Android 14)
android.api = 34
android.minapi = 21
android.sdk = 34
android.ndk = 25b
android.archs = arm64-v8a

# Экспорт папки с инженерными базами данных внутрь APK
android.meta_data = data=data

# Оформление экрана загрузки
presplash.filename = %(source.dir)s/background.png
presplash_color = #000000

[buildozer]
log_level = 2
warn_on_root = 1




