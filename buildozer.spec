[app]
# 1. ОСНОВНЫЕ ДАННЫЕ ПРИЛОЖЕНИЯ ДЛЯ RUSTORE
title = WELL CONTROL
package.name = wellcontrol
package.domain = ru.urtushkov
version = 1.1.0
orientation = portrait
source.dir = .

# КОРРЕКТНЫЙ ЗНАЧКА ПРИЛОЖЕНИЯ НА ЭКРАНЕ ТЕЛЕФОНА
icon.filename = %(source.dir)s/icon.png

# 2. ВКЛЮЧЕНИЕ ВСЕХ ФАЙЛОВ И ТАБЛИЦ EXCEL
source.include_exts = py,png,jpg,kv,atlas,xlsx,ttf,json
source.include_patterns = data/*.xlsx, *.png, *.json, *.ttf

# 3. РАБОЧИЙ СТЕК REQUIREMENTS (БЕЗ РУЧНОГО УКАЗАНИЯ ВЕРСИЙ ПИТОНА)
requirements = python3, kivy, kivymd, pillow, openpyxl, jdcal, et_xmlfile, sqlite3

# ХАКИ ДЛЯ БЕЗОПАСНОЙ СБОРКИ В КОЛЛАБЕ
p4a.skip_pip_dependencies_check = True
android.add_compile_options = true

# 4. РАЗРЕШЕНИЯ (ДЛЯ КОРРЕКТНОГО ДОСТУПА К ФАЙЛАМ НА ANDROID 14)
android.permissions = INTERNET, READ_EXTERNAL_STORAGE, WRITE_EXTERNAL_STORAGE, MANAGE_EXTERNAL_STORAGE

# 5. СБОРКА ПОД СВЕЖИЕ ТРЕБОВАНИЯ GOOGLE PLAY И RUSTORE (ANDROID 14)
android.api = 34
android.minapi = 21
android.sdk = 34
android.ndk = 25b
# Отлаженная кровью и потом рабочая 64-битная архитектура
android.archs = arm64-v8a

# Экспорт папки с инженерными базами данных внутрь APK
android.meta_data = data=data

# Оформление экрана загрузки
presplash.filename = %(source.dir)s/background.png
presplash_color = #000000

# 6. ЖЕСТКАЯ ДИРЕКТИВА АВТОПОДПИСИ РЕЛИЗА КЛЮЧОМ (ДЛЯ RUSTORE)
# Теперь Buildozer сам в фоне заберет эти параметры, подпишет и выдаст готовый APK!
android.release_artifact = apk
android.keystore = %(source.dir)s/new_release.p12
android.keystore_type = PKCS12
android.keyalias = oilmate
android.keystoreresp = 120208
android.keypassresp = 120208


[buildozer]
log_level = 2
warn_on_root = 1




