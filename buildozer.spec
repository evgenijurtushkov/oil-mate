[app]
title = OIL MATE
package.name = oilmate
package.domain = ru.urtushkov
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,xlsx,ttf
source.include_patterns = data/*.xlsx, *.png, *.json, *.ttf
version = 1.0.2

requirements = python3,kivy==2.2.1,kivymd==1.1.1,pillow,openpyxl,jdcal,et_xmlfile

orientation = portrait
android.permissions = WRITE_EXTERNAL_STORAGE, READ_EXTERNAL_STORAGE, INTERNET, MANAGE_EXTERNAL_STORAGE
android.api = 33
android.minapi = 21
android.sdk = 33
android.ndk = 25b
android.archs = arm64-v8a
android.meta_data = data=data

presplash.filename = 
android.presplash_color = #000000

[buildozer]
log_level = 2
warn_on_root = 1
