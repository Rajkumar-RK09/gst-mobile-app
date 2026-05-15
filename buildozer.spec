[app]

title = GST Billing

package.name = gstbilling

package.domain = org.gstbilling

source.dir = .

source.include_exts = py,png,jpg,kv,xlsx

version = 1.0

requirements = python3,kivy,openpyxl

orientation = portrait

fullscreen = 0

android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE

android.api = 33
android.minapi = 21
android.sdk = 33
android.ndk = 25b
android.accept_sdk_license = True

presplash.color = #121212


[buildozer]

log_level = 2

warn_on_root = 1