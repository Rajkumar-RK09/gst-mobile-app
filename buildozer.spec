[app]

title = GST Billing

package.name = gstbilling

package.domain = org.gstbilling

source.dir = .

source.include_exts = py,png,jpg,jpeg,kv,xlsx

version = 1.0

requirements = python3==3.10.11,kivy==2.3.0,openpyxl

orientation = portrait

fullscreen = 1

icon.filename = icon.png

android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE

android.api = 33

android.minapi = 21

android.ndk = 25b

android.accept_sdk_license = True

android.archs = arm64-v8a, armeabi-v7a

presplash.color = #101114

window.softinput_mode = below_target

source.exclude_dirs = tests, bin, venv, .git, __pycache__

source.exclude_patterns = *.pyc


[buildozer]

log_level = 2

warn_on_root = 1