[app]

# ---------------- BASIC ----------------

title = GST Billing

package.name = gstbilling

package.domain = org.gstbilling

source.dir = .

source.include_exts = py,png,jpg,jpeg,kv,xlsx

version = 1.0

# ---------------- REQUIREMENTS ----------------

requirements = python3,kivy,openpyxl

# ---------------- SCREEN ----------------

orientation = portrait

fullscreen = 1

# ---------------- ICON ----------------

icon.filename = icon.png

# ---------------- ANDROID ----------------

android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE

android.api = 33

android.minapi = 21

android.sdk = 33

android.ndk = 25b

android.accept_sdk_license = True

android.archs = arm64-v8a, armeabi-v7a

# ---------------- SPLASH ----------------

presplash.color = #101114

# ---------------- APP SETTINGS ----------------

window.softinput_mode = below_target

# ---------------- SOURCE EXCLUDE ----------------

source.exclude_dirs = tests, bin, venv, .git, __pycache__

source.exclude_patterns = *.pyc

# ---------------- BUILD ----------------

[buildozer]

log_level = 2

warn_on_root = 1