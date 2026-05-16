[app]

# ---------------------------------------------
# APP INFO
# ---------------------------------------------

title = GST Billing

package.name = gstbilling

package.domain = org.gstbilling

source.dir = .

source.include_exts = py,png,jpg,jpeg,kv,xlsx

version = 1.0


# ---------------------------------------------
# REQUIREMENTS
# ---------------------------------------------

requirements = python3,kivy==2.3.0,openpyxl


# ---------------------------------------------
# DISPLAY
# ---------------------------------------------

orientation = portrait

fullscreen = 1

window.softinput_mode = below_target


# ---------------------------------------------
# ICON & SPLASH
# ---------------------------------------------

icon.filename = icon.png

presplash.color = #101114


# ---------------------------------------------
# ANDROID SETTINGS
# ---------------------------------------------

android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE

android.api = 33

android.minapi = 21

android.ndk = 25b

android.accept_sdk_license = True

android.archs = arm64-v8a, armeabi-v7a


# ---------------------------------------------
# EXCLUDE UNUSED FILES
# ---------------------------------------------

source.exclude_dirs = tests, bin, venv, .git, __pycache__

source.exclude_patterns = *.pyc


# ---------------------------------------------
# BUILDOZER
# ---------------------------------------------

[buildozer]

log_level = 2

warn_on_root = 1