[app]

title = Smart Attendance Manager
package.name = attendanceapp
package.domain = org.smartattendance

source.dir = .
source.include_exts = py,png,jpg,kv,atlas

version = 1.0

requirements = python3,kivy,openpyxl

orientation = portrait
fullscreen = 0

icon.filename = %(source.dir)s/icon.png

android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE

android.api = 35
android.minapi = 21
android.accept_sdk_license = True

[buildozer]

log_level = 2
warn_on_root = 1
