[app]
title = NB综合查询
package.name = nbquery
package.domain = org.hczh

source.dir = .
source.include_exts = py,xlsx
source.main = main.py

version = 1.0

requirements = python3==3.12.1,kivy,openpyxl

presplash.filename =
icon.filename =

fullscreen = 0

android.permissions =
android.api = 33
android.minapi = 21
android.ndk = 25b
android.build_tools_version = 33.0.1
android.archs = arm64-v8a

[buildozer]
log_level = 2
warn_on_root = 1
