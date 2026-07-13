# -*- coding: utf-8 -*-
# UTF-8
#
# VSVersionInfo syntax: https://stackoverflow.com/questions/65910
# This file is consumed by PyInstaller's `version=` parameter.

VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(1, 0, 0, 0),
    prodvers=(1, 0, 0, 0),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0),
  ),
  kids=[
    StringFileInfo(
      [
        StringTable(
          '040904B0',
          [
            StringStruct('FileDescription', 'PCF 管道组件文件收集器'),
            StringStruct('FileVersion', '1.0.0.0'),
            StringStruct('LegalCopyright', 'Copyright (c) 2026 jiazheng'),
            StringStruct('OriginalFilename', 'PCF收集器.exe'),
            StringStruct('ProductName', 'PCF 收集器'),
            StringStruct('ProductVersion', '1.0.0.0'),
          ]
        )
      ]
    ),
    VarFileInfo([VarStruct('Translation', [0x0804, 1200])])
  ]
)
