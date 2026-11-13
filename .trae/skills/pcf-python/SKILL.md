---
name: pcf-python
description: PCF_Python 项目速查——模块地图、366 张全量渲染回归、ALIAS.asc 符号库同步、PyInstaller 打包 exe、中文一行式提交规范。Use when 用户要求回归渲染、同步符号库、打包 exe、提交代码或了解本项目结构。Do not use for 与本项目无关的通用 Python 问题。
---

# PCF_Python 项目速查

PCF → 单线轴测图（ISO）渲染器：`pcf/` 渲染引擎 + ttkbootstrap GUI + PyInstaller 单文件 exe。

## 模块地图

- `pcf/render.py` — 渲染引擎（符号调度、接图标注、尺寸、BOM/切割表、预览 SceneModel 双后端）
- `pcf/continuations.py` — 接图点判定：**纯声明驱动**，每条 END-CONNECTION-PIPELINE = 一个接图点（CO-ORDS=坐标、CONNECTION-REFERENCE=名字），无几何匹配
- `pcf/asc_symbols.py` + `symbols/ALIAS.asc` — Isogen .ASC 符号库单源（315 符号，先到先得）+ 阀体→阀柄 skey 厂商配对
- `pcf_renderer.py` — CLI 批量渲染与 GUI 预览适配层（render_set 按目录缓存整套文档）
- `pcf_app.py` — GUI（管线清单勾选出图、矢量预览）
- `PCF收集器.spec` + `version_info.py` — PyInstaller 打包配置与版本信息

## 全量渲染回归（改动渲染链后必跑）

```powershell
py -3 -X utf8 pcf_renderer.py "c:\新建文件夹\HQSH_Iso_Pipeline_PTFE" --out-dir preview
```

- 366 张，判定通过 = exit 0 且日志无 Traceback/Error
- 警告基线（全为既有类别，新增类别须解释）：87 continuation（声明指向本交付集外线号）、19 不连通管段组、9 切割表截断、2 尺寸过小
- 测试语料 `c:\新建文件夹\HQSH_Iso_Pipeline_PTFE` 为**只读**厂商数据，勿改

## 符号库同步（vendor 更新 BSF.asc 后）

1. vendor 源：`d:\Users\jiazheng100002\Desktop\新建文件夹 (2)\BSF.asc`（勿动源文件）
2. 复制为 `symbols/ALIAS.asc`；`_prev_keys.json` 存符号全集基线，同步后对差集即新增清单，更新基线
3. 跑全量回归确认零异常

## 打包 exe

```powershell
py -3 -m PyInstaller "PCF收集器.spec" --noconfirm
```

- onefile → `dist/PCF收集器.exe`；symbols/ 目录、ttkbootstrap 全量已随 spec 打入；pymupdf/fitz 显式排除
- 版本号在 `version_info.py`（当前 1.0.0.0，升级时同步四处 filevers/prodvers/FileVersion/ProductVersion）
- 打包前先跑全量回归

## 提交规范

- 中文一行式 commit message，聚焦"为什么"；只 `git add` 具体文件，禁止 `git add -A`
- `.gitignore` 已覆盖 preview/、dist/、build/、__pycache__/、*.pdf、.screenshots/——构建产物与回归输出不入库
- 深层历史规则（符号语义、渲染细节、用户定稿决策）在 Trae 记忆 `project_memory.md`，不在本 skill 重复
