# 进度跟踪报表工具

把按规则命名的**设计明细表**（地质设计 / 工程设计 / 工艺设计 …）自动同步进
**《井下作业设计节点跟踪大表》**，跨平台（Windows + macOS），GUI + CLI 双入口。

```
明细表 (*.xls, 按 <前缀><阶段>-<状态>.xls 命名)
        │  文件名 → 阶段 × 状态
        │  表头文本 → 井号 / 完成日期 / 当前审核人
        ▼
   ┌─────────────┐      diff CSV      ┌──────────────┐
   │  同步引擎    │ ─────────────────▶ │ 变更对照报告  │
   │ (决策矩阵)   │ ─────────────────▶ │ 直接写入原表  │  (--apply, 保格式)
   └─────────────┘                    └──────────────┘
        ▲
        │ 表头文本定位（不硬编码列号）
   跟踪大表 (.xls, 3 个 sheet)
```

---

## 1. 快速开始

```bash
# 1) 装依赖（xlrd 必须 1.2.0，2.0+ 读不了 .xls）
pip install -r requirements.txt

# 2) 图形界面
python main.py

# 3) 命令行（默认试运行，不动文件）
python main.py --cli \
    --source    "/path/to/井下作业报表生成" \
    --trackbook "/path/to/井下作业设计节点跟踪大表.xls" \
    --out-dir   ./output

# 4) 确认 diff 无误后写回
python main.py --cli ... --apply

# 5) 跑测试
python tests/test_core.py
```

> **xlrd 版本陷阱**：`xlrd >= 2.0` 移除了 .xls（BIFF8）支持。
> `main.py` 启动时会硬拦截并提示降级，不要装最新版。

---

## 2. 业务规则

### 2.1 文件名 → 阶段 × 状态

| 文件名片段 | 解析结果 | 从明细表取的值 | 写入跟踪大表 |
|-----------|---------|--------------|-------------|
| `…地质设计-已完成.xls` | stage=地质设计, status=已完成 | **完成日期** | 该阶段「实际完成」列 |
| `…地质设计-审核中.xls` | stage=地质设计, status=审核中 | **当前审核人** | 该阶段「实际完成」列，值形如 `张三（审核中）` |

阶段关键字见 `core/config.py:STAGE_KEYWORDS`，新增阶段**只追加该列表、不改代码**：

```python
STAGE_KEYWORDS = ["工程方案审查", "工艺设计", "工程设计", "地质设计",
                  "修前工程及概算", "修井作业开工", "修井作业完工"]
```

匹配细节：

- **长关键字优先**：正则按长度降序拼接，避免 `工程` 截胡 `工程方案审查`（见 `_compile()`）。
- 状态只有 `已完成` / `审核中` 两种。`已驳回 / 已退回 / 草稿 / 作废` 落入
  **状态存疑**桶，不自动处理，在日志里列出供人工确认。
- 跳过 `.` 开头的隐藏文件（macOS `.DS_Store`）与 `~$` 开头的 Excel 锁文件。

### 2.2 跟踪大表 → 列定位（严禁硬编码列号）

三个 sheet 的列布局**不一样**：`工艺设计` 只在 sheet 7 有；sheet 6/8 的 L 列是
`修前工程`。因此全部按**表头文本**定位：

| 行（0-based） | 内容 | 配置常量 |
|--------------|------|---------|
| R2 | 一级表头（阶段名） | `ROW_STAGE_HEADER` |
| R3 | 二级表头（计划完成 / 实际完成） | `ROW_SUB_HEADER` |
| R4+ | 数据 | `ROW_DATA_START` |

定位算法（`TrackSheet._locate_impl`）：

1. 在 R2 找 `井号` → 井号列；
2. 把 R2 按「连续非空」切成块 → `[(阶段名, [列号...]), ...]`；
3. 命中阶段名的块内，R3 == `实际完成` → 目标列；
4. 块内没有 `实际完成` 时，回退找 `当前进度`（v1.0 旧模板），并置
   `legacy_used=True`，日志会提示先执行列结构调整；
5. 都没有 → 返回 `None`，该 sheet 静默跳过该阶段。

### 2.3 冲突决策矩阵

`core/sync_engine.py:decide(payload, status, old, allow_recheck_overwrite)`
返回写入值，`None` = 跳过。

| # | 场景 | 旧值 | 新值 | 结果 |
|---|------|------|------|------|
| ① | 原值为空 / `/` | `""` `/` | 任意 | **写入** |
| ② | 审核中 → 已完成 | `张三（审核中）` | `2026-03-01` | **写入** |
| ③ | 已完成 → 审核中（回退） | `2026-03-01` | `李四（审核中）` | **跳过**（`--allow-recheck-overwrite` 可强制） |
| ④ | 幂等 | 同日期 / 同审核人 | — | **跳过** |
| ⑤ | 两个日期 | `2026-01-05` | `2025-12-20` | **保留较新者** |
| ⑥ | 明细表同井号多条 | — | — | 取**流程次数最大**者 |
| ⑦ | 同 sheet 同阶段重复井号 | — | — | 仅第一条写入，其余 `warn-dup-well` |
| ⑧ | 明细值为空 | `2026-03-01` | `""` | **跳过**（绝不清空已有日期） |
| 兜底 | 原值是脏数据 | `待定` / `2025/3/38` | 合法日期 | **覆盖** |

**⑦ 的去重是按阶段、不是按 sheet** —— 这是 v1.1 修掉的关键 bug：
同一口井合法地同时出现在「地质设计」和「工艺设计」两列，早期版本按 sheet 去重
导致工艺设计被误判为重复井号，产生 23 条假告警。

**脏日期判定**（`is_dirty_date`）除了格式校验，还做**日历校验**：
`2025/3/38` 能过正则但不是合法日期，会被识别为脏数据并覆盖——
已在真实数据上抓到一处（`6.上试井` 行6 卧85）。

---

## 3. 目录结构

```
井下报表工具/
├── main.py                 入口：依赖检查 + argparse（--cli / --auto / --apply / --restructure）
├── core/
│   ├── config.py           ★ 全部业务关键字、决策策略、行号常量
│   ├── pathutil.py         跨平台收敛：NFC 归一、短横线归一、casefold、非法字符
│   ├── filename_parser.py  文件名 → (stage, status)；scan_dir 分类
│   ├── detail_reader.py    明细表 → DetailRecord（按表头文本找列）
│   ├── trackbook.py        跟踪大表封装 + 列自适应定位 + 备份
│   ├── sync_engine.py      ★ 决策矩阵 + sync_sheet
│   ├── restructure.py      sheet 7 列结构调整（工艺设计补两列）
│   ├── writer.py           写回（xlwt 整表重建，改动的单元格标红）
│   ├── reports.py          diff CSV + 日志
│   └── runner.py           统一执行入口 run_sync(cfg, on_log, on_done, on_error)
├── gui/main_window.py      tkinter 界面
├── tests/test_core.py      152 条用例（A–J 分组），含端到端
├── build_mac.sh / build.bat / build_installer.bat / installer.iss / *.spec
└── .github/workflows/windows-build.yml
```

---

## 4. 关键设计约束

### 4.1 xlrd 1.2.0 钉死

`.xls` 是 BIFF8 老格式。`xlrd 2.0+` 只支持 `.xlsx`，`requirements.txt` 里写死
`xlrd==1.2.0`。`.xlsx` 明细表目前**不解析**（`load_detail` 直接返回空列表）。

### 4.2 原地写入与格式保全（`core/writer.py`）

xlwt 只有 `write(r, c, v)`，没有插入行/列的 API，也没有就地改单元格的能力。两种写回模式：

- **原地写入（默认，`write_inplace`，`inplace=True`）**
  用 `xlrd(formatting_info=True)` + `xlutils.copy` 把整份 .xls 克隆进内存，
  **只改目标单元格**（标红），再用 `os.replace` 原子替换原表。
  ✅ **合并单元格、字体、边框、列宽全部保留**，改动单元格标红方便扫一眼。
  **图形界面已极简化为唯一行为**：直接改原表、零新文件，不生成任何 `.bak` / diff / 日志。
  若需要 `.bak` / diff / 日志，仅命令行加 `--report`。
- **另存副本（谨慎模式，`write_back`，`inplace=False`）**
  用 xlwt 从头重建，产出 `.synced-{ts}.xls`，原表纹丝不动。
  ⚠️ 此模式会丢失合并单元格与条件格式，仅用于只读核对场景。

> v1.0 → v1.1 的关键变化：**默认不再是「另存 .synced 副本」，而是「直接写进原表并保格式」**。
> 这是本工具的核心承诺——用户给的跟踪大表结构复杂（多级合并表头），丢格式不可接受。
> v1.2 把**默认改成「真正零新文件」**：原地改原表、不落任何 `.bak` / diff / 日志。
> v1.3 进一步把图形界面**极简化**：所有运行选项（实际写入 / 谨慎模式 / 保留备份 / 允许覆盖 / 调整列）
> 全部从界面移除，点按钮即按规则直接写入原表；高级选项仅命令行可用。

**列结构调整**（restructure）仍是整表重建（`snapshot()` 读入 → 内存插列 → `write_snapshots()` 重写），
故那份 `_重列-` 文件会丢格式，仅作一次性整理用。

### 4.3 跨平台三处差异（`core/pathutil.py`）

| 差异 | 处理 |
|------|------|
| 路径分隔符 | 一律 `pathlib`，禁止硬编码 `\` 或 `/` |
| Unicode 归一化 | macOS 倾向 NFD、Windows 用 NFC → 比较前一律 `nfc()` |
| 大小写 | 一律 `casefold()` 比较 |

井号比较键 `canon()` = NFC → 剔除全角空格/NBSP → 全角破折号类归一为 `-` → casefold。
**全角破折号类是真实痛点**：`黄202H1－2`（U+FF0D）与 `黄202H1-2`（ASCII）在
Windows 和 macOS 上敲出来常常不一样。

### 4.4 GUI 用 tkinter，不用 PyQt6

对齐参考项目（报表转换工具）的既有打包链与 CI，避免引入 Qt 的授权与体积负担。

- macOS Aqua 下 `tk.Button` 会忽略 `bg`，所以主按钮用 **Frame + Label 自绘**
  实现 hover 变色（`gui/main_window.py:_accent_button`）。
- 后台线程 + `root.after(0, ...)` 切回主线程更新控件（`core/runner.py:run_sync`）。

---

## 5. 列结构调整（一次性）

v1.0 模板里 sheet 7 的「工艺设计」只有一列、二级表头叫 `当前进度`，
与地质设计/工程设计的「计划完成 + 实际完成」两列不一致。补列后：

| | 变更前（18 列） | 变更后（19 列） |
|---|---|---|
| L | 工艺设计 / **当前进度** | 工艺设计 / **实际完成** |
| M | 修前工程及概算 / 计划完成 | 工艺设计 / **计划完成**（新增） |
| N | 修前工程及概算 / 实际完成 | 修前工程及概算 / 计划完成 |
| O | … | 修前工程及概算 / 实际完成 |

```bash
python main.py --cli --restructure --out-dir ./output
```

产出 `井下作业设计节点跟踪大表_重列-{ts}.xls` + `restructure-report-{ts}.csv`
（列号 / 列字母 / 新旧表头对照）。**幂等**：已是标准结构时日志会提示无需调整，
重复执行不会再加列。

---

## 6. 输出物

| 文件 | 说明 | 是否默认生成 |
|------|------|-------------|
| `井下作业设计节点跟踪大表.xls`（被改的原表） | 原地写入结果，合并单元格/字体/边框保留、改动单元格标红 | ✅ **默认即改此文件** |
| `diff-{ts}.csv` | 变更对照：井号 / sheet / 行号 / 阶段 / 目标列 / 改前 / 改后 / 动作 / 说明。**UTF-8 BOM**，Excel 双击不乱码 | 仅 `--report` / 勾「保留备份与核对报告」 |
| `run-{ts}.log` | 完整运行日志 | 仅 `--report` |
| `*.bak-{ts}.xls` | 原地写入前的原表备份（在报告目录） | 仅 `--report` |
| `*.synced-{ts}.xls` | 谨慎模式（另存副本）的同步结果，原表不动、改动的单元格标红。⚠️ 丢失合并单元格/条件格式 | 仅「谨慎模式」 |

> **默认（`--apply` 不挂 `--report`）是「零新文件」**：源目录里只有被改的 `井下作业设计节点跟踪大表.xls`，
> 没有任何 `.bak` / diff / 日志。想留痕就加 `--report`。

---

## 7. 测试

```bash
python tests/test_core.py
```

152 条用例，10 个分组（A–J），全部在临时目录里造数据、不污染项目：

| 分组 | 覆盖 |
|------|------|
| A | 归一化：NFC / 全角破折号 / 全角空格 / 大小写 / None |
| B | 文件名解析：长关键字优先、非法字符、扩展名、状态存疑、锁文件跳过 |
| C | 日期：截断、补零、日历校验（闰年 / 越界） |
| D | **决策矩阵 8 条规则 + 兜底** |
| E | 明细聚合：流程次数最大、按阶段分组 |
| F | `sync_sheet`：**按阶段去重的回归**、无该阶段时静默跳过、幂等 |
| G | sheet 7 列改造：列数 18→19、数据右移、**重复改造幂等** |
| H | 报告：列字母、统计、BOM |
| I | **端到端**：造 .xls → 扫描 → 同步 → diff → 二次运行写入数归零 |
| J | **7 节点定位 + 原地保格式写回**：含换行归一化 / 二级表头变体、多源合并、节点过滤、`write_inplace` 合并单元格不变、标红、副本模式 |

---

## 8. 打包

### macOS（已验证，arm64，~37 MB）

```bash
# 一次性：建专用打包 venv（不动 envs/default，那里是 xlrd 2.0.2 会失败）
/Users/zoe/.workbuddy/binaries/python/versions/3.13.12/bin/python3 \
    -m venv /Users/zoe/.workbuddy/binaries/python/envs/wellsync-build
/Users/zoe/.workbuddy/binaries/python/envs/wellsync-build/bin/pip install \
    xlrd==1.2.0 xlwt openpyxl pyinstaller pillow

# 生成应用图标（蓝色圆角卡片 + 表格 + 绿色写入对勾）
/Users/zoe/.workbuddy/binaries/python/envs/wellsync-build/bin/python tools/make_icon.py

# 打包
./build_mac.sh
# 产物：dist/进度跟踪报表工具.app
# 运行：open "dist/进度跟踪报表工具.app"
```

打包脚本 `build_mac.sh` 按以下顺序探测 Python（首个可用即停）：

1. `$PYTHON_BIN` 环境变量
2. `./venv/bin/python`（项目内 venv）
3. `~/.../envs/wellsync-build/bin/python`（专用打包 venv）
4. `~/.../versions/3.13.12/bin/python3`（managed，xlrd 1.2.0）

**不会 fallback 到 `envs/default`**——那里是 xlrd **2.0.2**，会静默读不了 .xls。

Spec 文件 `进度跟踪报表工具.spec`（macOS）与 `进度跟踪报表工具_win.spec`（Windows）
维护**完全一致**的 hidden-imports
（`tkinter` / `tkinter.ttk` / `tkinter.filedialog` / `tkinter.messagebox` /
`threading` / `openpyxl` / `xlutils` / `xlutils.copy`）和 `--collect-submodules core`，
以及 macOS 的 bundle identifier `com.well-sync.app`。

> ⚠️ **PyInstaller 不能跨平台编译**——macOS 上只能出 `.app`，Windows 上才能出 `.exe`。
> 本机（macOS）无法生成 Windows 安装包；请在 Windows 机器或 CI 上执行下方步骤。

### Windows

```cmd
build.bat            :: PyInstaller 用 进度跟踪报表工具_win.spec 打包 → dist\进度跟踪报表工具.exe
build_installer.bat  :: 需先装 Inno Setup（iscc），产出 installer\进度跟踪报表工具_setup.exe
```

`requirements.txt` 已含 `xlrd==1.2.0 / xlwt / openpyxl / xlutils`，无需手动补装。

也可走 GitHub Actions（`.github/workflows/windows-build.yml`，push 或手动 `workflow_dispatch` 触发）：
Python 3.11 → 装依赖 → 校验 `xlrd < 2.0` 与源码布局 → PyInstaller 打包（含 `xlutils`/`xlutils.copy` hidden-import）→
Inno Setup 出安装包 → 上传 `windows-installer` artifact。

---

## 9. 已知限制 / 待办

- **`.xlsx` 明细表不解析**，只处理 `.xls`。需要时再引入 `openpyxl` 读路径。
- **默认原地写入已保格式**（xlutils 克隆），无需再用 WPS 手动替换；**默认零新文件**（不生成
  任何 `.bak` / diff / 日志，仅改原表）。GUI 已极简化为唯一行为，备份/副本/试运行/调整列
  等高级选项仅命令行可用（`--report` / `--copy-mode` / 不挂 `--apply` / `--restructure`）。
  仅「谨慎模式（另存 `.synced-` 副本）」与「列结构调整」会丢合并单元格/条件格式。
- **工艺设计的「计划完成」列目前不写入**（业务上未确认口径），只补了列。
- **未校验作业类型**：同一口井若在同 sheet 出现两次不同作业，目前按「仅第一条 +
  告警」处理，需人工判断。
- `工艺设计` 真实明细表到位后，需替换测试用的合成数据再跑一遍。
