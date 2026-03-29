# DataCopilot — Offline Desktop Excel AI Assistant
### 离线桌面 Excel AI 助手

> Talk to your spreadsheets in plain Chinese or English. No internet. No cloud. No installation.
> 用自然语言对话操作表格，完全离线，数据不出本机。

---

## What It Does / 功能简介

DataCopilot lets non-technical users query and transform Excel/CSV files using plain language — no formulas, no coding required. You describe what you want, the AI generates SQL, and the result appears instantly.

**Examples / 示例指令：**
- `毛利率的英文是什么，种类是什么？`
- `Show me all rows where sales > 10000`
- `统计各部门的平均薪资，按从高到低排列`
- `Filter the table to only keep rows from Q4 2024`
- `每个类别出现了多少次？`

---

## Key Features / 核心特性

| Feature | Detail |
|---|---|
| **100% Offline** | No Wi-Fi required. All data stays on your machine. |
| **Natural Language** | Supports Chinese and English input |
| **Auto SQL Generation** | AI converts your instruction to DuckDB SQL |
| **Self-Correction** | If SQL fails, the AI automatically retries up to 2 times |
| **Smart Output** | Small results (≤20 rows) shown in chat; large results saved to Excel |
| **Zero Install** | Double-click the `.exe`, no Python or dependencies needed |
| **Low Resource** | Runs on 8GB RAM laptops with no GPU (tested on Surface Pro 8 i5-1135G7) |

---

## System Requirements / 系统要求

| Item | Minimum |
|---|---|
| OS | Windows 10 / 11 (64-bit) |
| RAM | 8 GB |
| Storage | 2 GB free space |
| CPU | x64, AVX2 support (Intel 4th gen+ / AMD Zen+) |
| GPU | Not required |
| Internet | Not required |

---

## Folder Structure / 目录结构

```
DataCopilot/
├── DataCopilot.exe          # Main application (double-click to run)
├── model/
│   └── qwen2.5-coder-1.5b-instruct-q4_k_m.gguf   # AI model (1.1 GB)
└── engine/                  # Bundled inside the .exe (no action needed)
    ├── llama-server.exe
    ├── llama.dll
    ├── ggml.dll
    └── ggml-cpu-*.dll  (and other DLLs)
```

> **Note:** The `model/` folder must be in the same directory as `DataCopilot.exe`.

---

## Quick Start for End Users / 小白使用指南

### Step 1 — 你只需要这两样东西

从开发者处拿到压缩包解压后，你只需要关心以下两个文件/文件夹：

```
DataCopilot/
├── DataCopilot.exe      ← 主程序，双击运行
└── model/               ← AI 模型文件夹，必须和 .exe 放在一起
    └── qwen2.5-coder-1.5b-instruct-q4_k_m.gguf
```

> **重要：** `model/` 文件夹必须和 `DataCopilot.exe` 放在同一个目录下，不能移动或删除。其他文件夹不用管。

### Step 2 — 双击运行

双击 `DataCopilot.exe`。

- 首次启动需要约 **10~15 秒**加载 AI 模型，右上角会显示 `Loading AI model...`
- 看到右上角变为绿色的 **`AI Ready`** 后，即可使用

> 如果杀毒软件弹窗拦截，选择"允许运行"即可。这是正常现象，因为程序包含一个本地 AI 引擎。

### Step 3 — 加载你的表格

点击蓝色按钮 **"Open File / 打开文件"**，选择你的 `.xlsx`、`.xls` 或 `.csv` 文件。

加载成功后，聊天框会显示表格的行数、列数和列名信息。

### Step 4 — 用自然语言提问

在底部输入框输入你的问题，按 **Enter** 发送。

**可以这样问：**
```
毛利率的英文是什么？
统计各部门的平均薪资，按从高到低排列
只保留销售额大于 10000 的行
2024年Q4的数据有哪些？
每个类别分别出现了多少次？
```

**结果说明：**
- 结果在 **20 行以内** → 直接显示在聊天框
- 结果超过 **20 行** → 自动在你的表格同目录下生成 `Result_时间戳.xlsx` 文件

---

## How to Use / 使用方式（详细）

1. **Double-click** `DataCopilot.exe`
2. Wait for the status bar to show **"AI Ready"** (first launch ~10–15 seconds while model loads)
3. Click **"Open File / 打开文件"** and select your `.xlsx`, `.xls`, or `.csv` file
4. Type your question in the input box and press **Enter** (or click Send)
5. Results appear in the chat. For large results (>20 rows), an Excel file is auto-saved next to your source file

**Keyboard shortcut:**
- `Enter` — Send message
- `Shift + Enter` — New line in input box

---

## Output Behaviour / 输出规则

| Result Size | Output |
|---|---|
| ≤ 20 rows | Displayed directly in the chat window |
| > 20 rows | Saved as `Result_YYYYMMDD_HHMMSS.xlsx` next to your source file, with a 5-row preview in chat |

---

## What Is NOT Supported / 不支持的功能

To keep the AI fast and reliable on low-end hardware, the following are out of scope:

- **Charts / graphs** — No bar charts, pie charts, or visualizations
- **Cell formatting** — No font colors, bold, merged cells
- **Cross-file operations** — Only one file at a time (V1.0)
- **Subjective analysis** — Cannot answer "why did sales drop?" type questions
- **Formula writing** — Cannot write Excel formulas

---

## Performance / 性能指标

| Metric | Value |
|---|---|
| AI response time | 3–8 seconds (i5-1135G7, CPU only) |
| Peak RAM usage | ~1.5 GB |
| Model size | 1.1 GB (Q4_K_M quantized) |
| Total package size | ~1.2 GB |

---

## Tech Stack / 技术架构

```
┌─────────────────────────────────────────────┐
│  DataCopilot.exe  (Python + Tkinter UI)     │
│                                             │
│  ┌─────────────┐    ┌────────────────────┐  │
│  │  Pandas     │    │  requests          │  │
│  │  (load file │    │  (HTTP to local    │  │
│  │  into RAM)  │    │   AI server)       │  │
│  └──────┬──────┘    └────────┬───────────┘  │
│         │                    │              │
│  ┌──────▼──────┐    ┌────────▼───────────┐  │
│  │  DuckDB     │    │  llama-server.exe  │  │
│  │  (SQL on    │    │  (Qwen2.5-Coder    │  │
│  │  DataFrame) │    │   1.5B, port 8080) │  │
│  └─────────────┘    └────────────────────┘  │
└─────────────────────────────────────────────┘
```

| Component | Technology |
|---|---|
| UI | Python 3.10 + Tkinter |
| Data loading | Pandas + openpyxl |
| SQL engine | DuckDB (in-memory, queries Pandas DataFrame directly) |
| AI model | Qwen2.5-Coder-1.5B-Instruct (GGUF Q4_K_M, ~1.1 GB) |
| AI runtime | llama.cpp (`llama-server.exe`) |
| Packaging | PyInstaller (clean conda env) |

---

## For Developers / 开发者指南

### Run from source / 源码运行

**Prerequisites:**
```bash
# Create a clean environment (recommended)
conda create -n datacopilot python=3.10 -y
conda activate datacopilot
pip install pandas duckdb requests openpyxl
```

**Run:**
```bash
python main.py
```

**Project layout:**
```
excel helper/
├── main.py              # All application code (~430 lines)
├── engine/
│   ├── llama-server.exe
│   └── *.dll            # llama.cpp runtime DLLs
├── model/
│   └── qwen2.5-coder-1.5b-instruct-q4_k_m.gguf
├── plan/
│   ├── prd.md           # Product requirements document
│   └── 产品技术架构.md    # Technical architecture document
├── build.bat            # One-click build script
└── README.md
```

### Build the .exe / 打包

```bash
conda activate datacopilot
pip install pyinstaller
```

Then run the build script:
```
build.bat
```

Or manually:
```bash
pyinstaller --noconsole --onefile --name DataCopilot --add-data "engine;engine" main.py
```

> **Important:** Always build from the clean `datacopilot` conda environment, not from the base Anaconda environment. Building from Anaconda base will pull in torch, scipy, and other unneeded packages, causing the build to take 30+ minutes and produce a 1GB+ exe.

After building, copy the `model/` folder next to `dist/DataCopilot.exe`:
```
dist/
├── DataCopilot.exe   (~59 MB)
└── model/
    └── qwen2.5-coder-1.5b-instruct-q4_k_m.gguf
```

Zip the `dist/` folder for distribution (~1.2 GB total).

### Download the model / 下载模型

The model file is not included in this repository. Download it from Hugging Face:

- Model: `Qwen2.5-Coder-1.5B-Instruct`
- File: `qwen2.5-coder-1.5b-instruct-q4_k_m.gguf`
- Place in: `model/` folder

### Download llama.cpp engine / 下载推理引擎

Download the Windows prebuilt release from:
- https://github.com/ggerganov/llama.cpp/releases
- Choose: `llama-bXXXX-bin-win-avx2-x64.zip`
- Copy **all `.dll` files** + `llama-server.exe` into the `engine/` folder

---

## Architecture Decisions / 架构决策说明

**Why DuckDB + Pandas instead of direct DuckDB file reading?**
DuckDB cannot read `.xlsx` files offline without downloading an extension. Instead, Pandas loads the file into memory as a DataFrame, then DuckDB queries the DataFrame directly via Apache Arrow shared memory — zero data duplication, millisecond query speed.

**Why llama-server.exe as a separate process?**
Running the AI in a separate C++ process physically isolates it from the Python UI process. Neither can freeze the other. The server exposes an OpenAI-compatible HTTP API on `localhost:8080`.

**Why Qwen2.5-Coder-1.5B?**
It is the smallest model that reliably generates correct SQL and understands Chinese. At Q4_K_M quantization it fits in ~1.1 GB on disk and uses ~900 MB RAM at inference time, leaving headroom for the rest of the application on an 8 GB machine.

---

## License

Apache License 2.0

Copyright 2026 DataCopilot Contributors

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
