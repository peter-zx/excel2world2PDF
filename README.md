# 合同自动填写工具

读取Excel表格信息，自动批量填写到Word合同模板中，完美保留原格式。

## 技术架构

| 模块 | 技术 | 说明 |
|------|------|------|
| WebUI | Streamlit | 零前端代码，快速构建交互界面 |
| Word处理 | python-docx + docxtpl | 占位符提取 + 模板渲染 |
| Excel处理 | pandas + openpyxl | 高效处理表格数据 |
| 数据存储 | JSON | 模板配置持久化 |

## 核心设计

### 占位符格式

在Word模板中使用 `【变量名】` 格式标记填写位置：

```
姓名：【姓名】
身份证号：【身份证号】
入职日期：【入职日期】
```

### 工作流程

```
┌─────────────────────────────────────────────────────────────────┐
│  步骤1: 模板管理                                                  │
│  上传Word → 自动提取【】占位符 → 配置变量映射 → 保存模板           │
├─────────────────────────────────────────────────────────────────┤
│  步骤2: 数据导入                                                  │
│  选择模板 → 上传Excel → 预览数据 → 确认列映射                      │
├─────────────────────────────────────────────────────────────────┤
│  步骤3: 批量生成                                                  │
│  执行生成 → 下载ZIP压缩包                                         │
└─────────────────────────────────────────────────────────────────┘
```

## 项目结构

```
world2pdf/
├── venv/                      # 虚拟环境
├── src/
│   ├── config.py              # 配置文件
│   ├── models/
│   │   └── schemas.py         # 数据模型
│   ├── services/
│   │   ├── word_service.py    # Word处理服务
│   │   ├── excel_service.py   # Excel处理服务
│   │   └── template_service.py # 模板管理服务
│   └── storage/               # 文件存储
│       ├── templates/         # 模板文件
│       ├── configs/           # 模板配置JSON
│       └── outputs/           # 生成的合同
├── samples/                   # 示例文件（生成后）
├── app.py                     # Streamlit主应用
├── generate_samples.py        # 生成示例文件脚本
└── requirements.txt           # 依赖文件
```

## 快速开始

### 1. 激活虚拟环境

Windows PowerShell:
```powershell
.\venv\Scripts\Activate.ps1
```

Windows CMD:
```cmd
.\venv\Scripts\activate.bat
```

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 3. 生成示例文件

```bash
python generate_samples.py
```

这将生成：
- `samples/合同模板示例.docx` - 包含【】占位符的模板
- `samples/员工数据示例.xlsx` - 示例员工数据

### 4. 启动Web应用

```bash
streamlit run app.py
```

浏览器自动打开 `http://localhost:8501`

## 使用指南

### 步骤1：模板管理

1. 准备Word合同模板，使用 `【姓名】` 格式标记
2. 在WebUI上传模板
3. 系统自动提取占位符
4. 配置变量映射（Excel列名 → 模板变量）
5. 保存模板

### 步骤2：数据导入

1. 选择已保存的模板
2. 上传Excel数据文件
3. 预览数据，确认列映射正确

### 步骤3：批量生成

1. 点击"开始批量生成"
2. 下载ZIP压缩包

## 后续迭代方向

- [ ] 支持PDF模板处理
- [ ] 支持表格循环（如合同明细表）
- [ ] 多用户系统
- [ ] 模板版本管理
- [ ] API接口（迁移到FastAPI）
