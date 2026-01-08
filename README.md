# Doc Generator

批量将 Excel 数据填充到 Word 模板，生成多个 Word 文档。

## 功能特性

- **Excel 数据读取**: 支持 .xlsx 格式，自动识别列名
- **Word 模板渲染**: 使用 `{{占位符}}` 语法，支持段落和表格
- **灵活的映射配置**: 支持直接映射和表达式映射
- **表达式计算**: 支持数学运算、字符串操作、条件判断等
- **批量生成**: 自动为每行数据生成一个 Word 文档
- **配置保存**: 可保存和加载映射配置

## 安装

### 从源码安装

```bash
# 克隆仓库
git clone https://github.com/yourusername/doc-generator.git
cd doc-generator

# 创建虚拟环境（推荐）
python -m venv venv
source venv/bin/activate  # Linux/macOS
# 或 venv\Scripts\activate  # Windows

# 安装依赖
pip install -r requirements.txt

# 安装项目
pip install -e .
```

### 运行

```bash
# 使用命令行入口
doc-generator

# 或直接运行模块
python -m doc_generator.main
```

## 使用说明

### 1. 准备 Word 模板

在 Word 文档中使用 `{{占位符名}}` 作为占位符：

```
尊敬的 {{姓名}}：

您的订单编号是 {{订单号}}，金额为 {{金额}} 元。
```

### 2. 准备 Excel 数据

Excel 第一行作为列名（表头），从第二行开始为数据：

| 姓名 | 订单号 | 金额 |
|------|--------|------|
| 张三 | A001   | 100  |
| 李四 | A002   | 200  |

### 3. 配置映射

在程序中：
1. 选择 Excel 文件和 Word 模板
2. 程序会自动识别 Excel 列和 Word 占位符
3. 配置映射关系（同名会自动匹配）
4. 设置输出目录和文件名模式
5. 点击"生成文档"

### 表达式示例

支持在映射中使用表达式：

```
# 数学运算
{{单价}} * {{数量}}

# 字符串拼接
concat({{姓}}, {{名}})

# 四舍五入
round({{金额}} * 1.1, 2)

# 条件判断
if({{金额}} > 1000, "大额", "普通")

# 空值处理
ifempty({{备注}}, "无")
```

### 支持的函数

| 函数 | 说明 | 示例 |
|------|------|------|
| `concat(a, b, ...)` | 连接字符串 | `concat({{姓}}, {{名}})` |
| `sum(a, b, ...)` | 求和 | `sum({{价格1}}, {{价格2}})` |
| `avg(a, b, ...)` | 平均值 | `avg({{分数1}}, {{分数2}})` |
| `round(x, n)` | 四舍五入 | `round({{金额}}, 2)` |
| `if(cond, t, f)` | 条件判断 | `if({{年龄}} >= 18, "成人", "未成年")` |
| `ifempty(v, d)` | 空值替换 | `ifempty({{备注}}, "无")` |
| `upper(s)` | 转大写 | `upper({{代码}})` |
| `lower(s)` | 转小写 | `lower({{邮箱}})` |
| `left(s, n)` | 左截取 | `left({{电话}}, 3)` |
| `right(s, n)` | 右截取 | `right({{身份证}}, 4)` |

### 文件名模式

可以使用占位符自定义输出文件名：

- `{{姓名}}_合同.docx` - 使用姓名列的值
- `{{_index}}_output.docx` - 使用行号（从1开始）
- `{{订单号}}_{{姓名}}.docx` - 组合多个列

## 打包为 EXE

```bash
# 安装 PyInstaller
pip install pyinstaller

# 使用 spec 文件打包
pyinstaller build.spec

# 生成的 exe 在 dist 目录
```

## 项目结构

```
doc-generator/
├── src/doc_generator/
│   ├── main.py              # 程序入口
│   ├── gui/
│   │   ├── main_window.py   # 主窗口
│   │   └── mapping_widget.py # 映射配置组件
│   ├── core/
│   │   ├── excel_reader.py  # Excel 读取
│   │   ├── word_renderer.py # Word 渲染
│   │   ├── expression.py    # 表达式解析
│   │   └── mapping.py       # 映射规则
│   └── utils/
│       └── config.py        # 配置管理
├── tests/                   # 测试文件
├── requirements.txt
├── pyproject.toml
└── build.spec              # PyInstaller 配置
```

## 依赖

- Python 3.10+
- PyQt6 - GUI 框架
- openpyxl - Excel 读取
- python-docx - Word 处理
- simpleeval - 安全的表达式计算

## License

MIT License
