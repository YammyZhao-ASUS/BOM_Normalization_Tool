# BOM Basic Tool

这是一个最简单可运行的 BOM 归一化版本。

它会读取 BOM Excel/CSV，识别常见的电阻与电容写法，并生成一个 Excel 报告：

- `Summary`：BOM 行数、相同规格组数、相近值组数；
- `Normalized BOM`：元件的归一化值；
- `Same Specification`：相同规格但使用不同料号的元件；
- `Near Values`：阻值或容值接近、需要人工确认的元件。

## 运行

安装依赖：

```powershell
python -m pip install -r requirements.txt
```

双击 `Run_BOM_Tool.bat`，或执行：

```powershell
python src/gui.py
```

在窗口中选择 BOM 文件，选择报告输出位置，然后点击 `Generate Report`。

也可以使用命令行：

```powershell
python src/main.py input/bom.xlsx --output output/BOM_Basic_Report.xlsx
```

## 建议输入列

工具会自动识别以下常见列名：

| 用途 | 示例列名 |
| --- | --- |
| 料号 | `Part Number`, `MPN`, `Item Number` |
| 位号 | `Part Reference`, `Reference`, `RefDes` |
| 描述 | `Component_Name`, `Description`, `Item Description` |
| 数量 | `Quantity`, `Qty` |
| 厂商 | `Vendor`, `Manufacturer` |

最容易取得好结果的描述格式例如：

```text
MLCC 100NF 16V 0402 X7R 10%
RES 10K OHM 1/16W 0402 1% THICK FILM
```

当前版本只做基础提示，不会自动替换料号。相近值、相同规格仍需要工程人员确认后再使用。

## 测试

```powershell
python -m unittest discover -s src -p "test_*.py"
```