
# Excel 表格合并工具 📊⇨📈

![Python版本](https://img.shields.io/badge/Python-3.7%2B-blue)
![GitHub stars](https://img.shields.io/github/stars/your-username/python_Excel_merge_tool?style=social)

一款自动化合并多Excel文件的工具，专为处理表头相同的表格设计，支持.xlsx和.xls格式

## 🚀 核心功能

- **智能合并**：自动识别相同表头的多个Excel文件
- **格式保留**：保持原单元格样式和数据类型
- **批处理**：支持整个文件夹一键合并

## 🛠️ 技术实现

```python
# 核心合并逻辑示例
def merge_excels(file_list, output_path):
    combined_df = pd.DataFrame()
    for file in file_list:
        df = pd.read_excel(file)
        combined_df = pd.concat([combined_df, df], ignore_index=True)
    combined_df.to_excel(output_path, index=False)
```

**技术栈**：
- pandas（数据处理）
- openpyxl（Excel操作）
- tkinter（GUI界面）

## 📦 使用方式

### 图形界面操作
1. 双击运行 `excel合并工具.exe`
2. 选择包含Excel的文件夹
3. 设置输出文件路径
4. 点击"开始合并"按钮

### 命令行操作
```bash
python excel合并工具_code.py -i "输入文件夹路径" -o "输出文件.xlsx"
```

## 📂 文件结构
```
python_Excel_merge_tool/
├── excel合并工具.exe      # 可执行程序
├── excel合并工具_code.py  # 源代码
```

## 💡 典型应用场景

1. **月度报表合并**：将30个分表合并为季度总表
2. **多部门数据汇总**：整合销售/财务/生产等部门报表
3. **科研数据处理**：合并实验重复数据

## 🚨 注意事项

- 确保所有文件表头完全一致
- 建议提前备份原始文件
- 合并10万+行数据可能需要2分钟
- 中文路径需使用Python 3.7+


## 📄 开源协议
[MIT License](LICENSE)

---

⭐ **如果这个工具节省了您的时间，请点个Star！**  
🐛 **问题反馈**：附上样例文件和错误日志截图