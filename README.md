# Excel Tools

面向 Windows 的 Excel 辅助工具。

## 项目结构

| 路径 | 用途 |
| --- | --- |
| `ExcelSearch.py` | 在文件夹内搜索 Excel 单元格，展示品名和价格信息 |
| `ExcelCompare.py` | Excel 内容对比 |
| `ImageSearch.py` | 图片搜索与插入 Excel |
| `Updater.py` | 客户端更新程序 |
| `UsageReporter.py` | 应用退出时发送使用记录 |
| `server.py` | Flask 更新与使用记录接口 |
| `pack.py` | Windows 打包、归档与发布脚本 |
| `tests/` | ExcelSearch 自动测试、性能测量和完整性核对工具 |
| `legacy/` | 旧版源码 |

## ExcelSearch 搜索范围

- 支持 `.xls`、`.xlsx`、`.xlsm`，跳过 `~$` 开头的 Excel 锁文件。
- 按保存的单元格值进行不区分大小写的子串搜索，包含隐藏工作表。
- 不重新计算公式；公式结果取决于文件最后保存的计算值。
- 不模拟所有自定义显示格式，也不搜索批注、图片或形状。
- 品名、币种列和单价标记通过规则推断，复杂表格仍需人工核对。
- 目录无法访问时不弹窗，详情保留在进度条悬停提示中；文件读取失败和搜索异常会提示。

## 测试与性能检查

参见 [tests/README.md](tests/README.md)。测试用临时生成的工作簿，无需提供个人 Excel 文件，
也不依赖 Git 历史版本。测试依赖可单独安装：

```powershell
python -m pip install -r requirements-test.txt
python -m unittest discover -s tests -p "test_excel_search*.py" -v
```

本地报告统一输出到 `benchmarks/`，该目录不提交；报告可能包含文件名、单元格内容和本机路径。

## 服务端与发布

服务端依赖：`python -m pip install -r requirements-server.txt`。
部署前检查 `server.py`、`Updater.py` 和 `UsageReporter.py` 中的服务地址，并按自己的部署配置版本和凭据。

`pack.py` 需要另行安装 PyInstaller。**该脚本不仅打包，还会调用发布接口更新服务器版本**；
运行前应阅读脚本并确认目标服务。它会重新生成 `dist/`，归档写入 `release/`。
`release_token.txt` 是本地发布凭据，禁止提交。

正常启动桌面应用时，会尝试登记已有 `Updater.exe` 的启动项；退出时会发送使用记录。
使用记录可能包含搜索关键词、文件夹、结果概况，以及复制或打开文件的操作信息。
自动测试和核对工具不执行主入口，不登记启动项，也不发送这些记录。

许可见 [LICENSE](LICENSE)。
