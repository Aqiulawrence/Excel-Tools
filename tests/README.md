# ExcelSearch 测试与核对

在 Windows、Python 3.12 和项目根目录下运行。只测试搜索功能时，安装：

```powershell
python -m pip install -r requirements-test.txt
python -m unittest discover -s tests -p "test_excel_search*.py" -v
```

测试通过 `unittest` 运行，工作簿由 openpyxl/xlwt 写入临时目录，结束后自动清理。
Qt 使用离屏窗口，进程池使用真实的 Windows spawn 子进程。测试不依赖 Git、旧版脚本、
本机业务文件、Excel 安装或注册表中的个人设置。请使用上面的命令启动，
不要将测试入口通过标准输入管道传给 Python，Windows 子进程无法重新加载该入口。

## 自动测试

| 文件 | 覆盖内容 |
| --- | --- |
| `test_excel_search_results.py` | 固定地址和值、类型、公式缓存、隐藏和合并单元格、自定义列、停止检查 |
| `test_excel_search_completeness.py` | 错误工作表范围、特殊文件名、无法访问目录、XLS 类型、异常配置和数值文本 |
| `test_excel_search_worker.py` | 不同并发数、进程池复用、异常恢复、停止重搜、文件增删改、失败计数 |
| `test_excel_search_ui.py` | 结果表、进度、完成状态、弹窗、搜索记录和关闭流程 |

离屏测试不能替代真实桌面 Excel 自动定位或打包 EXE 的验收。
测试针对已保存的单元格值，不证明公式缓存最新或价格推断一定正确。

## 手动性能测量

```powershell
python tests/benchmark_excel_search.py --source ExcelSearch.py --folder "C:\ExcelSamples" --workers 8 --keyword JOY --repeats 3 --report benchmarks/local-performance.json
```

使用真实搜索线程和离屏结果表；计时包含目录收集、进程、读取、通信与结果插入。
默认在多轮搜索之间复用进程池，`--ephemeral-pool` 可用于逐轮销毁进程池的对照。
比较两个版本时，`--source` 分别指向各自的 Python 脚本；按顺序交替运行，
保持文件、关键词、并发数和机器负载一致，并核对完整结果指纹。
操作系统文件缓存会影响耗时。该测量不包含程序启动和真实屏幕绘制，不等同于 EXE 性能。

## 独立完整性核对

```powershell
python tests/audit_excel_search.py "C:\ExcelSamples" benchmarks/local-text-audit.json --keywords JOY USD 品名 --workers 1 12 16
python tests/audit_excel_xml_ranges.py "C:\ExcelSamples" benchmarks/local-xml-audit.json
```

第一个工具直接读取 XLSX/XLSM 的 XML 和 XLS 的 BIFF 文本值，按文件、工作表、地址、内容及重复次数，
与真实搜索结果比较。关键词应使用文字；它的独立预期结果不包含数字、日期、布尔值的格式化，
因此用数字关键词可能出现预期之外的差异。`--workers` 按关键词循环分配。
读取失败或结果不一致时以非零状态退出；无法访问的目录会记入报告，不能视为已检查。

第二个工具检查实际 XML 单元格是否超出声明范围，并统计公式保存结果；
它不运行公式，也不验证公式结果是否最新。

这些工具只读取指定文件夹，不修改输入工作簿；但报告包含本机路径，完整性核对报告还可能包含
命中内容。输出放在被忽略的 `benchmarks/` 中，分享前自行检查。

不要用 `pack.py` 运行测试：它会打包并发布版本。
