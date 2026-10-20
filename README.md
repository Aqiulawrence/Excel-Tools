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
| `legacy/` | 旧版源码 |

## ExcelSearch 搜索范围

- 支持 `.xls`、`.xlsx`、`.xlsm`，跳过 `~$` 开头的 Excel 锁文件。
- 按保存的单元格值进行不区分大小写的子串搜索，包含隐藏工作表。
- 不重新计算公式；公式结果取决于文件最后保存的计算值。
- 不模拟所有自定义显示格式，也不搜索批注、图片或形状。
- 品名、币种列和单价标记通过规则推断，复杂表格仍需人工核对。
- 目录无法访问时不弹窗，详情保留在进度条悬停提示中；文件读取失败和搜索异常会提示。

## 服务端与发布

服务端依赖：`python -m pip install -r requirements-server.txt`。
部署前检查 `server.py`、`Updater.py` 和 `UsageReporter.py` 中的服务地址，并按自己的部署配置版本和凭据。

`pack.py` 需要在运行它的 Python 环境中安装 PyInstaller 和 `cos-python-sdk-v5`。
运行后会重新生成 `dist/`，归档写入 `release/`，将 `dist/dist.zip` 上传至广州地域
`wublog-1328523094` 桶根目录的 `dist.zip`（覆盖上一版），上传成功后才调用发布接口更新服务器版本。
上传失败会保留本地压缩包并停止发布。使用简单上传，仅需 `cos:PutObject` 权限，最大支持 5GB。
COS 密钥从同目录的 `cos_credentials.json` 读取，字段为 `SecretId`、`SecretKey`，临时密钥还需 `Token`；
也支持 `COSSECRETID`、`COSSECRETKEY`、`COSSESSIONTOKEN` 环境变量，优先于配置文件。
`cos_credentials.json` 和 `release_token.txt` 是本地凭据，禁止提交。
`--publish-only` 仅重试更新服务器版本，不打包或上传，使用前应确认桶中的 `dist.zip` 已更新。

正常启动桌面应用时，会尝试登记已有 `Updater.exe` 的启动项；退出时会发送使用记录。
使用记录可能包含搜索关键词、文件夹、结果概况，以及复制或打开文件的操作信息。

许可见 [LICENSE](LICENSE)。
