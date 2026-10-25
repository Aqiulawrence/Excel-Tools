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
| `release.py` | Windows 打包、归档与发布脚本 |
| `.secrets/` | 本地发布令牌及 COS 凭据，不提交至 Git |
| `legacy/` | 旧版源码 |

## ExcelSearch 搜索范围

- “搜索方式”下拉框默认选择索引搜索，也可切换为原有常规搜索；会记住上次选择。
- `ExcelSearch.py` 内置目录扫描、`SearchEngine`、`DirectoryWatcher` 和 `LiveIndexService`，运行及打包不依赖 `ExcelSearchBenchmark.py`，可删除该测速文件。新用户先选择搜索目录，再自动开始后台核对和递归监听；已有用户启动时使用上次保存的有效目录，没有有效记录时等待用户选择，不默认扫描任何磁盘。切换为常规搜索时，后台索引仍保持监听。
- 固定复用一个 `ExcelSearch-index.sqlite3` 数据库，不提供新建数据库入口。缓存包含单元格内容；源文件需要可读，索引目录需要可写。SQLite 的 `-wal`、`-shm` 是同一数据库的工作文件。
- Windows 原生监听持续增量处理新增、修改、删除及目录重命名；通知溢出时补扫，监听中断后重连；暂时无法读取的文件按 1、3、10 秒策略重试。查询前先同步已收到的通知，日常查询不重新遍历整盘。
- 使用 SQLite / FTS5 trigram 子串搜索；短关键词或 SQLite 不支持 trigram 时回退到索引库内的子串查询。结果、表头、同行品名和价格从同一索引快照读取；按主程序习惯展示全部结果，分批更新表格。
- 界面保留搜索方式、开始/停止搜索和原有设置，后台自动维护索引。原进度条显示扫描目录数、发现文件数、索引更新进度/耗时及搜索结果数，不展示数据库路径和测速日志；完成搜索后保留结果概要。“停止搜索”仅取消本次查询和结果整理，保留已显示结果，不取消或暂停索引更新和监听，无需再次搜索即可继续处理文件变化；关闭程序会等待后台任务和监听退出。
- 索引机制、实现说明及已发现的限制见 [索引实现核对说明](ExcelSearchIndexReview.md)。
- 常规搜索仍使用原有多进程逐文件扫描算法，并沿用此前保存的并发文件数（未设置时使用默认值）。
- 两种搜索方式共用内置的文件发现规则，使用 `scandir` 保留 `DirEntry` 元数据，减少 Windows 下逐文件的额外 `stat` 查询，跳过回收站、系统保护目录和 Excel 锁文件；常规搜索按文件路径、工作表原顺序及行列顺序输出结果，完成统计与索引模式一致。目录访问详情仅保留在进度条悬停提示中。
- 支持 `.xls`、`.xlsx`、`.xlsm`，跳过 `~$` 开头的 Excel 锁文件。
- 按保存的单元格值进行不区分大小写的子串搜索，包含隐藏工作表。
- 不重新计算公式；公式结果取决于文件最后保存的计算值。
- 不模拟所有自定义显示格式，也不搜索批注、图片或形状。
- 品名、币种列和单价标记通过规则推断，复杂表格仍需人工核对。
- 个别文件读取失败或目录无法访问时不弹窗，文件统计显示在进度条中，目录访问及文件失败详情保留在悬停提示中；仅导致整个搜索失败的异常弹窗提示，主动停止时不弹窗。

## 服务端与发布

服务端依赖：`python -m pip install -r requirements-server.txt`。
部署前检查 `server.py`、`Updater.py` 和 `UsageReporter.py` 中的服务地址，并按自己的部署配置版本和凭据。

`release.py` 需要在运行它的 Python 环境中安装 PyInstaller 和 `cos-python-sdk-v5`。
运行后会重新生成 `dist/`，归档写入 `release/`，将 `dist/dist.zip` 上传至广州地域
`wublog-1328523094` 桶根目录的 `dist.zip`（覆盖上一版），上传成功后才调用发布接口更新服务器版本。
上传失败会保留本地压缩包并停止发布。使用简单上传，仅需 `cos:PutObject` 权限，最大支持 5GB。
COS 密钥从 `release.py` 所在目录下的 `.secrets/cos_credentials.json` 读取，字段为 `SecretId`、`SecretKey`，临时密钥还需 `Token`；
也支持 `COSSECRETID`、`COSSECRETKEY`、`COSSESSIONTOKEN` 环境变量，优先于配置文件。
发布令牌从 `.secrets/release_token.txt` 读取。服务端也需在 `server.py` 所在目录下创建 `.secrets/release_token.txt`，填写相同令牌。
`.secrets/` 内的文件是本地凭据，已通过 `.gitignore` 排除，禁止提交或随应用分发。
`--publish-only` 仅重试更新服务器版本，不打包或上传，使用前应确认桶中的 `dist.zip` 已更新。

正常启动桌面应用时，会尝试登记已有 `Updater.exe` 的启动项；退出时会发送使用记录。
使用记录可能包含搜索关键词、文件夹、结果概况，以及复制或打开文件的操作信息。

许可见 [LICENSE](LICENSE)。
