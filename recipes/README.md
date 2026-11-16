# 配方

每个配方是一段在工作簿事务内运行的脚本（见 `excel-kit help script`），只用公开的命令与操作，附可再分发的样本和说明。

进入配方目录，按该目录 README 的步骤执行；需要预置内容的先执行 `excel-kit run prepare.json`。输出写在配方目录里；同名输出已存在时拒绝覆盖，再试一次就换一个输出名。用于自己的文件时，改 `request.json` 的 `input`、`output` 和 `arguments`，保留 `scriptPath`。脚本什么时候发布见 `SKILL.md` 第 2 条：正常返回后发布成功的步骤，抛错则什么都不写。

| 编号 | 目录 | 做什么 | 组合一次性命令 |
|---|---|---|---|
| R1 | [roll-forward-period](roll-forward-period/README.md) | 滚动到下一期间 | 否 |
| R2 | [update-matching-records](update-matching-records/README.md) | 按条件批量修改记录并读回结果 | 否 |
| R3 | [import-clean-append](import-clean-append/README.md) | 导入 CSV、清洗、去重并追加到 Table | 否 |
| R4 | [fill-template](fill-template/README.md) | 填写模板、校验、加打印设置与保护 | 否 |
| R5 | [audit-and-repair](audit-and-repair/README.md) | 巡检公式问题并应用选定的修正 | 否 |
| R6 | [update-assumptions](update-assumptions/README.md) | 只改指定假设并返回结果的静态上游 | 否 |
| R7 | [style-matching-sheets](style-matching-sheets/README.md) | 对名称匹配的工作表批量设置格式、打印与视图 | 否 |
| R8 | [normalize-codes](normalize-codes/README.md) | 用正则在整本工作簿里规范化编号 | 否 |
| R9 | [pivot-by-slicer](pivot-by-slicer/README.md) | 追加明细、刷新透视，按切片器逐项读取下游公式 | 否 |
| R10 | [inventory-media](inventory-media/README.md) | 清点图片与附件并提取核对 | `image-extract`、`attachment-extract` |
| R11 | [edit-protected-workbook](edit-protected-workbook/README.md) | 保存无口令保护权限、编辑并恢复核对；带口令时拒绝 | 先 `run prepare.json` 准备示例 |

## 写自己的脚本时

- **能事先写全就不用脚本。** 一份 `run` 把改动、计算、读取写在一起；脚本只用在下一步取决于上一步结果的时候。
- **请求写 `strict:true`。** 配方的请求都带它：任何一次 `wb.*` 失败直接抛出，脚本整体回滚，原错误在 `error.details.cause`。要自己处理某一次失败，只捕获那一次并拿回原响应：`.catch(error => error.response)`（`fill-template`、`update-matching-records` 各有一处）。不带 `strict` 时失败的调用返回 `ok:false`，每次取值前都得自己判；配方的脚本两样都做了，所以换一份不带 `strict` 的请求调用也报原错误。
- **先数再改，改完看改了谁。** 用查询数出匹配数量，编辑时带 `expectedMatches`，数量不符就拒绝。它只核对个数；条件用到公式的值要写 `source:"calculation"`。
- **让步骤失败，而不是自己判断。** 要校验就用步骤的 `validate`，要整本算完才接受步骤，就明确写 `requireCompleteCalculation:true`；默认计算不完整也可接受步骤并发布可得值，须检查计算状态和不可得清单。失败的步骤不留下任何修改。
- **该放一步的放一步。** 互相依赖的修改（修循环引用、追加后重新应用筛选、改透视来源后刷新透视）必须在同一步；要读这一步算出的值，把 `read` 也写进这一步。
- **发布前自证。** 用同一个搜索证明零残留，用独立求和核对小计，用第二次审计证明干净；不满足就 `throw`，脚本整体回滚。
- **`wb` 没有通用文件接口。** CSV 用 `data` 操作导入暂存表再读回；图片与附件的提取交给一次性命令。脚本不是安全沙箱，宿主的文件与网络访问不受工作簿事务回滚保护。
