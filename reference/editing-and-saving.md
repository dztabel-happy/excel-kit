# 编辑与保存

一份请求怎样组合、输出文件怎样覆盖、`preflight` 与 `batch` 各保证什么、写入锁。事务的规则见 `SKILL.md` 第 1、2 条；请求外壳的字段查 `help run.request`，每个操作查 `help <操作名>`，批次查 `help batch`；保存之后怎样核对见 [deliverable.md](deliverable.md)。

## 请求怎样组合

- `operations` 按顺序执行，后一个看到前一个的结果。互相依赖的修改必须放在同一步，否则中间状态过不了计算或校验：同时修多个循环引用、向带筛选的 Table 追加后重新应用筛选、改透视来源后刷新透视。
- 公式写成 `{"formula":"=B2*1.1"}`；以 `=` 开头的普通字符串仍是文本。Table 的当前行引用写 `[@列]`。
- 步骤有两种写法。完整的：`{"operations":[…],"nativeEdits":[…],"read":[…],"validate":[…],"calculation":{…},"timeoutMs":…}`。简写：只在批次的步骤、`session.run`、`wb.run` 里，可以把单个操作 `{"op":…}` 直接当一步；用了简写就不能在旁边写 `calculation`、`read` 这些步骤字段，要它们就换成完整写法。一次性 `run` 只有完整写法。
- `workbook.settings` 操作自己的 `calculation` 是存进文件的 Excel 设置，写在这个操作里面；它不是这一步的计算方式。
- `validate` 让工作簿自己的数据验证规则判定：不通过，这一步失败、不留下任何修改。规则程序求不了值时也失败：这时去掉 `validate`，用 `query` 的 `rules` 读出规则自己核对，并在报告里说明。它不能与 `"mode":"none"` 同用。
- 不保留原工作簿、只要新表时用 `create` 代替 `input`，例如 `"create":{"sheets":["Data"]}`。

## 输出、覆盖与版本

- 输出目录要先存在，程序不创建目录。
- `output` 已存在报 `OUTPUT_EXISTS`，`details.currentVersion` 是现有文件的版本；确认覆盖就在请求外层写 `"overwrite":{"expectedVersion":"<版本>"}`。原地修改（`input` 与 `output` 是同一个文件）也要写。
- 版本是文件的 SHA-256，取自 `inspect` 的 `workbook.version`、读取与查询的 `workbookVersion`、上一步的 `output.version`。`expectedInputVersion` 让输入不是这个版本时拒绝执行。
- 发布前会再核对一次：输入在执行期间变了报 `SOURCE_CHANGED`，目标变了报 `TARGET_CHANGED`。
- 扩展名要与工作簿类型一致，否则 `WORKBOOK_FORMAT_MISMATCH`：`.xlsm` 仍存为 `.xlsm`，改扩展名不等于转换格式或去掉宏。
- 新增或改写的公式换成原生存储形式后超过 8192 字符，报 `FORMULA_TOO_LONG`、不发布；只数请求里的字符可能不够。原样保留的公式不受影响。

## `preflight`

把同一份 `run` 请求在内存里完整做一遍：解析、操作、所请求的计算、写出、读回核对，带 `read` 时返回这次计算的值。它只差最后的写文件，所以与 `run` 一样贵：大模型上不要先 `preflight` 再 `run`，`run` 失败本来就不写文件。

| | |
|---|---|
| 查 | 输出的扩展名与目录（`INVALID_OUTPUT`、`OUTPUT_UNAVAILABLE`）；`output` 已存在而没给版本（`OUTPUT_EXISTS`）、版本不对（`TARGET_CHANGED`） |
| 不查 | 目录能不能写；不占写入锁，不留任何文件 |
| 不保证 | 之后没人改输入与输出：`run` 时重新核对 |
| 给 | `executionRequest`：原请求加上输入的版本、外部来源的版本、这次用的时钟、随机种子与区域设置。里面没有"选择命中了哪些对象"，重放时起作用的仍是你写的 `expectedMatches` |

## 批次的结局

| 情形 | `ok` | `status` | `output.saved` | 顶层 `error` |
|---|---|---|---|---|
| 全部成功 | `true` | `completed` | `true` | 无 |
| 有步骤没成功：`commit:"atomic"`（默认），或 `successful` 而一步都没成功 | `false` | `failed` | `false` | 无 |
| 有步骤没成功，`commit:"successful"` 且至少一步成功 | `false` | `partial` | `true` | 无 |
| 被中止：取消、整批超时、内部错误、发布前版本变了、写文件失败 | `false` | `aborted` | `false` | 有 |
| 解析、检查输出目标或读入输入时就失败 | `false` | 没有 `batch` | | 有 |

部分发布时响应带 `continuationRequest`：只含没发布的步骤，绑定到已保存的版本。它们将在已保存的最终状态上执行，位置不会自动换算；先改对再发。`"preflight":true` 的批次不写文件。

## 写入锁

写入时程序持有目标文件的锁，并在输出旁留下 `.excel-kit-*.lock.guard.sqlite`。它不影响工作簿，不是死锁，也不是交付物；不要为了"解锁"删它。交付前要清理，先确认没有 excel-kit 进程在写这个目录，下次写入会重建。`FILE_BUSY` 是另一个写入者持有锁，或遗留的锁记录无法安全确认：看 `details` 的原因与恢复提示。确有写入者时等它结束；是遗留的锁记录时等待不一定解决，按提示处理。锁记录与 guard 是两种文件。

## 例：被拒之后带上版本重试

失败的响应只有 `ok` 与 `error`，修正要用的值在 `error.details` 里：

```sh
cat > inplace.json <<'EOF'
{"input":"input.xlsx","output":"input.xlsx",
 "operations":[{"op":"write","sheet":"Sheet1","range":"B2","values":[[1200]]}]}
EOF
./excel-kit run inplace.json > refused.json || true
python3 - <<'EOF'
import json
refused, request = json.load(open('refused.json')), json.load(open('inplace.json'))
if refused['ok'] or refused['error']['code'] != 'OUTPUT_EXISTS':
    raise SystemExit(f"unexpected: {json.dumps(refused.get('error'))}")
request['overwrite'] = {'expectedVersion': refused['error']['details']['currentVersion']}
json.dump(request, open('inplace.json', 'w'))
EOF
./excel-kit run inplace.json > inplace.out.json
```
