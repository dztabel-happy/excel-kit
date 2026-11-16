<h1 align="center">excel-kit</h1>
<p align="center">让 agent 读取、编辑、计算并交付可继续使用的 Excel 工作簿。</p>

把本地工作簿和任务要求交给能执行命令的 agent。excel-kit 负责读取、修改、计算及保存 XLSX／XLSM，返回结构化结果和需要处理的问题。处理文件时不依赖 Excel、LibreOffice 或云端服务。

**许可：仅限个人非商业使用，禁止商用和再分发。** 商业使用须另获书面许可；完整条款见 [LICENSE](LICENSE)。

- **批量整理与汇总**：清洗个人记录、去重、关联分类、追加 Table、更新汇总。
- **模型与预算调整**：调整个人预算或学习模型的假设、计算指定结果、滚动到下一期间，保留公式和格式。
- **模板与交付**：填写模板、设置验证和保护、检查公式问题，保存新的工作簿。

## 安装

### 1. 安装命令行程序

已有 Node.js 20 或更新版本时：

```sh
npm install -g @dztabel/excel-kit
```

npm 自动安装当前平台的程序。支持 macOS ARM64／x64、Windows x64；Windows ARM64 使用 x64 仿真程序。当前没有 Linux 发行包。

无需 Node.js 的安装方式：从 [Releases](https://github.com/dztabel-happy/excel-kit/releases) 下载对应平台的 ZIP，核对同页 `SHA256SUMS` 后解压。macOS 执行 `./excel-kit`，Windows 执行 `.\excel-kit.exe`，也可将所在目录加入 PATH。

### 2. 确认安装

```sh
excel-kit version
excel-kit help
```

当前版本应返回 `cli: "0.9.0"`、`protocol: 1`。帮助首页是索引，`excel-kit help <名称>` 查询具体命令、操作或函数。

### 3. 交给 agent 使用

让 agent 先阅读 [SKILL.md](SKILL.md)，再给出文件路径和任务，例如：

> 用 excel-kit 整理我的家庭支出，按编号去重，关联分类表，更新月度汇总。保留原件，另存为新的工作簿，并说明未能计算的内容。

> 把我的学习用预算模型滚动到下一期间，保留公式、格式和打印设置；更新给出的假设，并核对总金额。

`SKILL.md`、`reference/` 和 11 个[配方](recipes/README.md)也包含在 npm 安装目录中。全局安装目录可用 `npm root -g` 查找，其下的 `@dztabel/excel-kit/` 就是说明目录。使用配方前将其复制到工作目录，按对应 README 执行。

## 调用方式

| 方式 | 适合的任务 |
|---|---|
| `run`／`batch` | 事先能写全的编辑、计算与核对步骤 |
| `session` | 边看边改，跨调用保留工作簿，最后保存或另存 |
| `script` | 在一个工作簿事务中循环、分支、根据中间结果继续操作 |

```sh
excel-kit inspect input.xlsx
excel-kit run request.json
excel-kit help session
excel-kit help script
```

请求由 agent 按任务与帮助生成。命令返回 JSON，既要检查 `ok`，也要检查计算状态和值来源：`ok:true` 可以伴随不完整计算；文件里的历史缓存不代表本次算出的结果。要求整本算完才能保存时，显式设置 `requireCompleteCalculation:true`。

## 范围与版本

0.9.0 主要降低大工作簿的内存占用，并修正查找空值、透视结果读取与 EXP 舍入问题。实际验证、发行例外和资源边界见 [VALIDATION.md](VALIDATION.md)，版本变化见 [CHANGELOG.md](CHANGELOG.md)。

不执行 VBA，不编辑加密文件，不提供多工作簿依赖求值。具体函数边界见 [函数限制](reference/function-limits.md)，未支持范围见 [unsupported.md](reference/unsupported.md)。脚本不是安全沙箱，只运行可信代码。内存与时间取决于文件结构，不能保证所有大表都在固定预算内完成。

## 安装排查

- 提示缺少平台包：重新执行 `npm install -g --include=optional @dztabel/excel-kit`，不要省略 optional dependencies。
- 安装后找不到命令：确认 npm 全局可执行目录已加入 PATH，或使用 `npx @dztabel/excel-kit version`。
- 不使用 npm：ZIP 内含独立运行时，不需要安装 Bun、Python 或办公软件。
- Node.js 客户端通过子进程调用 CLI／[JSON-RPC 会话](reference/api.md)；本 npm 包不导出工作簿内核库。

本仓库用于公开构建、安装入口和使用说明，正式程序在 Release 和 npm 平台包中。构建仅由维护者手动触发，只传出已核对摘要的程序；核心源码始终保存在独立私有仓库。许可范围见 [LICENSE](LICENSE)，第三方声明见 [THIRD-PARTY-NOTICES.txt](THIRD-PARTY-NOTICES.txt)。问题反馈请附版本、操作和可公开的最小样本，勿上传敏感工作簿。
