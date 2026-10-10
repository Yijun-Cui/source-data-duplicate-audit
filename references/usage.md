# 安装与使用说明

本包让其他人的 Codex 按同一规范审查 Excel 数值源数据。它包含 skill、完整 Python 代码、中文规则、映射示例和模拟数据，不包含真实研究数据。

脚本需要 Python 3.10 或更新版本，只使用标准库。无需安装 Excel、`pandas` 或 `openpyxl`。检查过程在本机执行，不上传数据，也不执行工作簿中的文字指令。

## 1. 在其他人的 Codex 上使用

### 1.1 安装为 skill

解压 ZIP 后，将整个 `source-data-duplicate-audit` 文件夹放入 Codex 的技能目录。

如果从 GitHub 克隆项目，或下载了源码 ZIP，也可以在项目根目录运行安全安装脚本。

```text
python -B -X utf8 scripts/install-skill.py
```

它尊重 `CODEX_HOME`，默认使用当前用户的技能目录。目标已有同名技能时会停止，不覆盖文件。使用 `--skills-dir` 可以指定独立目录进行测试。

- Windows 默认位置：`%USERPROFILE%\.codex\skills\source-data-duplicate-audit`。
- macOS 和 Linux 默认位置：`~/.codex/skills/source-data-duplicate-audit`。
- 设置了 `CODEX_HOME` 时，使用该目录下的 `skills/source-data-duplicate-audit`。

不要只复制 `SKILL.md`，也不要把技能文件夹放成两层同名目录。重新打开 Codex 或新建会话后，确认技能列表中出现“论文源数据查重”。不同 Codex 版本的刷新方式可能不同。

在会话中提供自己的工作簿路径，并使用以下请求。

```text
请使用 $source-data-duplicate-audit 审查这份 Excel 的数值源数据。
先检查工作表和各子图的分组，再建立并核对分组映射。
按默认规则检查跨组完整重复、大量重复、高精度相同值、百分比近似、
变换关系、序列模式、末位集中、成套记录和元数据冲突。
原文件不要修改。请交付高亮副本、逐点证据、映射、参数和中文报告。
没有样本 ID 或其他必要字段时，请说明无法评估的项目。
```

Codex 会根据当前环境寻找 Python。绝对路径由使用者的电脑决定，本包没有绑定原作者的运行时或研究数据路径。

### 1.2 不安装，直接读取

也可以把解压后的整个目录放在任意项目中，让 Codex 读取其中的 `SKILL.md`。提供该文件的实际路径，再提出同样的审查请求。

```text
请完整读取这个目录中的 SKILL.md 及任务所需的 references 文档，
按其中的规范审查我提供的 Excel。脚本和模拟示例都在同一目录。
请重新识别我的分组，不要沿用模拟数据的组名或位置。
```

仅阅读文档不会自动完成查重。Codex 仍需运行脚本、建立工作簿专属映射，并核验输出。

## 2. 先用模拟数据验证

在终端进入 `source-data-duplicate-audit` 文件夹。Windows 使用已配置的 `python`，macOS 和 Linux 通常使用 `python3`。下面命令中的 `python` 应替换为当前电脑的实际解释器。

```text
python --version
python -B -X utf8 scripts/test-audit.py
python -B -X utf8 scripts/data-audit.py demo --out demo-run
python -B -X utf8 scripts/data-audit.py audit --input demo-run/demo.xlsx --map demo-run/demo-map.json --out demo-run/result
```

测试成功时显示 `OK`。模拟审查成功时显示 `complete_for_mapped_rules`。该状态表示声明范围内的规则运行成功，不表示数据已被认证为真实。

`demo-run` 和 `demo-run/result` 必须为新目录或空目录。重复运行时换一个输出目录，避免覆盖已有结果。

包内也附带生成好的模拟文件：

- `assets/demo/demo.xlsx`：模拟原始数据。
- `assets/demo/demo-map.json`：已核对的模拟分组映射。
- `assets/demo/expected.json`：22 种规则的指定预期示例和反例。
- `assets/demo/review.xlsx`：按默认规则生成的高亮示例。

预期示例不是命中数量清单。模拟数据可能在同一组对上触发多个规则，应按类型和位置检查，不要把所有命中相加。

## 3. 审查自己的工作簿

### 3.1 清点结构

将源文件保存在本机，关闭正在保存该文件的程序。脚本支持普通 `.xlsx`，不直接支持 `.xls`、`.xlsm`、加密文件和数字签名文件。

以下 PowerShell 示例在技能目录下执行，路径以示例工作目录表示。

```powershell
$source = "D:\research-audit\source.xlsx"
$work = "D:\research-audit\run-01"
python -B -X utf8 scripts/data-audit.py inspect --input "$source" --out "$work\inventory.json"
```

`inventory.json` 列出存储值、表头、公式、边框线型、合并范围和隐藏状态。Codex 用这些信息核对结构，必要时查看实际 Excel。清点文件包含数据，不应随公共技能包分享。

### 3.2 建立分组映射

如果工作簿是简单的列式表，可以先生成草稿。

```powershell
python -B -X utf8 scripts/data-audit.py draft-map --input "$source" --out "$work\groups-draft.json"
```

草稿只尝试未合并的简单表头列。它可能把时间列识别为测量值，也可能跨过不相关内容。复杂图块、横排数据、长表和时间矩阵，需要 Codex 按 [mapping-schema.md](mapping-schema.md) 重新建立映射。

检查完整后，将最终映射保存为 `groups.json`，设为 `reviewed: true`。主要测量组、排除理由、比较集合和元数据字段都必须有依据。

### 3.3 运行查重

```powershell
python -B -X utf8 scripts/data-audit.py audit --input "$source" --map "$work\groups.json" --out "$work\result"
```

需要调整阈值时，建立一个 JSON 设置文件，只写要覆盖的键。例如，严格要求小数位数大于 3，并把单点近似误差改为 0.5%。

```json
{
  "high_precision_decimals": 4,
  "near_relative_tolerance": 0.005,
  "near_similarity": 0.995
}
```

```powershell
python -B -X utf8 scripts/data-audit.py audit --input "$source" --map "$work\groups.json" --settings "$work\settings.json" --out "$work\result-stricter"
```

设置中的比例使用 `0.01 = 1%`，不是字符串 `"1%"`，也不是数字 `1`。未知参数会报错。阈值应在理解指标后确定，不能根据期望结论反复选择。

macOS 和 Linux 可使用同一流程。使用带引号的本机路径，将 `python` 改为 `python3`，将路径分隔符改为 `/`。

## 4. 读取结果

| 文件 | 用途 |
| --- | --- |
| `review.xlsx` | 源数据高亮副本，追加 6 个报告页 |
| `report.md` | 中文摘要，最多展示前 30 条需要复核的命中 |
| `analysis.json` | 完整规则命中、组对、匹配坐标、实际参数和检查范围 |
| `coverage.json` | 每个数值、公式和错误单元格的归属或排除理由 |
| `mapping-used.json` | 本次使用的完整映射 |
| `verification.json` | 输入保护、高亮验证、报告页名称和输出校验值 |

Excel 报告页为“审查总览”“查重明细”“逐点配对”“分组覆盖”“规则与限制”和“低信息量记录”。若输入已存在同名页，脚本自动加后缀。

先读“审查总览”，再按类别筛选“查重明细”。“逐点配对”提供可以跳回源单元格的内部链接。覆盖率、相似度和相对误差显示为百分比。

同一单元格可能命中多个规则，显示颜色取最高优先级。颜色只表示复核分类，不表示造假程度。低信息量记录不额外着色，但完整保留在单独报告页和 JSON 中。

结合实验记录复核红色完整重复、橙色多点或结构模式、黄色身份冲突。紫色末位和少量高精度线索、蓝色近似关系也需要解释，但不能直接定性。

## 5. 常见失败与处理

| 失败 | 处理 |
| --- | --- |
| `reviewed=true` 未设置 | 先核对结构，不要直接跳过检查 |
| 数值覆盖不完整 | 阅读 `coverage.json` 和 `incomplete.json`，补充分组或有依据的排除项 |
| 主要分组重叠 | 修正范围；确属辅助轨迹时改为 `auxiliary` |
| 组内存在文本型数字 | 检查导出格式；需要转换时在副本中记录转换依据 |
| 公式未允许或缓存非数值 | 明确排除，或在 Excel 中重算后保存副本并明确纳入缓存 |
| 日期进入测量组 | 核对日期格式和指标范围，日期一般作为元数据排除 |
| 输出目录已存在文件 | 换新的输出目录，保留旧结果 |
| 未知键或无效地址 | 修正拼写、工作表名称、列名或 A1 范围 |
| 超过比较上限 | 说明未完成范围；科学分区后分批，或提高有记录的设置 |
| 无样本 ID、稳定字段或配对列 | 不补造字段；报告对应项目无法评估 |
| 文件不是普通 XLSX | 先确认转换是否允许，再生成并记录转换副本 |

脚本失败时退出码为 `2`。不能把已有的部分输出当作完整审查。修正后使用新的目录重跑，并确认退出码、报告及 `verification.json` 均成功。

## 6. 隐私与分享

本包的模拟数据是人工生成的。完整代码在 `scripts/data-audit.py`，测试代码在 `scripts/test-audit.py`，均可直接审阅和运行。

不要把自己的 `inventory.json`、映射、报告和真实工作簿误打包进公共技能。它们可能含有样本身份、未发表数据或本机路径。对外分享前单独检查输出内容。

本包只验证了本机运行和隔离副本运行。其他机器应先执行自检，再处理真实数据。自动检测不能代替原始实验记录、仪器导出和独立复核。

GitHub 项目通过 Actions 配置了 Linux、Windows 和 macOS 的测试。实际执行结果以仓库的 Actions 页面为准，本机自检不等于所有平台均已验证。
