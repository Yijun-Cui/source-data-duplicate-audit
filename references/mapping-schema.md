# 分组映射说明

映射是工作簿结构与查重规则之间的接口。它使用 JSON，编码为 UTF-8。每份工作簿建立自己的映射，并将最终版本随报告交付。

## 1. 根结构

```json
{
  "schema_version": "1.0",
  "reviewed": true,
  "groups": [],
  "excluded": [],
  "records": [],
  "reuse": [],
  "derived_checks": []
}
```

`reviewed: true` 表示 Codex 已核对表头、分组位置和排除项。自动草稿默认设为 `false`。不要为了运行脚本，跳过结构检查直接修改该字段。

映射拒绝未知键，避免参数拼写错误被默默忽略。所有工作表名称必须与输入完全一致。单元格使用 A1 地址；记录字段的列名使用大写字母。

## 2. 数值分组

### 2.1 字段

| 字段 | 含义与默认值 |
| --- | --- |
| `id` | 必填；全局唯一的组标识，建议使用易读英文 |
| `sheet` | 必填；实际工作表名称 |
| `ranges` 或 `cells` | 必填且二选一；位置顺序参与比较 |
| `label` | 可选；实际组名，默认显示组标识 |
| `panel` | 可选；子图标记，不自动推断 |
| `comparison_set` | 可选；默认 `measurements`，只有同集合组之间比较 |
| `ordering` | `observed`、`unordered` 或 `time`，默认 `observed` |
| `role` | `primary` 或 `auxiliary`，默认 `primary` |
| `kind` | 可选；例如 `measurement`、`time_slice`、`trajectory` |
| `include_formulas` | 布尔值；默认 `false`，明确纳入时只读取保存的数值缓存 |

`ranges` 中的每个范围必须是一维行或列。一个组可以包含多个范围，但内部不能重复地址。`cells` 支持长表的离散位置，脚本按列表顺序处理。

内部空白属于位置的一部分，应留在范围中。不要把空白删掉后拼接数据，也不要把两个不同子图的同名组直接合为一组。

默认比较集合只是软件默认值，不代表所有测量都应该比较。Codex 应按指标和单位选择集合，并在报告中说明跨图比较的范围。

### 2.2 普通列式数据

以下两个组属于同一子图、同一指标。

```json
{
  "id": "fig1a-control",
  "sheet": "Figure 1",
  "panel": "Fig. 1A",
  "label": "Control",
  "ranges": ["B4:B13"],
  "comparison_set": "protein-expression",
  "ordering": "observed"
}
```

另一个实验组应建立独立条目，例如 `C4:C13`。不同子图中同一指标可以放在相同集合，以检查跨图复用。指标不同则使用不同集合。

### 2.3 横向组与合并表头

横排数据只需改为行范围。合并表头不是测量值，也不是单独的样本。

```json
{
  "id": "strain-x-drug",
  "sheet": "Horizontal",
  "panel": "Fig. 1A",
  "label": "Strain X / Drug",
  "ranges": ["G2:K2"],
  "comparison_set": "same-measurement"
}
```

`inspect` 输出每个单元格的存储内容、样式索引和边框线型，以及合并范围。粗框、表头和空行应一起核对。`draft-map` 不适用于复杂图块或合并表头。

### 2.4 长表

如果 A 列为组别，D 列为测量值，且属于同一组的行不连续，应明确列出位置。

```json
{
  "id": "long-control",
  "sheet": "Long table",
  "label": "Control",
  "cells": ["D2", "D5", "D8", "D11", "D14"],
  "comparison_set": "readout-1",
  "ordering": "unordered"
}
```

`unordered` 表示组内没有可用于拟合的顺序。脚本仍执行多重集合精确匹配和无序近似匹配，但不执行同序近似、组间变换或组内序列规则。

### 2.5 时间矩阵

主要分组可按每个时间点的实验组建立。辅助分组则按每个样本的时间轨迹建立。

```json
{
  "id": "control-day0",
  "sheet": "Time",
  "panel": "Time matrix",
  "ranges": ["B2:E2"],
  "kind": "time_slice",
  "role": "primary",
  "comparison_set": "time-slices"
}
```

```json
{
  "id": "control-subject-1",
  "sheet": "Time",
  "panel": "Time matrix",
  "ranges": ["B2:B7"],
  "kind": "trajectory",
  "role": "auxiliary",
  "ordering": "time",
  "comparison_set": "subject-trajectories"
}
```

同一时间矩阵的切片重合会降为低信息量。辅助轨迹不替代主要分组的覆盖。若缺少样本 ID，列位置只能表示表内轨迹，不能确认实际动物或受试者配对。

## 3. 排除项

每个排除项必须有理由，不能与测量分组重叠。`ranges` 中的每个范围要求一维。如需排除矩形区域，可分列列出，或使用 `cells` 明确地址。

```json
{
  "sheet": "Time",
  "ranges": ["A2:A7"],
  "reason": "天数轴，不是测量值"
}
```

```json
{
  "sheet": "Figure 1",
  "ranges": ["F4:F13"],
  "reason": "已核对的汇总均值，不作为独立原始点"
}
```

排除必须依据实验含义，不得仅因单元格被标记而排除。报告保留排除项和原始存储值，供他人复核。

## 4. 成套记录与身份字段

```json
{
  "sheet": "Records",
  "first_row": 2,
  "last_row": 101,
  "fields": {
    "sample_id": "A",
    "visit": "B",
    "sex": "C",
    "value_1": "D",
    "value_2": "E"
  },
  "id_fields": ["sample_id", "visit"],
  "stable_fields": ["sex"],
  "value_fields": ["value_1", "value_2"],
  "unique_observation": true
}
```

`id_fields` 是本记录表的复合身份键。`unique_observation: true` 表示该键在映射行范围内应当唯一。技术重复需要把重复编号加入键，或者明确说明该记录键不唯一。

`stable_fields` 表示同一键下不应改变的字段。缺少字段值时不能判断冲突。若希望核对跨访视的稳定性，应另建以受试者 ID 为键的记录映射，并设置 `unique_observation: false`；不要把访视编号误设为稳定字段。

`value_fields` 至少需要 2 个指标，才能执行成套数值重复检查。指标数值还需要通过 `groups` 归属主要分组；仅声明记录表不等于完成数值覆盖。

身份和分类字段自动记为排除项，不作为测量值。身份键缺失的行仍可参加完整指标的数值比较，但身份冲突检查会标记为部分完成。没有身份键或稳定字段时，相关核验标记为无法评估。

## 5. 合理复用与派生关系

### 5.1 共享对照

```json
{
  "group_a": "fig1a-control",
  "group_b": "fig2b-shared-control",
  "reason": "两张图共用同一次实验的对照",
  "evidence": "图注明确说明；原始实验编号 EXP-014"
}
```

声明放在 `reuse` 数组中。脚本保留重复证据，只把该组对的分类改为绿色解释项。没有图注或记录依据时，不应添加这种声明。

### 5.2 显式滞后差分

```json
{
  "id": "two-step-difference",
  "source_sheet": "Derived",
  "source_range": "B2:B11",
  "target_sheet": "Derived",
  "target_range": "C2:C11",
  "operation": "lag_difference",
  "lag": 2,
  "status": "confirmed",
  "reason": "方法说明使用两步差分，前两个点使用基线 0",
  "relative_tolerance": 0
}
```

源范围与目标范围长度必须相同，且没有空白。默认要求精确相同；若有记录支持的舍入误差，可以设置相对比例。未经确认的关系使用 `hypothesis`，并在结论中保持假设语气。

## 6. 映射核对清单

1. 工作表、子图、组名和边框是否一致？是否遗漏隐藏页？
2. 每组的范围是否只包含对应指标？表头、轴、ID 和汇总值是否分开？
3. 多个子图中的同名组是否仍为独立条目？跨图比较集合是否合理？
4. 空白位置是否保留？长表的 `cells` 是否按合理顺序列出？
5. 时间切片与样本轨迹是否使用不同集合？辅助轨迹是否误当独立样本？
6. 唯一键、稳定字段和配对指标是否有实验依据？
7. 所有数值、公式和错误是否归属或明确排除？是否存在主要分组重叠？
8. 合理复用和派生关系是否有证据，而不是为解释命中临时编写？

只有核对完成后，才设置 `reviewed: true` 并运行审查。映射不确定时，应说明不确定范围；不要把未识别的数字当作“没有问题”。
