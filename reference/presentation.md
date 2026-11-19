# 版面与对象

固定行高的换行文本可以用 `help query.layout` 的 `axis:"textClipping"` 查。观察报告中的地址、合并范围和以磅计的可用及估计所需高度。未取得当前公式结果的格会报告无法判断，需要时在同一查询对该格指定 targets。可先在受保护的表上检查，挑出值得调整的行，再决定是否修改行高。调整后用相同范围复查。

## 已有图表、透视与规则的局部编辑

先以 `query.charts` 定位工作表内的图表名称、系列 `index` 与来源。用 `chart.set.seriesEdit` 指定一个系列。例如 `{"index":1,"sources":{"values":"Data!D2:D4"},"marker":{"symbol":"diamond","size":6}}` 会同时更新系列来源和标记。线图和散点图可设置 `linearTrendline` 与 `fixedErrorBars`，传 `null` 可删除。点覆盖使用 `point.index` 和 `point.appearance`，传空外观值可清除覆盖。

在 `chart.create` 中指定 `type:"histogram"` 或 `type:"boxWhisker"`、本地单列或单行数值来源 `series:[{values:"Data!A1:A5"}]`，并传入绘图锚点。直方图用 `binning:{width:3}` 或 `binning:{count:4}`，箱线图用 `quartileMethod:"exclusive"` 或 `"inclusive"`。已有图可用 `chart.set` 改设置，用 `chart.refresh` 从当前工作簿刷新数值缓存。Excel 保存后可能把来源写为隐藏定义名称，查询返回文件实际绑定。对带分箱设置的直方图，Excel 使用分箱元素的 `val` 属性；固定版 Open XML SDK 的严格校验可能在该元素上同时报属性与文本诊断。

`conditional-format.set` 的 `condition` 支持 `{kind:"average",equal:true}` 和 `{kind:"time-period",period:"today"}`。数据条在同一条件对象传 `negativeColor`、`axisPosition` 和 `direction`。列表验证的 `showDropdown` 放在 `validation.set` 操作顶层，false 隐藏箭头。用 `query.rules` 查定义，`query.ruleResults` 查当前输入的匹配结果。
