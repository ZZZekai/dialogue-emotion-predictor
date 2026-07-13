# Excel Experiment Comparison Merger

用于合并两份情绪预测实验工作簿，并按同名工作表、相同轮次对齐预测结果。

合并后的对话工作表列顺序：

```text
轮次
提问人员
被谈话人
预测情绪标签1
预测情绪标签2
原始情绪标签1
原始情绪标签2
预测原因1
预测原因2
原始模型输出1
原始模型输出2
```

预测情绪标签不一致的两格会显示浅黄色背景。`index` 工作表保留文件1的内容。

从项目根目录启动桌面界面：

```bash
python3 tool/excel_comparison_merger/app.py
```

在界面中选择两份 `.xlsx` 文件，填写合并文件名，然后点击“开始合并”。结果固定输出到：

```text
tool/excel_comparison_merger/output/
```

运行测试：

```bash
.venv/bin/python -m unittest -v \
  tool/excel_comparison_merger/test_merge_excel.py
```
