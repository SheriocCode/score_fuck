## score-fuck武汉理工综测评定成绩单绩点计算
- 2025/1/15 🤤感谢您的厚爱，由于 新教务系统 打印的成绩单pdf貌似不能提取，所以该脚本暂时失效。同时新系统支持导出成绩表格，也简化了后续的开发（敬请期待）

```python
* 更新日志  
  2024/9/9 完善课程绩点统计  
  2024/9/7 修复已知问题
```

1、 依赖包
```python
pip install pandas
pip install pdfplumber
```

2、将学生成绩单pdf放到文件夹，并复制路径到`PATH`
![image](https://github.com/user-attachments/assets/1d70d171-d5e8-4a2b-a270-4d35bd6e95e4)


运行脚本后结果  
`pdf2excel`文件夹中存放每个学生成绩单pdf提取的excel  
`result.xlsx`为统计表  
![image](https://github.com/user-attachments/assets/046faa3e-6705-4d3d-8a69-3bf4b893888e)

