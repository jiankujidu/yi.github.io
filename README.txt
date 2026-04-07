# ELISA检测服务 - 样本信息填写系统

## 使用方法

### 方式一：双击启动（推荐）
1. 双击"启动系统.bat"
2. 系统会自动启动后端服务并打开填写页面
3. 按照步骤填写信息
4. 最后点击"导出Excel表格"下载填写好的Excel文件
5. 使用完毕后按任意键停止服务

### 方式二：手动启动
1. 命令行运行: python server.py
2. 双击打开 index.html
3. 填写并导出

## 文件说明
- index.html - 前端填写页面
- server.py - 后端服务（处理Excel导出）
- 优品Elisa代测表.xlsx - 原始Excel模板
- 启动系统.bat - 一键启动脚本

## 导出说明
导出的Excel文件完全保留原始模板的格式：
- 字体、颜色、边框
- 合并单元格
- 所有工作表结构

## 系统要求
- Python 3.x
- openpyxl库（pip install openpyxl）

## 技术支持
如有问题请联系开发人员
