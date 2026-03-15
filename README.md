# 考勤统计助手

用于按月汇总考勤、请假和年假数据，并生成年度合计 Excel。

当前项目包含 3 个界面入口：

- 主版本：`ttkbootstrap + litera`
- 附加版本：`CustomTkinter`
- 附加版本：`NiceGUI`

当前建议交付给同事使用的版本是主版本。

## 主要功能

- 自动识别考勤打卡表中的年月
- 按中国大陆法定节假日和调休计算工作日
- 统计月度考勤明细
- 统计月度汇总
- 汇总全年统计
- 处理请假记录和 11 类请假类型
- 生成正式员工月度统计 sheet
- 支持多个月份累计生成年度结果

生成结果文件：

- `考勤统计结果.xlsx`

## 数据来源

每个月需要 2 个文件：

1. `考勤打卡记录表.xls`
2. `请假记录表.xls`

另外还有 1 个单独维护的“当前年假表”：

3. `员工年假总数表.xlsx`

当前年假表通常放在 `data/月度文件/` 根目录下，只有员工年假信息发生变化时才需要更新。

推荐按月放在：

```text
data/
  月度文件/
    2026-02/
      考勤打卡记录表.xls
      请假记录表.xls
    2026-03/
      考勤打卡记录表.xls
      请假记录表.xls
    当前员工年假总数表.xlsx
```

## 项目结构

核心脚本：

- [`attendance_report.py`](/Users/leechain/project/attendance-tool/attendance_report.py)

桌面 GUI：

- [`attendance_gui.py`](/Users/leechain/project/attendance-tool/attendance_gui.py)
- [`attendance_customtkinter.py`](/Users/leechain/project/attendance-tool/attendance_customtkinter.py)

Web GUI：

- [`attendance_nicegui.py`](/Users/leechain/project/attendance-tool/attendance_nicegui.py)

Windows 脚本：

- [`运行考勤统计.bat`](/Users/leechain/project/attendance-tool/运行考勤统计.bat)
- [`运行考勤统计_CustomTkinter.bat`](/Users/leechain/project/attendance-tool/运行考勤统计_CustomTkinter.bat)
- [`运行考勤统计_NiceGUI.bat`](/Users/leechain/project/attendance-tool/运行考勤统计_NiceGUI.bat)
- [`打包Windows程序.bat`](/Users/leechain/project/attendance-tool/打包Windows程序.bat)
- [`打包Windows程序_CustomTkinter.bat`](/Users/leechain/project/attendance-tool/打包Windows程序_CustomTkinter.bat)
- [`打包Windows程序_NiceGUI.bat`](/Users/leechain/project/attendance-tool/打包Windows程序_NiceGUI.bat)
- [`一键环境检查.bat`](/Users/leechain/project/attendance-tool/一键环境检查.bat)

文档：

- [`使用说明-给同事看.txt`](/Users/leechain/project/attendance-tool/使用说明-给同事看.txt)
- [`非技术人员使用说明.md`](/Users/leechain/project/attendance-tool/docs/非技术人员使用说明.md)
- [`Windows打包说明.md`](/Users/leechain/project/attendance-tool/docs/Windows打包说明.md)

## 本地运行

### 1. 直接运行统计脚本

```bash
python3 attendance_report.py
```

### 2. 运行主版本 GUI

```bash
python3 attendance_gui.py
```

主版本依赖：

```bash
python3 -m pip install ttkbootstrap pandas openpyxl xlrd holidays chinese-calendar
```

### 3. 运行 CustomTkinter 版本

```bash
python3 attendance_customtkinter.py
```

依赖：

```bash
python3 -m pip install customtkinter pandas openpyxl xlrd holidays chinese-calendar
```

### 4. 运行 NiceGUI 版本

```bash
python3 attendance_nicegui.py
```

依赖：

```bash
python3 -m pip install nicegui pandas openpyxl xlrd holidays chinese-calendar
```

## Windows 打包

主版本：

- 双击 [`打包Windows程序.bat`](/Users/leechain/project/attendance-tool/打包Windows程序.bat)
- 输出：`dist\\考勤统计助手.exe`

CustomTkinter 版本：

- 双击 [`打包Windows程序_CustomTkinter.bat`](/Users/leechain/project/attendance-tool/打包Windows程序_CustomTkinter.bat)
- 输出：`dist\\考勤统计助手_CustomTkinter.exe`

NiceGUI 版本：

- 双击 [`打包Windows程序_NiceGUI.bat`](/Users/leechain/project/attendance-tool/打包Windows程序_NiceGUI.bat)
- 输出：`dist\\财务公司考勤统计助手_NiceGUI.exe`

详细说明见：

- [`Windows打包说明.md`](/Users/leechain/project/attendance-tool/docs/Windows打包说明.md)

## 当前主版本说明

当前主版本是：

- `attendance_gui.py`
- 主题：`ttkbootstrap litera`

当前主版本的交互流程是：

1. 先确认或上传 `当前年假表`
2. 选择年份
3. 选择月份
4. 点击 `上传所选月份2个表`
5. 点击 `生成结果文件`
6. 点击 `打开生成好的 Excel`

结果文件名固定为：

- `考勤统计结果.xlsx`

软件版本：

- `考勤统计助手 v1.0.0`

## 注意事项

- 一次只统计一个年份的数据
- 每个月文件夹里只需要考勤表和请假表，每类只能有 1 个
- 当前年假表单独维护，只有变更时才需要更新
- 如果某个月缺文件或重复文件，程序会阻止生成
- 年度合计基于当前数据目录下同一年份的所有月份文件重算

## 给同事使用

如果是发给非技术同事，优先提供：

1. `考勤统计助手.exe`
2. `data/月度文件/`
3. `使用说明-给同事看.txt`

## 版本状态

- 主版本：稳定，建议交付
- CustomTkinter：可运行，可单独打包，适合继续对比视觉效果
- NiceGUI：可运行，可单独打包，更适合做现代化界面原型
