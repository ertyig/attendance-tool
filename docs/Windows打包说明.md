# Windows 打包说明

这份说明是给“负责打包程序的人”看的，不是给最终使用者看的。

目标：

- 把当前考勤统计助手打包成 Windows 可双击运行的 `.exe`
- 打包后发给不懂计算机的同事直接使用

---

## 一、需要在哪台电脑上打包

请在 **Windows 电脑** 上打包。

不要在 macOS 上直接打 Windows 的 `.exe`。

原因：

- Windows 程序最好在 Windows 环境下打包
- 这样兼容性最稳

---

## 二、打包前需要准备什么

在 Windows 电脑上准备：

1. 安装 Python 3
2. 把整个项目文件夹拷贝到 Windows 电脑上

项目里已经准备好了这些文件：

- `attendance_gui.py`
- `attendance_gui.spec`
- `attendance_customtkinter.py`
- `attendance_customtkinter.spec`
- `attendance_nicegui.py`
- `attendance_nicegui.spec`
- `打包Windows程序.bat`
- `打包Windows程序_CustomTkinter.bat`
- `打包Windows程序_NiceGUI.bat`
- `一键环境检查.bat`

---

## 三、最简单的打包方法

建议先双击：

- `一键环境检查.bat`

先确认这台 Windows 电脑的 Python、pip、依赖包是否正常。

如果检查没问题，再双击：

在 Windows 电脑上，直接双击：

- `打包Windows程序.bat`

这个脚本会自动执行：

1. 安装/更新打包依赖
2. 调用 PyInstaller 打包
3. 同时把输出写入日志文件 `build_windows.log`

如果你要打包 NiceGUI 版本，双击：

- `打包Windows程序_NiceGUI.bat`

它会输出独立的 NiceGUI 版本 exe，并把日志写到：

- `build_windows_nicegui.log`

如果你要打包 CustomTkinter 版本，双击：

- `打包Windows程序_CustomTkinter.bat`

它会输出独立的 CustomTkinter 版本 exe，并把日志写到：

- `build_windows_customtkinter.log`

---

## 四、打包成功后，exe 在哪里

打包成功后，程序默认会生成在：

- `dist\考勤统计助手.exe`

如果打的是 NiceGUI 版本，默认会生成：

- `dist\财务公司考勤统计助手_NiceGUI.exe`

如果打的是 CustomTkinter 版本，默认会生成：

- `dist\考勤统计助手_CustomTkinter.exe`

这个文件就是给同事双击运行的程序。

---

## 五、如果双击 bat 没反应怎么办

你也可以手工打开 Windows 命令行，进入项目目录后执行：

```bat
py -m pip install --upgrade pip pyinstaller ttkbootstrap pandas openpyxl xlrd holidays chinese-calendar
py -m PyInstaller --noconfirm --clean attendance_gui.spec
```

如果是 CustomTkinter 版本，执行：

```bat
py -m pip install --upgrade pip pyinstaller customtkinter pandas openpyxl xlrd holidays chinese-calendar
py -m PyInstaller --noconfirm --clean attendance_customtkinter.spec
```

如果是 NiceGUI 版本，执行：

```bat
py -m pip install --upgrade pip pyinstaller nicegui ttkbootstrap pandas openpyxl xlrd holidays chinese-calendar
py -m PyInstaller --noconfirm --clean attendance_nicegui.spec
```

打包完成后，还是去：

- `dist\考勤统计助手.exe`

找生成的程序。

---

## 六、打包后应该发给同事哪些文件

建议至少发下面这些内容：

1. `考勤统计助手.exe`
2. 或 `考勤统计助手_CustomTkinter.exe`
2. 或 `财务公司考勤统计助手_NiceGUI.exe`
3. `data/月度文件/`
4. `使用说明-给同事看.txt`

建议整理成一个完整文件夹发给同事，例如：

```text
考勤统计助手/
  考勤统计助手.exe
  考勤统计结果.xlsx
  使用说明-给同事看.txt
  data/
    月度文件/
      使用说明.txt
```

其中：

- `考勤统计助手.exe` 是她双击运行的程序
- `data/月度文件/` 是她放每个月 Excel 文件的地方
- `使用说明-给同事看.txt` 是给她看的操作说明

---

## 七、发给同事之前，建议先自测一次

建议在你自己的 Windows 电脑上先测试：

1. 在 `data/月度文件/2026-02/` 下放入 3 个示例文件
2. 双击运行 `考勤统计助手.exe`
3. 或双击运行 `考勤统计助手_CustomTkinter.exe`
4. 或双击运行 `财务公司考勤统计助手_NiceGUI.exe`
5. 在程序里点击：
   - `上传所选月份3个表`
   - `生成结果文件`
6. 检查是否成功生成：
   - `考勤统计结果.xlsx`

如果这一步没问题，再把程序发给同事。

---

## 八、同事那边怎么使用

同事不需要安装 Python。

她只需要：

1. 双击 `考勤统计助手.exe`
2. 或双击 `考勤统计助手_CustomTkinter.exe`
3. 或双击 `财务公司考勤统计助手_NiceGUI.exe`
4. 在程序里选择年份和月份
5. 点击 `上传所选月份3个表`
6. 点击 `生成结果文件`

---

## 九、更新程序时怎么做

如果后面脚本又改了，重新打包的方法还是一样：

1. 用最新代码覆盖项目文件
2. 在 Windows 上重新双击 `打包Windows程序.bat`
3. 或重新双击 `打包Windows程序_CustomTkinter.bat`
4. 或重新双击 `打包Windows程序_NiceGUI.bat`
5. 把新的 exe 发给同事

---

## 十、常见问题

### 1. 打包时报错找不到模块

请先确认 Python 已安装，然后重新执行：

```bat
py -m pip install --upgrade pip pyinstaller ttkbootstrap pandas openpyxl xlrd holidays chinese-calendar
```

如果是 CustomTkinter 版本，先执行：

```bat
py -m pip install --upgrade pip pyinstaller customtkinter pandas openpyxl xlrd holidays chinese-calendar
```

---

### 2. exe 能打开，但没有结果

请检查：

1. `data/月度文件/` 下面是否有月份文件夹
2. 每个月份文件夹里是否有 3 个文件
3. 文件名是否正确：
   - `考勤打卡记录表.xls`
   - `请假记录表.xls`
   - `员工年假总数表.xlsx`

---

### 3. 同事电脑打开 exe 被拦截

有些 Windows 电脑会对未知来源程序弹出安全提示。

这时通常需要：

1. 点击“更多信息”
2. 再点击“仍要运行”

如果公司电脑安全策略更严格，可能需要 IT 协助。

### 4. 双击 bat 一闪而过，没反应

通常是下面几种原因：

1. 电脑没有安装 Python
2. `py` 命令不可用
3. 缺少依赖包

现在项目里的 bat 已经做了改进：

- 会自动尝试 `py`
- 如果没有，再尝试 `python`
- 会自动新开一个不会立刻关闭的命令行窗口
- 会把输出写入日志文件，方便排查

建议这样排查：

1. 先双击：
   - `一键环境检查.bat`
2. 再双击：
   - `运行考勤统计.bat`
   - 或 `打包Windows程序.bat`
3. 如果打包失败，优先查看：
   - `build_windows.log`
   - `build_windows_nicegui.log`
   - `env_check.log`
3. 如果仍失败，请手工打开 Windows 命令行，在项目目录执行：

```bat
py --version
```

如果报错，再试：

```bat
python --version
```

如果两个都不行，说明这台电脑没有正确安装 Python。

### 5. Windows 上看到中文乱码

这通常是命令行编码问题，不一定影响程序本身。

本项目里的 bat 已经加了：

```bat
chcp 65001
```

用来尽量减少乱码。

如果仍然有少量乱码，但程序能正常打包或正常运行，通常可以先忽略。

真正需要优先确认的是：

1. 程序有没有成功启动
2. 有没有成功生成 `考勤统计结果.xlsx`
3. 有没有成功生成 `dist\考勤统计助手.exe`
