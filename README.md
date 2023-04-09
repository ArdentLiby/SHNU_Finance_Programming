# SHNU_Finance_Programming

# 1、基本信息

## 1.1 版本与系统

本项目基于 Python 3.8.8 编写，支持 Windows 10/11 和 MacOS，对于低于 Windows 10 版本的 Windows 系统的支持情况未经测试。



## 1.2 程序说明

此项目的编写、测试和说明均由个人完成。由于个人对于 MacOS 的使用较少，少数功能在 Mac 版的程序中无法实现，将在下面标出。此项目的主要功能如下：

1、输入学生信息并根据选择的成绩表生成相应的平均分、文字结果和证明文档（以下所说生成的证明文档中均包含中、英文证明信息，生成的文档命名方式为“姓名_学号[学籍状态].docx(.pdf) ”，Mac版本无法生成 PDF 文档）；

2、根据选择的学生信息表和给定的学号导入学生的完整信息（包括中、英文信息），避免逐个输入学生信息，也可通过“重置”功能清空已经输入的值；

3、可自定义学生信息表和学生成绩表中相应信息所处的列，增加了灵活性；

4、可添加或删除专业及对应的英文名称；

5、可自选需要保存的文档类型，并根据选择的学生信息表和成绩表文件夹（文件夹中的成绩表名称须包含学号(如：张三_190123456.xlsx)，否则无法生成文档）批量生成证明文档；

6、可自选生成的证明文档的保存位置。



# 2、使用说明

本项目无需安装，打开目录中的可执行文件即可直接运行。初次运行或配置文件缺失时程序将自动创建 “config2.json” 、 “config-batch.json” 两个配置文件和“ ouputs” 文件夹。

本程序不会申请任何管理员权限，因此请将此项目移入无需管理员权限的文件夹。



## 2.1 设置内容介绍

由于实现各项功能之前需要进行简单设置，故首先对设置界面进行说明。

设置页面自上而下是：设置学生信息表路径、设置学生信息表中各项学生信息的所在列数、设置输出文档的位置、成绩条数及成绩表中重要信息的所在列、选择输出文档类型，最后是添加或删除专业及对应英文信息。



### 2.1.1 学生信息表的导入和设置

进行设置时，首先选择学生信息表路径作为“ 导入学生信息” 和“ 批量生成文档” 功能的参照表；并依次填写学生姓名、学生学号、学生性别、班级或专业信息、年级信息（如班级信息中包含年级可留空）、学籍信息所在的列数（Excel 中从左往右数的列数）。

注1：班级或专业信息支持的格式包括：“ 年级+专业信息+本科x班”、“ 专业信息+本科x班”  、“ 仅专业信息” 和“ 年级+专业信息” 。例：“ 2019经济学(中美合作)本科1班” 、“ 经济学(中美合作)本科1班” 、“ 经济学(中美合作)” 、“ 2019级经济学(中美合作)” 均是可行的。其余格式均不支持。

注2：如果班级或专业信息中包含年级，则年级信息可留空；否则年级信息必填。其余**所有信息**均为必填。



### 2.1.2 其他设置

导出文档位置：请选择一个**没有管理员权限**的文件夹作为自定义导出文件夹；或使用默认文件夹（程序目录下的 outputs 文件夹）。

成绩表相关设置：成绩表所可以包含的最大成绩条数默认为200，如需修改请输入正整数。成绩Excel表中的学分和分数的所在列也可根据实际需求更改。

输出文档类型：根据需求选择，但在无命令行模式下（w 模式）或 Mac 版中，无法生成 PDF 文档。

增加或删除专业：第一空填写要添加的专业的中文名称（如包含括号请使用中文括号），第二空填写对应英文名即可新增专业；第一空选择专业名并在第二空输入内容即可更改专业英文名称；第一空选择专业名称后第二空留空即可删除相应的专业。增加、更改、删除专业名称均在点击“ 设为默认” 后生效。

\*外观设置：将 PNG 图片放在主程序的同一目录下将显示背景。如有需要，请尽量选择与窗口尺寸（545\*680) 比例相似的图片。



## 2.2 基本功能介绍

可以逐个输入学生信息，或基于选择的学生信息表通过学号导入学生的信息；计分方式和成绩表需要自行选择。

所有信息填写完毕后，点击“ 计算并输出文字结果” 将生成平均分和中英文说明文字的可选文本；点击“导出文档” 后将根据填写的学生信息和成绩表及设置中选择的导出文档类型导出相应格式的文档，同时也会输出文字结果。

点击“ 重置” 将会清空所有的输入。



## 2.3 批量生成文档相关

选择学生成绩表存放的文件夹和计分方式后，将根据设置中的学生信息表的内容批量生成成绩证明文档。

注：如需生成 PDF 文档，请使用 Windows 版中“ main.exe” 文件。PDF 文档的生成所需时间较长，请耐心等待。



# 3、程序代码简要说明

本项目基于 Python 版本 3.8.8 编写，需要 Python >= 3.7 以运行，过低的版本可能会出现不兼容问题。

需要的第三方库：pandas、pypinyin、python-docx、docx2pdf、pillow。如未安装可在 cmd 中安装所需要的第三方库：

```
pip install -i https://pypi.tuna.tsinghua.edu.cn/simple pandas
pip install -i https://pypi.tuna.tsinghua.edu.cn/simple pypinyin
pip install -i https://pypi.tuna.tsinghua.edu.cn/simple python-docx
pip install -i https://pypi.tuna.tsinghua.edu.cn/simple docx2pdf
pip install -i https://pypi.tuna.tsinghua.edu.cn/simple pillow
```





## 3.1 程序源文件

### 3.1.1 main.py

该文件是项目的主程序，包含项目的主体逻辑和两个主要功能（根据学号导入学生信息和保存设置），运行该文件即可开启项目。



### 3.1.2 assist.py

该文件是程序的主要功能函数的集成文件。项目的大部分功能如选择文件（夹）的处理、计算成绩、显示弹窗等。



### 3.1.3 genDocx.py

该文件控制生成文档，会获取当前日期并根据选择生成 Word 和 PDF 两种文档并保存在指定路径下。



### 3.1.4 batchExport.py

是控制批量生成文档的文件，根据设置中选择的学生信息表和自定义内容识别学生，并根据成绩表文件名批量生成对应学生的成绩证明文档。



### 3.1.5 initConfig.py

保存初始的设置，即第一次运行主程序时将以该文件中的相关设置作为初始默认设置。



### \*3.1.6 main-mac.py

使用 Mac 系统的主程序文件，功能与“ main.py” 相同。



## 3.2 配置文件

### 3.2.1 config2.json

是本项目的主要配置文件，保存了中英文专业名称、文档保存路径等信息。当该文件不存在时将以 initConfig.py 中的内容生成初始配置文件。



### 3.2.2 config-batch.json

控制批量生成的配置文件，只保存了批量生成文件的文档类型和需要批处理的文件两项设置。批处理功能的其他设置均以主要配置文件为准。



# 4、其他说明

项目 GitHub 地址：https://github.com/ArdentLiby/SHNU_Finance_Programming.git

所有代码、配置文件和测试数据均已上传。

