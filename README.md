# Simple-RPA
## Description:
建立在由b站up主 不高兴就喝水 的RPA框架上\
由Xingyu Qiu改进了鼠标点击功能、command line interface，并增加了自定义按键事件、鼠标移动事件、鼠标位置监测、退出选项

## Setup & Usage:
1.安装python3.4以上版本，并配置环境变量\
教程：https://www.runoob.com/python3/python3-install.html

2.安装依赖包\
方法：在cmd中（win+R  输入cmd  回车）或打开powershell 输入\
pip install pyperclip 回车\
pip install xlrd 回车\
pip install pyautogui==0.9.50 回车\
pip install opencv-python 回车

3.把每一步要操作的图标、区域截图保存至pictures文件夹  png格式（注意如果同屏有多个相同图标，回默认找到最左上的一个，因此怎么截图，截多大的区域，是个学问，如输入框只截中间空白部分肯定是不行的，宗旨就是“唯一”）

4.在cmd.xls 的sheet中，配置每一步的指令\
指令1238对应的内容填截图文件名或坐标如p(500,500)、p330,550\
指令4对应的内容是输入内容\
指令5对应的内容是等待时长（单位秒）\
指令6对应的内容是滚轮滚动的距离，正数表示向上滚，负数表示向下滚\
指令7对应的内容是自定义按键，如enter、shift、tab\
重试次数和重复次数现只针对指令123有用

5.保存文件

6.双击simpleRPA.py打开程序，按照提示输入以执行程序

7.如果报错不能运行用vscode或pycharm运行看看报错内容

8.开始程序后请将程序框最小化，不然程序框挡住的区域是无法识别和操作的

9.如果程序开始后因为你选择了无限重复而鼠标被占用停不下来，alt+F4吧~