# WordToHtml
POI将word文件转成HTML进行在线预览<br />
#### 目前测试过word创建的文件,apache poi 创建的文件可能会存在bug<br />
Constant：存放部分静态变量<br/>
Format：docx文件转成html后如果带有复杂的表格，那么会打乱格式，用于统一表格格式。<br/>
ClearAreaCode、WordToHtml：doc文件转html后如果原文件带有目录域的域代码，会一同显示，用于去除域代码。<br/>
ReadWordTable:word表格转化成html方法，可修改样式。<br/>
代码不断完善中，欢迎提出意见
