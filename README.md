# ClubManage
社团管理

### 程序功能
读取从教务系统下载的xls文件, 根据通讯录, 制作个人, 部门, 大x, 全体的无课表<br>
可以根据成员的课表, 制作成员的空闲时间, 调用浏览器打开activity.html, 以为活动安排人员<br>
还可以添加, 修改成员的信息(粗糙)

### 适用高校
使用强智教务系统的高校

### 使用环境
python 3.x
使用到的库: xlrd, xlwt

### 文件目录说明
目录: "教务xls课表"(需要自己创建), 用于存放成员的xlsx课表文件

文件: "通讯录.xlsx", 程序会自动寻找含有"通讯录"的xlsx文件, **只允许存在一个通讯录文件**<br>
通讯录, 必须拥有 **部门	职务	姓名	性别	学号	长号	宿舍号**
其中, 宿舍号的格式为"**xx栋yy**"

### 一些问题
xls课表制作完成后, 需要自己设置换行, 详见**必看.docx**

### 非常重要
这是一个非常粗糙的作品
