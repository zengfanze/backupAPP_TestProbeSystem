Attribute VB_Name = "value"
Option Explicit
'shell 对象
Global shellobject As New Shell
'文件目录地址对象
Global openfolder As Folder
Global savefolder As Folder
'目标搜索目录
Global searchPath As String
'目标保存目录
Global savePath As String
'第一级文件夹集合数组，即探头型号数组
Global first() As String
'第二级文件夹集合数组，即探头对应的序列号数组
Global second() As String
'保存读取的文件内容
Global filedata() As String
'比较日期
Global checkdate As String
'撤销移动
Global rebackFolders As Boolean

