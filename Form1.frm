VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "测试系统探头数据移动备份程序"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   7080
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox filelistbox 
      Height          =   1860
      Left            =   360
      TabIndex        =   11
      Top             =   5760
      Width           =   6135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "撤销移动"
      Height          =   615
      Left            =   4440
      TabIndex        =   10
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox datecheck 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy-MM-dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Text            =   "2014-12-30"
      Top             =   4080
      Width           =   3060
   End
   Begin VB.CommandButton Command3 
      Caption         =   "开始移动（备份）"
      Height          =   615
      Left            =   4440
      TabIndex        =   4
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "打开"
      Height          =   615
      Left            =   4440
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox TextOpenPath 
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox savefolderPath 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存"
      Height          =   615
      Left            =   4440
      TabIndex        =   0
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   7800
      Width           =   6135
   End
   Begin VB.Label status 
      Caption         =   "状态：就绪"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "目标文件路径：请选择测试系统文件夹下的“test1”为目标文件路径"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "保存的文件路径"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "日期比较规则：                    如果对应探头序列号的测试日期小于这个日期，那么该文件夹会被移动 "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   6
      Top             =   3120
      Width           =   3135
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim strPath As String
    Dim m_ifileNum As Integer
    Dim m_size As Integer
    '获取保存的文件目录
    Set savefolder = shellobject.BrowseForFolder(0, "选择文件夹", 0)
    '获取路径
    If savefolder Is Nothing Then
        MsgBox "请指定目标系统文件目录！"
        searchPath = ""
    ElseIf savefolder = "桌面" Then
        MsgBox "zm!"
        searchPath = "C:\Documents and Settings\Administrator\桌面"
    Else
        Me.savefolderPath.Text = savefolder.Items.Item.path
        savePath = savefolder.Items.Item.path
    End If
    
End Sub

Private Sub Command2_Click()
    Dim strPath As String
    Dim m_ifileNum As Integer
    Dim m_size As Integer
    Set openfolder = shellobject.BrowseForFolder(0, "请选择目标文件夹", 0, 0)
    '获取路径
    If openfolder Is Nothing Then
        MsgBox "请指定目标系统文件目录！"
        searchPath = ""
    ElseIf openfolder = "桌面" Then
        MsgBox "zm!"
        searchPath = "C:\Documents and Settings\Administrator\桌面"
    Else
        Me.TextOpenPath.Text = openfolder.Items.Item.path
        searchPath = openfolder.Items.Item.path
    End If
    '释放openfolder
    Set openfolder = Nothing
'        If searchPath <> "" Then
'        m_ifileNum = FreeFile()
'        Open searchPath & "\" & "123.txt" For Output As #m_ifileNum
'        Print #m_ifileNum, StrTemp
'        Print #m_ifileNum, StrTemp
'        Print #m_ifileNum, StrTemp
'        Print #m_ifileNum, StrTemp
'        关闭打开的文件
'        Close m_ifileNum
'    End If
End Sub
Private Sub Command3_Click()
        
    searchPath = Me.TextOpenPath.Text
    savePath = Me.savefolderPath.Text
    Dim checkover As Boolean
    checkdate = Me.datecheck.Text
    If datecheck = "" Then
        MsgBox "请填写日期"
    Else
        checkover = True
    End If

    If searchPath <> "" And checkover Then
        Call searchAndMoveObjectFolder(searchPath, "Probe")
        rebackFolders = True
    End If
    
End Sub

Private Sub Command4_Click()
    searchPath = Me.TextOpenPath.Text
    savePath = Me.savefolderPath.Text
    If rebackFolders Then
        Call rebackTheFolder(searchPath, savePath, "Probe")
    End If
End Sub

