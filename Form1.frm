VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "����ϵͳ̽ͷ�����ƶ����ݳ���"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   7080
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ListBox filelistbox 
      Height          =   1860
      Left            =   360
      TabIndex        =   11
      Top             =   5760
      Width           =   6135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�����ƶ�"
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
      Caption         =   "��ʼ�ƶ������ݣ�"
      Height          =   615
      Left            =   4440
      TabIndex        =   4
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��"
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
      Caption         =   "����"
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
      Caption         =   "״̬������"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Ŀ���ļ�·������ѡ�����ϵͳ�ļ����µġ�test1��ΪĿ���ļ�·��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "������ļ�·��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���ڱȽϹ���                    �����Ӧ̽ͷ���кŵĲ�������С��������ڣ���ô���ļ��лᱻ�ƶ� "
      BeginProperty Font 
         Name            =   "����"
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
    '��ȡ������ļ�Ŀ¼
    Set savefolder = shellobject.BrowseForFolder(0, "ѡ���ļ���", 0)
    '��ȡ·��
    If savefolder Is Nothing Then
        MsgBox "��ָ��Ŀ��ϵͳ�ļ�Ŀ¼��"
        searchPath = ""
    ElseIf savefolder = "����" Then
        MsgBox "zm!"
        searchPath = "C:\Documents and Settings\Administrator\����"
    Else
        Me.savefolderPath.Text = savefolder.Items.Item.path
        savePath = savefolder.Items.Item.path
    End If
    
End Sub

Private Sub Command2_Click()
    Dim strPath As String
    Dim m_ifileNum As Integer
    Dim m_size As Integer
    Set openfolder = shellobject.BrowseForFolder(0, "��ѡ��Ŀ���ļ���", 0, 0)
    '��ȡ·��
    If openfolder Is Nothing Then
        MsgBox "��ָ��Ŀ��ϵͳ�ļ�Ŀ¼��"
        searchPath = ""
    ElseIf openfolder = "����" Then
        MsgBox "zm!"
        searchPath = "C:\Documents and Settings\Administrator\����"
    Else
        Me.TextOpenPath.Text = openfolder.Items.Item.path
        searchPath = openfolder.Items.Item.path
    End If
    '�ͷ�openfolder
    Set openfolder = Nothing
'        If searchPath <> "" Then
'        m_ifileNum = FreeFile()
'        Open searchPath & "\" & "123.txt" For Output As #m_ifileNum
'        Print #m_ifileNum, StrTemp
'        Print #m_ifileNum, StrTemp
'        Print #m_ifileNum, StrTemp
'        Print #m_ifileNum, StrTemp
'        �رմ򿪵��ļ�
'        Close m_ifileNum
'    End If
End Sub
Private Sub Command3_Click()
        
    searchPath = Me.TextOpenPath.Text
    savePath = Me.savefolderPath.Text
    Dim checkover As Boolean
    checkdate = Me.datecheck.Text
    If datecheck = "" Then
        MsgBox "����д����"
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

