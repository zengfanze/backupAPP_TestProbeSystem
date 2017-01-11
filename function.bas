Attribute VB_Name = "function"
Option Explicit

Public Function searchAndMoveObjectFolder(path As String, SearchFolderName As String)
    
    Dim folderParent
    Dim folderfirstCount, firstNum, folderfirst
    Dim folderSecondCount, secondNum
    Dim folderobject As Object
    Dim bool_exists As Boolean
    bool_exists = False
    Dim f
    Set folderobject = CreateObject("Scripting.FileSystemObject")
    '��ȡĿ���ļ���·��
    Set folderParent = folderobject.GetFolder(path & "\" & SearchFolderName)
    '��ȡĿ���ļ����µ��ļ�������Ŀ���������ͺŵ��ļ��м���
    Set folderfirstCount = folderParent.subfolders
    Dim firstcnt As Long
    '�洢�ͺŸ���
    firstcnt = folderfirstCount.Count
    ReDim first(0 To firstcnt) As String
    
    '��ȡ��һ��Ŀ¼�µ������ļ���·�����ƣ�������������first()��
    Dim I As Long
    I = 0
    For Each firstNum In folderfirstCount
        'first(i) = path & "\" & SearchName & "\" & firstNum.Name
        first(I) = firstNum.Name
        I = I + 1
    Next
   
    '����
    Dim secondcnt As Long
    Dim StrNewPath As String
    Dim j As Long
    j = 0
    '�ļ��д��ڱ�־
    Dim folderflog As Boolean
    '���ڱȽϱ�־
    Dim comparison As Boolean
    On Error GoTo ErrorHandler:
    '��һ��Ŀ¼ѭ��
    For I = 0 To firstcnt - 1
        '��ȡĿ���ͺŵ�·��
        On Error GoTo ErrorHandler:
        StrNewPath = createfolders(searchPath, savePath, first(I))
        On Error GoTo 0
        On Error Resume Next
        '��ȡ��һ��Ŀ¼�µ������ļ���·��,��ȡ�ͺ�·��
         '��ȡ�ڶ����ļ��еļ��ϣ���ÿ���ͺţ��������к��ļ��еļ���
        Set folderSecondCount = folderobject.GetFolder(path & "\" & SearchFolderName & "\" & first(I)).subfolders
        '�洢���кŸ���
        secondcnt = folderSecondCount.Count
        ReDim second(0 To secondcnt) As String
        '�����ͺţ��������е��ͺŵ����кţ�����������second()
        For Each secondNum In folderSecondCount
            second(j) = path & "\" & SearchFolderName & "\" & first(I) & "\" & secondNum.Name
             '�ȼ���ļ����Ƿ�Ϊ�գ��������ɾ��,���˳��ô�ѭ��
            folderflog = checkthefolders(second(j))
            If Not folderflog Then
                '�ļ������ڣ���ô�˳�ѭ��
                Exit For
            End If
            '����ļ��Ƿ���ڣ���������ڣ���ɾ�������ļ��У����˳��ô�ѭ��
            folderflog = checkthefiles(second(j), "Summary.Data")
            If Not folderflog Then
                '�ļ������ڣ���ô�˳�ѭ��
                Exit For
            End If
            '�ļ������ڵ��жϳ����󣬶�ȡ�ļ�
            Call ReadTheFile(second(j), "Summary.Data")
            '�ж�����
            comparison = datecomparison(filedata(7), checkdate)
            '���Ϊ�棬��ô�ƶ��ļ��е�����ĵط���
            If comparison = True Then
                Call movethefolders(second(j), StrNewPath)
            End If
            MainForm.Label4.Caption = "i:" & I & "%%" & "j:" & j
            MainForm.Refresh
            If j < 8 Then
                MainForm.filelistbox.AddItem first(I) & "\" & secondNum.Name, j
            Else
                MainForm.filelistbox.RemoveItem 0
                MainForm.filelistbox.AddItem first(I) & "\" & secondNum.Name, 8
            End If
            j = j + 1
        Next
        j = 0
        Call Progress_move_files(firstcnt, I + 1)
    Next I
    rebackFolders = False
    MsgBox "����ƶ���"
    Exit Function
ErrorHandler:
    Select Case Err.Number
        Case 9
            MsgBox "�±�Խ��9"
            Exit Function
        End Select
    Resume
    
End Function
Public Function rebackTheFolder(openpath As String, savePath As String, SearchFolderName As String)
    Dim folderParent
    Dim folderfirstCount, firstNum, folderfirst
    Dim folderSecondCount, secondNum
    Dim folderobject As Object
    Dim bool_exists As Boolean
    bool_exists = False
    Dim f
    Set folderobject = CreateObject("Scripting.FileSystemObject")
    '��ȡĿ���ļ���·��
    Set folderParent = folderobject.GetFolder(savePath & "\����ϵͳ" & "\test1" & "\" & SearchFolderName)
    '��ȡĿ���ļ����µ��ļ�������Ŀ���������ͺŵ��ļ��м���
    Set folderfirstCount = folderParent.subfolders
    Dim firstcnt As Long
    '�洢�ͺŸ���
    firstcnt = folderfirstCount.Count
    ReDim first(0 To firstcnt) As String
    
    '��ȡ��һ��Ŀ¼�µ������ļ���·�����ƣ�������������first()��
    Dim I As Long
    I = 0
    For Each firstNum In folderfirstCount
        'first(i) = path & "\" & SearchName & "\" & firstNum.Name
        first(I) = firstNum.Name
        I = I + 1
    Next
    '����
    Dim secondcnt As Long
    Dim StrNewPath As String
    Dim j As Long
    j = 0
    '�ļ��д��ڱ�־
    Dim folderflog As Boolean
    '���ڱȽϱ�־
    Dim comparison As Boolean
    '��һ��Ŀ¼ѭ��
    For I = 0 To firstcnt - 1
        '��ȡĿ���ͺŵ�·��
        'StrNewPath = createfolders(searchPath, savePath, first(i))
        '��ȡ��һ��Ŀ¼�µ������ļ���·��,��ȡ�ͺ�·��
         '��ȡ�ڶ����ļ��еļ��ϣ���ÿ���ͺţ��������к��ļ��еļ���
        Set folderSecondCount = folderobject.GetFolder(savePath & "\����ϵͳ" & "\test1" & "\" & SearchFolderName & "\" & first(I)).subfolders
        '�洢���кŸ���
        secondcnt = folderSecondCount.Count
        ReDim second(0 To secondcnt) As String
        '�����ͺţ��������е��ͺŵ����кţ�����������second()
        For Each secondNum In folderSecondCount
            second(j) = savePath & "\����ϵͳ" & "\test1" & "\" & SearchFolderName & "\" & first(I) & "\" & secondNum.Name
             '�ȼ���ļ����Ƿ�Ϊ�գ��������ɾ��,���˳��ô�ѭ��
            folderflog = checkthefolders(second(j))
            If Not folderflog Then
                '�ļ������ڣ���ô�˳�ѭ��
                Exit For
            End If
'            '����ļ��Ƿ���ڣ���������ڣ���ɾ�������ļ��У����˳��ô�ѭ��
'            folderflog = checkthefiles(second(j), "Summary.Data")
'            If Not folderflog Then
'                '�ļ������ڣ���ô�˳�ѭ��
'                Exit For
'            End If
            '�ļ������ڵ��жϳ����󣬶�ȡ�ļ�
            'Call ReadTheFile(second(j), "Summary.Data")
            '�ж�����
'            comparison = datecomparison(filedata(7), checkdate)
'            '���Ϊ�棬��ô�ƶ��ļ��е�����ĵط���
'            If comparison = True Then
'                Call movethefolders(second(j), StrNewPath)
'            End If
            Call movethefolders(second(j), openpath & "\" & SearchFolderName & "\" & first(I))
            j = j + 1
        Next
        j = 0
        Call Progress_move_files(firstcnt, I + 1)
    Next I
    MsgBox "��ɳ����ƶ�"
End Function
'����ļ��Ƿ����
Public Function checkthefiles(path As String, filename As String) As Boolean
    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(path)
    '������ļ����ڣ���ȡ���ļ��е�ֵ
    If fs.fileExists(path & "\" & filename) Then
        checkthefiles = True
    Else
        '��������ڣ���ôɾ���ļ�
        f.Delete
        checkthefiles = False
    End If
End Function
'����ļ����Ƿ�Ϊ�գ����Ϊ�գ���ôɾ���ļ���
Public Function checkthefolders(path As String) As Boolean
    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(path)
    If f.Size = 0 Then
        f.Delete
        checkthefolders = False
    Else
        checkthefolders = True
    End If
    
End Function
'��ȡ�ļ����ݣ�
Public Function ReadTheFile(path As String, filename As String)
    Dim m_fileNum As Integer
    m_fileNum = FreeFile()
    Dim I As Integer
    I = 0
    Dim strtemp As String
    '��ȡ�ļ�������
    Open path & "\" & filename For Input As #m_fileNum
        Do While Not EOF(1)
        Line Input #m_fileNum, strtemp
            I = I + 1
        Loop
    Close m_fileNum
    '��ȡÿһ�е�����
    Dim j As Integer
    j = 0
    m_fileNum = FreeFile()
    ReDim filedata(0 To I) As String
    Open path & "\" & filename For Input As #m_fileNum
        Do While j < I
        Line Input #m_fileNum, filedata(j)
            j = j + 1
        Loop
    Close m_fileNum
End Function

'�Ƚ�����,����������ڱ����õ�У������С����ô�ƶ���Ŀ¼
Public Function datecomparison(testdate As String, checkdate As String) As Boolean
    '�Ƚ����ڣ������У�����ڴ���ô���ظ�ֵ�������У������С��������ֵ
    Dim t
    If DateDiff("y", testdate, checkdate) < 0 Then
        datecomparison = False
    Else
        datecomparison = True
    End If
    
End Function
'�ƶ��ļ���
Public Function movethefolders(path As String, object As String)
    Dim obj
    Set obj = CreateObject("Scripting.FileSystemObject")
    On Error GoTo ErrorHandler:
        obj.movefolder path, object & "\"
        
ErrorHandler:
    Exit Function
    
End Function

'������Ŀ¼�������ظ�Ŀ¼·��
Public Function createfolders(openpath As String, savePath As String, filename As String)
    Dim parentPath As String
    parentPath = "����ϵͳ"
    Dim obj
    Set obj = CreateObject("Scripting.FileSystemObject")
    parentPath = savePath & "\" & parentPath
    '����ļ��в����ڣ��򴴽�
    If Not obj.FolderExists(parentPath) Then
        obj.CreateFolder (parentPath)
    End If
    parentPath = parentPath & "\test1"
    If Not obj.FolderExists(parentPath) Then
        obj.CreateFolder (parentPath)
        '����һЩ�ļ�
'        Call copyfiles(openPath, parentPath, "*.jpg")
'        Call copyfiles(openPath, parentPath, "*.png")
'        Call copyfiles(openPath, parentPath, "*.bmp")
'        Call copyfiles(openPath, parentPath, "*.exe")
         Call copyfiles(openpath, parentPath, "*.*")
    End If
    If Not obj.FolderExists(parentPath & "\Wafer") Then
        obj.CreateFolder (parentPath & "\Wafer")
    End If
    
    parentPath = parentPath & "\Probe"
    If Not obj.FolderExists(parentPath) Then
        obj.CreateFolder (parentPath)
        '����һЩ�ļ�
'       Call copyfiles(openPath & "\Probe", parentPath, "*.xls")
'       Call copyfiles(openPath & "\Probe", parentPath, "*.data")
        Call copyfiles(openpath & "\Probe", parentPath, "*.*")
    End If
    
    parentPath = parentPath & "\" & filename
    If Not obj.FolderExists(parentPath) Then
        obj.CreateFolder (parentPath)
'        Call copyfiles(openPath & "\Probe" & "\" & filename, parentPath, "*.xls")
'        Call copyfiles(openPath & "\Probe" & "\" & filename, parentPath, "*.data")
        Call copyfiles(openpath & "\Probe" & "\" & filename, parentPath, "*.*")
    End If
    createfolders = parentPath
End Function
Public Function copyfiles(path As String, savePath As String, filename As String)
    Dim obj
    Set obj = CreateObject("Scripting.FileSystemObject")
    obj.copyfile path & "\" & filename, savePath & "\"
End Function

Public Function Progress_move_files(max_cnt As Long, cnt As Long)
    Dim percent As Single
    percent = Format((cnt / max_cnt) * 100, "0.00")
    MainForm.status.Caption = "״̬�� " & percent & "%"
    MainForm.Refresh
End Function
Public Function saveTheParam(appPath As String, filename As String)
    Dim m_fileNum As Integer
    m_fileNum = FreeFile()
    Dim obj
    Set obj = CreateObject("Scripting.FileSystemObject")
    Dim parentPath As String
    parentPath = "config"
    parentPath = appPath & "\" & parentPath
    '����ļ��в����ڣ��򴴽�
    If Not obj.FolderExists(parentPath) Then
        obj.CreateFolder (parentPath)
    End If
    Open parentPath & "\" & filename For Output As m_fileNum
        
    
    Close m_fileNum
    
End Function
