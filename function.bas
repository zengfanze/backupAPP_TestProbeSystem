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
    '获取目标文件夹路径
    Set folderParent = folderobject.GetFolder(path & "\" & SearchFolderName)
    '获取目标文件夹下的文件夹总数目，即所有型号的文件夹集合
    Set folderfirstCount = folderParent.subfolders
    Dim firstcnt As Long
    '存储型号个数
    firstcnt = folderfirstCount.Count
    ReDim first(0 To firstcnt) As String
    
    '获取第一级目录下的所有文件夹路径名称，并保存在数组first()中
    Dim I As Long
    I = 0
    For Each firstNum In folderfirstCount
        'first(i) = path & "\" & SearchName & "\" & firstNum.Name
        first(I) = firstNum.Name
        I = I + 1
    Next
   
    '声明
    Dim secondcnt As Long
    Dim StrNewPath As String
    Dim j As Long
    j = 0
    '文件夹存在标志
    Dim folderflog As Boolean
    '日期比较标志
    Dim comparison As Boolean
    On Error GoTo ErrorHandler:
    '第一级目录循环
    For I = 0 To firstcnt - 1
        '获取目标型号的路径
        On Error GoTo ErrorHandler:
        StrNewPath = createfolders(searchPath, savePath, first(I))
        On Error GoTo 0
        On Error Resume Next
        '获取第一级目录下的所有文件夹路径,获取型号路径
         '获取第二级文件夹的集合，即每个型号，所有序列号文件夹的集合
        Set folderSecondCount = folderobject.GetFolder(path & "\" & SearchFolderName & "\" & first(I)).subfolders
        '存储序列号个数
        secondcnt = folderSecondCount.Count
        ReDim second(0 To secondcnt) As String
        '根据型号，查找所有的型号的序列号，并存入数组second()
        For Each secondNum In folderSecondCount
            second(j) = path & "\" & SearchFolderName & "\" & first(I) & "\" & secondNum.Name
             '先检查文件夹是否为空，如果空则删除,并退出该次循环
            folderflog = checkthefolders(second(j))
            If Not folderflog Then
                '文件不存在，那么退出循环
                Exit For
            End If
            '检查文件是否存在，如果不存在，就删除整改文件夹，并退出该次循环
            folderflog = checkthefiles(second(j), "Summary.Data")
            If Not folderflog Then
                '文件不存在，那么退出循环
                Exit For
            End If
            '文件都存在的判断成立后，读取文件
            Call ReadTheFile(second(j), "Summary.Data")
            '判断日期
            comparison = datecomparison(filedata(7), checkdate)
            '如果为真，那么移动文件夹到另外的地方。
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
    MsgBox "完成移动！"
    Exit Function
ErrorHandler:
    Select Case Err.Number
        Case 9
            MsgBox "下标越界9"
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
    '获取目标文件夹路径
    Set folderParent = folderobject.GetFolder(savePath & "\测试系统" & "\test1" & "\" & SearchFolderName)
    '获取目标文件夹下的文件夹总数目，即所有型号的文件夹集合
    Set folderfirstCount = folderParent.subfolders
    Dim firstcnt As Long
    '存储型号个数
    firstcnt = folderfirstCount.Count
    ReDim first(0 To firstcnt) As String
    
    '获取第一级目录下的所有文件夹路径名称，并保存在数组first()中
    Dim I As Long
    I = 0
    For Each firstNum In folderfirstCount
        'first(i) = path & "\" & SearchName & "\" & firstNum.Name
        first(I) = firstNum.Name
        I = I + 1
    Next
    '声明
    Dim secondcnt As Long
    Dim StrNewPath As String
    Dim j As Long
    j = 0
    '文件夹存在标志
    Dim folderflog As Boolean
    '日期比较标志
    Dim comparison As Boolean
    '第一级目录循环
    For I = 0 To firstcnt - 1
        '获取目标型号的路径
        'StrNewPath = createfolders(searchPath, savePath, first(i))
        '获取第一级目录下的所有文件夹路径,获取型号路径
         '获取第二级文件夹的集合，即每个型号，所有序列号文件夹的集合
        Set folderSecondCount = folderobject.GetFolder(savePath & "\测试系统" & "\test1" & "\" & SearchFolderName & "\" & first(I)).subfolders
        '存储序列号个数
        secondcnt = folderSecondCount.Count
        ReDim second(0 To secondcnt) As String
        '根据型号，查找所有的型号的序列号，并存入数组second()
        For Each secondNum In folderSecondCount
            second(j) = savePath & "\测试系统" & "\test1" & "\" & SearchFolderName & "\" & first(I) & "\" & secondNum.Name
             '先检查文件夹是否为空，如果空则删除,并退出该次循环
            folderflog = checkthefolders(second(j))
            If Not folderflog Then
                '文件不存在，那么退出循环
                Exit For
            End If
'            '检查文件是否存在，如果不存在，就删除整改文件夹，并退出该次循环
'            folderflog = checkthefiles(second(j), "Summary.Data")
'            If Not folderflog Then
'                '文件不存在，那么退出循环
'                Exit For
'            End If
            '文件都存在的判断成立后，读取文件
            'Call ReadTheFile(second(j), "Summary.Data")
            '判断日期
'            comparison = datecomparison(filedata(7), checkdate)
'            '如果为真，那么移动文件夹到另外的地方。
'            If comparison = True Then
'                Call movethefolders(second(j), StrNewPath)
'            End If
            Call movethefolders(second(j), openpath & "\" & SearchFolderName & "\" & first(I))
            j = j + 1
        Next
        j = 0
        Call Progress_move_files(firstcnt, I + 1)
    Next I
    MsgBox "完成撤销移动"
End Function
'检查文件是否存在
Public Function checkthefiles(path As String, filename As String) As Boolean
    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(path)
    '如果该文件存在，读取该文件中的值
    If fs.fileExists(path & "\" & filename) Then
        checkthefiles = True
    Else
        '如果不存在，那么删除文件
        f.Delete
        checkthefiles = False
    End If
End Function
'检查文件夹是否为空，如果为空，那么删除文件夹
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
'读取文件内容，
Public Function ReadTheFile(path As String, filename As String)
    Dim m_fileNum As Integer
    m_fileNum = FreeFile()
    Dim I As Integer
    I = 0
    Dim strtemp As String
    '获取文件的行数
    Open path & "\" & filename For Input As #m_fileNum
        Do While Not EOF(1)
        Line Input #m_fileNum, strtemp
            I = I + 1
        Loop
    Close m_fileNum
    '获取每一行的内容
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

'比较日期,如果测试日期比设置的校验日期小，那么移动该目录
Public Function datecomparison(testdate As String, checkdate As String) As Boolean
    '比较日期，如果比校验日期大，那么返回负值，如果比校验日期小，返回正值
    Dim t
    If DateDiff("y", testdate, checkdate) < 0 Then
        datecomparison = False
    Else
        datecomparison = True
    End If
    
End Function
'移动文件夹
Public Function movethefolders(path As String, object As String)
    Dim obj
    Set obj = CreateObject("Scripting.FileSystemObject")
    On Error GoTo ErrorHandler:
        obj.movefolder path, object & "\"
        
ErrorHandler:
    Exit Function
    
End Function

'创建根目录，并返回该目录路径
Public Function createfolders(openpath As String, savePath As String, filename As String)
    Dim parentPath As String
    parentPath = "测试系统"
    Dim obj
    Set obj = CreateObject("Scripting.FileSystemObject")
    parentPath = savePath & "\" & parentPath
    '如果文件夹不存在，则创建
    If Not obj.FolderExists(parentPath) Then
        obj.CreateFolder (parentPath)
    End If
    parentPath = parentPath & "\test1"
    If Not obj.FolderExists(parentPath) Then
        obj.CreateFolder (parentPath)
        '复制一些文件
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
        '复制一些文件
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
    MainForm.status.Caption = "状态： " & percent & "%"
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
    '如果文件夹不存在，则创建
    If Not obj.FolderExists(parentPath) Then
        obj.CreateFolder (parentPath)
    End If
    Open parentPath & "\" & filename For Output As m_fileNum
        
    
    Close m_fileNum
    
End Function
