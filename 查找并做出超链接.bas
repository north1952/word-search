Attribute VB_Name = "查找并做出超链接"
Dim oWord As Object
Dim oDic As Object
Sub 查找并做出超链接()
    Const wdParagraph = 4
   Const wdExtend = 1
    Set oWord = VBA.CreateObject("word.application")
    oWord.Visible = True
    '创建字典对象
    Set oDic = CreateObject("Scripting.Dictionary")
    Dim sPath As String
    '获取文件或者文件夹的路径
    sPath = GetPath()
    If Len(sPath) Then
        Call EnuAllFiles(sPath)
    End If
    '查找到的段落内容
    arrKeys = oDic.keys
    '查找到的段落所在文件路径
    arrItems = oDic.iTems
    Set oDoc = oWord.Documents.Add
    With oDoc
        For i = 0 To UBound(arrKeys)
            oWord.Selection.TypeText Left(arrKeys(i), Len(arrKeys(i)) - 1)
            oWord.Selection.MoveUp wdParagraph, 1, wdExtend
            .Hyperlinks.Add oWord.Selection.Range, arrItems(i)
            oWord.Selection.TypeParagraph
        Next i
'        For i = 1 To .Paragraphs.Count - 1
'            If Len(.Paragraphs(i).Range.Text) > 2 Then
'                .Paragraphs(i).Range.Hyperlinks.Add .Paragraphs(i).Range, arrItems(k)
'                k = k + 1
'            End If
'        Next i
    End With
    Dim oWK As Worksheet
    Set oWK = Sheets("Sheet1")
    With oWK
        .Range("a2:b65536").Clear
        For i = 2 To 2 + UBound(arrKeys)
            .Cells(i, "A") = arrKeys(i - 2)
            .Cells(i, "B") = arrItems(i - 2)
            .Hyperlinks.Add anchor:=.Cells(i, "B"), Address:=arrItems(i - 2)
        Next i
        .Columns.AutoFit
    End With
    Set oDic = Nothing
    '释放word应用程序对象
    Set oWord = Nothing
    MsgBox "处理完成!!!"
End Sub
Function GetPath() As String
    '声明一个FileDialog对象变量
    Dim oFD As FileDialog
'    '创建一个选择文件对话框
'    Set oFD = Application.FileDialog(msoFileDialogFilePicker)
    '创建一个选择文件夹对话框
    Set oFD = Application.FileDialog(msoFileDialogFolderPicker)
    '声明一个变量用来存储选择的文件名
    Dim vrtSelectedItem As Variant
    With oFD
        '允许选择多个文件
        .AllowMultiSelect = True
        '使用Show方法显示对话框，如果单击了确定按钮则返回-1。
        If .Show = -1 Then
            '遍历所有选择的文件
            For Each vrtSelectedItem In .SelectedItems
                '获取所有选择的文件的完整路径,用于各种操作
                GetPath = vrtSelectedItem
            Next
            '如果单击了取消按钮则返回0
        Else
        End If
    End With
    '释放对象变量
    Set oFD = Nothing
End Function
Sub EnuAllFiles(ByVal sPath As String, Optional bEnuSub As Boolean = False)
    Dim oWK As Worksheet
    Set oWK = Excel.Worksheets("Sheet1")
    用户输入 = InputBox("请输入要查询的关键词：")
    '要查找的关键字
    With oWK
       sText = 用户输入
    End With
    '定义文件系统对象
    Dim oFso As Object
    Set oFso = CreateObject("Scripting.FileSystemObject")
    '定义文件夹对象
    Dim oFolder As Object
    Set oFolder = oFso.GetFolder(sPath)
    '定义文件对象
    Dim oFile As Object
    '如果指定的文件夹含有文件
    If oFolder.Files.Count Then
        For Each oFile In oFolder.Files
            With oFile
                '输出文件所在的盘符
                Dim sDrive As String
                sDrive = .Drive
                '输出文件的类型
                Dim sType As String
                sType = .Type
                '输出含后缀名的文件名称
                Dim sName As String
                sName = .Name
                '输出含文件名的完整路径
                Dim sFilePath As String
                sFilePath = .Path
                '如果文件是Word文件且不是隐藏文件
                If sType Like "*ord*" And .Attributes <> 2 Then
                    '以下是对每个文件进行处理的代码
                    '*********************************
                    Debug.Print sFilePath
                    '打开word文档
                    Set oDoc = oWord.Documents.Open(sFilePath)
                    With oDoc
                        Const wdReplaceAll = 2
                        Dim oRng
                        Dim oRng1
                        Set oRng = oWord.Selection.Range
                        '先判断是否有选中区域，没有选中则表示整个文档
                        If oRng.Start = oRng.End Then
                            Set oRng = .Content
                        End If
                        '获取要执行操作的区域的起点和终点，用于查找替换时判断是否超出了选定区域
                        iStart = oRng.Start
                        iEnd = oRng.End
                        Debug.Print oRng.Text
                        Set oRng1 = oRng
                        With oRng1.Find
                            .ClearFormatting
                            .MatchWildcards = True
                            .Text = sText
                            '每执行一次查找，只要找到了结果，oRng对象会自动变成被找到的内容所在的区域
                            Do Until .Execute() = False Or oRng1.Start > iEnd Or oRng1.End < iStart
                                sFindText = oRng1.Paragraphs(1).Range.Text
                                With oDic
                                    If .exists(sFindText) Then
                                        Else
                                        .Add sFindText, sFilePath
                                    End If
                                End With
                            Loop
                        End With
                        '保存word文档
    '                    .Save
                        '关闭word文档
                        .Close
                    End With
                Else
                End If
            End With
        Next
    '如果指定的文件夹不含有文件
    Else
    End If
    '如果要遍历子文件夹
    If bEnuSub = True Then
        '定义子文件夹集合对象
        Dim oSubFolders As Object
        Set oSubFolders = oFolder.SubFolders
        If oSubFolders.Count > 0 Then
            For Each oTempFolder In oSubFolders
                sTempPath = oTempFolder.Path
                Call EnuAllFiles(sTempPath, True)
            Next
        End If
        Set oSubFolders = Nothing
    End If
    Set oFile = Nothing
    Set oFolder = Nothing
    Set oFso = Nothing
End Sub

