Attribute VB_Name = "LIB_File" '定义模块名称

Rem 文件操作类
  ' 来源:https://github.com/emilefyon/Excel-VBA-Productivity-Libraries
  ' 函数简介
  ' writeFile 写文件
  ' readFile 读文件
  ' readFileAndTruncate 读文件并按指定长度截断
  ' fileExists 检测文件是否存在
  ' getFileUpdateTime 获取文件更新时间

Rem 写文件（新建或更新）
  ' File 完整文件路径，例如"D:\123.txt"
  ' content 要写入的内容，例如"我是内容"
Function writeFile(ByVal File As String, ByVal content As String) As String
    
   Open File For Output As #1
   Print #1, content
   Close #1
   
   writeFile = "写入成功"

End Function

Rem 读文件
  ' File 完整文件路径
  ' createFile 文件不存在时是否创建文件，默认为否
Function readFile(ByVal File As String, Optional createFile As Boolean) As String
    
    If (IsMissing(createFile)) Then createFile = False
    
    If (fileExists(File) = False) Then
        If (createFile = True) Then
            temp = writeFile(File, "")
        Else
            readFile = "错误:文件不存在"
            Exit Function
        End If
    End If
    
    Dim MyString, MyNumber
    Open File For Input As #1 ' Open file for input.
    fileContent = ""
    Do While Not EOF(1) ' Loop until end of file.
        Line Input #1, MyString
        ' Debug.Print MyString
        fileContent = fileContent & MyString & " "
    Loop
    Close #1 ' Close file.
    readFile = fileContent
End Function

Rem 读文件并按指定长度截断
  ' File 完整文件路径
  ' createFile 文件不存在时是否创建文件，默认为否
  ' strLen 截断长度，默认为30000
Function readFileAndTruncate(ByVal file As String, Optional createFile As Boolean, Optional strLen As Integer) As String

        If (IsMissing(createFile)) Then createFile = False
        If (IsMissing(strLen)) Then strLen = 30000
    readFileAndTruncate = Left(readFile(file, createFile), strLen)

End Function

Rem 检测文件是否存在
  ' File 完整文件路径
Function fileExists(file As String) As Boolean

	fileExists = false
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.fileExists(file) Then fileExists = True
    Set objFSO = Nothing

End Function

Rem 获取文件更新时间
  ' File 完整文件路径
Function getFileUpdateTime(ByVal file As String) As Double
    
    getFileUpdateTime = FileDateTime(file)

End Function
