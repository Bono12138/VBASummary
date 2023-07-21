VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub renameWorkbook()
    For r = 4 To Worksheets("重命名文件").Range("A" & Rows.Count).End(xlUp).Row
        filePath = Range("A" & r)
        oldFileNm = Range("B" & r)
        newFileNm = Range("C" & r)
        
        If Right(filePath, 1) <> "\" Then filePath = MyPath & "\"
        Name filePath & oldFileNm As filePath & newFileNm
    Next r
    MsgBox "整好了"
End Sub
