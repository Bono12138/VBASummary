VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub changeCellValue()

    Dim MyPath          As String
    Dim MyFile          As String
    Dim Wkb             As Workbook
    Dim Cnt             As Long
    Dim fileType        As String
    Dim cellLocation    As String
    Dim cellContent     As String
    
    Application.ScreenUpdating = False
    
    MyPath = Worksheets("指定位置替换").Range("B2").Value 'change the path accordingly
    
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    fileType = Range("B3").Value
    MyFile = Dir(MyPath & "*." & fileType)
    
    Cnt = 0
    Do While Len(MyFile) > 0
        Cnt = Cnt + 1
        Set Wkb = Workbooks.Open(Filename:=MyPath & MyFile, editable:=True)
        cellLocation = Range("B4").Value
        cellContent = Range("B5").Value
        Wkb.Worksheets(1).Range(cellLocation).Value = cellContent 'change the new value accordingly
        Wkb.Close savechanges:=True
        MyFile = Dir
    Loop
    
    If Cnt > 0 Then
        MsgBox "整完了！", vbExclamation
    Else
        MsgBox "没有找到要修改的文件啊!", vbExclamation
    End If
    
    Application.ScreenUpdating = True
    
End Sub

