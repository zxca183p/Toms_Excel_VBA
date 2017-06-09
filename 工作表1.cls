VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "工作表1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim temp As Double

Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False '暫時停止事件觸發
    If Not Intersect(Target, Range("B1")) Is Nothing Then
        Add
        enter_bool = False
    End If
    Application.EnableEvents = True '啟用事件觸發
    
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Not Intersect(Target, Range("B1")) Is Nothing Then
        temp = Cells(1, 2).Value
    End If
End Sub

Private Sub Add()
    If Not enter_bool Then
        Cells(1, 2).Value = temp
        Exit Sub
    End If
    
    If Cells(1, 2).Value = 0 Then
        Cells(1, 2).Value = 0
        Exit Sub
    End If
    
    If TypeName(Cells(1, 2).Value) = "Double" Then
        Cells(1, 2).Value = Cells(1, 2).Value + temp
    Else
        Cells(1, 2).Value = temp
    End If
End Sub

