Attribute VB_Name = "Module1"
Option Explicit

'回路の図面を作成するコード
'参考図面 以下リンク
' https://akizukidenshi.com/catalog/g/g108241/
Sub main()
    Dim generatepath As String
    Dim dxf As New Dxfclass
    Dim i As Integer
    Dim j As Integer
    if (dxf.Init_dxfclass = "") Then exit sub
    dxf.header
    Call dxf.dxfRect(0, 0, 36, 47)
    For i = 0 To 13
        For j = 0 To 16
            Call dxf.dxfCircle(0.5, 2.98 / 2 + (i * 2.54), 6.36 / 2 + (j * 2.54))
        Next j
    
    Next i
    Call dxf.dxfCircle(1.6, 18, 2.5)
    Call dxf.dxfCircle(1.6, 18, 47 - 2.5)
    dxf.footer
End Sub

