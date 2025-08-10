Attribute VB_Name = "CreateTimeScheduler1"
Sub main()
'A4:210*297
    ptime ("start time")
    Dim i, n, deg As Integer
    Dim generatepath As String
    Dim dxf As New Dxfclass
    n = 24 ' 1d = 24h
    deg = 360 / n
    ChDir dxf.DesktopFilepath
    If (dxf.Init_dxfclass = "") Then Exit Sub
    dxf.header
    
    Call dxf.dxfRect(0, 0, 210, 297)
    Call dxf.dxfCircle(90, 105, 120)
    For i = 1 To n
        Call dxf.dxfLine(105 + 85 * sin(deg * (i - 1) * pi / 180), 120 + 85 * cos(deg * (i - 1) * pi / 180), _
        105 + 90 * sin(deg * (i - 1) * pi / 180), 120 + 90 * cos(deg * (i - 1) * pi / 180))
    Next i
    Call dxf.dxfLine(40, 230, 170, 230)
    Call dxf.dxfCircle(1, 105, 120)
    
    dxf.footer
End Sub
Private Function ptime(mes As String) As String
    ptime = Format(Now, "yyyy-mm-dd-hh-nn-ss")
    Debug.Print mes & " : " & ptime
End Function
Private Function pi() As Double
    pi = Atn(1) * 4
End Function
