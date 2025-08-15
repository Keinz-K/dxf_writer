Attribute VB_Name = "Generate_Diagram_dxf"
Option Explicit

Sub main()
    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    Dim an As Integer
    Dim deg As Integer
    Dim fpath As String
    Dim dxf As New Dxfclass
    For j = 1 To 10
        n = j
        deg = 360 / n
        fpath = dxf.DesktopFilepath & "take1¥" & j & ".dxf"
        
        dxf.header (fpath)
        
        For i = 1 To n
            Open fpath For Append As #2
                Print #2, "0"
                Print #2, "LINE"
                Print #2, "8"
                Print #2, "_1-0_"
                Print #2, "6" & vbCrLf & "CONTINUOUS"
                Print #2, "62"
                Print #2, "7"
                Print #2, "10"
                Print #2, 100 + 100 * Sin(deg * (i - 1) * pi / 180)
                Print #2, "20"
                Print #2, 100 + -100 * Cos(deg * (i - 1) * pi / 180)
                Print #2, "11"
                Print #2, 100 + 100 * Sin(deg * i * pi / 180)
                Print #2, "21"
                Print #2, 100 + -100 * Cos(deg * i * pi / 180)
            Close #2
        Next i
        dxf.footer (fpath)
    Next j
End Sub
Private Function pi() As Double
pi = Atn(1) * 4
End Function

