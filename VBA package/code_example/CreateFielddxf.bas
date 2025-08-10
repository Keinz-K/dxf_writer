Attribute VB_Name = "CreateFielddxf"
Option Explicit
Sub main()
    ptime ("start time")
    Dim generatepath As String
    Dim dxf As New Dxfclass
    If (dxf.Init_dxfclass = "") Then Exit Sub
    dxf.header
    Call dxf.dxfRect(0, 0, 4000, 4000)
    Call dxf.dxfRect(250, 250, 250, 250)
    Call dxf.dxfRect(750, 250, 250, 250)
    Call dxf.dxfRect(250, 750, 250, 250)
    Call dxf.dxfRect(750, 750, 500, 500)
    Call dxf.dxfRect(3750, 3750, -500, -500)
    Call dxf.dxfRect(3750, 250, -500, 500)
    Call dxf.dxfRect(250, 3750, 500, -500)
    Call dxf.dxfRect(1990, 0, 20, 4000)
    Call dxf.dxfRect(0, 1990, 4000, 20)
    dxf.footer
End Sub
Private Function ptime(mes As String) As String
    ptime = Format(Now, "yyyy-mm-dd-hh-nn-ss")
    Debug.Print mes & " : " & ptime
End Function
Private Function pi() As Double
    pi = Atn(1) * 4
End Function
