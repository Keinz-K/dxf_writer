Attribute VB_Name = "Create_dxf_scad"
Option Explicit
'コードを生成した後、openscadで読み込むためのコード(サンプル)
'linear_extrude(height = 10, convexity = 1)
'  import(file = "output.dxf", layer = "0");      

Sub Main_scad()
    Dim generatepath As String
    Dim dxf As New Dxfclass_scad
    If (dxf.Init_dxfclass = "") Then Exit Sub
    dxf.header_scad
    Call dxf.dxfRect_scad(0, 0, 4000, 4000)
    Call dxf.dxfRect_scad(250, 250, 250, 250)
    Call dxf.dxfRect_scad(750, 250, 250, 250)
    Call dxf.dxfRect_scad(250, 750, 250, 250)
    Call dxf.dxfRect_scad(750, 750, 500, 500)
    Call dxf.dxfRect_scad(3750, 3750, -500, -500)
    Call dxf.dxfRect_scad(3750, 250, -500, 500)
    Call dxf.dxfRect_scad(250, 3750, 500, -500)
    Call dxf.dxfRect_scad(1990, 0, 20, 4000)
    Call dxf.dxfRect_scad(0, 1990, 4000, 20)
    dxf.footer_scad
End Sub

