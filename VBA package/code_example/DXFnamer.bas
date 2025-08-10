Attribute VB_Name = "DXFnamer"
Option Explicit
Dim originpath As String
Dim generatepath As String
Sub main()
'jwcad -> .dxf に変換した後のデータについて、レイヤの名前を変更するプログラム。
    ptime ("process started")
    Dim i As Integer
    Dim getline As String
    Dim txtline() As String
    Dim dxf As New Dxfclass
    originpath = dxf.getOriginpath() 'レイヤ名の無いファイルを選択する�
    
    If (originpath = "") Then Exit Sub
    
    generatepath = dxf.getSaveAspath() 'ファイルの保存先を指定する�
    
    If (generatepath = originpath) Then Exit Sub
    
    Open originpath For Input As #1
        Do While Not EOF(1): Line Input #1, getline
            If (getline = "ENTITIES") Then
                Do While Not EOF(1)
                    i = i + 1: Line Input #1, getline
                Loop
            End If
        Loop
    Close #1
    
    Open originpath For Input As #1
        Do While Not EOF(1): Line Input #1, getline
            If (getline = "ENTITIES") Then
                ReDim txtline(i): i = 0
                Do While Not EOF(1)
                    i = i + 1: Line Input #1, txtline(i)
                Loop
            End If
        Loop
    Close #1
    dxf.header (generatepath): i = 0
    Open generatepath For Append As #4
        Do While (1) 'レイヤが0グループの0から7レイヤである時、名前を変更する
            i = i + 1:
            If (Left(txtline(i), 5) = "_0-0_") Then txtline(i) = Left(txtline(i), 5) & "ORIGIN"
            If (Left(txtline(i), 5) = "_0-1_") Then txtline(i) = Left(txtline(i), 5) & "CAM01"
            If (Left(txtline(i), 5) = "_0-2_") Then txtline(i) = Left(txtline(i), 5) & "CAM02"
            If (Left(txtline(i), 5) = "_0-3_") Then txtline(i) = Left(txtline(i), 5) & "CAM03"
            If (Left(txtline(i), 5) = "_0-4_") Then txtline(i) = Left(txtline(i), 5) & "CAM04"
            If (Left(txtline(i), 5) = "_0-5_") Then txtline(i) = Left(txtline(i), 5) & "CAM05"
            If (Left(txtline(i), 5) = "_0-6_") Then txtline(i) = Left(txtline(i), 5) & "CAM06"
            If (Left(txtline(i), 5) = "_0-7_") Then txtline(i) = Left(txtline(i), 5) & "CAM07"
            Print #4, txtline(i)
            If (txtline(i) = "EOF") Then Exit Do
        Loop
    Close #4
    ptime ("process finished")
End Sub
Private Function ptime(message As String) As String
    ptime = Format(Now, "yyyy-mm-dd-hh-nn-ss")
    Debug.Print message & " : " & ptime
End Function
