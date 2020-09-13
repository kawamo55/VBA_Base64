Attribute VB_Name = "B64Module"
'
'   Base64 Encode -----------> MakeBase64(b() as byte)
'   Base64 Decode -----------> DecBase64(s as string)
'   Make Base64 Encode File -> MakeB64F(fname As String)
'                               Create fname & ".txt"
'   Make Base64 Decode File -> ReadB64F(fname As String)
'                               fname: original file name + Encode data file
'                               Create: original file name
'
'   Ver 1.00.1
'   Coder: M.kawase Embed AI Laboratory Inc
'   Web page:  https://embed-ai.com
'   e-mail:    m_kawase@embed-ai.com
'   Released under the GPL License
'


'
'   Public Padding Count
'
Public Base64_Padding_Count As Long
'
Private Const chartbl As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Private Const padchar As String = "="
'

' ----------------------------------------------------------------> Test Code
'  Encode Test
Sub EncTest()
    Dim bt(3) As Byte
    
    Debug.Print UBound(bt)
    bt(0) = Asc("a")
    bt(1) = Asc("b")
    bt(2) = Asc("c")
    bt(3) = 10
    Debug.Print MakeBase64(bt)
End Sub

'   String Encode Decode Test
Private Sub Base64_Encode_Decode_test()
    Dim b() As Byte, s As String
    
    b = "abcdefg"
    Debug.Print b
    s = MakeBase64(b)
    Debug.Print s
    Debug.Print DecBase64(s)
End Sub

'　Decode From File Test
Sub ts()
    ReadB64F "Your Test file path"
End Sub
'
' <-------------------------------------------------------------------------

'
' Base64File
'
Sub MakeB64F(fname As String)
    Dim i As Long, l As Long, r() As Byte
    Dim of As String, o As String, fary As Variant

    l = FileLen(fname) - 1
    ReDim r(l)
    Open fname For Binary As #1
        Get #1, , r
    Close #1
    of = fname & ".txt"
    o = MakeBase64(r)
    l = Len(o)
    fary = Split(fname, "\")
    i = UBound(fary)
    Open of For Output As #1
        Print #1, fary(i)
        For i = 1 To l Step 76
            Print #1, Mid(o, i, 76)
        Next
    Close #1
End Sub

'
' 呼び出す前にカレントディレクトリを指定しておくこと
' でないとどこにファイルができたかわからない
'
Sub ReadB64F(fname As String)
    Dim tbuf As String, b() As Byte, c() As Byte, m As Long, i As Long
    Dim outfname As String
    ' file name
    Open fname For Input As #1
        Line Input #1, outfname
'       テスト用に変更
'        Open "C:\Users\kawam\Documents\" & outfname For Binary Access Write As #2
        Open outfname For Binary Access Write As #2
        Do While Not EOF(1)
            Line Input #1, tbuf
            b = DecBase64(tbuf)
            If Base64_Padding_Count > 0 Then
                m = UBound(b) - Base64_Padding_Count
                ReDim c(m)
                For i = 0 To m
                    c(i) = b(i)
                Next i
                Put #2, , c
            Else
                Put #2, , b
            End If
        Loop
        Close #2
    Close #1
End Sub

'
Function MakeBase64(s() As Byte) As String
    Dim m As Long, i As Long, l As Long, b As String
    m = UBound(s)
    For i = 0 To m Step 3
        If (i + 2) > m Then
            l = (s(i) And &HFC) / 4 + 1
            b = Mid(chartbl, l, 1)
            Select Case (m Mod 3)
            Case 0
                l = (s(i) And 3) * 16 + 1
                b = b & Mid(chartbl, l, 1)
                b = b & "="
                b = b & "="
            Case 1
                l = (s(i) And 3) * 16 + (s(i + 1) And &HF0) / 16 + 1
                b = b & Mid(chartbl, l, 1)
                l = (s(i + 1) And &HF) * 4 + 1
                b = b & Mid(chartbl, l, 1)
                b = b & "="
            Case 2
                b = b & Mid(chartbl, l, 1)
                l = (s(i + 1) And &HF) * 4 + (s(i + 2) And &HC0) / 64 + 1
                b = b & Mid(chartbl, l, 1)
                l = (s(i + 2) And &H3F) + 1
                b = b & Mid(chartbl, l, 1)
            End Select
        Else
            l = (s(i) And &HFC) / 4 + 1
            b = Mid(chartbl, l, 1)
            l = (s(i) And 3) * 16 + (s(i + 1) And &HF0) / 16 + 1
            b = b & Mid(chartbl, l, 1)
            l = (s(i + 1) And &HF) * 4 + (s(i + 2) And &HC0) / 64 + 1
            b = b & Mid(chartbl, l, 1)
            l = (s(i + 2) And &H3F) + 1
            b = b & Mid(chartbl, l, 1)
        End If
        MakeBase64 = MakeBase64 & b
        b = ""
    Next i
End Function

'
Function DecBase64(s As String) As Byte()
    Dim l As Long, i As Long, bm As Long, bp As Long, lw As Long
    Dim d() As Byte
    
    Base64_Padding_Count = 0
    l = Len(s)
    bm = l / 4 * 3 - 1
    ReDim d(bm)
    For i = 1 To l Step 4
        d(Int((i - 1) / 4) * 3 + 0) = getpos(Mid(s, i, 1)) * 4 + Int(getpos(Mid(s, i + 1, 1)) / 16)
        d(Int((i - 1) / 4) * 3 + 1) = (getpos(Mid(s, i + 1, 1)) And &HF) * 16 + Int((getpos(Mid(s, i + 2, 1)) And &H3C) / 4)
        d(Int((i - 1) / 4) * 3 + 2) = (getpos(Mid(s, i + 2, 1)) And 3) * 64 + getpos(Mid(s, i + 3, 1))
    Next i
    If Right(s, 1) = "=" Then
        Base64_Padding_Count = 0
        If Right(s, 2) = "==" Then
            Base64_Padding_Count = 2
        Else
            Base64_Padding_Count = 1
        End If
    End If
    DecBase64 = d
End Function

'
Private Function getpos(c As String)
    If c = "=" Then
        getpos = 0
    Else
        getpos = InStr(chartbl, c) - 1
    End If
End Function

