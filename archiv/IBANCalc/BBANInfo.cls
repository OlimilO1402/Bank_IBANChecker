VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BBANInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_infoW As String 'W -> wie von der Wikipedia Seite in die Datei kopiert
Private m_infoR As String 'R -> wie aus der Datei gelesen und umgewandelt
Private m_col   As New Collection

Public Sub New_(ByVal infoW As String)
    If InStr(1, infoW, " ") > 0 Then infoW = Replace(infoW, " ", "")
    m_infoW = infoW
    m_infoR = ParseInfoR
    Parse
End Sub
Public Function ToStr(Optional br As Boolean = False) As String
    ToStr = ToInfoW
    If br Then ToStr = ToInfoR
End Function

Public Function BBANLength() As Byte
    BBANLength = Len(m_infoW)
End Function
Public Function ToInfoW(Optional bGroupedBy4 As Boolean = False) As String
    'so wie es von Wikipedia gelesen wurde
    ToInfoW = m_infoW
    If bGroupedBy4 Then ToInfoW = GroupBy4(ToInfoW)
End Function
Private Sub Parse()
'die collection von BBANParts aufbauen
    'Dim r As String: r = m_infoR
    Dim sArr() As String: sArr = Split(m_infoR, "; ")
    Dim Pos As Byte: Pos = 1
    Dim i As Integer
    For i = 0 To UBound(sArr) - 1 ' Index - 1
        Dim lArr() As String: lArr = Split(sArr(i), ": ")
        m_col.Add MNew.BBANPart(lArr(0), Pos, Trim(lArr(1)))
        Pos = Pos + Trim(lArr(1))
    Next
    'lArr = Split(sArr(Index), ": ")
    'Set BBANPart = MNew.BBANPart(lArr(0), Pos, Trim(lArr(1)))
End Sub
Public Property Get BBANPart(ByVal Index As Integer) As BBANPart
    Set BBANPart = m_col(Index + 1)
'    Dim r As String: r = ToInfoR
'    Dim sArr() As String: sArr = Split(r, "; ")
'    Dim lArr() As String
'    Dim Pos As Byte, i As Integer
'    Pos = 1
'    For i = 0 To Index - 1
'        lArr = Split(sArr(i), ": ")
'        Pos = Pos + Trim(lArr(1))
'    Next
'    lArr = Split(sArr(Index), ": ")
'    Set BBANPart = MNew.BBANPart(lArr(0), Pos, Trim(lArr(1)))
End Property
Public Property Get CountBBANParts() As Byte
    CountBBANParts = m_col.Count
'    Dim r As String: r = ToInfoR
'    CountBBANParts = UBound(Split(r, "; "))
End Property
Private Function GroupBy4(ByVal s_in As String) As String
    s_in = Replace(s_in, " ", "")
    Dim i As Long, s As String
    For i = 1 To Len(s_in)
        s = s & Mid(s_in, i, 1)
        If (i Mod 4) = 0 Then s = s & " "
    Next
    GroupBy4 = s
End Function
Public Function ToInfoR() As String
    ToInfoR = m_infoR
End Function
Private Function ParseInfoR() As String
    'jetzt umbauen, damit es leichter verarbeitet werden kann
    'zur Aktivierung der entspr. Eingabefelder und MaxLength Eintragung
    'Dim b, d, k, KK, r, s, X
    Dim char As String
    Dim c As Long
    Dim bi As String: bi = m_infoW & " " '1 dranhängen, zur Vereinfachung am Ende
    Dim sLine As String
    'c = 1
    Do Until c >= Len(bi)
        c = c + 1
        char = Mid(bi, c, 1)
        Select Case char 'Mid(bi, c, 1)
        Case "b", "d", "k", "K", "r", "s", "X"
            sLine = sLine & char & ": " & CountBlockChars(bi, c, char) & "; "
        'Case : sLine = sLine & char & ": " & CountBlockChars(bi, c, char) & "; "
        'Case "k": sLine = sLine & char & ": " & CountBlockChars(bi, c, char) & "; "
        'Case "K": sLine = sLine & char & ": " & CountBlockChars(bi, c, char) & "; "
        'Case "r": sLine = sLine & char & ": " & CountBlockChars(bi, c, char) & "; "
        'Case "s": sLine = sLine & char & ": " & CountBlockChars(bi, c, char) & "; "
        'Case "X": sLine = sLine & char & ": " & CountBlockChars(bi, c, char) & "; "
        End Select
    Loop
    ParseInfoR = sLine
End Function
Private Function CountBlockChars(sLine As String, pos_inout As Long, ByVal char As String) As Long
    'gibt die Anzahl von char im string sLine ab der Position pos_inout zurück
    'und liefert in pos_inout die aktuelle Position.
    Dim curchar As String
    Dim n As Long, l As Long
    l = Len(sLine)
    n = 1
    Do
        pos_inout = pos_inout + 1
        If pos_inout <= l Then
            curchar = Mid(sLine, pos_inout, 1)
            If curchar = char Then
                n = n + 1
            Else
                'einen zurückrollen
                pos_inout = pos_inout - 1
                Exit Do
            End If
        Else
            'einen zurückrollen
            pos_inout = pos_inout - 1
            Exit Do
        End If
    Loop
    CountBlockChars = n
End Function
'Private Function CInfoRToW() As String
'    Dim s As String
'    Dim sArr() As String: sArr = Split(Me.ToInfoR, "; ")
'    Dim i As Long
'    For i = 0 To UBound(sArr)
'        If Len(sArr(i)) Then
'            Dim lArr() As String: lArr = Split(sArr(i), ": ")
'            s = s & String(CLng(Trim(lArr(1))), Left(lArr(0), 1))
'        End If
'    Next
'    CInfoRToW = s
'End Function

'Public Function Check() As Boolean
'    Check = m_infoW = CInfoRToW
'End Function
