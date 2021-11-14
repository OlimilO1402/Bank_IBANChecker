VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BlzBics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_BlzBics() As BlzBic

Public Sub ParseFile(aFNm As String)
    Dim FNr As Integer: FNr = FreeFile
Try: On Error GoTo Finally
    Open aFNm For Binary As FNr
    Dim sFile As String: sFile = Space(LOF(FNr))
    Get FNr, , sFile
    Close FNr
    Dim sArr() As String: sArr = Split(sFile, vbCrLf)
    Dim u As Long: u = UBound(sArr)
    ReDim m_BlzBics(u)
    Dim i As Long, j As Long
    For i = 0 To u
        Dim bb As BlzBic: Set bb = MNew.BlzBic(sArr(i))
        Set m_BlzBics(i) = bb
    Next
    Exit Sub
Finally:
    Close FNr
End Sub

'Public Function ContainsBLZ(aBLZ As String) As Boolean
'    ContainsBLZ = Contains(m_BLZcol, aBLZ)
'End Function
'Public Function ContainsBIC(aBIC As String) As Boolean
'    ContainsBLZ = Contains(m_BICcol, aBIC)
'End Function

Public Function BLZcol(ByVal BLZ As String) As Collection
    'liefert zu einer BLZ eine Collection mit allen Elementen ähnlicher BLZ
    Set BLZcol = New Collection
    Dim v, bb As BlzBic
    For Each v In m_BlzBics()
        Set bb = v
        If Not bb Is Nothing Then
            If InStr(1, bb.BLZ, BLZ, vbTextCompare) Then BLZcol.Add bb
            If InStr(1, BLZ, bb.BLZ, vbTextCompare) Then BLZcol.Add bb
        End If
    Next
End Function
Public Function PLZcol(ByVal PLZ As String) As Collection
    'liefert zu einer PLZ eine Collection mit allen Elementen ähnlicher PLZ
    Set PLZcol = New Collection
    Dim v, bb As BlzBic
    For Each v In m_BlzBics()
        Set bb = v
        If InStr(1, bb.PLZ, PLZ, vbTextCompare) Then PLZcol.Add bb
        If InStr(1, PLZ, bb.PLZ, vbTextCompare) Then PLZcol.Add bb
    Next
End Function
Public Function ORTcol(ByVal Ort As String) As Collection
    'liefert zu einer PLZ eine Collection mit allen Elementen ähnlichen Ortes
    Set ORTcol = New Collection
    Dim v, bb As BlzBic
    For Each v In m_BlzBics()
        Set bb = v
        If InStr(1, bb.Ort, Ort, vbTextCompare) Then ORTcol.Add bb
        If InStr(1, Ort, bb.Ort, vbTextCompare) Then ORTcol.Add bb
    Next
End Function
Public Function BANKcol(ByVal BanknameLok As String) As Collection
    'liefert zu einem Banknamen eine Collection mit allen Elementen ähnlichen Banknamens
    Set BANKcol = New Collection
    Dim v, bb As BlzBic
    For Each v In m_BlzBics()
        Set bb = v
        If InStr(1, bb.BanknameLok, BanknameLok, vbTextCompare) Then BANKcol.Add bb
        If InStr(1, BanknameLok, bb.BanknameLok, vbTextCompare) Then BANKcol.Add bb
    Next
End Function
Public Function BICcol(ByVal BIC As String) As Collection
    'liefert zu einer BIC eine Collection mit allen Elementen ähnlicher BIC
    Set BICcol = New Collection
    Dim v, bb As BlzBic
    For Each v In m_BlzBics()
        Set bb = v
        If InStr(1, bb.BIC, BIC, vbTextCompare) Then BICcol.Add bb
        If InStr(1, BIC, bb.BIC, vbTextCompare) Then BICcol.Add bb
    Next
End Function


