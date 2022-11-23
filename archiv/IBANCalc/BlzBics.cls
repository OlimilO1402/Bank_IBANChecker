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
'Private m_BLZcol As New Collection
'Private m_BICcol As New Collection
Private m_BlzBics() As BlzBic

Public Sub ParseFile(aFNm As String)
    Dim FNr As Integer: FNr = FreeFile
    'Dim FNm As String:  FNm = App.Path & "\" & "blzBIC3_2015.txt"
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
    Dim bb 'As BlzBic
    For Each bb In m_BlzBics()
        If InStr(1, bb.BLZ, BLZ) Then BLZcol.Add bb
    Next
End Function
Public Function PLZcol(ByVal PLZ As String) As Collection
    'liefert zu einer PLZ eine Collection mit allen Elementen ähnlicher PLZ
    Set PLZcol = New Collection
    Dim bb 'As BlzBic
    For Each bb In m_BlzBics()
        If InStr(1, bb.PLZ, PLZ) Then PLZcol.Add bb
    Next
End Function
Public Function ORTcol(ByVal Ort As String) As Collection
    'liefert zu einer PLZ eine Collection mit allen Elementen ähnlichen Ortes
    Set ORTcol = New Collection
    Dim bb 'As BlzBic
    For Each bb In m_BlzBics()
        If InStr(1, bb.Ort, Ort) Then ORTcol.Add bb
    Next
End Function
Public Function BANKcol(ByVal BanknameLok As String) As Collection
    'liefert zu einem Banknamen eine Collection mit allen Elementen ähnlichen Banknamens
    Set BANKcol = New Collection
    Dim bb 'As BlzBic
    For Each bb In m_BlzBics()
        If InStr(1, bb.BanknameLok, BanknameLok) Then BANKcol.Add bb
    Next
End Function
Public Function BICcol(ByVal BIC As String) As Collection
    'liefert zu einer BIC eine Collection mit allen Elementen ähnlicher BIC
    Set BICcol = New Collection
    Dim bb 'As BlzBic
    For Each bb In m_BlzBics()
        If InStr(1, bb.BIC, BIC) Then BICcol.Add bb
    Next
End Function


