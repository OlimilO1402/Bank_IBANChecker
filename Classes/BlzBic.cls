VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BlzBic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public BLZ As String
Public PLZ As String
Public Ort As String
Public BanknameLok As String
Public BIC As String

Public Sub ParseLine(sLine As String)
    If Len(sLine) = 0 Then Exit Sub
    Dim lArr() As String: lArr = Split(sLine, vbTab)
    Dim j As Long, u As Long: u = UBound(lArr)
    BLZ = lArr(j): j = j + 1: If j > u Then Exit Sub
    PLZ = lArr(j): j = j + 1: If j > u Then Exit Sub
    Ort = lArr(j): j = j + 1: If j > u Then Exit Sub
    BanknameLok = lArr(j): j = j + 1: If j > u Then Exit Sub
    BIC = lArr(j): j = j + 1
End Sub

Friend Function ToStr() As String
    ToStr = BLZ & " " & BIC & " " & BanknameLok & " " & PLZ & " " & Ort
End Function
