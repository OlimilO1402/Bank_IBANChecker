VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BBAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_BInfo As BBANInfo
Private m_sBBAN As String

'wird von der Klasse IBAN selber erzeugt
Public Sub New_(BBANInfo As BBANInfo, ByVal sBBAN As String)
    m_sBBAN = sBBAN
    Set m_BInfo = BBANInfo
    If m_BInfo Is Nothing Then Exit Sub
    Dim d As Integer: d = Len(m_sBBAN) - m_BInfo.BBANLength
    If d <> 0 Then
        MsgBox "IBAN ist um " & Abs(d) & " zu " & IIf(d > 0, "lang", "kurz")
    End If
End Sub

'die Funktion DecodeAlpha hierher

Public Property Get Prop(ByVal Index As Byte) As BBANValue
    Dim bbp As BBANPart: Set bbp = m_BInfo.BBANPart(Index)
    Set Prop = MNew.BBANValue(bbp, Mid(m_sBBAN, bbp.Pos, bbp.Length))
End Property

Public Property Get CountParts() As Byte
    CountParts = m_BInfo.CountBBANParts
End Property

Public Function Check() As Boolean
    'hier kann man nur die gesamte Länge, die Länge der Bestandteile, und die Bestandteile <> 0 testen
    'naja das ist doch schon was
    Check = True
    Dim i As Byte
    For i = 0 To CountParts - 1
        Check = Check And Prop(i).Check
    Next
    If Not Check Then Exit Function
    Check = m_BInfo.BBANLength = Len(Me.ToStr)
End Function

Public Function ToStr() As String
    Dim s As String
    Dim i As Byte
    For i = 0 To CountParts - 1
        s = s & Prop(i).ToStr
    Next
    ToStr = s
End Function

