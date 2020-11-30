VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BBANPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum EBBANPart
    Bankleitzahl    '"b" 'bank identifier
    Kontotyp        '"d" 'type of account
    Kontonummer     '"k" 'bank account number
    Kontrollziffer  '"K" 'control code
    Regionalcode    '"r" 'region code
    Filialnummer    '"s" 'branch identifier
    SonstFunktion   '"X" 'other functions
    Kontrollziffer2 '
End Enum

Private m_BBANProp As EBBANPart
Private m_Pos      As Byte
Private m_Length   As Byte

Public Sub New_(aChar As String, ByVal aPos As Byte, ByVal aLength As Byte)
    m_BBANProp = EBBANPart_Parse(aChar)
    m_Pos = aPos
    m_Length = aLength
End Sub

Private Function EBBANPart_Parse(ByVal s As String) As EBBANPart
    Dim e As EBBANPart
    If Len(s) > 1 Then s = Left(s, 1)
    Select Case s
    Case "b": e = Bankleitzahl    '"b" 'Bank Code
    Case "d": e = Kontotyp        '"d" '
    Case "k": e = Kontonummer     '"k" '
    Case "K": e = Kontrollziffer  '"K" '
    Case "r": e = Regionalcode    '"r" '
    Case "s": e = Filialnummer    '"s" 'Branch Code
    Case "X": e = SonstFunktion   '"X" '
    End Select
    EBBANPart_Parse = e
End Function

Private Function EBBANPart_ToKeyChar(e As EBBANPart) As String
    Dim s As String
    Select Case e
    Case Bankleitzahl:   s = "b"  '"b" 'Bank Code
    Case Kontotyp:       s = "d"  '"d" '
    Case Kontonummer:    s = "k"  '"k" '
    Case Kontrollziffer: s = "K"  '"K" '
    Case Regionalcode:   s = "r"  '"r" '
    Case Filialnummer:   s = "s"  '"s" 'Branch Code
    Case SonstFunktion:  s = "X"  '"X" '
    End Select
    EBBANPart_ToKeyChar = s
End Function
Private Function EBBANPart_ToStr(e As EBBANPart) As String
    Dim s As String
    Select Case e
    Case Bankleitzahl:   s = "Bankleitzahl"   '"b" 'Bank Code
    Case Kontotyp:       s = "Kontotyp"       '"d" '
    Case Kontonummer:    s = "Kontonummer"    '"k" '
    Case Kontrollziffer: s = "Kontrollziffer" '"K" '
    Case Regionalcode:   s = "Regionalcode"   '"r" '
    Case Filialnummer:   s = "Filialnummer"   '"s" 'Branch Code
    Case SonstFunktion:  s = "SonstFunktion"  '"X" '
    End Select
    EBBANPart_ToStr = s
End Function
Public Property Get EBBANPart() As EBBANPart
    EBBANPart = m_BBANProp
End Property
Public Property Get Key() As String
    Key = EBBANPart_ToKeyChar(m_BBANProp)
End Property

Public Property Get Name() As String
    Name = EBBANPart_ToStr(m_BBANProp)
End Property
Public Property Get Pos() As Byte
    Pos = m_Pos
End Property
Public Property Get Length() As Byte
    Length = m_Length
End Property

