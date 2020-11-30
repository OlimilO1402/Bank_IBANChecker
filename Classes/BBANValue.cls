VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BBANValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_BBpart As BBANPart
Private m_Value  As String

Friend Sub New_(bbp As BBANPart, ByVal Value As String)
    Set m_BBpart = bbp
    'führende Nullen entfernen
    'm_Value = Str(CDec(Value))
    m_Value = Value
End Sub

Public Property Get BBANPart() As BBANPart
    Set BBANPart = m_BBpart
End Property

Public Property Get Value() As String
    Value = RemoveLeading0(m_Value)
End Property

Public Function Check() As Boolean
    'hier kann man nur prüfen ob der Wert ungleich 0
    Check = Len(Value) > 0
End Function
Public Function ToStr() As String
    ToStr = PadLeft0(m_Value, m_BBpart.Length)
End Function
