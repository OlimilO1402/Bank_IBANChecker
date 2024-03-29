VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NamedIBAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Name As String ' Persons Name
Private m_IBAN As IBAN   ' IBAN from Person

Friend Sub New_(aName As String, aIBAN As IBAN)
    m_Name = aName: Set m_IBAN = aIBAN
End Sub

Public Property Get Name() As String
    Name = m_Name
End Property
Public Property Let Name(ByVal Value As String)
    m_Name = Value
End Property

Public Property Get IBAN() As IBAN
    Set IBAN = m_IBAN
End Property

Public Property Get Key() As String
    Key = IBAN.ToStr
End Property
