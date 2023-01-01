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
Public Name As String ' Persons Name
Public IBAN As IBAN   ' IBAN from Person

Friend Sub New_(aName As String, aIBAN As IBAN)
    Name = aName: Set IBAN = aIBAN
End Sub

Public Property Get Key() As String
    Key = IBAN.ToStr
End Property
