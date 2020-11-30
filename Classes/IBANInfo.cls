VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBANInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Name  As String 'CountryName, Ländername Ungarn, Frankreich, etc.
Private m_LC    As String 'CountryID,   Ländercode DE, GB, FR etc.
Private m_bbinf As BBANInfo

Public Sub New_(Name As String, lc As String, bbinf As BBANInfo)
    m_Name = Name: m_LC = lc
    Set m_bbinf = bbinf
End Sub
Public Property Get IBANLength() As Byte
    IBANLength = m_bbinf.BBANLength + 4
End Property

Public Function ToStr(Optional br As Boolean = False) As String
    ToStr = m_LC & " " & m_bbinf.ToInfoW
    If br Then ToStr = m_LC & " " & m_bbinf.ToInfoR
End Function

Public Property Get CountryName() As String
    CountryName = m_Name
End Property

Public Property Get CountryID() As String
    CountryID = m_LC
End Property

Public Property Get Key() As String
    Key = m_Name & ", " & m_LC
End Property

Public Property Get BBANInfo() As BBANInfo
    Set BBANInfo = m_bbinf
End Property

'Public Sub FillComboBox(aCB As ComboBox)
'    With aCB
'        .AddItem Key
'    End With
'End Sub
