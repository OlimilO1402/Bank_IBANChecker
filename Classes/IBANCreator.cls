VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBANCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'erzeugt aus einem Array von Strings und der IBANInfo eine IBAN
Private m_IBANInfos As IBANInfos
Private m_IBANInfo  As IBANInfo
'Private m_sArr()    As String
Private m_list As Collection 'of String, String

Friend Sub New_(aIBANInfos As IBANInfos, aIBANInfo As IBANInfo, list As Collection) 'sArr() As String)
    'aus der IBANInfo
    'm_sArr = sArr
    Set m_list = list
    Set m_IBANInfos = aIBANInfos
    Set m_IBANInfo = aIBANInfo
End Sub
'Public Property Get BBAN() As BBAN
'    Set BBAN = MNew.BBAN(m_IBANInfo.BBANInfo, s)
'End Property
Public Property Get IBAN() As IBAN
    
    Dim siban As String
    siban = m_IBANInfo.CountryID & "00"
    Dim binfo As BBANInfo: Set binfo = m_IBANInfo.BBANInfo
    If binfo Is Nothing Then Exit Property
    Dim i As Long, w As Long
    'If UBound(m_sArr) + 1 <> binfo.CountBBANParts Then
    If m_list.Count <> binfo.CountBBANParts Then
        MsgBox "Halt"
    End If
    Dim bbp As BBANPart
    For i = 1 To binfo.CountBBANParts ' UBound(m_sArr)
        Set bbp = binfo.BBANPart(i)
        If Not bbp Is Nothing Then
            w = bbp.Length
            'siban = siban + PadLeft0(m_sArr(i), w)
            siban = siban + PadLeft0(bbp.Name, w)
        End If
    Next
    Set IBAN = MNew.IBAN(m_IBANInfos, siban)
End Property
