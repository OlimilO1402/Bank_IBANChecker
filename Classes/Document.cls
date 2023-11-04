VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_PFN        As PathFileName
Private m_NamedIBANs As List 'Of NamedIBAN

Friend Sub New_(aPFN As String)
    Set m_NamedIBANs = MNew.List(vbObject, , True)
End Sub

Public Function Read(aPFN As String) As Boolean
    Set m_PFN = MNew.PathFileName(aPFN)
    If m_NamedIBANs Is Nothing Then
        Set m_NamedIBANs = MNew.List(EDataType.vbObject, , True)
    End If
Try: On Error GoTo Catch
    Dim lines() As String
    If Not m_PFN.TryReadAllLines(lines) Then Exit Function
    Dim i As Long, line As String, sa() As String
    Dim siban As String, sName As String, ani As NamedIBAN
    For i = 0 To UBound(lines)
        line = lines(i)
        sa = Split(line, vbTab)
        siban = sa(0)
        sName = sa(1)
        Set ani = MNew.NamedIBAN(sName, MNew.IBAN(MApp.IBANInfos, siban))
        m_NamedIBANs.Add ani, ani.Key
    Next
    GoTo Finally
Catch:
    
Finally:
    m_PFN.CloseFile
End Function

Public Property Get NamedIBANs() As List 'Of NamedIBAN
    Set NamedIBANs = m_NamedIBANs
End Property

Public Function Save() As Boolean
    
End Function

Public Function SaveAs() As Boolean
    
End Function
