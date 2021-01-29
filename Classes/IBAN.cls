VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Klasse liest eine IBAN ein und teilt anhand der IBANInfo in ihre Bestandteile auf
'hmm es muss auch eine Klasse BBAN geben!
'denn IBAN = LC+PZ+BBAN
Private m_IInfo As IBANInfo 'enthält Landname und LC
Private m_sIBAN As String 'der ursprünglich übergebene String, wird hier gespeichert zum Check (=Vergleich mit dem errechneten Wert)
Private m_BBAN  As BBAN

Friend Sub New_(IBANInfos As IBANInfos, ByVal siban As String)
    m_sIBAN = Replace(siban, " ", "")
    Dim lc As String: lc = Left(Trim(m_sIBAN), 2)
    Set m_IInfo = IBANInfos.ItemByLC(UCase(lc))
    If m_IInfo Is Nothing Then
        MsgBox "Error, could not locate countrycode: " & lc
        Exit Sub
    End If
    Set m_BBAN = MNew.BBAN(m_IInfo.BBANInfo, Mid(m_sIBAN, 5))
    Dim d As Integer: d = Len(m_sIBAN) - m_IInfo.IBANLength
    If d <> 0 Then
        MsgBox "IBAN ist um " & Abs(d) & " zu " & IIf(d > 0, "lang", "kurz")
    End If
End Sub
Public Sub NewX(aIBANInfo As IBANInfo)
    Set m_IInfo = aIBANInfo
End Sub
Public Property Get BBAN() As BBAN
    Set BBAN = m_BBAN
End Property

'Private Function CalcPZ(BBANwLCPZ0 As String) As Long
'für die Funktion CalcPZ siehe Modul MIBAN
'End Function
Public Property Get IBANInfo() As IBANInfo
    Set IBANInfo = m_IInfo
End Property
Public Function PZExtracted() As Long
    Dim siban As String: siban = Replace(m_sIBAN, " ", "")
    PZExtracted = Mid(siban, 3, 2)
End Function
Private Function IBANForPZCalc() As String
    'nimmt den vorhandenen IBAN-String heftet die darus ausgelesene LC hinten an
    Dim siban As String: siban = Replace(m_sIBAN, " ", "")
    IBANForPZCalc = DecodeAlphas(Mid(siban, 5) & Left(siban, 2) & "00")
End Function
Public Function PZCalculated() As Long
    PZCalculated = CalcPZ(IBANForPZCalc)
End Function
Public Function CheckPZ() As Boolean
    'untersucht ob die übergebene Prüfziffer gleich der berechneten ist
    'die übergebene und rausgelesene PZ
    CheckPZ = PZExtracted = PZCalculated
End Function
Public Function Check() As Boolean
    Check = (CheckPZ And BBAN.Check)
End Function
Friend Function ToStr() As String
    Dim s As String
    s = IBANInfo.CountryID & PadLeft0(CStr(PZCalculated), 2) & BBAN.ToStr
    ToStr = s
End Function
