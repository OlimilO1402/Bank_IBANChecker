VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBANInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_PFN   As PathFileName
Private m_col() As IBANInfo ' As Collection 'of IBANInfo

Public Sub ReadFromFile(aPFN As String)
    Set m_PFN = MNew.PathFileName(aPFN)
    'm_PFN = aPFN
    'Dim FNr As Integer: FNr = FreeFile
Try: On Error GoTo Finally
    'Open m_PFN For Binary As FNr
    'Dim File As String: File = Space(LOF(FNr))
    'Get FNr, , File
    'Close FNr
    'Dim sArr() As String: sArr = Split(File, vbCrLf)
    Dim lines() As String
    If Not m_PFN.TryReadAllLines(lines) Then Exit Sub
    ReDim m_col(0 To UBound(lines))
    Dim i As Long, c As Long, sLine As String
    'Dim ii As IBANINfo
    Dim lArr() As String
    Dim ln As String, lc As String, bi As String
    For i = 0 To UBound(lines)
        'Set ii = New IBANINfo
        lArr = Split(lines(i), vbTab)
        ln = Trim(lArr(0)) 'Ländername
        sLine = Trim(lArr(1))
        lc = Left(sLine, 2) 'Ländercode
        bi = Trim(Mid(sLine, 6))
        'bi = Replace(bi, " ", "")
        'ii.New_ ln, lc, bi
        Set m_col(i) = MNew.IBANInfo(ln, lc, MNew.BBANInfo(bi))
    Next
    Exit Sub
Finally:
    'Close FNr
    m_PFN.CloseFile
End Sub

Public Sub FillComboBox(aCB As ComboBox)
    Dim ii 'As IBANINfo
    For Each ii In m_col
        If Not ii Is Nothing Then aCB.AddItem ii.Key
    Next
End Sub

Public Property Get Item(ByVal Index As Long) As IBANInfo
    Set Item = m_col(Index)
End Property

Public Property Get Index(ByVal lc As String) As Long
    Dim i As Long
    For i = 0 To UBound(m_col)
        If m_col(i).CountryID = lc Then Index = i: Exit Property
    Next
End Property
Public Property Get ItemByLC(ByVal CountryID As String) As IBANInfo
    Dim i
    For Each i In m_col
        If Not i Is Nothing Then
            Set ItemByLC = i
            If ItemByLC.CountryID = CountryID Then
                Exit Property
            End If
        End If
    Next
End Property

Public Function Check() As Boolean
    'überprüft ob die Wiki-Zeile und die gelesene geparste und wieder zurückgewandelte Information gleich sind
    Dim ii 'As BBANINfo
    For Each ii In m_col
        If Not ii Is Nothing Then
            Check = ii.BBANInfo.Check
            If Not Check Then
                Debug.Print ii.Key
                Debug.Print ii.BBANInfo.ToInfoR
                Debug.Print ii.BBANInfo.ToInfoW
                Exit Function
            End If
        End If
    Next
End Function
