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

Friend Sub New_(ByVal aPFN As String)
    If Not ReadFromFile(aPFN) Then
        Init
    End If
End Sub

Public Function ReadFromFile(aPFN As String) As Boolean
    Set m_PFN = MNew.PathFileName(aPFN)
    If Not m_PFN.Exists Then Exit Function
    'm_PFN = aPFN
    'Dim FNr As Integer: FNr = FreeFile
Try: On Error GoTo Finally
    'Open m_PFN For Binary As FNr
    'Dim File As String: File = Space(LOF(FNr))
    'Get FNr, , File
    'Close FNr
    'Dim sArr() As String: sArr = Split(File, vbCrLf)
    Dim lines() As String
    If Not m_PFN.TryReadAllLines(lines) Then Exit Function
    ReDim m_col(0 To UBound(lines))
    Dim i As Long, c As Long, sLine As String
    'Dim ii As IBANINfo
    Dim lArr() As String
    Dim line As String
    Dim LN As String, lc As String, bi As String
    For i = 0 To UBound(lines)
        'Set ii = New IBANINfo
        line = lines(i)
        If Len(line) Then
            lArr = Split(line, vbTab)
            LN = Trim(lArr(0)) 'Ländername
            sLine = Trim(lArr(1))
            lc = Left(sLine, 2) 'Ländercode
            bi = Trim(Mid(sLine, 6))
            'bi = Replace(bi, " ", "")
            'ii.New_ ln, lc, bi
            Set m_col(i) = MNew.IBANInfo(LN, lc, MNew.BBANInfo(bi))
        End If
    Next
    ReadFromFile = True
    Exit Function
Finally:
    'Close FNr
    m_PFN.CloseFile
End Function

'Public Sub WriteToFile(aPFN As String)
'    Dim s As String
'    Dim i As Long, ibi As IBANInfo
'    For i = 0 To UBound(m_col)
'        Set ibi = m_col(i)
'        If Not ibi Is Nothing Then
'            s = s & "    Set m_col(i) = MNew.IBANInfo(" & ibi.ToFile & "): i = i + 1" & vbCrLf
'        End If
'    Next
'    Dim pfn As PathFileName: Set pfn = MNew.PathFileName(aPFN)
'    If Not pfn.WriteStr(s) Then
'        MsgBox "Error"
'    End If
'End Sub

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

Public Sub Init()
    Dim i As Long
    Set m_col(i) = MNew.IBANInfo("Ägypten", "EG", MNew.BBANInfo("kkkk kkkk kkkk kkkk kkkk kkk")):                                   i = i + 1
    Set m_col(i) = MNew.IBANInfo("Albanien", "AL", MNew.BBANInfo("bbbs sssK kkkk kkkk kkkk kkkk")):                                 i = i + 1
    Set m_col(i) = MNew.IBANInfo("Algerien", "DZ", MNew.BBANInfo("kkkk kkkk kkkk kkkk kkkk")):                                      i = i + 1
    Set m_col(i) = MNew.IBANInfo("Andorra", "AD", MNew.BBANInfo("bbbb ssss kkkk kkkk kkkk")):                                       i = i + 1
    Set m_col(i) = MNew.IBANInfo("Angola", "AO", MNew.BBANInfo("bbbb ssss kkkk kkkk kkkK K")):                                      i = i + 1
    Set m_col(i) = MNew.IBANInfo("Aserbaidschan", "AZ", MNew.BBANInfo("bbbb kkkk kkkk kkkk kkkk kkkk")):                            i = i + 1
    Set m_col(i) = MNew.IBANInfo("Bahrain", "BH", MNew.BBANInfo("bbbb kkkk kkkk kkkk kk")):                                         i = i + 1
    Set m_col(i) = MNew.IBANInfo("Belgien", "BE", MNew.BBANInfo("bbbk kkkk kkKK")):                                                 i = i + 1
    Set m_col(i) = MNew.IBANInfo("Benin", "BJ", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kkKK")):                                    i = i + 1
    Set m_col(i) = MNew.IBANInfo("Bosnien und Herzegowina", "BA", MNew.BBANInfo("bbbs sskk kkkk kkKK")):                            i = i + 1
    Set m_col(i) = MNew.IBANInfo("Brasilien", "BR", MNew.BBANInfo("bbbb bbbb ssss skkk kkkk kkkk k")):                              i = i + 1
    Set m_col(i) = MNew.IBANInfo("Britische Jungferninseln", "VG", MNew.BBANInfo("bbbb kkkk kkkk kkkk kkkk")):                      i = i + 1
    Set m_col(i) = MNew.IBANInfo("Bulgarien", "BG", MNew.BBANInfo("bbbb ssss ddkk kkkk kk")):                                       i = i + 1
    Set m_col(i) = MNew.IBANInfo("Burkina Faso", "BF", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):                              i = i + 1
    Set m_col(i) = MNew.IBANInfo("Burundi", "BI", MNew.BBANInfo("kkkk kkkk kkkk")):                                                 i = i + 1
    Set m_col(i) = MNew.IBANInfo("Costa Rica", "CR", MNew.BBANInfo("bbbk kkkk kkkk kkkk k")):                                       i = i + 1
    Set m_col(i) = MNew.IBANInfo("Côte d'Ivoire (Elfenbeinküste)", "CI", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kkKK")):           i = i + 1
    Set m_col(i) = MNew.IBANInfo("Dänemark", "DK", MNew.BBANInfo("bbbb kkkk kkkk kK")):                                             i = i + 1
    Set m_col(i) = MNew.IBANInfo("Deutschland", "DE", MNew.BBANInfo("bbbb bbbb kkkk kkkk kk")):                                     i = i + 1
    Set m_col(i) = MNew.IBANInfo("Dominikanische Republik", "DO", MNew.BBANInfo("bbbb kkkk kkkk kkkk kkkk kkkk")):                  i = i + 1
    Set m_col(i) = MNew.IBANInfo("Estland", "EE", MNew.BBANInfo("bbkk kkkk kkkk kkkK")):                                            i = i + 1
    Set m_col(i) = MNew.IBANInfo("Färöer", "FO", MNew.BBANInfo("bbbb kkkk kkkk kK")):                                               i = i + 1
    Set m_col(i) = MNew.IBANInfo("Finnland", "FI", MNew.BBANInfo("bbbb bbkk kkkk kK")):                                             i = i + 1
    Set m_col(i) = MNew.IBANInfo("Frankreich", "FR", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):                                i = i + 1
    Set m_col(i) = MNew.IBANInfo("Französisch-Guayana", "FR", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):                       i = i + 1
    Set m_col(i) = MNew.IBANInfo("Französisch-Polynesien", "FR", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):                    i = i + 1
    Set m_col(i) = MNew.IBANInfo("Französische Süd- und Antarktisgebiete", "FR", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):    i = i + 1
    Set m_col(i) = MNew.IBANInfo("Guadeloupe", "FR", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):                                i = i + 1
    Set m_col(i) = MNew.IBANInfo("Martinique", "FR", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):                                i = i + 1
    Set m_col(i) = MNew.IBANInfo("Réunion", "FR", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):                                   i = i + 1
    Set m_col(i) = MNew.IBANInfo("Mayotte", "FR", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):                                   i = i + 1
    Set m_col(i) = MNew.IBANInfo("Neukaledonien", "FR", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):                             i = i + 1
    Set m_col(i) = MNew.IBANInfo("Saint-Barthélemy", "FR", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):                          i = i + 1
    Set m_col(i) = MNew.IBANInfo("Saint-Martin", "FR", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):                              i = i + 1
    Set m_col(i) = MNew.IBANInfo("Saint-Pierre und Miquelon", "FR", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):                 i = i + 1
    Set m_col(i) = MNew.IBANInfo("Wallis und Futuna", "FR", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):                         i = i + 1
    Set m_col(i) = MNew.IBANInfo("Gabun", "GA", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):                                     i = i + 1
    Set m_col(i) = MNew.IBANInfo("Georgien", "GE", MNew.BBANInfo("bbkk kkkk kkkk kkkk kk")):                                        i = i + 1
    Set m_col(i) = MNew.IBANInfo("Gibraltar", "GI", MNew.BBANInfo("bbbb kkkk kkkk kkkk kkk")):                                      i = i + 1
    Set m_col(i) = MNew.IBANInfo("Griechenland", "GR", MNew.BBANInfo("bbbs sssk kkkk kkkk kkkk kkk")):                              i = i + 1
    Set m_col(i) = MNew.IBANInfo("Grönland", "GL", MNew.BBANInfo("bbbb kkkk kkkk kK")):                                             i = i + 1
    Set m_col(i) = MNew.IBANInfo("Guatemala", "GT", MNew.BBANInfo("bbbb kkkk kkkk kkkk kkkk kkkk")):                                i = i + 1
    Set m_col(i) = MNew.IBANInfo("Iran", "IR", MNew.BBANInfo("kkkk kkkk kkkk kkkk kkkk kk")):                                       i = i + 1
    Set m_col(i) = MNew.IBANInfo("Irland", "IE", MNew.BBANInfo("bbbb ssss sskk kkkk kk")):                                          i = i + 1
    Set m_col(i) = MNew.IBANInfo("Island", "IS", MNew.BBANInfo("bbbb sskk kkkk XXXX XXXX XX")):                                     i = i + 1
    Set m_col(i) = MNew.IBANInfo("Israel", "IL", MNew.BBANInfo("bbbs sskk kkkk kkkk kkk")):                                         i = i + 1
    Set m_col(i) = MNew.IBANInfo("Italien", "IT", MNew.BBANInfo("Kbbb bbss sssk kkkk kkkk kkk")):                                   i = i + 1
    Set m_col(i) = MNew.IBANInfo("Jordanien", "JO", MNew.BBANInfo("bbbb ssss kkkk kkkk kkkk kkkk kk")):                             i = i + 1
    Set m_col(i) = MNew.IBANInfo("Kamerun", "CM", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):                                   i = i + 1
    Set m_col(i) = MNew.IBANInfo("Kap Verde", "CV", MNew.BBANInfo("bbbb ssss kkkk kkkk kkkK K")):                                   i = i + 1
    Set m_col(i) = MNew.IBANInfo("Kasachstan", "KZ", MNew.BBANInfo("bbbk kkkk kkkk kkkk")):                                         i = i + 1
    Set m_col(i) = MNew.IBANInfo("Katar", "QA", MNew.BBANInfo("bbbb kkkk kkkk kkkk kkkk kkkk k")):                                  i = i + 1
    Set m_col(i) = MNew.IBANInfo("Kongo (Brazzaville)", "CG", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):                       i = i + 1
    Set m_col(i) = MNew.IBANInfo("Kosovo", "XK", MNew.BBANInfo("bbbb kkkk kkkk kkkk")):                                             i = i + 1
    Set m_col(i) = MNew.IBANInfo("Kroatien", "HR", MNew.BBANInfo("bbbb bbbk kkkk kkkk k")):                                         i = i + 1
    Set m_col(i) = MNew.IBANInfo("Kuwait", "KW", MNew.BBANInfo("bbbb kkkk kkkk kkkk kkkk kkkk kk")):                                i = i + 1
    Set m_col(i) = MNew.IBANInfo("Lettland", "LV", MNew.BBANInfo("bbbb kkkk kkkk kkkk k")):                                         i = i + 1
    Set m_col(i) = MNew.IBANInfo("Libanon", "LB", MNew.BBANInfo("bbbb kkkk kkkk kkkk kkkk kkkk")):                                  i = i + 1
    Set m_col(i) = MNew.IBANInfo("Liechtenstein", "LI", MNew.BBANInfo("bbbb bkkk kkkk kkkk k")):                                    i = i + 1
    Set m_col(i) = MNew.IBANInfo("Litauen", "LT", MNew.BBANInfo("bbbb bkkk kkkk kkkk")):                                            i = i + 1
    Set m_col(i) = MNew.IBANInfo("Luxemburg", "LU", MNew.BBANInfo("bbbk kkkk kkkk kkkk")):                                          i = i + 1
    Set m_col(i) = MNew.IBANInfo("Madagaskar", "MG", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):                                i = i + 1
    Set m_col(i) = MNew.IBANInfo("Mali", "ML", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kkKK")):                                     i = i + 1
    Set m_col(i) = MNew.IBANInfo("Malta", "MT", MNew.BBANInfo("bbbb ssss skkk kkkk kkkk kkkk kkk")):                                i = i + 1
    Set m_col(i) = MNew.IBANInfo("Mauretanien", "MR", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):                               i = i + 1
    Set m_col(i) = MNew.IBANInfo("Mauritius", "MU", MNew.BBANInfo("bbbb bbss kkkk kkkk kkkk kkkK KK")):                             i = i + 1
    Set m_col(i) = MNew.IBANInfo("Mazedonien", "MK", MNew.BBANInfo("bbbk kkkk kkkk kKK")):                                          i = i + 1
    Set m_col(i) = MNew.IBANInfo("Moldawien", "MD", MNew.BBANInfo("bbkk kkkk kkkk kkkk kkkk")):                                     i = i + 1
    Set m_col(i) = MNew.IBANInfo("Monaco", "MC", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):                                    i = i + 1
    Set m_col(i) = MNew.IBANInfo("Montenegro", "ME", MNew.BBANInfo("bbbk kkkk kkkk kkkk KK")):                                      i = i + 1
    Set m_col(i) = MNew.IBANInfo("Mosambik", "MZ", MNew.BBANInfo("bbbb ssss kkkk kkkk kkkK K")):                                    i = i + 1
    Set m_col(i) = MNew.IBANInfo("Niederlande", "NL", MNew.BBANInfo("bbbb kkkk kkkk kk")):                                          i = i + 1
    Set m_col(i) = MNew.IBANInfo("Norwegen", "NO", MNew.BBANInfo("bbbb kkkk kkK")):                                                 i = i + 1
    Set m_col(i) = MNew.IBANInfo("Österreich", "AT", MNew.BBANInfo("bbbb bkkk kkkk kkkk")):                                         i = i + 1
    Set m_col(i) = MNew.IBANInfo("Osttimor", "TL", MNew.BBANInfo("bbbk kkkk kkkk kkkk kKK")):                                       i = i + 1
    Set m_col(i) = MNew.IBANInfo("Pakistan", "PK", MNew.BBANInfo("bbbb rrkk kkkk kkkk kkkk")):                                      i = i + 1
    Set m_col(i) = MNew.IBANInfo("Palästinensische Autonomiegebiete", "PS", MNew.BBANInfo("bbbb rrrr rrrr rkkk kkkk kkkk k")):      i = i + 1
    Set m_col(i) = MNew.IBANInfo("Polen", "PL", MNew.BBANInfo("bbbs sssK kkkk kkkk kkkk kkkk")):                                    i = i + 1
    Set m_col(i) = MNew.IBANInfo("Portugal", "PT", MNew.BBANInfo("bbbb ssss kkkk kkkk kkkK K")):                                    i = i + 1
    Set m_col(i) = MNew.IBANInfo("Rumänien", "RO", MNew.BBANInfo("bbbb kkkk kkkk kkkk kkkk")):                                      i = i + 1
    Set m_col(i) = MNew.IBANInfo("San Marino", "SM", MNew.BBANInfo("Kbbb bbss sssk kkkk kkkk kkk")):                                i = i + 1
    Set m_col(i) = MNew.IBANInfo("São Tomé und Príncipe", "ST", MNew.BBANInfo("bbbb ssss kkkk kkkk kkkK K")):                       i = i + 1
    Set m_col(i) = MNew.IBANInfo("Saudi-Arabien", "SA", MNew.BBANInfo("bbkk kkkk kkkk kkkk kkkk")):                                 i = i + 1
    Set m_col(i) = MNew.IBANInfo("Schweden", "SE", MNew.BBANInfo("bbbk kkkk kkkk kkkk kkkK")):                                      i = i + 1
    Set m_col(i) = MNew.IBANInfo("Schweiz", "CH", MNew.BBANInfo("bbbb bkkk kkkk kkkk k")):                                          i = i + 1
    Set m_col(i) = MNew.IBANInfo("Senegal", "SN", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kkKK")):                                  i = i + 1
    Set m_col(i) = MNew.IBANInfo("Serbien", "RS", MNew.BBANInfo("bbbk kkkk kkkk kkkk KK")):                                         i = i + 1
    Set m_col(i) = MNew.IBANInfo("Slowakei", "SK", MNew.BBANInfo("bbbb ssss sskk kkkk kkkk")):                                      i = i + 1
    Set m_col(i) = MNew.IBANInfo("Slowenien", "SI", MNew.BBANInfo("bbss skkk kkkk kKK")):                                           i = i + 1
    Set m_col(i) = MNew.IBANInfo("Spanien", "ES", MNew.BBANInfo("bbbb ssss KKkk kkkk kkkk")):                                       i = i + 1
    Set m_col(i) = MNew.IBANInfo("Tschechien", "CZ", MNew.BBANInfo("bbbb kkkk kkkk kkkk kkkk")):                                    i = i + 1
    Set m_col(i) = MNew.IBANInfo("Tunesien", "TN", MNew.BBANInfo("bbss skkk kkkk kkkk kkKK")):                                      i = i + 1
    Set m_col(i) = MNew.IBANInfo("Türkei", "TR", MNew.BBANInfo("bbbb brkk kkkk kkkk kkkk kk")):                                     i = i + 1
    Set m_col(i) = MNew.IBANInfo("Ungarn", "HU", MNew.BBANInfo("bbbs sssK kkkk kkkk kkkk kkkK")):                                   i = i + 1
    Set m_col(i) = MNew.IBANInfo("Vereinigte Arabische Emirate", "AE", MNew.BBANInfo("bbbk kkkk kkkk kkkk kkk")):                   i = i + 1
    Set m_col(i) = MNew.IBANInfo("Vereinigtes Königreich", "GB", MNew.BBANInfo("bbbb ssss sskk kkkk kk")):                          i = i + 1
    Set m_col(i) = MNew.IBANInfo("Jersey", "GB", MNew.BBANInfo("bbbb ssss sskk kkkk kk")):                                          i = i + 1
    Set m_col(i) = MNew.IBANInfo("Guernsey", "GB", MNew.BBANInfo("bbbb ssss sskk kkkk kk")):                                        i = i + 1
    Set m_col(i) = MNew.IBANInfo("Isle of Man", "GB", MNew.BBANInfo("bbbb ssss sskk kkkk kk")):                                     i = i + 1
    Set m_col(i) = MNew.IBANInfo("Zypern", "CY", MNew.BBANInfo("bbbs ssss kkkk kkkk kkkk kkkk")):                                   i = i + 1
    Set m_col(i) = MNew.IBANInfo("Zentralafrikanische Republik", "CF", MNew.BBANInfo("bbbb bsss sskk kkkk kkkk kKK")):              i = i + 1
End Sub
