Attribute VB_Name = "MIBANUtil"
Option Explicit
'IBAN = International Bank Account Number
'BIC = Business Identifier Code
'SWIFT = Society for Worldwide Interbank Financial Telecommunications
'
'der BIC hat 8 oder 11 Stellen und ist wie folgt aufgebaut:
'
'* bank code
'  4 Stellen Alphazeichen frei gewählt (Bundesbank z.B. MARK)
'
'* country code
'  2 Stellen Alphazeichen, ISO-Code des Landes (in Deutschland also DE)
'
'* location code
'  2 Stellen alphanumerisch zur Ortsangabe (z.B. FF für Frankfurt)
'
'* branch code
'  wahlweise 3 Stellen alphanumerisch zur Bezeichnung von Filialen
'
'BYLA DE M1 FSI
'Public Enum IBANLänderCode
'    Andorra = 1013          '   A D
'    Austria = 1029          '   A T
'    Belgien = 1114          '   B E
'    'Bulgarien = 0           '   0 0
'    Dänemark = 1320         '   D K
'    Deutschland = 1314      '   D E
'    Estland = 1414          '   E E
'    Finnland = 1518         '   F I
'    Frankreich = 1527       '   F R
'    Gibraltar = 1618        '   G I
'    Griechenland = 1627     '   G R
'    Großbritannien = 1611   '   G B
'    Irland = 1814           '   I E
'    Island = 1828           '   I S
'    Italien = 1829          '   I T
'    Lettland = 2131         '   L V
'    Litauen = 2129          '   L T
'    Luxemburg = 2130        '   L U
'    'Malta = 0               '   0 0
'    Niederlande = 2321      '   N L
'    Norwegen = 2324         '   N O
'    Polen = 2521            '   P L
'    Portugal = 2529         '   P T
'    'Rumänien = 0            '   0 0
'    Schweden = 2814         '   S E
'    Schweiz = 1217          '   C H
'    Slowakei = 2820         '   S K
'    Slowenien = 2818        '   S I
'    Spanien = 1428          '   E S
'    Tschechien = 1235       '   C Z
'    Ungarn = 1730           '   H U
'    Zypern = 1234           '   C Y
'End Enum
'Const IBANTbl1$ = "A: 1; D: 1; p: 1; p: 1; BLZ: 4; Bereich: 4; Kontonummer: 12; : 10" & vbCrLf & _
'                  "A: 1; T: 1; p: 1; p: 1; BLZ: 5; Kontonummer: 11; : 14" & vbCrLf & _
'                  "B: 1; E: 1; p: 1; p: 1; BLZ: 3; Kontonummer: 7; p: 1; p: 1; : 18" & vbCrLf & _
'                  "D: 1; K: 1; p: 1; p: 1; BLZ: 4; Kontonummer: 9; p: 1; : 16" & vbCrLf & _
'                  "D: 1; E: 1; p: 1; p: 1; BLZ: 8; Kontonummer: 10; : 12" & vbCrLf & _
'                  "E: 1; E: 1; p: 1; p: 1; BLZ: 2; Bereich: 2; Kontonummer: 11; p: 1; : 14" & vbCrLf & _
'                  "F: 1; I: 1; p: 1; p: 1; BLZ: 6; Kontonummer: 7; p: 1; : 16" & vbCrLf & _
'                  "F: 1; R: 1; p: 1; p: 1; BLZ: 5; Bereich: 5; Kontonummer: 11; p: 1; p: 1; : 7" & vbCrLf & _
'                  "G: 1; I: 1; p: 1; p: 1; BLZ: 4; Kontonummer: 15; : 11" & vbCrLf & _
'                  "G: 1; R: 1; p: 1; p: 1; BLZ: 3; Bereich: 4; Kontonummer: 16; : 7" & vbCrLf & _
'                  "G: 1; B: 1; p: 1; p: 1; BLZ: 4; Bereich: 6; Kontonummer: 8; : 12" & vbCrLf & _
'                  "I: 1; E: 1; p: 1; p: 1; BLZ: 4; Bereich: 6; Kontonummer: 8; : 12" & vbCrLf & _
'                  "I: 1; S: 1; p: 1; p: 1; BLZ: 4; Typ: 2; Kontonummer: 6; Identifikationsnr.: 10; : 8" & vbCrLf & _
'                  "I: 1; T: 1; p: 1; p: 1; p: 1; BLZ: 5; Bereich: 5; Kontonummer: 12; : 7" & vbCrLf & _
'                  "L: 1; V: 1; p: 1; p: 1; BLZ: 4; Kontonummer: 13; : 13" & vbCrLf & _
'                  "L: 1; T: 1; p: 1; p: 1; BLZ: 5; Kontonummer: 11; : 14"
'
'Const IBANTbl2$ = "L: 1; U: 1; p: 1; p: 1; BLZ: 3; Kontonummer: 13; : 14" & vbCrLf & _
'                  "N: 1; L: 1; p: 1; p: 1; BLZ: 4; Kontonummer: 10; : 16" & vbCrLf & _
'                  "N: 1; O: 1; p: 1; p: 1; BLZ: 4; Kontonummer: 6; p: 1; : 19" & vbCrLf & _
'                  "P: 1; L: 1; p: 1; p: 1; BLZ: 8; Kontonummer: 16; : 6" & vbCrLf & _
'                  "P: 1; T: 1; p: 1; p: 1; BLZ: 4; Bereich: 4; Kontonummer: 11; p: 1; p: 1; : 9" & vbCrLf & _
'                  "S: 1; E: 1; p: 1; p: 1; BLZ: 3; Kontonummer: 16; p: 1; : 10" & vbCrLf & _
'                  "C: 1; H: 1; p: 1; p: 1; BLZ: 5; Kontonummer: 12; : 11; : 1; : 1" & vbCrLf & _
'                  "S: 1; K: 1; p: 1; p: 1; BLZ: 4; Kto.nr. 1.Teil: 6; Kto.nr. 2.Teil: 10; : 10" & vbCrLf & _
'                  "S: 1; I: 1; p: 1; p: 1; BLZ: 5; Kontonummer: 8; p: 1; p: 1; : 15" & vbCrLf & _
'                  "E: 1; S: 1; p: 1; p: 1; BLZ: 4; Bereich: 4; p: 1; p: 1; Kontonummer: 10; : 10" & vbCrLf & _
'                  "C: 1; Z: 1; p: 1; p: 1; BLZ: 4; Kto. 1. Teil: 6; Kto. 2. Teil: 10; : 10" & vbCrLf & _
'                  "H: 1; U: 1; p: 1; p: 1; BLZ: 3; Bereich: 4; p: 1; Kontonummer: 15; p: 1; : 5; : 1" & vbCrLf & _
'                  "C: 1; Y: 1; p: 1; p: 1; BLZ: 3; Bereich: 5; Kontonummer: 16; : 6"
'Const IBANTbl$ = IBANTbl1 & IBANTbl2 'in zwei aufteilen, weil sonst Grenze von Zeilenfortsetzung erreicht wird
'Private Type TIBANInfo
'    CName As String
'    CCode As Long
'    BInfo As String
'End Type
'Private m_IBANInfo() As String 'verknüpft über den Index

'Ägypten   EGpp kkkk kkkk kkkk kkkk kkkk kkk
'Albanien  ALpp bbbs sssK kkkk kkkk kkkk kkkk
'Algerien  DZpp kkkk kkkk kkkk kkkk kkkk
'Andorra   ADpp bbbb ssss kkkk kkkk kkkk
'Angola    AOpp bbbb ssss kkkk kkkk kkkK K

'AD, BE, ...     Länderkennzeichen   Country Code
'
'b   Bankleitzahl    Bank Code
'd   Kontotyp
'k   Kontonummer
'K   Kontrollziffer
'r   Regionalcode
's   Filialnummer    Branch Code
'X   sonst. Funkt.
'Function CalcIBAN(IBANInfo As String, codes() As String) As String
'    Dim sArr() As String: sArr = Split(IBANInfo, "; ")
'    Dim lc: lc = Trim(sArr(1))
'    Dim iLC: iLC = (Asc(Left(lc, 1)) - 55) * 100 + Asc(Right(lc, 1)) - 55
'    Dim BBAN
'    Dim pc As Long, elms
'    Dim i As Long, j As Long, l
'    For i = 2 To UBound(sArr)
'        If Len(sArr(i)) > 0 Then
'            elms = Split(sArr(i), ": ")
'            l = elms(1)
'            BBAN = BBAN & PadLeft0(codes(pc), l)
'            pc = pc + 1
'        End If
'    Next
'    Dim PZ: PZ = "00": PZ = CalcPZ(BBAN & iLC & PZ)
'    CalcIBAN = lc & PZ & BBAN
'End Function
'Function CalcPZ(ByVal IBANwPZis0 As String) As String
'    CalcPZ = CStr(98 - modDecimal(CDec(IBANwPZis0), 97))
'End Function
'
Public Function PadLeft0(ByVal s As String, ByVal w As Long) As String
'shitty shitty bad function!!!!!!!
    'PadLeft0 = String$(w - Len(s), "0") & s
    PadLeft0 = MString.PadLeft(s, w, "0")
End Function

'Function PadLeft(this As String, _
'                 ByVal totalWidth As Long, _
'                 Optional ByVal paddingChar As String) As String
'    If LenB(paddingChar) Then
'        If Len(this) < totalWidth Then
'            PadLeft = String$(totalWidth, paddingChar)
'            MidB$(PadLeft, totalWidth * 2 - LenB(this) + 1) = this
'        Else
'            PadLeft = this
'        End If
'    Else
'        PadLeft = Space$(totalWidth)
'        RSet PadLeft = this
'    End If
'End Function

Public Function CalcPZ(BBANwLCPZ0 As String) As String
    CalcPZ = PadLeft0(CStr(98 - Modulo(BBANwLCPZ0, 97)), 2)
End Function

Public Function Modulo(ByVal Dividend As String, ByVal Divisor As Double)
    'Thanks to Hondo Alias Andreas
    Dim a As Variant
    Dim b As Variant
    Do While Len(Dividend) > 0
        a = b & Left(Dividend, 9 - Len(CStr(b))): Dividend = Mid(Dividend, 10 - Len(CStr(b)))
        b = a Mod Divisor
    Loop
    Modulo = b
End Function

Public Function DecodeAlphas(ByVal str As String) As String
    Dim i As Long, c As String
    Do While i < Len(str)
        i = i + 1
        c = Mid(str, i, 1)
        If InStr(1, "0123456789", c) = 0 Then
            str = Replace(str, c, CStr(Asc(c) - 55))
        End If
    Loop
    DecodeAlphas = str
End Function

Public Function RemoveLeading0(ByVal str As String) As String
    Dim i As Long
    RemoveLeading0 = str
    For i = 1 To Len(str) '- 1
        If Mid(str, i, 1) <> "0" Then
            'i = i - 1
            Exit For
        End If
    Next
    If i >= 1 Then RemoveLeading0 = Mid(str, i)
End Function

Public Function Group4(ByVal s As String) As String
    s = StringClean(s)
    Dim sout As String
    Do While Len(s) > 3
        sout = sout & " " & Left(s, 4)
        s = Right(s, Len(s) - 4)
    Loop
    If Len(s) > 0 Then sout = sout & " " & s
    Group4 = sout
End Function

Public Function StringClean(ByVal s As String) As String
    StringClean = Trim(MString.ReplaceAll(s, " .-,", ""))
    'StringClean = Trim$(MString.RecursiveReplace(s, " .-,", "")) 'Aarg it does not remove whitespaces " " but why
End Function

Public Function Contains(col As Collection, elem As String) As Boolean
    'https://www.vb-tec.de/collctns.htm
    On Error Resume Next

    If IsEmpty(col(elem)) Then: 'DoNothing
    Contains = (Err.Number = 0)

    On Error GoTo 0

End Function

'Sub ReadFileIBANcodes() 'ByRef sArr() As String)
'    Dim FNm As String:  FNm = App.Path & "\IBANcodes.txt"
'    Dim FNr As Integer: FNr = FreeFile
'Try: On Error GoTo Finally
'    Open FNm For Binary As FNr
'    Dim File As String: File = Space(LOF(FNr))
'    Get FNr, , File
'    Close FNr
'    Form1.Text1 = File
'    Dim sArr() As String: sArr = Split(File, vbCrLf)
'    ReDim m_IBANInfo(UBound(sArr))
'    Dim i As Long, c As Long, sLine As String
'    For i = 0 To UBound(sArr)
'        Dim lArr, ln, lc, bi, b, d, k, KK, r, s, X
'        lArr = Split(sArr(i), vbTab)
'        ln = Trim(lArr(0)) 'Ländername
'        sLine = Trim(lArr(1))
'        lc = Left(sLine, 2) 'Ländercode
'        bi = Mid(sLine, 6)
'        bi = Trim(bi)
'        bi = Replace(bi, " ", "")
''        For c = 1 To Len(bi)
''            Select Case Mid(bi, c, 1)
''            Case "b": b = b + 1
''            Case "k": k = k + 1
''            Case "K": KK = KK + 1
''            Case "r": r = r + 1
''            Case "s": s = s + 1
''            Case "X": X = X + 1
''            End Select
''        Next
'        'und genau hier isses quatsch, weil die Reihenfolge jetzt nicht mehr stimmt.
'        'd.h. man muss den String wieder so zusammenbasteln, dass die Reihenfolge wieder stimmt
'        'also nochmal die Schleife
'        sLine = ln & "; " & lc & "; "
'        Dim bol, char
'        'For c = 1 To Len(bi)
'        c = 1
'        Do Until c >= Len(bi)
'            char = Mid(bi, c, 1)
'            c = c + 1
'            Select Case char 'Mid(bi, c, 1)
'            Case "b":
'                Do Until c > Len(bi)
'                    char = Mid(bi, c, 1)
'                    c = c + 1
'                    b = b + 1
'                    If char <> "b" Or c > Len(bi) Then
'                        If c > Len(bi) Then b = b + 1
'                        c = c - 1
'                        sLine = sLine & IIf(b > 0, "b: " & b & "; ", "")
'                        b = 0
'                        Exit Do
'                    End If
'                Loop
'            Case "k":
'                Do Until c > Len(bi)
'                    char = Mid(bi, c, 1)
'                    c = c + 1
'                    k = k + 1
'                    If char <> "k" Or c > Len(bi) Then
'                        If c > Len(bi) Then k = k + 1
'                        c = c - 1
'                        sLine = sLine & IIf(k > 0, "k: " & k & "; ", "")
'                        k = 0
'                        Exit Do
'                    End If
'                Loop
'            Case "K":
'                Do Until c > Len(bi)
'                    char = Mid(bi, c, 1)
'                    c = c + 1
'                    KK = KK + 1
'                    If char <> "K" Or c > Len(bi) Then
'                        If c > Len(bi) Then KK = KK + 1
'                        c = c - 1
'                        sLine = sLine & IIf(KK > 0, "KK: " & KK & "; ", "")
'                        KK = 0
'                        Exit Do
'                    End If
'                Loop
'            Case "r":
'                Do Until c > Len(bi)
'                    char = Mid(bi, c, 1)
'                    c = c + 1
'                    r = r + 1
'                    If char <> "r" Or c > Len(bi) Then
'                        If c > Len(bi) Then r = r + 1
'                        c = c - 1
'                        sLine = sLine & IIf(r > 0, "r: " & r & "; ", "")
'                        r = 0
'                        Exit Do
'                    End If
'                Loop
'            Case "s":
'                Do Until c > Len(bi)
'                    char = Mid(bi, c, 1)
'                    c = c + 1
'                    s = s + 1
'                    If char <> "s" Or c > Len(bi) Then
'                        If c > Len(bi) Then s = s + 1
'                        c = c - 1
'                        sLine = sLine & IIf(s > 0, "s: " & s & "; ", "")
'                        s = 0
'                        Exit Do
'                    End If
'                Loop
'            Case "X":
'                Do Until c > Len(bi)
'                    char = Mid(bi, c, 1)
'                    c = c + 1
'                    X = X + 1
'                    If char <> "X" Or c > Len(bi) Then
'                        If c > Len(bi) Then X = X + 1
'                        c = c - 1
'                        sLine = sLine & IIf(X > 0, "X: " & X & "; ", "")
'                        X = 0
'                        Exit Do
'                    End If
'                Loop
'            End Select
'        Loop
''        sLine = ln & "; " & LC & "; " & _
''          IIf(b > 0, "b: " & b & "; ", "") & _
''          IIf(d > 0, "d: " & d & "; ", "") & _
''          IIf(k > 0, "k: " & k & "; ", "") & _
''          IIf(KK > 0, "KK: " & KK & "; ", "") & _
''          IIf(r > 0, "r: " & r & "; ", "") & _
''          IIf(s > 0, "s: " & s & "; ", "") & _
''          IIf(X > 0, "X: " & X, "")
'        m_IBANInfo(i) = sLine
'        b = 0: d = 0: k = 0: KK = 0: r = 0: s = 0: X = 0
'    Next
'    Exit Sub
'Finally:
'    Close FNr
'End Sub
'Public Property Get IBANInfo(ByVal Index As Long) As String
'    IBANInfo = m_IBANInfo(Index)
'End Property
'Public Sub IBANInfoFillCombo(aCB As ComboBox)
'    With aCB
'        .Clear
'        Dim sLine
'        For Each sLine In m_IBANInfo
'            If Len(sLine) Then
'                .AddItem Split(sLine, "; ")(0)
'            End If
'        Next
'    End With
'End Sub
'Public Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (ByRef Arr() As Any) As Long
'Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dst As Any, ByRef Src As Any, ByVal BytLength As Long)
'Public Function StrArrPtr(ByRef strArr As Variant) As Long
'    Call RtlMoveMemory(StrArrPtr, ByVal VarPtr(strArr) + 8, 4)
'End Function
'Public Property Get SAPtr(ByVal pArr As Long) As Long
'    Call RtlMoveMemory(SAPtr, ByVal pArr, 4)
'End Property
'Public Property Let SAPtr(ByVal pArr As Long, ByVal RHS As Long)
'    Call RtlMoveMemory(ByVal pArr, RHS, 4)
'End Property

'Property Get IBANTable() As Collection
'    If IBANTableCol Is Nothing Then
'        Set IBANTableCol = New Collection
'        Dim sLine, key, vlu
'        For Each sLine In Split(IBANTbl, vbCrLf)
'            key = CStr(Asc(Mid(sLine, 1, 1)) - 55) & CStr(Asc(Mid(sLine, 7, 1)) - 55)
'            vlu = Mid(sLine, 13)
'            IBANTableCol.Add vlu, key
'        Next
'    End If
'    Set IBANTable = IBANTableCol
'End Property
'Private Property Get IBANTable() 'As Collection
'    If m_IBANTable Is Nothing Then
'        Set m_IBANTable = New Collection
'        Dim File As String: File = ReadFileIBANcodes
'        Dim lines: lines = Split(File, vbCrLf)
'        Dim line
'        For Each line In lines
'            Dim elms: elms = Split(line, vbTab)
'            Dim Name: Name = elms(0)
'            Dim LC:     LC = Left(elms(1), 2)
'            Dim BBAN: BBAN = Replace(Mid(elms(1), 5), " ", "")
'            Dim s As String: s = Name & "; " & LC & "; " & BBAN
'            m_IBANTable.Add s, LC
'        Next
'    End If
'    Set IBANTable = m_IBANTable
'End Property
'Public Function CalcIBAN_DE(ByVal Blz, ByVal KtoNr) As String
'    CalcIBAN_DE = CalcIBAN(Deutschland, Blz, KtoNr)
'End Function
'Function IBANLänderCodeToStr(Land As IBANLänderCode) As String
'    Dim s As String: s = CStr(Land)
'    IBANLänderCodeToStr = Chr$(CInt(Left$(s, 2)) + 55) & Chr$(CInt(Mid$(s, 3)) + 55)
'End Function
'Public Function CalcIBAN(Land As IBANLänderCode, ByVal Blz, ByVal KtoNr, Optional ByVal Bereich = 0, Optional ByVal Typ = 0, Optional ByVal KtoNr2 = 0, Optional ByVal ID = 0) As String
'    Dim s: s = IBANTable.Item(CStr(Land))
'    Dim arr: arr = Split(s, "; ")
'    Dim i, BBAN, pcnt, pz0, pz1, pz2
'    Dim sw As Boolean
'    For i = 0 To UBound(arr)
'        s = Split(arr(i), ": ")
'        If s(0) = "p" Then
'            If sw Then pcnt = pcnt + 1
'            Select Case pcnt
'            Case 0: pz0 = pz0 & "0"
'            Case 1: pz1 = pz1 & "0"
'            Case 2: pz2 = pz2 & "0"
'            End Select
'            sw = False
'        Else
'            sw = True
'            Select Case s(0)
'            Case "BLZ":                BBAN = BBAN & PadLeft0(CStr(Blz), s(1))
'            Case "Bereich":            BBAN = BBAN & PadLeft0(CStr(Bereich), s(1))
'            Case "Typ":                BBAN = BBAN & PadLeft0(CStr(Typ), s(1))
'            Case "Kontonummer":        BBAN = BBAN & PadLeft0(CStr(KtoNr), s(1))
'            Case "Kto. 1. Teil":       BBAN = BBAN & PadLeft0(CStr(KtoNr), s(1))
'            Case "Kto. 2. Teil":       BBAN = BBAN & PadLeft0(CStr(KtoNr2), s(1))
'            Case "Identifikationsnr.": BBAN = BBAN & PadLeft0(CStr(ID), s(1))
'            End Select
'        End If
'    Next
'    'jetzt ist die IBAN zusammengestellt und fertig zum Berechnen der PZ
'    Select Case Land
'    Case Estland:  pz1 = "5"
'    Case Belgien:  pz1 = "34"
'    Case Dänemark: pz1 = "3"
'    Case Finnland: pz1 = "5"
'    End Select
'    Dim PZ As String
'    PZ = CalcPZ(BBAN & pz1 & CStr(Land) & pz0) '& pz1 & pz2)
'    CalcIBAN = IBANLänderCodeToStr(Land) & PZ & BBAN & pz1
'End Function
'
'Function CalcPZ(ByVal IBANwPZis0 As String) As String
'    CalcPZ = CStr(98 - modDecimal(CDec(IBANwPZis0), 97))
'End Function
'
'Public Function PadLeft0(ByVal s As String, ByVal w As Long) As String
'    PadLeft0 = String$(w - Len(s), "0") & s
'End Function
'
'' Hilfsfunktion: Modulo für große Zahlen
'Public Function modDecimal(Dividend, Divisor)
'    If Divisor = 0 Then
'        modDecimal = -1
'    Else
'        modDecimal = Dividend - Divisor * (Round(Dividend / Divisor))
'        If modDecimal < 0 Then modDecimal = Divisor + modDecimal
'    End If
'End Function
'
'Function IBANLänderCode_ToStr(LC As IBANLänderCode) As String
'    Dim s As String
'    Select Case LC
'    Case Andorra: s = "" ' = 1013          '   A D
'    Case Austria: s = "" ' = 1029          '   A T
'    Case Belgien: s = "" ' = 1114          '   B E
'    Case Dänemark: s = "" ' = 1320         '   D K
'    Case Deutschland: s = "" ' = 1314      '   D E
'    Case Estland: s = "" ' = 1414          '   E E
'    Case Finnland: s = "" ' = 1518         '   F I
'    Case Frankreich: s = "" ' = 1527       '   F R
'    Case Gibraltar: s = "" ' = 1618        '   G I
'    Case Griechenland: s = "" ' = 1627     '   G R
'    Case Großbritannien: s = "" ' = 1611   '   G B
'    Case Irland: s = "" ' = 1814           '   I E
'    Case Island: s = "" ' = 1828           '   I S
'    Case Italien: s = "" ' = 1829          '   I T
'    Case Lettland: s = "" ' = 2131         '   L V
'    Case Litauen: s = "" ' = 2129          '   L T
'    Case Luxemburg: s = "" ' = 2130        '   L U
'    Case Niederlande: s = "" ' = 2321      '   N L
'    Case Norwegen: s = "" ' = 2324         '   N O
'    Case Polen: s = "" ' = 2521            '   P L
'    Case Portugal: s = "" ' = 2529         '   P T
'    Case Schweden: s = "" ' = 2814         '   S E
'    Case Schweiz: s = "" ' = 1217          '   C H
'    Case Slowakei: s = "" ' = 2820         '   S K
'    Case Slowenien: s = "" ' = 2818        '   S I
'    Case Spanien: s = "" ' = 1428          '   E S
'    Case Tschechien: s = "" ' = 1235       '   C Z
'    Case Ungarn: s = "" ' = 1730           '   H U
'    Case Zypern: s = "" ' = 1234           '   C Y
'    End Select
'    IBANLänderCode_ToStr = s
'End Function
'Andorra '24  A   D   p   p   BLZ Bereich Kontonummer
'Belgien '16  B   E   p   p   BLZ Kontonummer p   p
''Bulgarien               p   p
'Dänemark   ' 18  D   K   p   p   BLZ Kontonummer p
'Deutschland '22  D   E   p   p   BLZ Kontonummer
'Estland '20  E   E   p   p           Kontonummer p
'Finnland   ' 18  F   I   p   p   BLZ Kontonummer p
'Frankreich ' 27  F   R   p   p   BLZ Bereich Kontonummer p   p
'Gibraltar  ' 23  G   I   p   p   BLZ Kontonummer
'Griechenland   ' 27  G   R   p   p   BLZ Bereich Kontonummer
'Großbritannien ' 22  G   B   p   p   BLZ Bereich Kontonummer
'Irland  '22  I   E   p   p   BLZ Bereich Kontonummer
'Island  '26  I   S   p   p   BLZ Typ Kontonummer Identifikationsnr.
'Italien '27  I   T   p   p   p   BLZ Bereich Kontonummer
'Lettland    '21  L   V   p   p   BLZ Kontonummer
'Litauen '20  L   T   p   p   BLZ Kontonummer
'Luxemburg '20  L   U   p   p   BLZ Kontonummer
''Malta               p   p
'Niederlande ' 18  N   L   p   p   BLZ Kontonummer
'Norwegen '15  N   O   p   p   BLZ Kontonummer p
'Österreich  '20  A   T   p   p   BLZ Kontonummer
'Polen   '28  P   L   p   p   BLZ Kontonummer
'Portugal   ' 25  P   T   p   p   BLZ Bereich Kontonummer p   p
''Rumänien                p   p
'Schweden    '24  S   E   p   p   BLZ Kontonummer p
'Schweiz '21  C   H   p   p   BLZ Kontonummer
'Slowakei    '24  S   K   p   p   BLZ Kto.nr. 1.Teil  Kto.nr. 2.Teil
'Slowenien   '19  S   I   p   p   BLZ Kontonummer p   p
'Spanien '24  E   S   p   p   BLZ Bereich p   p   Kontonummer
'Tschechien  '24  C   Z   p   p   BLZ Kto. 1. Teil    Kto. 2. Teil
'Ungarn  '28  H   U   p   p   BLZ Bereich p   Kontonummer p
'Zypern  '28  C   Y   p   p   BLZ Bereich Kontonummer
'Private Type TIBAN
'    PZ      As Byte 'enthält die Längen der Zeichen der einzelnen Bestandteile
'    BLZ     As Byte
'    Bereich As Byte
'    KtoNr   As Byte
'    KtoNr2  As Byte
'    ID      As Byte
'    PZ2     As Byte
'    Land    As IBANLC
'End Type
'Private TIBANs(31) As TIBAN
'
'Public Sub FillTIBANs()
'    Dim i
'    TIBAN i, Array(2, 4, 4, 12, 0, 0, 0), Andorra: i = i + 1
'    TIBAN i, Array(2, 3, 7, 0, 0, 0, 2), Belgien: i = i + 1
'    TIBAN i, Array(2, 4, 0, 9, 0, 0, 0), Dänemark: i = i + 1
'    TIBAN i, Array(2, 8, 0, 10, 0, 0, 0), Deutschland: i = i + 1
'    TIBAN i, Array(2, 2, 0, 11, 0, 0, 1), Estland: i = i + 1
'    TIBAN i, Array(2, 6, 0, 7, 0, 0, 1), Finnland: i = i + 1
'    TIBAN i, Array(2, 5, 5, 11, 0, 0, 2), Frankreich: i = i + 1
'    TIBAN i, Array(2, 4, 0, 15, 0, 0, 0), Gibraltar: i = i + 1
'    TIBAN i, Array(2, 3, 4, 16, 0, 0, 0), Griechenland: i = i + 1
'    TIBAN i, Array(2, 4, 6, 8, 0, 0, 0), Großbritannien: i = i + 1
'    TIBAN i, Array(2, 4, 6, 8, 0, 0, 0), Irland: i = i + 1
'    TIBAN i, Array(2, 4, 0, 8, 10, 0, 0), Island: i = i + 1
'    TIBAN i, Array(3, 5, 5, 12, 0, 0, 0), Italien: i = i + 1
'    TIBAN i, Array(2, 4, 0, 0, 0, 0, 0), Lettland: i = i + 1
'    TIBAN i, Array(2, 5, 0, 0, 0, 0, 0), Litauen: i = i + 1
'    TIBAN i, Array(2, 3, 0, 0, 0, 0, 0), Luxemburg: i = i + 1
'    TIBAN i, Array(2, 4, 0, 0, 0, 0, 0), Niederlande: i = i + 1
'    TIBAN i, Array(2, 4, 0, 0, 0, 0, 0), Norwegen: i = i + 1
'    TIBAN i, Array(2, 5, 0, 0, 0, 0, 0), Österreich: i = i + 1
'    TIBAN i, Array(2, 8, 0, 0, 0, 0, 0), Polen: i = i + 1
'    TIBAN i, Array(2, 4, 4, 0, 0, 0, 0), Portugal: i = i + 1
'    TIBAN i, Array(2, 3, 0, 0, 0, 0, 0), Schweden: i = i + 1
'    TIBAN i, Array(2, 5, 0, 0, 0, 0, 0), Schweiz: i = i + 1
'    TIBAN i, Array(2, 4, 0, 0, 0, 0, 0), Slowakei: i = i + 1
'    TIBAN i, Array(2, 5, 0, 0, 0, 0, 0), Slowenien: i = i + 1
'    TIBAN i, Array(2, 4, 4, 0, 0, 0, 0), Spanien: i = i + 1
'    TIBAN i, Array(2, 4, 0, 0, 0, 0, 0), Tschechien: i = i + 1
'    TIBAN i, Array(2, 3, 4, 0, 0, 0, 0), Ungarn: i = i + 1
'    TIBAN i, Array(2, 3, 5, 0, 0, 0, 0), Zypern: i = i + 1
'End Sub
'Sub TIBAN(i, arr, lc As IBANLC)
'    With TIBANs(i)
'        .PZ = arr(0)
'        .BLZ = arr(1)
'        .Bereich = arr(2)
'        .KtoNr = arr(3)
'        .KtoNr2 = arr(4)
'        .ID = arr(5)
'        .PZ2 = arr(6)
'        .Land = lc
'    End With
'End Sub

'BBANExamples
'Italy
'     X 05428 11101 000000123456
'IT60 X 05428 11101 000000123456
'
'Hungary
'3!n 4!n 1!n 15!n 1!n
'     117 7301 6 111110180000000 0
'HU42 117 7301 6 111110180000000 0
'
'United Kingdom
'GB 2!n 4!a 6!n 8!n
'     NWBK 601613 31926819
'GB29 NWBK 601613 31926819
'
'Turkey
'TR 2!n 5!n 1!c 16!c
'     00061 0 0519786457841326
'TR33 00061 0 0519786457841326
