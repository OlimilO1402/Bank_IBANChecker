Attribute VB_Name = "MNew"
Option Explicit

Public Function IBANInfo(Name As String, lc As String, bbinf As BBANInfo) As IBANInfo
    Set IBANInfo = New IBANInfo: IBANInfo.New_ Name, lc, bbinf
End Function

Public Function IBANInfos(aPFN As String) As IBANInfos
    Set IBANInfos = New IBANInfos: IBANInfos.ReadFromFile aPFN
End Function

Public Function BBANInfo(ByVal infoW As String) As BBANInfo
    Set BBANInfo = New BBANInfo: BBANInfo.New_ infoW
End Function

Public Function BBANPart(aChar As String, ByVal aPos As Byte, ByVal aLength As Byte) As BBANPart
    Set BBANPart = New BBANPart: BBANPart.New_ aChar, aPos, aLength
End Function

Public Function BBANValue(bbp As BBANPart, ByVal Value As String) As BBANValue
    Set BBANValue = New BBANValue: BBANValue.New_ bbp, Value
End Function

Public Function IBAN(IBANInfos As IBANInfos, siban As String) As IBAN
    Set IBAN = New IBAN: IBAN.New_ IBANInfos, siban
End Function

Public Function BBAN(BBANInfo As BBANInfo, sBBAN As String) As BBAN
    Set BBAN = New BBAN: BBAN.New_ BBANInfo, sBBAN
End Function

Public Function IBANCreator(IBANInfos As IBANInfos, aIBANInfo As IBANInfo, List As Collection) As IBANCreator 'sArr() As String) As IBANCreator
    Set IBANCreator = New IBANCreator: IBANCreator.New_ IBANInfos, aIBANInfo, List 'sArr
End Function

Public Function BlzBic(sLine As String) As BlzBic
    Set BlzBic = New BlzBic: BlzBic.ParseLine sLine
End Function

Public Function BlzBics(aFNm As String) As BlzBics
    Set BlzBics = New BlzBics: BlzBics.ParseFile aFNm
End Function

Public Function PathFileName(ByVal aPathOrPFN As String, _
                    Optional ByVal aFileName As String, _
                    Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathOrPFN, aFileName, aExt
End Function

Public Function List(Of_T As EDataType, _
                     Optional ArrColStrTypList, _
                     Optional ByVal IsHashed As Boolean = False, _
                     Optional ByVal Capacity As Long = 32, _
                     Optional ByVal GrowRate As Single = 2, _
                     Optional ByVal GrowSize As Long = 0) As List
    Set List = New List: List.New_ Of_T, ArrColStrTypList, IsHashed, Capacity, GrowRate, GrowSize
End Function

Public Function NamedIBAN(ByVal aName As String, aIBAN As IBAN) As NamedIBAN
    Set NamedIBAN = New NamedIBAN: NamedIBAN.New_ aName, aIBAN
End Function

Public Function Document(aPFN As String) As Document
    Set Document = New Document: Document.New_ aPFN
End Function

