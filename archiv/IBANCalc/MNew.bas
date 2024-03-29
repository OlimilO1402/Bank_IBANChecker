Attribute VB_Name = "MNew"
Option Explicit

Public Function IBANInfo(Name As String, lc As String, bbinf As BBANInfo) As IBANInfo
    Set IBANInfo = New IBANInfo
    IBANInfo.New_ Name, lc, bbinf
End Function

Public Function BBANInfo(ByVal infoW As String) As BBANInfo
    Set BBANInfo = New BBANInfo
    BBANInfo.New_ infoW
End Function

Public Function BBANPart(aChar As String, ByVal aPos As Byte, ByVal aLength As Byte) As BBANPart
    Set BBANPart = New BBANPart
    BBANPart.New_ aChar, aPos, aLength
End Function

Public Function BBANValue(bbp As BBANPart, ByVal Value As String) As BBANValue
    Set BBANValue = New BBANValue
    BBANValue.New_ bbp, Value
End Function

Public Function IBAN(IBANInfos As IBANInfos, siban As String) As IBAN
    Set IBAN = New IBAN
    IBAN.New_ IBANInfos, siban
End Function

Public Function BBAN(BBANInfo As BBANInfo, sBBAN As String) As BBAN
    Set BBAN = New BBAN
    BBAN.New_ BBANInfo, sBBAN
End Function

Public Function IBANCreator(IBANInfos As IBANInfos, aIBANInfo As IBANInfo, sArr() As String) As IBANCreator
    Set IBANCreator = New IBANCreator
    IBANCreator.New_ IBANInfos, aIBANInfo, sArr
End Function

Public Function BlzBic(sLine As String) As BlzBic
    Set BlzBic = New BlzBic
    BlzBic.ParseLine sLine
End Function

Public Function BlzBics(aFNm As String) As BlzBics
    Set BlzBics = New BlzBics
    BlzBics.ParseFile aFNm
End Function

