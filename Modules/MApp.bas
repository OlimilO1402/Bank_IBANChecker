Attribute VB_Name = "MApp"
Option Explicit
Private m_IbInfos As IBANInfos
Private m_BlzBics As BlzBics
Private m_Doc     As Document
Private m_DefPFN  As String

Sub Main()
    Set m_IbInfos = MNew.IBANInfos(App.Path & "\Data\IBANcodes.txt")
    Set m_BlzBics = MNew.BlzBics(App.Path & "\Data\blzBIC3_2015_DE.txt")
    'Set m_BlzBics = MNew.BlzBics(App.Path & "\Data\blzBIC4_2022_AT.txt")
    m_DefPFN = App.Path & "\.." & "\Bank_IBANChecker_accessory\Bankaccounts.txt"
    GetNewDoc
    FMain.Show
End Sub

Public Property Get IBANInfos() As IBANInfos
    Set IBANInfos = m_IbInfos
End Property

Public Property Get BlzBics() As BlzBics
    Set BlzBics = m_BlzBics
End Property

Public Function GetNewDoc() As Document
    Set m_Doc = MNew.Document(m_DefPFN)
    Set GetNewDoc = m_Doc
End Function
