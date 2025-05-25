Attribute VB_Name = "PrintColumnCounter"
'@Folder "RSAPRigidLinks.Utilities"
Option Explicit

'Private Const NUMBER_OF_ELEMENTS_TITLE_NAMED_RANGE As String = "NumberOfBarsTitle"
'Private Const ELEMENT_NUMBER_TITLE_NAMED_RANGE As String = "BarNumberTitle"
Private Const NUMBER_OF_ELEMENTS_NAMED_RANGE As String = "K2"
Private Const ELEMENT_NUMBER_NAMED_RANGE As String = "J2"

'Private Const NUMBER_OF_BARS_TITLE As String = "Number of Elements"
'Private Const BAR_NUMBER_TITLE As String = "Element Number"

Sub PrintCounters(ByVal count As Long)
    'Range(NUMBER_OF_ELEMENTS_TITLE_NAMED_RANGE).value = NUMBER_OF_BARS_TITLE
'    Range(ELEMENT_NUMBER_TITLE_NAMED_RANGE).value = BAR_NUMBER_TITLE
    Range(NUMBER_OF_ELEMENTS_NAMED_RANGE).value = count
End Sub

Sub UpdateBarNumberCounter(ByVal counter As Integer)
    Range(ELEMENT_NUMBER_NAMED_RANGE).value = counter
End Sub

Sub CleanCounters()
    'Range(NUMBER_OF_ELEMENTS_TITLE_NAMED_RANGE).ClearContents
    'Range(ELEMENT_NUMBER_TITLE_NAMED_RANGE).ClearContents
    Range(NUMBER_OF_ELEMENTS_NAMED_RANGE).ClearContents
    Range(ELEMENT_NUMBER_NAMED_RANGE).ClearContents
End Sub
