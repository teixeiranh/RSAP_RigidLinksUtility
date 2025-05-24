Attribute VB_Name = "CommonUtils"
'@Folder "RSAPRigidLinks.Utilities"
Option Explicit


Public Function IsSelectedInExcel(value As String) As Boolean
    IsSelectedInExcel = Range(value).value = "True"
End Function

