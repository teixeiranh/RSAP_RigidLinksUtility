Attribute VB_Name = "MathUtility"
'@Folder "RSAPRigidLinks.Utilities"
Option Explicit

Public Function ConvertToRadians(ByVal angle As Double) As Double
    ConvertToRadians = angle * AppConst.PI_VALUE / 180
End Function

