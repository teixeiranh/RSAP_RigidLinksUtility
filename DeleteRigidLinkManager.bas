Attribute VB_Name = "DeleteRigidLinkManager"
'@Folder "RSAPRigidLinks.Manager"
'CommumRigidLinkManager
Option Explicit

Sub DeleteAllRigidLinks()
    RobotAPIHelper.PrepareEnvironment
    
    Dim NumberOfRigidLinks As Long
    NumberOfRigidLinks = RobotAPIHelper.RigidiLinkServer.count
    
    Dim i As Long
    Dim j As Long
    j = 0
    PrintColumnCounter.PrintCounters NumberOfRigidLinks
    'Necessário iterar para não saltar índices
    For i = NumberOfRigidLinks To 1 Step -1
        RobotAPIHelper.RigidiLinkServer.Remove i
        j = j + 1
        PrintColumnCounter.UpdateBarNumberCounter j
    Next i
    
    RobotAPIHelper.CleanVariables
    PrintColumnCounter.CleanCounters
End Sub

Sub DeleteForSelectedColumns()
    RobotAPIHelper.PrepareEnvironmentWithBarSelection

    Dim i As Long
    Dim barNumber As Long
    Dim bar As IRobotBar
    PrintColumnCounter.PrintCounters RobotAPIHelper.barSelection.count
    For i = 1 To RobotAPIHelper.barSelection.count
        barNumber = RobotAPIHelper.barSelection.Get(i)
        Set bar = RobotAPIHelper.barServer.Get(barNumber)

        DeleteRigidLinkIfMasterNode bar.StartNode
        DeleteRigidLinkIfMasterNode bar.EndNode
        PrintColumnCounter.UpdateBarNumberCounter i
    Next i

    RobotAPIHelper.CleanVariables
    PrintColumnCounter.CleanCounters
End Sub

Private Sub DeleteRigidLinkIfMasterNode(nodeNumber As Long)
    Dim i As Long
    Dim robRigidLink As RobotNodeRigidLinkDef
    For i = 1 To RobotAPIHelper.RigidiLinkServer.count
        Set robRigidLink = RobotAPIHelper.RigidiLinkServer.Get(i)
        
        If robRigidLink.PrimaryNode = nodeNumber Then
            RobotAPIHelper.RigidiLinkServer.Remove i
            Exit For
        End If
    Next i
End Sub

