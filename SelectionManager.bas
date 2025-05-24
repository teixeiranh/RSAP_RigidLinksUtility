Attribute VB_Name = "SelectionManager"
'@Folder "RSAPRigidLinks.Manager"
'SelectionManager
Option Explicit

Sub SelectRigidLinksMasterNodes()
    RobotAPIHelper.PrepareEnvironment
    RobotAPIHelper.NodeSelection.Clear
    
    Dim i As Integer
    Dim PrimaryNodeNumber As Long
    For i = 1 To RobotAPIHelper.RigidiLinkServer.count
        PrimaryNodeNumber = RobotAPIHelper.RigidiLinkServer.Get(i).PrimaryNode
        RobotAPIHelper.NodeSelection.AddOne PrimaryNodeNumber
    Next i
    
End Sub


Sub SelectAllRigidLinksSlaveNodes()
    RobotAPIHelper.PrepareEnvironment
    RobotAPIHelper.NodeSelection.Clear
    
    Dim i As Integer
    Dim SecondaryNodeNumber As String
    For i = 1 To RobotAPIHelper.RigidiLinkServer.count
        SecondaryNodeNumber = RobotAPIHelper.RigidiLinkServer.Get(i).SecondaryNodes
        RobotAPIHelper.NodeSelection.AddText SecondaryNodeNumber
    Next i
End Sub
