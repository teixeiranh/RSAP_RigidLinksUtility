Attribute VB_Name = "RobotAPIHelper"
'@Folder "RSAPRigidLinks.RSAP-API"
'RobotAPIHelper
Option Explicit

'Global
Public RobotApp As RobotApplication
Public Project As RobotProject

Public RobotNodesServer As RobotNodeServer
Public RigidiLinkServer As RobotNodeRigidLinkServer
Public barServer As RobotBarServer

Public NodeSelection As RobotSelection
Public AllNodesCol As RobotNodeCollection
Public barSelection As RobotSelection

Public Label As RobotLabel

Public RigidLinkData As RobotNodeRigidLinkData

'Constants
Public Const G_FIXED_RIGID_LINK As String = "rl_fixed"

Private Const MESSAGE_TO_START As String = "Start Robot and Load Model!"
Private Const MESSAGE_TO_SELECT_STRUCTURE As String = "Structure type should be Frame3D or Shell or Building!"
Private Const MESSAGE_TO_CREATE_NODES As String = "Please create nodes in Robot!"
Private Const MESSAGE_TO_SELECT_NODES As String = "Please select nodes in Robot!"
Private Const ERROR_MESSAGE As String = "Error!"

Public Sub VerifyIfRobotIsOpened()
    Set RobotApp = New RobotApplication
    If Not RobotApp.Visible Then
        Set RobotApp = Nothing
        MsgBox MESSAGE_TO_START, vbOKOnly, ERROR_MESSAGE
        End
    Else
        If (RobotApp.Project.Type <> I_PT_FRAME_3D) And _
           (RobotApp.Project.Type <> I_PT_SHELL) And _
           (RobotApp.Project.Type <> I_PT_BUILDING) Then
            MsgBox MESSAGE_TO_SELECT_STRUCTURE, vbOKOnly, ERROR_MESSAGE
            End
        End If
    End If
End Sub

Public Sub VerifyIfNodesWereCreated()
    Set RobotApp = New RobotApplication
    Set AllNodesCol = RobotApp.Project.Structure.Nodes.GetAll
    If AllNodesCol.count = 0 Then
        MsgBox MESSAGE_TO_CREATE_NODES, vbOKOnly, ERROR_MESSAGE
        End
    End If
End Sub

Public Sub VerifyIfNodesAreSelected()
    Set RobotApp = New RobotApplication
    Set NodeSelection = RobotApp.Project.Structure.Selections.Get(I_OT_NODE)
    If NodeSelection.count = 0 Then
        MsgBox MESSAGE_TO_SELECT_NODES, vbOKOnly, ERROR_MESSAGE
        End
    End If
End Sub

Public Sub CleanVariables()
    RobotApp.Project.ViewMngr.Refresh

    Set RobotApp = Nothing
    Set AllNodesCol = Nothing
    Set RigidiLinkServer = Nothing
    Set RigidLinkData = Nothing
    Set RobotNodesServer = Nothing
    Set NodeSelection = Nothing
End Sub

Public Sub SetStates()
    Set RobotApp = New RobotApplication
    RobotApp.Visible = True
    RobotApp.Interactive = 1
    RobotApp.UserControl = True
End Sub

Public Sub RigidLinkLabelCreate()
    Set Label = RobotApp.Project.Structure.Labels.Create(I_LT_NODE_RIGID_LINK, G_FIXED_RIGID_LINK)
    Set RigidLinkData = Label.Data
    RigidLinkData.UX = True
    RigidLinkData.UY = True
    RigidLinkData.UZ = True
    RigidLinkData.RX = True
    RigidLinkData.RY = True
    RigidLinkData.RZ = True
    RobotApp.Project.Structure.Labels.Store Label
End Sub

Public Sub PrepareEnvironmentWithNodeSelection()
    PrepareEnvironment
    VerifyNodesCreationAndSelection
End Sub

Public Sub PrepareEnvironmentWithBarSelection()
    PrepareEnvironment
    Set barSelection = Project.Structure.Selections.Get(IRobotObjectType.I_OT_BAR)
End Sub

Public Sub PrepareEnvironment()
    Set RobotApp = New RobotApplication
    Set RobotNodesServer = RobotApp.Project.Structure.Nodes
    Set RigidiLinkServer = RobotApp.Project.Structure.Nodes.RigidLinks
    Set NodeSelection = RobotApp.Project.Structure.Selections.Get(I_OT_NODE)
    Set Project = RobotApp.Project
    Set barServer = Project.Structure.Bars
    RobotAPIHelper.VerifyIfRobotIsOpened
    RobotAPIHelper.SetStates
    RobotAPIHelper.RigidLinkLabelCreate
End Sub

Private Sub VerifyNodesCreationAndSelection()
    RobotAPIHelper.VerifyIfNodesWereCreated
    RobotAPIHelper.VerifyIfNodesAreSelected
End Sub


