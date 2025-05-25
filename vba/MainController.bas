Attribute VB_Name = "MainController"
'@Folder "RSAPRigidLinks.Controller"
Option Explicit

Public Sub HandleButtonAction(Action As String)
    Select Case Action
    Case RigidLinkConst.CREATE_RIGIDLINKS_NODES_PLANE
        NodePlaneRigidLinkManager.CreateRigidLinksForNodesPlane
    Case RigidLinkConst.CREATE_RIGIDLINKS_FOR_SELECTION_BARS_PLANE
        BarPlaneRigidLinkManager.CreateRigidLinksForBarsPlane
    Case RigidLinkConst.DELETE_ALL_RIGIDLINKS
        DeleteRigidLinkManager.DeleteAllRigidLinks
    Case RigidLinkConst.DELETE_RIGIDLINKS_FOR_SELECTION_BARS_PLANE
        DeleteRigidLinkManager.DeleteForSelectedColumns
    Case RigidLinkConst.REDO_RIGIDLINKS_FOR_SELECTION_BARS_PLANE
        BarPlaneRigidLinkManager.RedoRigidLinks
    Case RigidLinkConst.SELECT_ALL_MASTER_NODES_BARS_PLANE
        SelectionManager.SelectRigidLinksMasterNodes
    Case RigidLinkConst.SELECT_ALL_SLAVE_NODES_BARS_PLANE
        SelectionManager.SelectAllRigidLinksSlaveNodes
    Case RigidLinkConst.CREATE_RIGIDLINKS_FOR_SELECTION_BARS_LINE
        BarLinearRigidLinkManager.CreateRigidLinksForBarsLine
    Case RigidLinkConst.REDO_RIGIDLINKS_FOR_SELECTION_BARS_LINE
        BarLinearRigidLinkManager.CreateRigidLinksForBarsLine
    Case Else
        MsgBox AppConst.UNNOWKN_ACTION & Action, vbExclamation
    End Select
End Sub

