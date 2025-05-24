Attribute VB_Name = "UIModule"
'@Folder "RSAPRigidLinks.UI"
'UIModule
Option Explicit

Public Sub ButtonCreateRigidLinks_NodesPlane_Click()
    MainController.HandleButtonAction RigidLinkConst.CREATE_RIGIDLINKS_NODES_PLANE
End Sub

Public Sub ButtonCreateRigidLinks_BarsPlane_Click()
    MainController.HandleButtonAction RigidLinkConst.CREATE_RIGIDLINKS_FOR_SELECTION_BARS_PLANE
End Sub

Public Sub ButtonRedoRigidLinks_BarsPlane_Click()
    MainController.HandleButtonAction RigidLinkConst.REDO_RIGIDLINKS_FOR_SELECTION_BARS_PLANE
End Sub

Public Sub ButtonDeleteRigidLinks_BarsPlane_Click()
    MainController.HandleButtonAction RigidLinkConst.DELETE_RIGIDLINKS_FOR_SELECTION_BARS_PLANE
End Sub

Public Sub ButtonSelectAllMasterNodes_BarsPlane_Click()
    MainController.HandleButtonAction RigidLinkConst.SELECT_ALL_MASTER_NODES_BARS_PLANE
End Sub

Public Sub ButtonSelectAllSlaveNodes_BarsPlane_Click()
    MainController.HandleButtonAction RigidLinkConst.SELECT_ALL_SLAVE_NODES_BARS_PLANE
End Sub

Public Sub ButtonDeleteAllRigidLinks_Click()
    MainController.HandleButtonAction RigidLinkConst.DELETE_ALL_RIGIDLINKS
End Sub

Public Sub ButtonCreateRigidLinks_BarLinear_Click()
    MainController.HandleButtonAction RigidLinkConst.CREATE_RIGIDLINKS_FOR_SELECTION_BARS_LINE
End Sub

Public Sub ButtonRedoRigidLinks_BarsLine_Click()
    MainController.HandleButtonAction RigidLinkConst.REDO_RIGIDLINKS_FOR_SELECTION_BARS_LINE
End Sub

