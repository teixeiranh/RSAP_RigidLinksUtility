Attribute VB_Name = "NodePlaneRigidLinkManager"
'@Folder "RSAPRigidLinks.Manager"
'NodePlaneRigidLinkManager
Option Explicit

Private Const XL_SECTION_WIDTH As String = "C3"
Private Const XL_SECTION_HEIGHT As String = "C4"
Private Const XL_ANGLE_OF_ROTATION As String = "C5"
Private Const XL_DIAMETER As String = "C3"
Private Const XL_NODES_GEOMETRY_TYPE As String = "NODES_RECTANGULAR_TYPE"

Private Const TYPE_RECTANGULAR As String = "Rectangular"
Private Const TYPE_CIRCULAR As String = "Circular"

Public Sub CreateRigidLinksForNodesPlane()
    Dim rigidLinkNet As INodeFinder
    Set rigidLinkNet = CreateRigidLink()
    
    If Not rigidLinkNet Is Nothing Then
        RobotAPIHelper.PrepareEnvironmentWithNodeSelection
        ConfigureAndApplyRigidLink rigidLinkNet
        RobotAPIHelper.CleanVariables
    End If
End Sub

Private Function CreateRigidLink() As INodeFinder
    Select Case Range(XL_NODES_GEOMETRY_TYPE).value
        Case TYPE_RECTANGULAR
            Set CreateRigidLink = New RectangularColumnRigidLink
        Case TYPE_CIRCULAR
            Set CreateRigidLink = New CircularColumnRigidLink
        Case Else
            MsgBox AppConst.UNNOWKN_RIGIDLINK_TYPE & Range(XL_NODES_GEOMETRY_TYPE).value, vbExclamation
            Set CreateRigidLink = Nothing
    End Select
End Function

Private Sub ConfigureAndApplyRigidLink(ByRef rigidLinkNet As INodeFinder)
    If TypeOf rigidLinkNet Is RectangularColumnRigidLink Then
        Dim rectangularLink As RectangularColumnRigidLink
        Set rectangularLink = rigidLinkNet
        FillRectangularRigidLinkData rectangularLink
    ElseIf TypeOf rigidLinkNet Is CircularColumnRigidLink Then
        Dim circularLink As CircularColumnRigidLink
        Set circularLink = rigidLinkNet
        FillCircularRigidLinkData circularLink
    End If
    
    Dim ii As Long
    For ii = 1 To RobotAPIHelper.NodeSelection.count
        rigidLinkNet.centerNode = RobotAPIHelper.NodeSelection.Get(ii)
        rigidLinkNet.FindListOfNodes
        RobotAPIHelper.RigidiLinkServer.Set rigidLinkNet.centerNode, rigidLinkNet.nodesListString, RobotAPIHelper.G_FIXED_RIGID_LINK
    Next ii
End Sub

Private Sub FillRectangularRigidLinkData(ByRef rigidLinkNet As RectangularColumnRigidLink)
    With rigidLinkNet
        .angleOfRotation = MathUtility.ConvertToRadians(Range(XL_ANGLE_OF_ROTATION).value)
        .sectionWidth = Range(XL_SECTION_WIDTH).value
        .sectionHeight = Range(XL_SECTION_HEIGHT).value
        .radiusBX = .sectionWidth / 2
        .radiusHy = .sectionHeight / 2
        .radiusDiagonal = Sqr(.radiusBX ^ 2 + .radiusHy ^ 2)
        .diagonalAngle = Atn(.radiusHy / .radiusBX)
    End With
End Sub

Private Sub FillCircularRigidLinkData(ByRef rigidLinkNet As CircularColumnRigidLink)
    rigidLinkNet.Diameter = Range(XL_DIAMETER).value
End Sub
