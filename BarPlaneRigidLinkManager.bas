Attribute VB_Name = "BarPlaneRigidLinkManager"
'@Folder "RSAPRigidLinks.Manager"
'BarPlaneRigidLinkManager
Option Explicit

Sub CreateRigidLinksForBarsPlane()
    RobotAPIHelper.PrepareEnvironmentWithBarSelection
    
    Dim i As Long
    PrintColumnCounter.PrintCounters RobotAPIHelper.barSelection.count
    For i = 1 To RobotAPIHelper.barSelection.count
        Dim sectionData As IRobotBarSectionData
        Dim bar As IRobotBar
        GetSectionData RobotAPIHelper.barServer, RobotAPIHelper.barSelection, i, bar, sectionData
        
        Select Case sectionData.ShapeType
            Case I_BSST_CONCR_COL_R
                SetupAndApplyRectangularRigidLink RobotAPIHelper.G_FIXED_RIGID_LINK, sectionData, bar
            Case I_BSST_CONCR_COL_C
                SetupAndApplyCircularRigidLink RobotAPIHelper.G_FIXED_RIGID_LINK, sectionData, bar
        End Select
        
        PrintColumnCounter.UpdateBarNumberCounter i
    Next i
    
    RobotAPIHelper.CleanVariables
    PrintColumnCounter.CleanCounters
End Sub

Private Sub GetSectionData(ByVal barServer As IRobotBarServer, ByVal barSelection As IRobotSelection, ByVal index As Long, ByRef bar As IRobotBar, ByRef sectionData As IRobotBarSectionData)
    Set bar = barServer.Get(barSelection.Get(index))
    
    Dim sectionLabelId As IRobotLabel
    Set sectionLabelId = bar.GetLabel(IRobotLabelType.I_LT_BAR_SECTION)
    
    Dim sectionLabel As IRobotLabel
    Set sectionLabel = RobotAPIHelper.Project.Structure.Labels.Get(I_LT_BAR_SECTION, sectionLabelId.Name)
    Set sectionData = sectionLabel.Data
End Sub

Private Sub SetupAndApplyRectangularRigidLink(ByVal rigidLinkLabel As String, ByVal sectionData As IRobotBarSectionData, ByVal bar As IRobotBar)
    Dim rectangularRigidLink As RectangularColumnRigidLink
    Set rectangularRigidLink = New RectangularColumnRigidLink
    ConfigureRectangularRigidLink sectionData, bar, rectangularRigidLink
    
    If CommonUtils.IsSelectedInExcel(AppConst.TOP_SECTION) Then
        ApplyRigidLinkToTargetNode bar.EndNode, rigidLinkLabel, rectangularRigidLink
    End If
    
    If CommonUtils.IsSelectedInExcel(AppConst.TOP_SECTION) Then
        ApplyRigidLinkToTargetNode bar.StartNode, rigidLinkLabel, rectangularRigidLink
    End If
End Sub

Private Sub ConfigureRectangularRigidLink(ByVal sectionData As IRobotBarSectionData, ByVal bar As IRobotBar, ByRef rectangularRigidLink As RectangularColumnRigidLink)
    With rectangularRigidLink
        .gammaAngle = bar.Gamma
        .sectionWidth = sectionData.GetValue(IRobotBarSectionDataValue.I_BSDV_BF)
        .sectionHeight = sectionData.GetValue(IRobotBarSectionDataValue.I_BSDV_D)
        .angleOfRotation = MathUtility.ConvertToRadians(.gammaAngle - 90)
        .radiusBX = .sectionWidth / 2
        .radiusHy = .sectionHeight / 2
        .radiusDiagonal = Sqr(.radiusBX ^ 2 + .radiusHy ^ 2)
        .diagonalAngle = Atn(.radiusHy / .radiusBX)
    End With
End Sub

Private Sub SetupAndApplyCircularRigidLink(ByVal rigidLinkLabel As String, ByVal sectionData As IRobotBarSectionData, ByVal bar As IRobotBar)
    Dim circularRigidLink As CircularColumnRigidLink
    Set circularRigidLink = New CircularColumnRigidLink
    circularRigidLink.Diameter = sectionData.GetValue(IRobotBarSectionConcreteDataValue.I_BSCDV_COL_N) * 2
    
    If CommonUtils.IsSelectedInExcel(AppConst.TOP_SECTION) Then
        ApplyRigidLinkToTargetNode bar.EndNode, rigidLinkLabel, circularRigidLink
    End If
    
    If CommonUtils.IsSelectedInExcel(AppConst.TOP_SECTION) Then
        ApplyRigidLinkToTargetNode bar.StartNode, rigidLinkLabel, circularRigidLink
    End If
End Sub

Private Sub ApplyRigidLinkToTargetNode(ByVal targetNode As Long, ByVal rigidLinkLabel As String, ByRef rigidLink As INodeFinder)
    rigidLink.centerNode = targetNode
    rigidLink.FindListOfNodes
    RobotAPIHelper.RigidiLinkServer.Set rigidLink.centerNode, rigidLink.nodesListString, rigidLinkLabel
End Sub

Sub RedoRigidLinks()
    DeleteRigidLinkManager.DeleteForSelectedColumns
    CreateRigidLinksForBarsPlane
End Sub

