Attribute VB_Name = "BarLinearRigidLinkManager"
'@Folder "RSAPRigidLinks.Manager"
'BarLinearRigidLinkManager
Option Explicit

Sub CreateRigidLinksForBarsLine()
    RobotAPIHelper.PrepareEnvironmentWithBarSelection
    
    Dim i As Long
    PrintColumnCounter.PrintCounters RobotAPIHelper.barSelection.count
    For i = 1 To RobotAPIHelper.barSelection.count
        Dim sectionData As IRobotBarSectionData
        Dim bar As IRobotBar
        GetSectionData RobotAPIHelper.barServer, RobotAPIHelper.barSelection, i, bar, sectionData
        
        If sectionData.ShapeType = I_BSST_CONCR_COL_R Or _
        sectionData.ShapeType = I_BSST_CONCR_BEAM_RECT Then
            MakeLinearRigidLink RobotAPIHelper.G_FIXED_RIGID_LINK, sectionData, bar
        End If
        
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

Private Sub MakeLinearRigidLink(ByVal rigidLinkLabel As String, ByVal sectionData As IRobotBarSectionData, ByVal bar As IRobotBar)
    Dim LinearRigidLink As LinearRigidLink
    Set LinearRigidLink = New LinearRigidLink
    FillLinearSectionData sectionData, bar, LinearRigidLink
    
    If CommonUtils.IsSelectedInExcel(AppConst.START_SECTION) Then
        ApplyLinearRigidLinkAtTargetNode bar.StartNode, rigidLinkLabel, LinearRigidLink
    End If
    
    If CommonUtils.IsSelectedInExcel(AppConst.END_SECTION) Then
        ApplyLinearRigidLinkAtTargetNode bar.EndNode, rigidLinkLabel, LinearRigidLink
    End If
End Sub

Private Sub FillLinearSectionData(ByVal sectionData As IRobotBarSectionData, ByVal bar As IRobotBar, ByRef LinearRigidLink As LinearRigidLink)
    With LinearRigidLink
        .sectionWidth = sectionData.GetValue(IRobotBarSectionDataValue.I_BSDV_BF)
        .sectionHeight = sectionData.GetValue(IRobotBarSectionDataValue.I_BSDV_D)
        .gammaAngle = bar.Gamma
        .meshSize = Range(MESH_SIZE).value
        .directionOfBar = Range(DIRECTION).value
        .VerifyIsColumn RobotAPIHelper.RobotNodesServer.Get(bar.StartNode).Z, RobotAPIHelper.RobotNodesServer.Get(bar.EndNode).Z
    End With
End Sub

Private Sub ApplyLinearRigidLinkAtTargetNode(ByVal targetNode As Long, ByVal rigidLinkLabel As String, ByVal LinearRigidLink As LinearRigidLink)
    LinearRigidLink.centerNode = targetNode
    LinearRigidLink.INodeFinder_FindListOfNodes
    RobotAPIHelper.RigidiLinkServer.Set LinearRigidLink.centerNode, LinearRigidLink.nodesListString, rigidLinkLabel
End Sub

Sub RedoRigidLinks()
    DeleteRigidLinkManager.DeleteForSelectedColumns
    CreateRigidLinksForBarsLine
End Sub

