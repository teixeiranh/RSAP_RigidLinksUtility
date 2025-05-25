VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinearRigidLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "RSAPRigidLinks.Classes"
Option Explicit

Implements INodeFinder

Private m_centerNode As Long
Private m_nodesListString As String
Public sectionWidth As Double
Public sectionHeight As Double
Public usedNumberOfNodes As Long
Public gammaAngle As Double
Public isColumn As Boolean
Public meshSize As Double
Public directionOfBar As String

Private nodeListCollection As Collection

Public Property Get centerNode() As Double
    centerNode = m_centerNode
End Property

Public Property Let centerNode(ByVal value As Double)
    m_centerNode = value
End Property

Public Property Get nodesListString() As String
    nodesListString = m_nodesListString
End Property

Public Property Let nodesListString(ByVal value As String)
    m_nodesListString = value
End Property

Public Property Get INodeFinder_centerNode() As Double
    INodeFinder_centerNode = m_centerNode
End Property

Public Property Let INodeFinder_centerNode(ByVal value As Double)
    m_centerNode = value
End Property

Public Property Get INodeFinder_nodesListString() As String
    INodeFinder_nodesListString = m_nodesListString
End Property

Public Property Let INodeFinder_nodesListString(ByVal value As String)
    m_nodesListString = value
End Property

Public Sub VerifyIsColumn(startNodeZ As Double, endNodeZ As Double)
    isColumn = (startNodeZ <> endNodeZ)
End Sub

Private Sub AddNode(ByVal nodeValue As Double)
    nodeListCollection.Add nodeValue
End Sub

Public Function GetNode(ByVal index As Long) As Double
    If index > 0 And index <= nodeListCollection.count Then
        GetNode = nodeListCollection(index)
    Else
        Err.Raise vbObjectError + 513, "GetNode", "Index out of bounds"
    End If
End Function

Public Function NodeCount() As Long
    NodeCount = nodeListCollection.count
End Function

Public Sub INodeFinder_FindListOfNodes()
    Set nodeListCollection = New Collection
    
    Dim isMainDirection As Boolean
    isMainDirection = (directionOfBar = "Z")
    Dim calculatedNumberOfNodes As Double
    calculatedNumberOfNodes = WorksheetFunction.Ceiling(IIf(isMainDirection, sectionHeight, sectionWidth) / meshSize, 1)
    If calculatedNumberOfNodes Mod 2 = 0 Then calculatedNumberOfNodes = calculatedNumberOfNodes + 1
    usedNumberOfNodes = Application.WorksheetFunction.Max(3, calculatedNumberOfNodes)
    Dim nodeSpacing As Double
    nodeSpacing = IIf(isMainDirection, sectionHeight, sectionWidth) / (usedNumberOfNodes - 1)
    Dim halfNodes As Long
    halfNodes = Int((usedNumberOfNodes - 1) / 2)
    
    Dim nodeNumber As Double
    MakeSlaveNodes nodeSpacing, halfNodes, nodeNumber, isMainDirection, -1
    MakeSlaveNodes nodeSpacing, halfNodes, nodeNumber, isMainDirection, 1

    Dim nodesArray() As String
    ReDim nodesArray(1 To nodeListCollection.count)
    
    Dim i As Long
    For i = 1 To nodeListCollection.count
        nodesArray(i) = CStr(nodeListCollection(i))
    Next i

    m_nodesListString = Join(nodesArray, " ")
End Sub

Private Sub MakeSlaveNodes(ByVal nodeSpacing As Double, ByVal halfOfNodes As Long, ByRef nodeNumber As Double, isMainDirection As Boolean, ByVal directionMultiplier As Long)
    Dim angle As Double
    If isMainDirection Then
        angle = MathUtility.ConvertToRadians(gammaAngle)
    Else
        angle = MathUtility.ConvertToRadians(gammaAngle + AppConst.PI_VALUE / 2)
    End If
    
    Dim i As Long
    For i = 1 To halfOfNodes
        nodeNumber = RobotAPIHelper.RobotNodesServer.FreeNumber
        
        If Not isColumn Then
            RobotAPIHelper.RobotNodesServer.Create nodeNumber, _
                                                   RobotAPIHelper.RobotNodesServer.Get(m_centerNode).X, _
                                                   RobotAPIHelper.RobotNodesServer.Get(m_centerNode).Y, _
                                                   RobotAPIHelper.RobotNodesServer.Get(m_centerNode).Z + directionMultiplier * nodeSpacing * i
        Else
            RobotAPIHelper.RobotNodesServer.Create nodeNumber, _
                                                   RobotAPIHelper.RobotNodesServer.Get(m_centerNode).X + directionMultiplier * nodeSpacing * i * Cos(angle), _
                                                   RobotAPIHelper.RobotNodesServer.Get(m_centerNode).Y + directionMultiplier * nodeSpacing * i * Sin(angle), _
                                                   RobotAPIHelper.RobotNodesServer.Get(m_centerNode).Z
            
        End If
        nodeListCollection.Add nodeNumber
    Next i
End Sub


