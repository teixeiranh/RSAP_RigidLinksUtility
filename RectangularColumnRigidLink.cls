VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RectangularColumnRigidLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "RSAPRigidLinks.Classes"
'@ModuleDescription "Class representing the rigid link of a rectangular section (8 nodes)."

Option Explicit

Implements INodeFinder

Public radiusBX As Double
Public angleOfRotation As Double
Public radiusDiagonal As Double
Public diagonalAngle As Double
Public radiusHy As Double
Public gammaAngle As Double
Public sectionWidth As Double
Public sectionHeight As Double

Public iNode1 As Double
Public iNode2 As Double
Public iNode3 As Double
Public iNode4 As Double
Public iNode5 As Double
Public iNode6 As Double
Public iNode7 As Double
Public iNode8 As Double

Private m_centerNode As Double
 
Private m_nodesListString As String

' Getter for m_centerNode
Public Property Get centerNode() As Double
    centerNode = m_centerNode
End Property

' Setter for m_centerNode
Public Property Let centerNode(ByVal value As Double)
    m_centerNode = value
End Property

' Implement centerNode property for INodeFinder
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

Public Sub INodeFinder_FindListOfNodes()
    iNode1 = CreateNode(radiusBX, angleOfRotation)
    iNode2 = CreateNode(radiusDiagonal, angleOfRotation + diagonalAngle)
    iNode3 = CreateNode(radiusHy, AppConst.PI_VALUE / 2 + angleOfRotation)
    iNode4 = CreateNode(radiusDiagonal, AppConst.PI_VALUE - diagonalAngle + angleOfRotation)
    iNode5 = CreateNode(radiusBX, AppConst.PI_VALUE + angleOfRotation)
    iNode6 = CreateNode(radiusDiagonal, AppConst.PI_VALUE + angleOfRotation + diagonalAngle)
    iNode7 = CreateNode(radiusHy, 3 * AppConst.PI_VALUE / 2 + angleOfRotation)
    iNode8 = CreateNode(radiusDiagonal, 2 * AppConst.PI_VALUE + angleOfRotation - diagonalAngle)
    
    m_nodesListString = Join(Array(iNode1, iNode2, iNode3, iNode4, iNode5, iNode6, iNode7, iNode8), " ")
End Sub

Private Function CreateNode(ByVal radius As Double, ByVal angle As Double) As Double
    Dim nodeNumber As Double
    nodeNumber = RobotAPIHelper.RobotNodesServer.FreeNumber
    RobotAPIHelper.RobotNodesServer.Create nodeNumber, _
                      RobotAPIHelper.RobotNodesServer.Get(m_centerNode).X + radius * Cos(angle), _
                      RobotAPIHelper.RobotNodesServer.Get(m_centerNode).Y + radius * Sin(angle), _
                      RobotAPIHelper.RobotNodesServer.Get(m_centerNode).Z
    CreateNode = nodeNumber
End Function


