VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CircularColumnRigidLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Class representing the rigid link of a circular section (8 nodes). "
'@Folder "RSAPRigidLinks.Classes"
'CircularColumnRigidLink
Option Explicit

Implements INodeFinder

Public Diameter As Double

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
    iNode1 = CreateNode(0)
    iNode2 = CreateNode(AppConst.PI_VALUE / 4)
    iNode3 = CreateNode(AppConst.PI_VALUE / 2)
    iNode4 = CreateNode(3 * AppConst.PI_VALUE / 4)
    iNode5 = CreateNode(AppConst.PI_VALUE)
    iNode6 = CreateNode(5 * AppConst.PI_VALUE / 4)
    iNode7 = CreateNode(3 * AppConst.PI_VALUE / 2)
    iNode8 = CreateNode(7 * AppConst.PI_VALUE / 4)
    
    m_nodesListString = Join(Array(iNode1, iNode2, iNode3, iNode4, iNode5, iNode6, iNode7, iNode8), " ")
End Sub

Private Function CreateNode(ByVal angle As Double) As Double
    Dim nodeNumber As Double
    Dim radius As Double
    radius = Diameter / 2
    nodeNumber = RobotAPIHelper.RobotNodesServer.FreeNumber
    RobotAPIHelper.RobotNodesServer.Create nodeNumber, _
                      RobotAPIHelper.RobotNodesServer.Get(m_centerNode).X + Cos(angle) * radius, _
                      RobotAPIHelper.RobotNodesServer.Get(m_centerNode).Y + Sin(angle) * radius, _
                      RobotAPIHelper.RobotNodesServer.Get(m_centerNode).Z
    CreateNode = nodeNumber
End Function


