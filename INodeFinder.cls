VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INodeFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "RSAPRigidLinks.Classes"
' INodeFinder Interface
Option Explicit

' Method to find and populate the list of nodes
Public Sub FindListOfNodes()
End Sub

' Property to get or set the list of nodes as a Collection
Public Property Get nodesListString() As String
End Property

Public Property Let nodesListString(ByVal value As String)
End Property

' Property for the center node, required for positioning
Public Property Get centerNode() As Double
End Property

Public Property Let centerNode(ByVal value As Double)
End Property
