Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
#End Region

Namespace Core

	Public Class AdNodes

		Inherits List(Of AdNode)

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private _parent As Adnode
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Friend Sub New(ByVal parent As Adnode)
			_parent = parent
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		Friend Overloads ReadOnly Property Item _
		(ByVal distinguishedName As DistinguishedName) As Adnode
		Get
			Return Me.Item(distinguishedName.Value)
		End Get
		End Property

		Friend Overloads ReadOnly Property Item _
		(ByVal distinguishedName As String) As Adnode
		Get
			Return Me.Where(Function(n) n.DistinguishedName.Value = distinguishedName).FirstOrDefault
		End Get
		End Property
#End Region	'{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region	'{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region	'{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
		Public Shadows Function Add(ByVal grantNode As Adnode) As Adnode
			grantNode.SetParent(_parent)
			MyBase.Add(grantNode)
			Return grantNode
		End Function

    Public Shadows Function Add(ByVal distinguishedName As DistinguishedName) As AdNode
      Dim n = New AdNode(distinguishedName, _parent)
      MyBase.Add(n)
      Return n
    End Function

    Public Shadows Sub AddRange(ByVal grantNodes As IEnumerable(Of Adnode))
			grantNodes.ToList.ForEach(Sub(node) node.SetParent(_parent))
			MyBase.AddRange(grantNodes)
		End Sub

    Public Shadows Sub AddRange(ByVal distinguishedNames As IEnumerable(Of DistinguishedName))
      MyBase.AddRange(distinguishedNames.Select(Function(s) New AdNode(s, _parent)))
    End Sub

    Public Shadows Sub RemoveByDistinguishedName(ByVal distinguishedName As String)
			Dim node = Me.Item(distinguishedName)
			If node Is Nothing Then Return
			Me.Remove(node)
		End Sub

		Public Shadows Sub RemoveByDistinguishedName(ByVal distinguishedName As DistinguishedName)
			RemoveByDistinguishedName(distinguishedName.Value)
		End Sub
#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace
