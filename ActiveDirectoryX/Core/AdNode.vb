Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
#End Region

Namespace Core

	Public Class AdNode

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private _distinguishedName As DistinguishedName
		Private _parent As AdNode
		Private _nodes As New AdNodes(Me)
#End Region  '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
    'Friend Sub New(ByVal distinguishedName As String)
    '	_distinguishedName = New DistinguishedName(distinguishedName)
    'End Sub

    'Friend Sub New(ByVal distinguishedName As String, ByVal parent As AdNode)
    '	_distinguishedName = New DistinguishedName(distinguishedName)
    '	_parent = parent
    'End Sub

    'Friend Sub New(ByVal distinguishedName As DistinguishedName)
    '	Me.New(distinguishedName.Value)
    'End Sub

    'Friend Sub New(ByVal distinguishedName As DistinguishedName, ByVal parent As AdNode)
    '	Me.New(distinguishedName.Value, parent)
    'End Sub

    Friend Sub New(ByVal distinguishedName As DistinguishedName)
      _distinguishedName = distinguishedName
    End Sub

    Friend Sub New(ByVal distinguishedName As DistinguishedName, ByVal parent As AdNode)
      _distinguishedName = distinguishedName
      _parent = parent
    End Sub
#End Region  '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
    ''' <summary>
    ''' Liefert alle untergeordneten Nodes.
    ''' </summary>
    Public ReadOnly Property Nodes As AdNodes 
		Get
			Return _nodes
		End Get
		End Property

		''' <summary>
		''' Liefert den übergeordneten Nodes.
		''' </summary>
		Public ReadOnly Property Parent As AdNode 
		Get
			Return _parent
		End Get
		End Property

		''' <summary>
		''' Liefert den zugrunde liegenden DistinguishedName des Nodes.
		''' </summary>
		Public ReadOnly Property DistinguishedName As DistinguishedName 
		Get
			Return _distinguishedName
		End Get
		End Property
#End Region	'{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region	'{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region	'{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
		Friend Sub SetParent(ByVal value As AdNode)
			_parent = value
		End Sub

		Public Overrides Function ToString() As String
			Return Me.DistinguishedName.Value
		End Function
#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace
