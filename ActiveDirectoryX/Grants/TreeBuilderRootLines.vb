Option Explicit On
Option Infer On
 Option Strict On

#Region " --------------->> Imports/ usings "
Imports SSP.ActiveDirectoryX.Grants.Enums
Imports SSP.ActiveDirectoryX.Grants.Exceptions
#End Region

Namespace Grants

	Friend Class TreeBuilderRootLines

		'┌───┬───┐
		'│   │   │
		'├───┼───┤
		'│   │   │
		'└───┴───┘

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private _regularElement As String = " ├─ "
		Private _endElement As String = " └─ "
		Private _parentLine As String = " │ "
		Private _whiteSpace As String = "   "
		Private _treeBuilderRootLineType As TreeBuilderRootLineTypes
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
	Public Sub New()
		Me.New(TreeBuilderRootLineTypes.Regular)
	End Sub

	Public Sub New(treeBuilderRootLineType As TreeBuilderRootLineTypes)
		_treeBuilderRootLineType = treeBuilderRootLineType
		Select Case _treeBuilderRootLineType
		Case TreeBuilderRootLineTypes.None
			_regularElement = ""
			_endElement = ""
			_parentLine = ""
			_whiteSpace = ""
		Case TreeBuilderRootLineTypes.Tabs
			_regularElement = vbTab
			_endElement = vbTab
			_parentLine = vbTab
			_whiteSpace = vbTab
		Case Else	'TreeBuilderRootLineTypes.Regular
			_regularElement = " ├─ "
			_endElement = " └─ "
			_parentLine = " │ "
			_whiteSpace = "   "
		End Select
	End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		Public Property RegularElement As String
		Get
			Return _regularElement
		End Get
		Set(value As String)
			If Not _treeBuilderRootLineType = TreeBuilderRootLineTypes.Custom Then
				Throw New TreeBuilderRootLinesNotCustomType
			End If
			_regularElement = value
		End Set
		End Property

		Public Property EndElement As String
		Get
			Return _endElement
		End Get
		Set(value As String)
			If Not _treeBuilderRootLineType = TreeBuilderRootLineTypes.Custom Then
				Throw New TreeBuilderRootLinesNotCustomType
			End If
			_endElement = value
		End Set
		End Property

		Public Property ParentLine As String
		Get
			Return _parentLine
		End Get
		Set(value As String)
			If Not _treeBuilderRootLineType = TreeBuilderRootLineTypes.Custom Then
				Throw New TreeBuilderRootLinesNotCustomType
			End If
			_parentLine = value
		End Set
		End Property

		Public Property WhiteSpace As String
		Get
			Return _whiteSpace
		End Get
		Set(value As String)
			If Not _treeBuilderRootLineType = TreeBuilderRootLineTypes.Custom Then
				Throw New TreeBuilderRootLinesNotCustomType
			End If
			_whiteSpace = value
		End Set
		End Property
#End Region	'{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region	'{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region	'{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace
