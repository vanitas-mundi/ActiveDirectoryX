Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core
#End Region

Namespace Grants.Administration

	Public Class GrantGroupsAdministration

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private Shared _instance As GrantGroupsAdministration
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Shared Sub New()
			_instance = New GrantGroupsAdministration
		End Sub

		Private Sub New()
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		Public Shared ReadOnly Property Instance As GrantGroupsAdministration
		Get
			Return _instance
		End Get
		End Property

		''' <summary>
		''' Stellt Funktionen zur Administration von Berechtigungtabellen (Applikationen) zur Verfügung.
		''' </summary>
		Public ReadOnly Property GrantTables As GrantTablesAdministration
		Get
			Return GrantTablesAdministration.Instance
		End Get
		End Property

		''' <summary>
		''' Stellt Funktionen zur Administration von Berechtigungen zur Verfügung.
		''' </summary>
		Public ReadOnly Property Grants As GrantsAdministration
		Get
			Return GrantsAdministration.Instance
		End Get
		End Property


		''' <summary>
		''' Stellt Funktionen zur Administration von Rollen zur Verfügung.
		''' </summary>
		Public ReadOnly Property Roles As RolesAdministration
		Get
			Return RolesAdministration.Instance
		End Get
		End Property

		''' <summary>
		''' Stellt Funktionen zur Administration von Mappings zur Verfügung.
		''' </summary>
		Public ReadOnly Property Mappings As MappingsAdministration
		Get
			Return MappingsAdministration.Instance
		End Get
		End Property

		''' <summary>
		''' Stellt Funktionen zur Administration von GroupManagern zur Verfügung.
		''' </summary>
		Public ReadOnly Property GroupManagers As GroupManagersAdministration
		Get
			Return GroupManagersAdministration.Instance
		End Get
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


