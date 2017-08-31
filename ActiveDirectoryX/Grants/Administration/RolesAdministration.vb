Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core
#End Region

Namespace Grants.Administration

	Public Class RolesAdministration

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private Shared _instance As RolesAdministration
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Shared Sub New()
			_instance = New RolesAdministration
		End Sub

		Private Sub New()
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		Public Shared ReadOnly Property Instance As RolesAdministration
		Get
			Return _instance
		End Get
		End Property
#End Region	'{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region	'{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region	'{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "

#Region " --> CreateRole "
	''' <summary>
	''' Erstellt eine neue Rolle.
	''' </summary>
	Public Function CreateRole(ByVal parentOrganizationalUnitName As String _
	, ByVal groupName As String) As AdManipulationResults

		Return CreateRole(DistinguishedName.GetByOu(parentOrganizationalUnitName), groupName)
	End Function

	''' <summary>
	''' Erstellt eine neue Rolle.
	''' </summary>
	Public Function CreateRole(ByVal parentOrganizationalUnitDn As DistinguishedName _
	, ByVal groupName As String) As AdManipulationResults

		Return Administrations.Instance.CreateGroup _
		(parentOrganizationalUnitDn, groupName, SpecialDistinguishedNameKeys.Roles)
	End Function
#End Region

#Region " --> DeleteRole "
	''' <summary>
	''' Löscht eine Rolle.
	''' </summary>
	Public Function DeleteRole(ByVal groupName As String) As AdManipulationResults

		Return DeleteRole(DistinguishedName.GetByGroupName(groupName))
	End Function

	''' <summary>
	''' Löscht eine Rolle.
	''' </summary>
	Public Function DeleteRole(ByVal grouptDn As DistinguishedName) As AdManipulationResults

		Return Administrations.Instance.DeleteGroup(grouptDn, SpecialDistinguishedNameKeys.Roles)
	End Function
#End Region

#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace


