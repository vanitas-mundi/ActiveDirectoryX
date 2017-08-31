Option Explicit On
Option Infer On
Option Strict On


#Region " --------------->> Imports "

Imports System.DirectoryServices
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core.Manipulation
Imports SSP.ActiveDirectoryX.Core

#End Region

Namespace Grants.Administration

	Public Class Administrations

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private Shared _instance As Administrations
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Shared Sub New()
			_instance = New Administrations
		End Sub

		Private Sub New()
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		Public Shared ReadOnly Property Instance As Administrations
		Get
			Return _instance
		End Get
		End Property

		''' <summary>
		''' Stellt Funktionen zur Administration von Rollen zur Verfügung.
		''' </summary>
		Public ReadOnly Property AdministrationGroups As AdministrationGroupsAdministration
		Get
			Return AdministrationGroupsAdministration.Instance
		End Get
		End Property

		''' <summary>
		''' Stellt Funktionen zur Administration von Berechtigungsgruppen zur Verfügung.
		''' </summary>
		Public ReadOnly Property GrantGroups As GrantGroupsAdministration
		Get
			Return GrantGroupsAdministration.Instance
		End Get
		End Property

		''' <summary>
		''' Stellt Funktionen zur Administration von Organisationsgruppen zur Verfügung.
		''' </summary>
		Public ReadOnly Property OrganizationGroups As OrganizationGroupsAdministration
		Get
			Return OrganizationGroupsAdministration.Instance
		End Get
		End Property

		''' <summary>
		''' Stellt Funktionen zur Administration von Usern zur Verfügung.
		''' </summary>
		Public ReadOnly Property GrantUsers As GrantUsersAdministration
		Get
			Return GrantUsersAdministration.Instance
		End Get
		End Property

		''' <summary>
		''' Stellt Funktionen zur Administration von Administrationsorganisationseinheiten zur Verfügung.
		''' </summary>
		Public ReadOnly Property AdministrationOrganizationalUnits As OrganizationalUnitsAdministration
		Get
			Return OrganizationalUnitsAdministration.Instance
		End Get
		End Property

		''' <summary>
		''' Stellt Funktionen zum Auflösen von Administrationstypen zur Verfügung.
		''' </summary>
		Public ReadOnly Property TypeResolver As AdministrationTypeResolver
		Get
			Return AdministrationTypeResolver.Instance
		End Get
		End Property
#End Region	'{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region	'{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region	'{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "

#Region " --> CreateGroup "
	''' <summary>
	''' Erstellt eine neue Domänen-Gruppe.
	''' </summary>
	Public Function CreateGroup(ByVal parentOrganizationalUnitName As String _
	, ByVal groupName As String) As AdManipulationResults

		Return CreateGroup(DistinguishedName.GetByOu _
		(parentOrganizationalUnitName), groupName)
	End Function

	''' <summary>
	''' Erstellt eine neue Domänen-Gruppe.
	''' </summary>
	Public Function CreateGroup(ByVal parentOrganizationalUnitDn As DistinguishedName _
	, ByVal groupName As String) As AdManipulationResults

		Return CreateGroup(parentOrganizationalUnitDn, groupName, SpecialDistinguishedNameKeys.Domain)
	End Function

	''' <summary>
	''' Erstellt eine neue Domänen-Gruppe und prüft zuvor, ob Sie vom Typ groupTypDn ist.
	''' </summary>
	Public Function CreateGroup(ByVal parentOrganizationalUnitDn As DistinguishedName _
	, ByVal groupName As String, ByVal groupTypDn As DistinguishedName) As AdManipulationResults

		Select Case True
		Case (parentOrganizationalUnitDn Is Nothing) OrElse (Not parentOrganizationalUnitDn.IsOrganizationalUnit)
			Return AdManipulationResults.GroupIsNotOrganizationUnit
		Case Not parentOrganizationalUnitDn.ContainsDn(groupTypDn)
			Select Case True
			Case groupTypDn.Equals(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Roles))
				Return AdManipulationResults.GroupIsNotRole
			Case groupTypDn.Equals(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.OrganizationGroups))
				Return AdManipulationResults.GroupIsNotOrganizationGroup
			Case groupTypDn.Equals(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Administration))
				Return AdManipulationResults.GroupIsNotAdminstrationGroup
			Case Else
				Return AdManipulationResults.GroupIsNotDomainGroup
			End Select
		Case Else
			Try
				AdGroups.CreateGroup(parentOrganizationalUnitDn, groupName, True)
				Return AdManipulationResults.Successful
			Catch ex As DirectoryServicesCOMException
				Return AdManipulationResults.MemberAlreadyExist
			Catch ex As UnauthorizedAccessException
				Return AdManipulationResults.AccesDenied
			Catch ex As System.Exception
				Return AdManipulationResults.UnknownError
			End Try
		End Select
	End Function

	''' <summary>
	''' Erstellt eine neue Domänen-Gruppe und prüft zuvor, ob Sie vom Typ specialDistinguishedName ist.
	''' </summary>
	Public Function CreateGroup(ByVal parentOrganizationalUnitDn As DistinguishedName _
	, ByVal groupName As String, ByVal specialDistinguishedName As SpecialDistinguishedNameKeys) As AdManipulationResults

		Return CreateGroup(parentOrganizationalUnitDn, groupName, SpecialDistinguishedNames.Item(specialDistinguishedName))
	End Function

#End Region

#Region " --> DeleteGroup "
	''' <summary>
	''' Löscht die angegebene Domänen-Gruppe.
	''' </summary>
	Public Function DeleteGroup(ByVal groupName As String) As AdManipulationResults
		Return DeleteGroup(DistinguishedName.GetByGroupName(groupName))
	End Function

	''' <summary>
	''' Löscht die angegebene Domänen-Gruppe.
	''' </summary>
	Public Function DeleteGroup(ByVal groupDn As DistinguishedName) As AdManipulationResults

		Return DeleteGroup(groupDn, SpecialDistinguishedNameKeys.Domain)
	End Function

	''' <summary>
	''' Löscht die angegebene Domänen-Gruppe und prüft zuvor, ob es sich um eine Gruppe vom Typ groupTypDn handelt.
	''' </summary>
	Public Function DeleteGroup(ByVal groupDn As DistinguishedName, ByVal groupTypDn As DistinguishedName) As AdManipulationResults

		Select Case True
		Case groupDn Is Nothing
			Return AdManipulationResults.MemberNotExist
		Case Not groupDn.IsGroup
			Return AdManipulationResults.IsNotGroup
		Case Not groupDn.ContainsDn(groupTypDn)
			Select Case True
			Case groupTypDn.Equals(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Mappings))
				Return AdManipulationResults.GroupIsNotMapping
			Case groupTypDn.Equals(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Roles))
				Return AdManipulationResults.GroupIsNotRole
			Case groupTypDn.Equals(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.OrganizationGroups))
				Return AdManipulationResults.GroupIsNotOrganizationGroup
			Case groupTypDn.Equals(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Administration))
				Return AdManipulationResults.GroupIsNotAdminstrationGroup
			Case Else
				Return AdManipulationResults.GroupIsNotDomainGroup
			End Select
		Case Else
			Try
				AdGroups.DeleteGroup(groupDn, True)
				Return AdManipulationResults.Successful
			Catch ex As UnauthorizedAccessException
				Return AdManipulationResults.AccesDenied
			Catch ex As System.Exception
				Return AdManipulationResults.UnknownError
			End Try
		End Select
	End Function

	''' <summary>
	''' Löscht die angegebene Domänen-Gruppe und prüft zuvor, ob es sich um eine Gruppe vom Typ specialDistinguishedName handelt.
	''' </summary>
	Public Function DeleteGroup(ByVal groupDn As DistinguishedName _
	, ByVal specialDistinguishedName As SpecialDistinguishedNameKeys) As AdManipulationResults

		Return DeleteGroup(groupDn, SpecialDistinguishedNames.Item(specialDistinguishedName))
	End Function
#End Region

#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace



