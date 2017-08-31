Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core
#End Region

Namespace Grants.Administration

	Public Class OrganizationGroupsAdministration

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private Shared _instance As OrganizationGroupsAdministration
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Shared Sub New()
			_instance = New OrganizationGroupsAdministration
		End Sub

		Private Sub New()
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		Public Shared ReadOnly Property Instance As OrganizationGroupsAdministration
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

#Region " --> CreateOrganizationGroup "
	''' <summary>
	''' Legt eine neue Organisationsgruppe an.
	''' </summary>
	Public Function CreateOrganizationGroup(ByVal parentOrganizationalUnitName As String _
	, ByVal groupName As String) As AdManipulationResults

		Return CreateOrganizationGroup(DistinguishedName.GetByOu(parentOrganizationalUnitName), groupName)
	End Function

	''' <summary>
	''' Legt eine neue Organisationsgruppe an.
	''' </summary>
	Public Function CreateOrganizationGroup(ByVal parentOrganizationalUnitDn As DistinguishedName _
	, ByVal groupName As String) As AdManipulationResults

		Return Administrations.Instance.CreateGroup _
		(parentOrganizationalUnitDn, groupName, SpecialDistinguishedNameKeys.OrganizationGroups)
	End Function

	''' <summary>
	''' Legt eine neue Organisationsgruppe, vom Typ specialDistinguishedName, an.
	''' </summary>
	Public Function CreateOrganizationGroup(ByVal parentOrganizationalUnitDn As DistinguishedName _
	, ByVal groupName As String, ByVal specialDistinguishedName As SpecialDistinguishedNameKeys) As AdManipulationResults

		Select Case True
		Case Not parentOrganizationalUnitDn.ContainsDn(SpecialDistinguishedNames.Item _
		(SpecialDistinguishedNameKeys.OrganizationGroups))
			Return AdManipulationResults.GroupIsNotOrganizationGroup
		Case Not parentOrganizationalUnitDn.ContainsDn(SpecialDistinguishedNames.Item(specialDistinguishedName))
			Return AdManipulationResults.InvalidOrganizationGroupType
		Case Else
			Return Administrations.Instance.CreateGroup _
			(parentOrganizationalUnitDn, groupName, specialDistinguishedName)
		End Select
	End Function
#End Region

#Region " --> DeleteOrganizationGroup "
	''' <summary>
	''' Löscht eine Organisationsgruppe.
	''' </summary>
	Public Function DeleteOrganizationGroup(ByVal groupName As String) As AdManipulationResults

		Return DeleteOrganizationGroup(DistinguishedName.GetByGroupName(groupName))
	End Function

	''' <summary>
	''' Löscht eine Organisationsgruppe.
	''' </summary>
	Public Function DeleteOrganizationGroup(ByVal grouptDn As DistinguishedName) As AdManipulationResults

		Return Administrations.Instance.DeleteGroup(grouptDn, SpecialDistinguishedNameKeys.OrganizationGroups)
	End Function

	''' <summary>
	''' Löscht eine Organisationsgruppe vom Typ specialDistinguishedName.
	''' </summary>
	Public Function DeleteOrganizationGroup(ByVal grouptDn As DistinguishedName _
	, ByVal specialDistinguishedName As SpecialDistinguishedNameKeys) As AdManipulationResults

		Select Case True
		Case Not grouptDn.ContainsDn(SpecialDistinguishedNames.Item _
		(SpecialDistinguishedNameKeys.OrganizationGroups))
			Return AdManipulationResults.GroupIsNotOrganizationGroup
		Case Not grouptDn.ContainsDn(SpecialDistinguishedNames.Item(specialDistinguishedName))
			Return AdManipulationResults.InvalidOrganizationGroupType
		Case Else
			Return Administrations.Instance.DeleteGroup _
			(grouptDn, SpecialDistinguishedNameKeys.OrganizationGroups)
		End Select
	End Function

#End Region

#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace


