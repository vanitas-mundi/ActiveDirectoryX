Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports/ usings "
Imports SSP.ActiveDirectoryX.Core
Imports SSP.ActiveDirectoryX.Grants.Enums
#End Region

Namespace Grants.Interfaces

  Public Interface IGrantsManagerBase(Of TGrantNamesEnum As Structure)
		ReadOnly Property Routines As GrantsBaseRoutines
    ReadOnly Property ApplicationGrants() As GrantTable
    ReadOnly Property LoginUserPersonId() As Int64
		ReadOnly Property GrantUser As GrantUser

		Function IsGranted(ByVal grantName As TGrantNamesEnum) As Boolean
		Function GetGroupManagerDistinguishedNamesByManagerPersonId(ByVal managerPersonId As Int64) As DistinguishedName()
		Function GetGroupManagerDistinguishedNamesByManagerPersonId(ByVal managerPersonId As Int64, ByVal includeDeputies As Boolean) As DistinguishedName()
		Function GetGrantAssignedUsers(ByVal grantName As TGrantNamesEnum) As DistinguishedName()
		Function GetAssignedUsersOfGroupManager(ByVal groupManagerDn As DistinguishedName) As DistinguishedName()
		Function GetAssignedAdministrationGroupsOfGroupManager(ByVal groupManagerDn As DistinguishedName) As AdministrationGroup()
		Function GetAssignedAdministrationGroupsByManagerPersonId(ByVal managerPersonId As Int64) As AdministrationGroup()
		Function GetAssignedAdministrationGroupsByManagerPersonId(ByVal managerPersonId As Int64, ByVal includeDeputies As Boolean) As AdministrationGroup()
		Function GetManagedUsersOfManager(ByVal managerPersonId As Int64) As DistinguishedName()
		Function GetManagedUsersOfManager(ByVal managerPersonId As Int64, ByVal includeDeputies As Boolean) As DistinguishedName()
		Function IsLoginUserManagerOf(ByVal groupType As OrganizationGroupTypes, ByVal personId As Int64) As Boolean
		Function IsLoginUserManagerOf(ByVal groupType As OrganizationGroupTypes, ByVal personId As Int64, ByVal includeDeputies As Boolean) As Boolean
		Function GetManagerDistinguishedNamesOfGrantUser(ByVal groupType As OrganizationGroupTypes) As DistinguishedName()
		Function GetManagerDistinguishedNamesOfGrantUser(ByVal groupType As OrganizationGroupTypes, ByVal includeDeputies As Boolean) As DistinguishedName()
		Function GetManagerDistinguishedNamesOfPersonId(ByVal personId As Int64, ByVal groupType As OrganizationGroupTypes) As DistinguishedName()
		Function GetManagerDistinguishedNamesOfPersonId(ByVal personId As Int64, ByVal groupType As OrganizationGroupTypes, ByVal includeDeputies As Boolean) As DistinguishedName()
		Function GetGrantedInstitutesOf() As String()
	End Interface

End Namespace
