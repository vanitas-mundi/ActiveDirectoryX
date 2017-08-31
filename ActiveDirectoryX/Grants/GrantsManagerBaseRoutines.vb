Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.ActiveDirectoryX.Core
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Data
Imports SSP.ActiveDirectoryX.Grants.Enums
Imports SSP.ActiveDirectoryX.Data.Repositories
Imports System.Security.Principal
#End Region

Namespace Grants

	Public Class GrantsBaseRoutines

#Region " --------------->> Enumerationen der Klasse "
#End Region  '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private Shared _instance As GrantsBaseRoutines
#End Region

#Region " --------------->> Konstruktor der Klasse "
		Private Sub New()
		End Sub

		Shared Sub New()
			_instance = New GrantsBaseRoutines
		End Sub
#End Region

#Region " --------------->> Zugriffsmethoden der Klasse "
		Public Shared ReadOnly Property Instance As GrantsBaseRoutines
			Get
				Return _instance
			End Get
		End Property
#End Region

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
		'Private Function GetMembersFromGroupByGroupManagerDn _
		'(ByVal groupManagerDn As DistinguishedName _
		', ByVal memberType As GetMembersRecursiveTypes) As DistinguishedName()

		'  Dim sb = New AdSelectBuilder
		'  sb.Select.Add(AdProperties.distinguishedName.ToString)
		'  sb.From.Add("'" & SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Administration).ToUrl & "'")
		'  sb.Where.Add(String.Format("{0} = '{1}'", AdProperties.managedBy.ToString, groupManagerDn.Value))

		'  Dim distinguishedNames = DistinguishedNameRepository.Instance.GetByDistinguishedNames(sb)

		'  Dim members = New List(Of DistinguishedName)

		'  If Not memberType = GetMembersRecursiveTypes.UsersOnly Then
		'    members.AddRange(distinguishedNames)
		'  End If

		'  distinguishedNames.ToList.ForEach(Sub(groupDn) members.AddRange(groupDn.GetMembersRecursive(memberType)))

		'  Return members.ToArray
		'End Function
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
		'''<summary>Liefert die Personen-Id anhand des angemeldeten Windows-Benutzers.</summary>
		Public Function GetPersonIdByWindowsUser() As Int64
			Dim userName = WindowsIdentity.GetCurrent.Name.Split("\"c).Last
			Return DistinguishedName.GetByUserName(userName).BaseProperties.PersonId
		End Function

		'''<summary>
		'''Liefert die DistinguishedNames der Gruppenmanager, bei welchen der Manager oder 
		'''die Stellvertreter der angegebenen managerPersonId entsprechen
		'''</summary>
		Public Function GetGroupManagerDistinguishedNamesByManagerPersonId(ByVal managerPersonId As Int64) As DistinguishedName()

			Return GetGroupManagerDistinguishedNamesByManagerPersonId(managerPersonId, True)
		End Function

		'''<summary>
		'''Liefert die DistinguishedNames der Gruppenmanager, bei welchen der Manager der angegebenen managerPersonId entspricht.
		'''Wird für includeDeputies 'true' übergeben kann die angegebene managerPersonId auch ein Stellvertreter sein.
		'''</summary>
		Public Function GetGroupManagerDistinguishedNamesByManagerPersonId _
		(ByVal managerPersonId As Int64, ByVal includeDeputies As Boolean) As DistinguishedName()

			Return GroupManagers.Instance.ToList.Where(Function(x) (x.ManagerPersonId = managerPersonId) _
			OrElse (If(includeDeputies, x.DeputyPersonIds.Contains(managerPersonId), False))).Select _
			(Function(x) x.ManagerGroupDn).ToArray
		End Function

		'''<summary>Liefert anhand des angegebenen Distinguishednames des Gruppenmanagers alle unterstellten User (rekursiv).</summary>
		Public Function GetAssignedUsersOfGroupManager(ByVal groupManagerDn As DistinguishedName) As DistinguishedName()

			Return GroupManagers.Instance.Item(groupManagerDn).AssignedUsers
		End Function

		'''<summary>Liefert anhand des angegebenen Distinguishednames des Gruppenmanagers alle unterstellten Administrationsgruppen (rekursiv).</summary>
		Public Function GetAssignedAdministrationGroupsOfGroupManager(ByVal groupManagerDn As DistinguishedName) As AdministrationGroup()

			Return GroupManagers.Instance.GetAdministrationGroupsOf(groupManagerDn)
		End Function

		'''<summary>
		'''Liefert alle Administrationsgruppen, bei welchen der Manager der angegebenen managerPersonId entspricht oder einem Stellvertreter.
		'''</summary>
		Public Function GetAssignedAdministrationGroupsByManagerPersonId(ByVal managerPersonId As Int64) As AdministrationGroup()

			Return GetAssignedAdministrationGroupsByManagerPersonId(managerPersonId, True)
		End Function

		'''<summary>
		'''Liefert alle Administrationsgruppen, bei welchen der Manager der angegebenen managerPersonId entspricht.
		'''Wird 'true' für includeDeputies übergeben, dann kann managerPersonId auch einem Stellvertreter gehören.
		'''</summary>
		Public Function GetAssignedAdministrationGroupsByManagerPersonId _
		(ByVal managerPersonId As Int64, ByVal includeDeputies As Boolean) As AdministrationGroup()

			Dim result = New List(Of AdministrationGroup)
			Dim groupManagersList = GetGroupManagerDistinguishedNamesByManagerPersonId(managerPersonId, includeDeputies)

			groupManagersList.ToList.ForEach(Sub(x) result.AddRange(GroupManagers.Instance.GetAdministrationGroupsOf(x)))

			Return result.ToArray
		End Function

		'''<summary>
		'''Liefert rekursiv alle User aller Gruppen, bei welchen der Manager der 
		'''angegebenen managerPersonId entspricht oder einem Stellvertreter.
		'''</summary>
		Public Function GetManagedUsersOfManager(ByVal managerPersonId As Int64) As DistinguishedName()
			Return GetManagedUsersOfManager(managerPersonId, True)
		End Function

		'''<summary>
		'''Liefert rekursiv alle User aller Gruppen, bei welchen der Manager der angegebenen managerPersonId entspricht.
		'''Wird 'true' für includeDeputies übergeben, dann kann managerPersonId auch einem Stellvertreter gehören.
		'''</summary>
		Public Function GetManagedUsersOfManager _
		(ByVal managerPersonId As Int64, ByVal includeDeputies As Boolean) As DistinguishedName()

			Return GroupManagers.Instance.GetManagedUsersOf(managerPersonId, includeDeputies)
		End Function

		'''<summary>Prüft, ob der LoginUser Gruppenmanager der angegebenen Personen-Id ist - vom angegebenen Gruppentyp.</summary>
		Public Function IsLoginUserManagerOf(ByVal groupType As OrganizationGroupTypes _
		, ByVal loginUserPersonId As Int64, ByVal personId As Int64) As Boolean

			Return IsLoginUserManagerOf(groupType, loginUserPersonId, personId, True)
		End Function

		'''<summary>Prüft, ob der LoginUser Gruppenmanager der angegebenen Personen-Id ist vom angegebenen Gruppentyp.</summary>
		'''Wird 'true' für includeDeputies übergeben, dann kann managerPersonId auch einem Stellvertreter gehören.
		Public Function IsLoginUserManagerOf(ByVal groupType As OrganizationGroupTypes _
		, ByVal loginUserPersonId As Int64, ByVal personId As Int64, ByVal includeDeputies As Boolean) As Boolean

			Dim result = New List(Of DistinguishedName)

			GroupManagers.Instance.GetManagedGroupsOf(loginUserPersonId, includeDeputies).Where _
			(Function(x) x.OrganizationGroupType = groupType).ToList.ForEach _
			(Sub(x) result.AddRange(x.GroupDistinguishedName.GetMembers))

			Return result.Any(Function(x) x.BaseProperties.PersonId = personId)
		End Function

		'''<summary>
		'''Liefert alle Manager-DistinguishedNames, des angegebenen Gruppentyps,
		'''welchen der GrantUser unterstellt ist, sowie die DistinguishedNames derer Stellvertreter.
		'''</summary>
		Public Function GetManagerDistinguishedNamesOfGrantUser _
		(ByVal grantUser As GrantUser, ByVal groupType As OrganizationGroupTypes) As DistinguishedName()

			Return GetManagerDistinguishedNamesOfGrantUser(grantUser, groupType, True)
		End Function

		'''<summary>
		'''Liefert alle Manager-DistinguishedNames, des angegebenen Gruppentyps,
		'''welchen der GrantUser unterstellt ist.
		'''Wird 'true' für includeDeputies übergeben, dann werden auch die Stellvertreter geliefert.
		'''</summary>
		Public Function GetManagerDistinguishedNamesOfGrantUser _
		(ByVal grantUser As GrantUser, ByVal groupType As OrganizationGroupTypes, ByVal includeDeputies As Boolean) As DistinguishedName()

			Dim result = New List(Of DistinguishedName)

			Dim managerAndDeputiesList = grantUser.OrganizationGroups.OrganizationGroups _
			(groupType).Select(Function(x) If(includeDeputies _
			, x.GroupManager.ManagerAndDeputies, New DistinguishedName() {x.GroupManager.ManagerDn})).ToList

			managerAndDeputiesList.ForEach(Sub(x) result.AddRange(x))

			Return result.GroupBy(Function(x) x.BaseProperties.PersonId).Select(Function(x) x.First).ToArray
		End Function

		'''<summary>Liefert die Manager-DistinguishedNames des angegebenen Gruppentyps, welche der Person personId zugeordnet sind.</summary>
		Public Function GetManagerDistinguishedNamesOfPersonId _
		(ByVal personId As Int64, ByVal groupType As OrganizationGroupTypes) As DistinguishedName()

			Return GetManagerDistinguishedNamesOfPersonId(personId, groupType, True)
		End Function

		'''<summary>
		'''Liefert die Manager-DistinguishedNames des angegebenen Gruppentyps, welche der Person personId zugeordnet sind.
		'''Wird 'true' für includeDeputies übergeben, werden ebenfalls die Stellvertreter berücksichtigt.
		'''</summary>
		Public Function GetManagerDistinguishedNamesOfPersonId(ByVal personId As Int64 _
		, ByVal groupType As OrganizationGroupTypes, ByVal includeDeputies As Boolean) As DistinguishedName()

			Return GetManagerDistinguishedNamesOfGrantUser(New GrantUser(personId), groupType, includeDeputies)
		End Function

		'''<summary>Liefert dem LoginUser zugewiesene Institute.</summary>
		Public Function GetGrantedInstitutesOf(ByVal grantUser As GrantUser) As String()

			Return grantUser.GrantTables.Table("granted_institute").GrantNames
		End Function

#End Region '{Öffentliche Methoden der Klasse}

	End Class

End Namespace
