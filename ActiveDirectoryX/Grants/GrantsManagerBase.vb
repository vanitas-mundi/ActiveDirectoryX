Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.ActiveDirectoryX.Core
Imports SSP.ActiveDirectoryX.Grants.Enums
Imports SSP.ActiveDirectoryX.Grants.Interfaces
#End Region

Namespace Grants

	Public MustInherit Class GrantsManagerBase(Of TGrantNamesEnum As Structure) 

		Implements IGrantsManagerBase(of TGrantNamesEnum)

#Region " --------------->> Enumerationen der Klasse "
#End Region '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
    Protected _grantsAppName As String
#End Region

#Region " --------------->> Konstruktor der Klasse "
    Public Sub New(ByVal grantsAppName As String, ByVal loginUserId As Int64)

      _Routines = GrantsBaseRoutines.Instance
      _grantsAppName = grantsAppName
      Me.LoginUserPersonId = loginUserId
      _GrantUser = New GrantUser(Me.LoginUserPersonId)
    End Sub

    Public Sub New(ByVal grantsAppName As String)

      Me.New(grantsAppName, GetPersonIdByWindowsUser)
    End Sub
#End Region

#Region " --------------->> Zugriffsmethoden der Klasse "
    Protected ReadOnly Property Routines As GrantsBaseRoutines _
    Implements IGrantsManagerBase(Of TGrantNamesEnum).Routines

    '''<summary>Liefert die Applikationsberechtigungen des angemeldeten Benutzers.</summary>
    Public ReadOnly Property ApplicationGrants() As GrantTable _
    Implements IGrantsManagerBase(Of TGrantNamesEnum).ApplicationGrants
      Get
        Return Me.GrantUser.GrantTables.Table(_grantsAppName)
      End Get
    End Property

    '''<summary>Liefert die Personen-Id des angemeldeten Benutzers.</summary>
    Public ReadOnly Property LoginUserPersonId() As Int64 _
    Implements IGrantsManagerBase(Of TGrantNamesEnum).LoginUserPersonId

    '''<summary>Liefert das GrantUser-Objekt des angemeldeten Benutzers.</summary>
    Public ReadOnly Property GrantUser As GrantUser _
    Implements IGrantsManagerBase(Of TGrantNamesEnum).GrantUser
#End Region

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
    '''<summary>Liefert die Personen-Id anhand des angemeldeten Windows-Benutzers.</summary>
    Public Shared Function GetPersonIdByWindowsUser() As Int64
			Return GrantsBaseRoutines.Instance.GetPersonIdByWindowsUser()
		End Function

    '''<summary>Prüft, ob die Berechtigung grantName des angemeldeten Benutzers in der Applikation gewährt ist.</summary>
    Public Function IsGranted(ByVal grantName As TGrantNamesEnum) As Boolean _
    Implements IGrantsManagerBase(Of TGrantNamesEnum).IsGranted
      Return Me.ApplicationGrants.Grant(grantName.ToString).IsGranted
    End Function

    '''<summary>
    '''Liefert die DistinguishedNames der Gruppenmanager, bei welchen der Manager oder 
    '''die Stellvertreter der angegebenen managerPersonId entsprechen
    '''</summary>
    Public Function GetGroupManagerDistinguishedNamesByManagerPersonId _
    (ByVal managerPersonId As Int64) As DistinguishedName() _
    Implements IGrantsManagerBase(Of TGrantNamesEnum).GetGroupManagerDistinguishedNamesByManagerPersonId

      Return Me.Routines.GetGroupManagerDistinguishedNamesByManagerPersonId(managerPersonId, True)
    End Function

    '''<summary>
    '''Liefert die DistinguishedNames der Gruppenmanager, bei welchen der Manager der angegebenen managerPersonId entspricht.
    '''Wird für includeDeputies 'true' übergeben kann die angegebene managerPersonId auch ein Stellvertreter sein.
    '''</summary>
    Public Function GetGroupManagerDistinguishedNamesByManagerPersonId _
		(ByVal managerPersonId As Int64, ByVal includeDeputies As Boolean) As DistinguishedName() _
		Implements IGrantsManagerBase(of TGrantNamesEnum).GetGroupManagerDistinguishedNamesByManagerPersonId

			Return Me.Routines.GetGroupManagerDistinguishedNamesByManagerPersonId(managerPersonId, includeDeputies)
		End Function

    '''<summary>Liefert anhand des angegebenen Berechtigungsnamen alle berechtigten User (rekursiv).</summary>
    Public Function GetGrantAssignedUsers _
    (ByVal grantName As TGrantNamesEnum) As DistinguishedName() _
    Implements IGrantsManagerBase(Of TGrantNamesEnum).GetGrantAssignedUsers

      Return (New GrantTable(_grantsAppName)).Grant(grantName.ToString).AssignedUsers
    End Function

    '''<summary>Liefert anhand des angegebenen Distinguishednames des Gruppenmanagers alle unterstellten User.</summary>
    Public Function GetAssignedUsersOfGroupManager _
    (ByVal groupManagerDn As DistinguishedName) As DistinguishedName() _
    Implements IGrantsManagerBase(Of TGrantNamesEnum).GetAssignedUsersOfGroupManager

      Return Me.Routines.GetAssignedUsersOfGroupManager(groupManagerDn)
    End Function

    '''<summary>
    '''Liefert anhand des angegebenen Distinguishednames des Gruppenmanagers 
    '''alle unterstellten Administrationsgruppen (rekursiv).
    '''</summary>
    Public Function GetAssignedAdministrationGroupsOfGroupManager _
    (ByVal groupManagerDn As DistinguishedName) As AdministrationGroup() _
    Implements IGrantsManagerBase(Of TGrantNamesEnum).GetAssignedAdministrationGroupsOfGroupManager

      Return Me.Routines.GetAssignedAdministrationGroupsOfGroupManager(groupManagerDn)
    End Function

    '''<summary>
    '''Liefert alle Administrationsgruppen, bei welchen der Manager 
    '''der angegebenen managerPersonId entspricht oder einem Stellvertreter.
    '''</summary>
    Public Function GetAssignedAdministrationGroupsByManagerPersonId _
    (ByVal managerPersonId As Int64) As AdministrationGroup() _
    Implements IGrantsManagerBase(Of TGrantNamesEnum).GetAssignedAdministrationGroupsByManagerPersonId

      Return Me.Routines.GetAssignedAdministrationGroupsByManagerPersonId(managerPersonId)
    End Function

    '''<summary>
    '''Liefert alle Administrationsgruppen, bei welchen der Manager der angegebenen managerPersonId entspricht.
    '''Wird 'true' für includeDeputies übergeben, dann kann managerPersonId auch einem Stellvertreter gehören.
    '''</summary>
    Public Function GetAssignedAdministrationGroupsByManagerPersonId _
		(ByVal managerPersonId As Int64, ByVal includeDeputies As Boolean) As AdministrationGroup() _
		Implements IGrantsManagerBase(of TGrantNamesEnum).GetAssignedAdministrationGroupsByManagerPersonId

			Return Me.Routines.GetAssignedAdministrationGroupsByManagerPersonId(managerPersonId, includeDeputies)
		End Function

		'''<summary>
		'''Liefert rekursiv alle User aller Gruppen, bei welchen der Manager der 
		'''angegebenen managerPersonId entspricht oder einem Stellvertreter.
		'''</summary>
		Public Function GetManagedUsersOfManager(ByVal managerPersonId As Int64) As DistinguishedName() _
		Implements IGrantsManagerBase(of TGrantNamesEnum).GetManagedUsersOfManager

			Return Me.Routines.GetManagedUsersOfManager(managerPersonId)
		End Function

		'''<summary>
		'''Liefert rekursiv alle User aller Gruppen, bei welchen der Manager der angegebenen managerPersonId entspricht.
		'''Wird 'true' für includeDeputies übergeben, dann kann managerPersonId auch einem Stellvertreter gehören.
		'''</summary>
		Public Function GetManagedUsersOfManager _
		(ByVal managerPersonId As Int64, ByVal includeDeputies As Boolean) As DistinguishedName() _
		Implements IGrantsManagerBase(of TGrantNamesEnum).GetManagedUsersOfManager

			Return Me.Routines.GetManagedUsersOfManager(managerPersonId, includeDeputies)
		End Function

		'''<summary>Prüft, ob der LoginUser Gruppenmanager der angegebenen Personen-Id ist - vom angegebenen Gruppentyp.</summary>
		Public Function IsLoginUserManagerOf(ByVal groupType As OrganizationGroupTypes, ByVal personId As Int64) As Boolean _
		Implements IGrantsManagerBase(of TGrantNamesEnum).IsLoginUserManagerOf

			Return Me.Routines.IsLoginUserManagerOf(groupType, Me.LoginUserPersonId, personId)
		End Function

    '''<summary>
    '''Prüft, ob der LoginUser Gruppenmanager der angegebenen Personen-Id ist vom angegebenen Gruppentyp.
    '''Wird 'true' für includeDeputies übergeben, dann kann managerPersonId auch einem Stellvertreter gehören.
    '''</summary>
    Public Function IsLoginUserManagerOf _
    (ByVal groupType As OrganizationGroupTypes, ByVal personId As Int64, ByVal includeDeputies As Boolean) As Boolean _
    Implements IGrantsManagerBase(Of TGrantNamesEnum).IsLoginUserManagerOf

      Return Me.Routines.IsLoginUserManagerOf(groupType, Me.LoginUserPersonId, personId, includeDeputies)
    End Function

    '''<summary>
    '''Liefert alle Manager-DistinguishedNames, des angegebenen Gruppentyps,
    '''welchen der LoginUser unterstellt ist, sowie die DistinguishedNames derer Stellvertreter.
    '''</summary>
    Public Function GetManagerDistinguishedNamesOfGrantUser(ByVal groupType As OrganizationGroupTypes) As DistinguishedName() _
		Implements IGrantsManagerBase(of TGrantNamesEnum).GetManagerDistinguishedNamesOfGrantUser

			Return Me.Routines.GetManagerDistinguishedNamesOfGrantUser(Me.GrantUser, groupType)
		End Function

		'''<summary>
		'''Liefert alle Manager-DistinguishedNames, des angegebenen Gruppentyps,
		'''welchen der LoginUser unterstellt ist.
		'''Wird 'true' für includeDeputies übergeben, dann werden auch die Stellvertreter geliefert.
		'''</summary>
		Public Function GetManagerDistinguishedNamesOfGrantUser _
		(ByVal groupType As OrganizationGroupTypes, ByVal includeDeputies As Boolean) As DistinguishedName() _
		Implements IGrantsManagerBase(of TGrantNamesEnum).GetManagerDistinguishedNamesOfGrantUser

			Return Me.Routines.GetManagerDistinguishedNamesOfGrantUser(Me.GrantUser, groupType, includeDeputies)
		End Function

    '''<summary>
    '''Liefert die Manager-DistinguishedNames des angegebenen Gruppentyps, 
    '''welche der Person personId zugeordnet sind.
    '''</summary>
    Public Function GetManagerDistinguishedNamesOfPersonId _
		(ByVal personId As Int64, ByVal groupType As OrganizationGroupTypes) As DistinguishedName() _
		Implements IGrantsManagerBase(of TGrantNamesEnum).GetManagerDistinguishedNamesOfPersonId

			Return Me.Routines.GetManagerDistinguishedNamesOfPersonId(personId, groupType)
		End Function

    '''<summary>
    '''Liefert die Manager-DistinguishedNames des angegebenen Gruppentyps, 
    '''welche der Person personId zugeordnet sind.
    '''Wird 'true' für includeDeputies übergeben, werden ebenfalls die Stellvertreter berücksichtigt.
    '''</summary>
    Public Function GetManagerDistinguishedNamesOfPersonId _
    (ByVal personId As Int64, ByVal groupType As OrganizationGroupTypes _
    , ByVal includeDeputies As Boolean) As DistinguishedName() _
    Implements IGrantsManagerBase(Of TGrantNamesEnum).GetManagerDistinguishedNamesOfPersonId

      Return Me.Routines.GetManagerDistinguishedNamesOfPersonId(personId, groupType, includeDeputies)
    End Function

    '''<summary>Liefert dem LoginUser zugewiesene Institute.</summary>
    Public Function GetGrantedInstitutesOf() As String() _
    Implements IGrantsManagerBase(Of TGrantNamesEnum).GetGrantedInstitutesOf

      Return Me.Routines.GetGrantedInstitutesOf(Me.GrantUser)
    End Function
#End Region '{Öffentliche Methoden der Klasse}

  End Class

End Namespace
