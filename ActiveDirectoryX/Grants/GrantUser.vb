Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports System.DirectoryServices
Imports System.DirectoryServices.AccountManagement
Imports SSP.ActiveDirectoryX.Grants.Administration
Imports SSP.ActiveDirectoryX.Core
Imports SSP.ActiveDirectoryX.Grants.Exceptions
Imports SSP.ActiveDirectoryX.Grants.Enums
#End Region

Namespace Grants

  Public Class GrantUser

#Region " --------------->> Enumerationen der Klasse "
#End Region '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
    Private _administration As GrantUserAdministration
    Private _organizationGroups As OrganizationGroups
    Private _grantGroups As GrantGroups
    Private _userName As String
    Private _userDirectoryEntry As DirectoryEntry
    Private _userDistinguishedName As DistinguishedName
    Private _grantTree As GrantTree
    Private _userPrincipal As UserPrincipal
#End Region  '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
    Public Sub New(ByVal personId As Int64)

      Me.New(DistinguishedName.GetByPersonId(personId))
    End Sub

    Public Sub New(ByVal userName As String)

      Me.New(DistinguishedName.GetByUserName(userName))
    End Sub

    Public Sub New(ByVal userDn As DistinguishedName)

      If userDn Is Nothing Then
        Throw New PersonIdNotExistsException
      Else
        Initialize(New GrantTree(userDn))
      End If
    End Sub

    Public Sub New(ByVal gt As GrantTree)
      If gt Is Nothing Then
        Throw New GrantTableIsNullException
      Else
        Initialize(gt)
      End If
    End Sub

    ''' <summary>
    ''' Liefert eine GrantUser-Instanz oder Null, wenn personId nicht vorhanden.
    ''' </summary>
    Public Shared Function CreateGrantUser(ByVal personId As Int64) As GrantUser
      Try
        Return New GrantUser(personId)
      Catch ex As System.Exception
        Return Nothing
      End Try
    End Function

    ''' <summary>
    ''' Liefert eine GrantUser-Instanz oder Null, wenn userName nicht vorhanden.
    ''' </summary>
    Public Shared Function CreateGrantUser(ByVal userName As String) As GrantUser
      Try
        Return New GrantUser(userName)
      Catch ex As System.Exception
        Return Nothing
      End Try
    End Function

    ''' <summary>
    ''' Liefert eine GrantUser-Instanz oder Null, wenn userDn eine Null-Referenz besitzt.
    ''' </summary>
    Public Shared Function CreateGrantUser(ByVal userDn As DistinguishedName) As GrantUser
      Try
        Return New GrantUser(userDn)
      Catch ex As System.Exception
        Return Nothing
      End Try
    End Function

    ''' <summary>
    ''' Liefert eine GrantUser-Instanz oder Null, wenn gt eine Null-Referenz besitzt.
    ''' </summary>
    Public Shared Function CreateGrantUser(ByVal gt As GrantTree) As GrantUser
      Try
        Return New GrantUser(gt)
      Catch ex As System.Exception
        Return Nothing
      End Try
    End Function
#End Region '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
    ''' <summary>
    ''' Stellt User-Administrationsfunktionalität zur Verfügung.
    ''' </summary>
    Public ReadOnly Property Administration As GrantUserAdministration
      Get
        If _administration Is Nothing Then
          _administration = New GrantUserAdministration(Me)
        End If

        Return _administration
      End Get
    End Property

    ''' <summary>
    ''' Liefert sofern vorhanden den zugrunde liegenden GrantTree oder NULL.
    ''' </summary>
    Public ReadOnly Property GrantTree As GrantTree
      Get
        Return _grantTree
      End Get
    End Property

    ''' <summary>
    ''' DistinguishedName des Users.
    ''' </summary>
    Public ReadOnly Property UserDistinguishedName As DistinguishedName
      Get
        Return _userDistinguishedName
      End Get
    End Property

    ''' <summary>
    ''' Liefert das zugrunde liegende DirectoryEntry.
    ''' </summary>
    Public ReadOnly Property UserDirectoryEntry As DirectoryEntry
      Get

        If _userDirectoryEntry Is Nothing Then
          _userDirectoryEntry = _userDistinguishedName.ToDirectoryEntry(False)
        End If
        Return _userDirectoryEntry
      End Get
    End Property

    ''' <summary>
    ''' Liefert das zugrunde liegende UserPrincipal-Objekt.
    ''' Haben sich in der Zwischenzeit seit dem ersten Property-Aufruf Daten im AD geändert,
    ''' bekommt die Property diese Änderungen nicht mit.
    ''' </summary>
    Public ReadOnly Property UserPrincipal As UserPrincipal
      Get
        If _userPrincipal Is Nothing Then
          _userPrincipal = AdPrincipals.GetUserPrincipal(Me.UserName)
        End If
        Return _userPrincipal
      End Get
    End Property

    ''' <summary>
    ''' UserName/ samAccountName des Users.
    ''' </summary>
    Public ReadOnly Property UserName As String
      Get
        If String.IsNullOrEmpty(_userName) AndAlso (_grantTree IsNot Nothing) Then
          _userName = _grantTree.DistinguishedNameContext.BaseProperties.Name
        End If
        Return _userName
      End Get
    End Property

    ''' <summary>
    ''' Organisationsgruppen des Users.
    ''' </summary>
    Public ReadOnly Property OrganizationGroups As OrganizationGroups
      Get
        Return _organizationGroups
      End Get
    End Property

    ''' <summary>
    ''' Berechtigungsgruppen des Users.
    ''' </summary>
    Public ReadOnly Property GrantGroups As GrantGroups
      Get
        Return _grantGroups
      End Get
    End Property

    ''' <summary>
    ''' Rollen des Users.
    ''' </summary>
    Public ReadOnly Property Roles As Roles
      Get
        Return _grantGroups.Roles()
      End Get
    End Property

    ''' <summary>
    ''' Mappings des Users.
    ''' </summary>
    Public ReadOnly Property Mappings As Mappings
      Get
        Return _grantGroups.Mappings
      End Get
    End Property

    ''' <summary>
    ''' GrantTables des Users.
    ''' </summary>
    Public ReadOnly Property GrantTables As GrantTables
      Get
        Return _grantGroups.GrantTables
      End Get
    End Property

    ''' <summary>
    ''' Liefert dem Benutzer zugewiesene Gruppenmanager-Manager (Bewilliger).
    ''' </summary>
    Public ReadOnly Property AssignedManagers As DistinguishedName()
      Get
        Return Me.OrganizationGroups.OrganizationGroups _
        (OrganizationGroupTypes.HolidayGroup).Select _
        (Function(x) GroupManagers.Instance.GetGroupManagerOf(x).ManagerDn).Distinct.ToArray
      End Get
    End Property

    ''' <summary>
    ''' Liefert dem Benutzer zugewiesene Gruppenmanager-Setllvertreter (Bewilliger-Stellvertreter).
    ''' </summary>
    Public ReadOnly Property AssignedDeputies As DistinguishedName()
      Get
        Dim deputies = New List(Of DistinguishedName)

        Me.OrganizationGroups.OrganizationGroups _
        (OrganizationGroupTypes.HolidayGroup).ToList.ForEach _
        (Sub(x) deputies.AddRange(GroupManagers.Instance.GetGroupManagerOf(x).Deputies))

        Return deputies.Distinct.ToArray
      End Get
    End Property

    ''' <summary>
    ''' Liefert dem Benutzer zugewiesene Gruppenmanager-Manager (Bewilliger) uind
    ''' Gruppenmanager-Setllvertreter (Bewilliger-Stellvertreter).
    ''' </summary>
    Public ReadOnly Property AssignedManagersAndDeputies As DistinguishedName()
      Get
        Dim result = New List(Of DistinguishedName)
        Return result.Union(Me.AssignedManagers).Union(Me.AssignedDeputies).ToArray
      End Get
    End Property

    ''' <summary>
    ''' Liefert die vom User verwalteten (zu bewilligenden) Administrationsgruppen.
    ''' </summary>
    Public ReadOnly Property ManagedGroups As AdministrationGroup()
      Get
        Return GroupManagers.Instance.GetManagedGroupsOf(Me.UserDistinguishedName)
      End Get
    End Property

    ''' <summary>
    ''' Liefert die vom User verwalteten (zu bewilligenden) Administrationsgruppen.
    ''' </summary>
    Public ReadOnly Property ManagedUsers As DistinguishedName()
      Get
        Return GroupManagers.Instance.GetManagedUsersOf(Me.UserDistinguishedName)
      End Get
    End Property
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
    ''' <summary>
    ''' Initialisiert den GrantUser.
    ''' </summary>
    Private Sub Initialize(ByVal gt As GrantTree)

      _grantTree = gt
      If Not gt.IsGenerated Then gt.Generate()

      _organizationGroups = New OrganizationGroups
      _organizationGroups.FillByGrantTree(gt)

      _grantGroups = New GrantGroups
      _grantGroups.FillByGrantTree(gt)

      _userDistinguishedName = gt.DistinguishedNameContext
    End Sub
#End Region  '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
    ''' <summary>
    ''' Liefert den Username.
    ''' </summary>
    Public Overrides Function ToString() As String
      Return Me.UserName
    End Function
#End Region '{Öffentliche Methoden der Klasse}

  End Class

End Namespace


