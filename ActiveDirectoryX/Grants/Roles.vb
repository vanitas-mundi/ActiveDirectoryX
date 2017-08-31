Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "

Imports SSP.ActiveDirectoryX.Grants.Administration
Imports SSP.ActiveDirectoryX.Grants.Enums
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core
Imports System.Collections.ObjectModel

#End Region

Namespace Grants

	Public Class Roles

		Implements IEnumerable(Of AdministrationGroup)

#Region " --------------->> Enumerationen der Klasse "
#End Region '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
    Private _fillType As FillTypes = FillTypes.FillAll
    Private _roles As New List(Of AdministrationGroup)
    Private _grantTree As GrantTree = Nothing
#End Region '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
#End Region '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
    ''' <summary>
    ''' Stellt Funktionen zur Administration von Rollen zur Verfügung.
    ''' </summary>
    Public ReadOnly Property Administration As RolesAdministration
      Get
        Return RolesAdministration.Instance
      End Get
    End Property

    Public ReadOnly Property Roles As ReadOnlyCollection(Of AdministrationGroup)
      Get
        Return _roles.AsReadOnly
      End Get
    End Property

    Public ReadOnly Property Roles(ByVal grantRoleType As GrantTypes) As List(Of AdministrationGroup)
      Get
        Return Me.Where(Function(r) r.GrantType = grantRoleType).ToList
      End Get
    End Property

    Public ReadOnly Property NoRoles As List(Of AdministrationGroup)
      Get
        Return Me.Roles(GrantTypes.NoRole)
      End Get
    End Property

    Public ReadOnly Property CommonRoles As List(Of AdministrationGroup)
      Get
        Return Me.Roles(GrantTypes.CommonRole)
      End Get
    End Property

    Public ReadOnly Property DepartmentRoles As List(Of AdministrationGroup)
      Get
        Return Me.Roles(GrantTypes.DepartmentRole)
      End Get
    End Property

    Public ReadOnly Property ApplicationRoles As List(Of AdministrationGroup)
      Get
        Return Me.Roles(GrantTypes.ApplicationRole)
      End Get
    End Property

    Public ReadOnly Property BaseRoles As List(Of AdministrationGroup)
      Get
        Return Me.Roles(GrantTypes.BaseRole)
      End Get
    End Property

    Public ReadOnly Property ExtraRoles As List(Of AdministrationGroup)
      Get
        Return Me.Roles(GrantTypes.ExtraRole)
      End Get
    End Property

    Public ReadOnly Property TeamRoles As List(Of AdministrationGroup)
      Get
        Return Me.Roles(GrantTypes.TeamRole)
      End Get
    End Property

    Public ReadOnly Property DenialRoles As List(Of AdministrationGroup)
      Get
        Return Me.Roles(GrantTypes.DenialRole)
      End Get
    End Property

    ''' <summary>
    ''' Liefert eine Liste mit allen Rollentypen.
    ''' </summary>
    Public ReadOnly Property RoleTypesArray As GrantTypes()
      Get
        Return CType(System.Enum.GetValues(GetType(GrantTypes)), GrantTypes())
      End Get
    End Property

    ''' <summary>
    ''' Liefert eine Liste mit allen Rollentypnamen.
    ''' </summary>
    Public ReadOnly Property RoleTypeNames As List(Of String)
      Get
        Return GetRoleTypeNames()
      End Get
    End Property

    ''' <summary>
    ''' Liefert eine Liste mit allen Rollentyp-DistinguishedNames.
    ''' </summary>
    Public Shared ReadOnly Property RoleTypeDistinguishedNames As List(Of DistinguishedName)
      Get
        Return GetRoleTypeDistinguishedNames()
      End Get
    End Property
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
    Private Sub CheckFilled()

      If _roles.Count = 0 Then

        Select Case _fillType
          Case FillTypes.FillAll
            Dim adminGroups = New AdministrationGroups
            adminGroups.FillAllGrantGroups()
            _roles.AddRange(adminGroups)
          Case FillTypes.FillByGrantTree
            Dim rolesDn = SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Roles)
            _roles.AddRange(_grantTree.Items.Where _
            (Function(dn) dn.ContainsDn(rolesDn)).Select _
            (Function(dn) New AdministrationGroup(dn)).ToList)
        End Select
      End If
    End Sub
#End Region  '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
    Public Function GetEnumerator() As IEnumerator(Of AdministrationGroup) Implements IEnumerable(Of AdministrationGroup).GetEnumerator
      CheckFilled()
      Return _roles.GetEnumerator
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
      CheckFilled()
      Return _roles.GetEnumerator
    End Function

    Public Function Item(ByVal index As Int32) As AdministrationGroup
      Return _roles.Item(index)
    End Function

    Public Function Role(ByVal index As Int32) As AdministrationGroup
      CheckFilled()
      Return _roles.Item(index)
    End Function

    Public Function Item(ByVal dn As DistinguishedName) As AdministrationGroup
      CheckFilled()
      Return _roles.Where(Function(r) r.GroupDistinguishedName.Value = dn.Value).FirstOrDefault
    End Function

    Public Function Role(ByVal dn As DistinguishedName) As AdministrationGroup
      Return Me.Item(dn)
    End Function

    ''' <summary>
    ''' Liefert eine Liste mit allen Rollentypennamen
    ''' </summary>
    Public Shared Function GetRoleTypeNames() As List(Of String)
      Return SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Roles).Children.Select(Function(dn) dn.ToRelativeName).ToList
    End Function

    ''' <summary>
    ''' Liefert eine Liste mit allen Rollentypen-DistinguishedNames
    ''' </summary>
    Public Shared Function GetRoleTypeDistinguishedNames() As List(Of DistinguishedName)
      Return SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Roles).Children.ToList
    End Function

    ''' <summary>
    ''' Füllt das Roles-Objekt mit allen vorhandenen Rollen-Objekten.
    ''' </summary>
    Public Sub FillAll()
      _roles.Clear()
      _grantTree = Nothing
      _fillType = FillTypes.FillAll

    End Sub

    Public Sub FillByGrantTree(ByVal gt As GrantTree)
      _roles.Clear()
      _grantTree = gt
      _fillType = FillTypes.FillByGrantTree
    End Sub
#End Region  '{Öffentliche Methoden der Klasse}

  End Class

End Namespace
