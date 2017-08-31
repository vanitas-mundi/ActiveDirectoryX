Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports System.ComponentModel
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Data.Repositories
#End Region

Namespace Core

  <DefaultProperty("Item")>
  Public Class SpecialDistinguishedNames

#Region " --------------->> Enumerationen der Klasse "
#End Region '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
    Private Shared _specialDistinguishedNames As New Dictionary(Of SpecialDistinguishedNameKeys, DistinguishedName)
#End Region '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
    Shared Sub New()
      Initialize()
    End Sub
#End Region '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
    Public Shared ReadOnly Property DomainDn As String
      Get
        Return String.Join(",", Settings.Instance.DomainName.Split("."c).Select(Function(item) "dc=" & item).ToArray)
      End Get
    End Property

    Private Shared ReadOnly Property AdministrationDn As String
      Get
        Return ToDistinguishedNameString(My.Settings.AdministrationDistinguishedName)
      End Get
    End Property

    Private Shared ReadOnly Property GrantsDn As String
      Get
        Return ToDistinguishedNameString(My.Settings.GrantsDistinguishedName)
      End Get
    End Property

    Private Shared ReadOnly Property RolesDn As String
      Get
        Return ToDistinguishedNameString(My.Settings.RolesDistinguishedName)
      End Get
    End Property

    Private Shared ReadOnly Property MappingsDn As String
      Get
        Return ToDistinguishedNameString(My.Settings.MappingsDistinguishedName)
      End Get
    End Property

    Private Shared ReadOnly Property DepartmentRolesDn As String
      Get
        Return ToDistinguishedNameString(My.Settings.DepartmentRolesDistinguishedName)
      End Get
    End Property

    Private Shared ReadOnly Property ApplicationRolesDn As String
      Get
        Return ToDistinguishedNameString(My.Settings.ApplicationRolesDistinguishedName)
      End Get
    End Property

    Private Shared ReadOnly Property BaseRolesDn As String
      Get
        Return ToDistinguishedNameString(My.Settings.BaseRolesDistinguishedName)
      End Get
    End Property

    Private Shared ReadOnly Property ExtraRolesDn As String
      Get
        Return ToDistinguishedNameString(My.Settings.ExtraRolesDistinguishedName)
      End Get
    End Property

    Private Shared ReadOnly Property TeamRolesDn As String
      Get
        Return ToDistinguishedNameString(My.Settings.TeamRolesDistinguishedName)
      End Get
    End Property

    Private Shared ReadOnly Property DenialRolesDn As String
      Get
        Return ToDistinguishedNameString(My.Settings.DenialRolesDistinguishedName)
      End Get
    End Property

    Private Shared ReadOnly Property OrganizationGroupsDn As String
      Get
        Return ToDistinguishedNameString(My.Settings.OrganizationGroupsDistinguishedName)
      End Get
    End Property

    Private Shared ReadOnly Property HolidayGroupsDn As String
      Get
        Return ToDistinguishedNameString(My.Settings.HolidayGroupsDistinguishedName)
      End Get
    End Property

    Private Shared ReadOnly Property WorkGroupsDn As String
      Get
        Return ToDistinguishedNameString(My.Settings.WorkGroupsDistinguishedName)
      End Get
    End Property

    Private Shared ReadOnly Property AccountingGroupsDn As String
      Get
        Return ToDistinguishedNameString(My.Settings.AccountingGroupsDistinguishedName)
      End Get
    End Property

    Private Shared ReadOnly Property GroupManagersDn As String
      Get
        Return ToDistinguishedNameString(My.Settings.GroupManagersDistinguishedName)
      End Get
    End Property

    Public Shared ReadOnly Property Item(ByVal specialDistinguishedNameKey As SpecialDistinguishedNameKeys) As DistinguishedName
      Get
        Return _specialDistinguishedNames.Item(specialDistinguishedNameKey)
      End Get
    End Property
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
    Private Shared Function ToDistinguishedNameString(ByVal specialDn As String) As String
      Return String.Format("{0},{1}", specialDn, DomainDn)
    End Function
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
    ''' <summary>
    ''' Verbindet einen relativen DistinguishedName mit einem SpecialDistinguishedName
    ''' und gibt diesen als String zurück.
    ''' </summary>
    Public Shared Function ConcatDnString _
    (ByVal specialDistinguishedNameKey As SpecialDistinguishedNameKeys _
    , ByVal relativeDistinguishedName As String) As String

      Return relativeDistinguishedName & "," & Item(specialDistinguishedNameKey).Value
    End Function

    ''' <summary>
    ''' Verbindet einen relativen DistinguishedName mit einem SpecialDistinguishedName
    ''' und gibt diesen als DistnguishedName zurück.
    ''' </summary>
    Public Shared Function ConcatDn _
    (ByVal specialDistinguishedNameKey As SpecialDistinguishedNameKeys _
    , ByVal relativeDistinguishedName As String) As DistinguishedName

      Return DistinguishedName.GetByDistinguishedName(ConcatDnString(specialDistinguishedNameKey, relativeDistinguishedName))
    End Function
#End Region '{Öffentliche Methoden der Klasse}


    Private Shared Sub Initialize()
      With _specialDistinguishedNames
        Dim specialDnList = DistinguishedNameRepository.Instance.GetByDistinguishedNames(New String() _
        {DomainDn, AdministrationDn, GrantsDn, RolesDn, MappingsDn, ApplicationRolesDn _
        , BaseRolesDn, DenialRolesDn, DepartmentRolesDn, ExtraRolesDn, TeamRolesDn _
        , OrganizationGroupsDn, HolidayGroupsDn, WorkGroupsDn, AccountingGroupsDn, GroupManagersDn}).ToList

        .Add(SpecialDistinguishedNameKeys.Domain, specialDnList.FirstOrDefault(Function(dn) String.Compare(dn.Value, DomainDn, True) = 0))
        .Add(SpecialDistinguishedNameKeys.Administration, specialDnList.FirstOrDefault(Function(dn) String.Compare(dn.Value, AdministrationDn, True) = 0))
        .Add(SpecialDistinguishedNameKeys.Grants, specialDnList.FirstOrDefault(Function(dn) String.Compare(dn.Value, GrantsDn, True) = 0))
        .Add(SpecialDistinguishedNameKeys.Roles, specialDnList.FirstOrDefault(Function(dn) String.Compare(dn.Value, RolesDn, True) = 0))
        .Add(SpecialDistinguishedNameKeys.Mappings, specialDnList.FirstOrDefault(Function(dn) String.Compare(dn.Value, MappingsDn, True) = 0))
        .Add(SpecialDistinguishedNameKeys.ApplicationRoles, specialDnList.FirstOrDefault(Function(dn) String.Compare(dn.Value, ApplicationRolesDn, True) = 0))
        .Add(SpecialDistinguishedNameKeys.BaseRoles, specialDnList.FirstOrDefault(Function(dn) String.Compare(dn.Value, BaseRolesDn, True) = 0))
        .Add(SpecialDistinguishedNameKeys.DenialRoles, specialDnList.FirstOrDefault(Function(dn) String.Compare(dn.Value, DenialRolesDn, True) = 0))
        .Add(SpecialDistinguishedNameKeys.DepartmentRoles, specialDnList.FirstOrDefault(Function(dn) String.Compare(dn.Value, DepartmentRolesDn, True) = 0))
        .Add(SpecialDistinguishedNameKeys.ExtraRoles, specialDnList.FirstOrDefault(Function(dn) String.Compare(dn.Value, ExtraRolesDn, True) = 0))
        .Add(SpecialDistinguishedNameKeys.TeamRoles, specialDnList.FirstOrDefault(Function(dn) String.Compare(dn.Value, TeamRolesDn, True) = 0))
        .Add(SpecialDistinguishedNameKeys.OrganizationGroups, specialDnList.FirstOrDefault(Function(dn) String.Compare(dn.Value, OrganizationGroupsDn, True) = 0))
        .Add(SpecialDistinguishedNameKeys.HolidayGroups, specialDnList.FirstOrDefault(Function(dn) String.Compare(dn.Value, HolidayGroupsDn, True) = 0))
        .Add(SpecialDistinguishedNameKeys.WorkGroups, specialDnList.FirstOrDefault(Function(dn) String.Compare(dn.Value, WorkGroupsDn, True) = 0))
        .Add(SpecialDistinguishedNameKeys.AccountingGroups, specialDnList.FirstOrDefault(Function(dn) String.Compare(dn.Value, AccountingGroupsDn, True) = 0))
        .Add(SpecialDistinguishedNameKeys.GroupManagers, specialDnList.FirstOrDefault(Function(dn) String.Compare(dn.Value, GroupManagersDn, True) = 0))

        '.Add(SpecialDistinguishedNameKeys.Domain, New DistinguishedName(DomainDn))
        '.Add(SpecialDistinguishedNameKeys.Administration, New DistinguishedName(AdministrationDn))
        '.Add(SpecialDistinguishedNameKeys.Grants, New DistinguishedName(GrantsDn))

        '.Add(SpecialDistinguishedNameKeys.Roles, New DistinguishedName(RolesDn))
        '.Add(SpecialDistinguishedNameKeys.Mappings, New DistinguishedName(MappingsDn))
        '.Add(SpecialDistinguishedNameKeys.ApplicationRoles, New DistinguishedName(ApplicationRolesDn))
        '.Add(SpecialDistinguishedNameKeys.BaseRoles, New DistinguishedName(BaseRolesDn))
        '.Add(SpecialDistinguishedNameKeys.DenialRoles, New DistinguishedName(DenialRolesDn))
        '.Add(SpecialDistinguishedNameKeys.DepartmentRoles, New DistinguishedName(DepartmentRolesDn))
        '.Add(SpecialDistinguishedNameKeys.ExtraRoles, New DistinguishedName(ExtraRolesDn))
        '.Add(SpecialDistinguishedNameKeys.TeamRoles, New DistinguishedName(TeamRolesDn))

        '.Add(SpecialDistinguishedNameKeys.OrganizationGroups, New DistinguishedName(OrganizationGroupsDn))
        '.Add(SpecialDistinguishedNameKeys.HolidayGroups, New DistinguishedName(HolidayGroupsDn))
        '.Add(SpecialDistinguishedNameKeys.WorkGroups, New DistinguishedName(WorkGroupsDn))
        '.Add(SpecialDistinguishedNameKeys.AccountingGroups, New DistinguishedName(AccountingGroupsDn))
        '.Add(SpecialDistinguishedNameKeys.GroupManagers, New DistinguishedName(GroupManagersDn))
      End With
    End Sub



  End Class

End Namespace
