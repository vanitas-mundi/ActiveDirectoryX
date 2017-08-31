Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.ActiveDirectoryX.Core
Imports SSP.ActiveDirectoryX.Core.Enums
#End Region

Namespace Grants

	Public Class AdministrationTypeResolver

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private Shared _instance As AdministrationTypeResolver
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Shared Sub New()
			_instance = New AdministrationTypeResolver
		End Sub

		Private Sub New()
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		Public Shared ReadOnly Property Instance As AdministrationTypeResolver
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
		'''<summary>
		'''Liefert den übergeordneten Administrationstyp zu einem DistinguishedName.
		'''Handelt es sich um keine Administrationsgruppe, wird Domain zurückgegeben.
		'''</summary>
		Public Function GetSuperordinateAdministrationType(ByVal dn As DistinguishedName) As SpecialDistinguishedNameKeys

			If dn.IsUser Then
				Return SpecialDistinguishedNameKeys.Domain
			Else
				Select Case True
				Case dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.OrganizationGroups))
					Return SpecialDistinguishedNameKeys.OrganizationGroups
				Case dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Grants))
					Return SpecialDistinguishedNameKeys.Grants
				Case dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Mappings))
					Return SpecialDistinguishedNameKeys.Mappings
				Case dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Roles))
					Return SpecialDistinguishedNameKeys.Roles
				Case dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Administration))
					Return SpecialDistinguishedNameKeys.Administration
				Case Else
					Return SpecialDistinguishedNameKeys.Domain
				End Select
			End If
		End Function

		'''<summary>
		'''Liefert den Administrationstyp zu einem DistinguishedName.
		'''Handelt es sich um keine Administrationsgruppe, wird Domain zurückgegeben.
		'''</summary>
		Public Function GetAdministrationType(ByVal dn As DistinguishedName) As SpecialDistinguishedNameKeys

			If dn.IsUser Then
				Return SpecialDistinguishedNameKeys.Domain
			Else
				Select Case True
				Case dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.AccountingGroups))
					Return SpecialDistinguishedNameKeys.AccountingGroups
				Case dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.HolidayGroups))
					Return SpecialDistinguishedNameKeys.HolidayGroups
				Case dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.GroupManagers))
					Return SpecialDistinguishedNameKeys.GroupManagers
				Case dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.WorkGroups))
					Return SpecialDistinguishedNameKeys.WorkGroups
				Case dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.ApplicationRoles))
					Return SpecialDistinguishedNameKeys.ApplicationRoles
				Case dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.BaseRoles))
					Return SpecialDistinguishedNameKeys.BaseRoles
				Case dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.DenialRoles))
					Return SpecialDistinguishedNameKeys.DenialRoles
				Case dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.DepartmentRoles))
					Return SpecialDistinguishedNameKeys.DepartmentRoles
				Case dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.ExtraRoles))
					Return SpecialDistinguishedNameKeys.ExtraRoles
				Case dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.TeamRoles))
					Return SpecialDistinguishedNameKeys.TeamRoles
				Case Else
					Return GetSuperordinateAdministrationType(dn)
				End Select
			End If
		End Function

		'''<summary>Liefert true, wenn DistinguishedName eine Organisationsgruppe ist.</summary>
		Public Function IsOrganizationGroup(ByVal dn As DistinguishedName) As Boolean

			Return dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.OrganizationGroups))
		End Function

		'''<summary>Liefert true, wenn DistinguishedName eine Berechtigungsgruppe ist.</summary>
		Public Function IsGrantGroup(ByVal dn As DistinguishedName) As Boolean

			Return dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Grants))
		End Function

		'''<summary>Liefert true, wenn DistinguishedName eine GrantTable (Applikation) ist.</summary>
		Public Function IsGrantTable(ByVal dn As DistinguishedName) As Boolean

			Return (IsGrantGroup(dn)) AndAlso (Not dn.Name.Contains("."))
		End Function

		'''<summary>Liefert true, wenn DistinguishedName eine Berechtigung ist.</summary>
		Public Function IsGrant(ByVal dn As DistinguishedName) As Boolean

			Return (IsGrantGroup(dn)) AndAlso (dn.Name.Contains("."))
		End Function

		'''<summary>Liefert true, wenn DistinguishedName ein Mapping ist.</summary>
		Public Function IsMapping(ByVal dn As DistinguishedName) As Boolean

			Return dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Mappings))
		End Function

		'''<summary>Liefert true, wenn DistinguishedName eine Rolle ist.</summary>
		Public Function IsRole(ByVal dn As DistinguishedName) As Boolean

			Return dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Roles))
		End Function

		'''<summary>Liefert true, wenn DistinguishedName eine Administrationsgruppe ist.</summary>
		Public Function IsAdministrationGroup(ByVal dn As DistinguishedName) As Boolean

			Return dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Administration))
		End Function

		'''<summary>Liefert true, wenn DistinguishedName eine Abrechnungsgruppe ist.</summary>
		Public Function IsAccountingGroup(ByVal dn As DistinguishedName) As Boolean

			Return dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.AccountingGroups))
		End Function

		'''<summary>Liefert true, wenn DistinguishedName eine Urlaubsgruppe ist.</summary>
		Public Function IsHolidayGroup(ByVal dn As DistinguishedName) As Boolean

			Return dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.HolidayGroups))
		End Function

		'''<summary>Liefert true, wenn DistinguishedName eine Managergruppe ist.</summary>
		Public Function IsGroupManager(ByVal dn As DistinguishedName) As Boolean

			Return dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.GroupManagers))
		End Function

		'''<summary>Liefert true, wenn DistinguishedName eine Arbeitsgruppe ist.</summary>
		Public Function IsWorkGroup(ByVal dn As DistinguishedName) As Boolean

			Return dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.WorkGroups))
		End Function

		'''<summary>Liefert true, wenn DistinguishedName eine Abteilungsrolle ist.</summary>
		Public Function IsApplicationRole(ByVal dn As DistinguishedName) As Boolean

			Return dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.ApplicationRoles))
		End Function

		'''<summary>Liefert true, wenn DistinguishedName eine Basisrolle ist.</summary>
		Public Function IsBaseRole(ByVal dn As DistinguishedName) As Boolean

			Return dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.BaseRoles))
		End Function

		'''<summary>Liefert true, wenn DistinguishedName eine Verweigerungsrolle ist.</summary>
		Public Function IsDenialRole(ByVal dn As DistinguishedName) As Boolean

			Return dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.DenialRoles))
		End Function

		'''<summary>Liefert true, wenn DistinguishedName eine Abteilungsrolle ist.</summary>
		Public Function IsDepartmentRole(ByVal dn As DistinguishedName) As Boolean

			Return dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.DepartmentRoles))
		End Function

		'''<summary>Liefert true, wenn DistinguishedName eine Extrarolle ist.</summary>
		Public Function IsExtraRole(ByVal dn As DistinguishedName) As Boolean

			Return dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.ExtraRoles))
		End Function

		'''<summary>Liefert true, wenn DistinguishedName eine Teamrolle ist.</summary>
		Public Function IsTeamRole(ByVal dn As DistinguishedName) As Boolean

			Return dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.TeamRoles))
		End Function

		'''<summary>Liefert true, wenn DistinguishedName eine Administrationsorganisationseinheit ist.</summary>
		Public Function IsOrganizationalUnit(ByVal dn As DistinguishedName) As Boolean
			Try
				Return dn.IsOrganizationalUnit AndAlso IsAdministrationGroup(dn)
			Catch ex As Exception
				Return False
			End Try
		End Function

		'''<summary>Liefert true, wenn DistinguishedName vom angegebenen Typ ist.</summary>
		Public Function IsType(ByVal dn As DistinguishedName, ByVal specialDistinguishedName As SpecialDistinguishedNameKeys) As Boolean

			Select Case specialDistinguishedName
			Case SpecialDistinguishedNameKeys.OrganizationGroups
				Return IsOrganizationGroup(dn)
			Case SpecialDistinguishedNameKeys.Grants
				Return IsGrantGroup(dn)
			Case SpecialDistinguishedNameKeys.Mappings
				Return IsMapping(dn)
			Case SpecialDistinguishedNameKeys.Roles
				Return IsRole(dn)
			Case SpecialDistinguishedNameKeys.Administration
				Return IsAdministrationGroup(dn)
			Case SpecialDistinguishedNameKeys.AccountingGroups
				Return IsAccountingGroup(dn)
			Case SpecialDistinguishedNameKeys.HolidayGroups
				Return IsHolidayGroup(dn)
			Case SpecialDistinguishedNameKeys.GroupManagers
				Return IsGroupManager(dn)
			Case SpecialDistinguishedNameKeys.WorkGroups
				Return IsWorkGroup(dn)
			Case SpecialDistinguishedNameKeys.ApplicationRoles
				Return IsApplicationRole(dn)
			Case SpecialDistinguishedNameKeys.BaseRoles
				Return IsBaseRole(dn)
			Case SpecialDistinguishedNameKeys.DenialRoles
				Return IsDenialRole(dn)
			Case SpecialDistinguishedNameKeys.DepartmentRoles
				Return IsDepartmentRole(dn)
			Case SpecialDistinguishedNameKeys.ExtraRoles
				Return IsExtraRole(dn)
			Case SpecialDistinguishedNameKeys.TeamRoles
				Return IsTeamRole(dn)
			Case Else
				Return False
			End Select
		End Function

#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace
