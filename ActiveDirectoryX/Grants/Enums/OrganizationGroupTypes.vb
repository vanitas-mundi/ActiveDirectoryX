Option Explicit On
Option Infer On
Option Strict On

Namespace Grants.Enums
	Public Enum OrganizationGroupTypes
		'''<summary>Gruppe ist keine Organisationsgruppe</summary>
		NoOrganizationGroup = 0
		'''<summary>Allgemeine Organisationsgruppe</summary>
		CommonOrganizationGroup = 1
		'''<summary>Urlaubsgruppen</summary>
		HolidayGroup = 2
		'''<summary>Arbeitsgruppen</summary>
		WorkGroup = 3
		'''<summary>Abrechnungsgruppen</summary>
		AccountingGroup = 4
	End Enum
End Namespace

