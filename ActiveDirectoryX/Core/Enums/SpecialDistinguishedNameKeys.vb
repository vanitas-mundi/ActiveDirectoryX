Option Explicit On
Option Infer On
Option Strict On

Namespace Core.Enums

	Public Enum SpecialDistinguishedNameKeys
		''' <summary>
		''' Liefert den DistinguishedName der Domäne.
		''' Bsp: dc=domain, dc=net
		''' </summary>
		Domain = 0
		''' <summary>
		''' Liefert den DistinguishedName der Administrationsstruktur.
		''' Bsp: dc=domain, dc=net
		''' </summary>
		Administration = 1
		''' <summary>
		''' Liefert den DistinguishedName der Berechtigungs-OU.
		''' Bsp: ou=verwaltung, dc=domain, dc=net
		''' </summary>
		Grants = 2
		''' <summary>
		''' Liefert den DistinguishedName der Rollen-OU.
		''' Bsp: ou=Berechtigungen, ou=verwaltung, dc=domain, dc=net
		''' </summary>
		Roles = 3

		''' <summary>
		''' Liefert den DistinguishedName der Abteilungsrollen-OU.
		''' Bsp: ou=Abteilungsrollen, ou=Berechtigungen, ou=verwaltung, dc=domain, dc=net
		''' </summary>
		DepartmentRoles = 4
		''' <summary>
		''' Liefert den DistinguishedName der Applikationsrollen-OU.
		''' Bsp: ou=Applikationsrollen, ou=Berechtigungen, ou=verwaltung, dc=domain, dc=net
		''' </summary>
		ApplicationRoles = 5
		''' <summary>
		''' Liefert den DistinguishedName der Basisrollen-OU.
		''' Bsp: ou=Basisrollen, ou=Berechtigungen, ou=verwaltung, dc=domain, dc=net
		''' </summary>
		BaseRoles = 6
		''' <summary>
		''' Liefert den DistinguishedName der Extrarollen-OU.
		''' Bsp: ou=Extrarollen, ou=Berechtigungen, ou=verwaltung, dc=domain, dc=net
		''' </summary>
		ExtraRoles = 7
		''' <summary>
		''' Liefert den DistinguishedName der Teamrollen-OU.
		''' Bsp: ou=Teamrollen, ou=Berechtigungen, ou=verwaltung, dc=domain, dc=net
		''' </summary>
		TeamRoles = 8
		''' <summary>
		''' Liefert den DistinguishedName der Verweigerungsrollen-OU.
		''' Bsp: ou=Verweigerungsrollen, ou=Berechtigungen, ou=verwaltung, dc=domain, dc=net
		''' </summary>
		DenialRoles = 9
		''' <summary>
		''' Liefert den DistinguishedName der Organisations-OU.
		''' Bsp: ou=Organisation, ou=verwaltung, dc=domain, dc=net
		''' </summary>
		OrganizationGroups = 10
		''' <summary>
		''' Liefert den DistinguishedName der Urlaubsgruppen-OU.
		''' Bsp: ou=Rollendefinitionen, ou=verwaltung, dc=domain, dc=net
		''' </summary>
		HolidayGroups = 11
		''' <summary>
		''' Liefert den DistinguishedName der Arbeitsgruppen-OU.
		''' Bsp: ou=Arbeitsgruppen, ou=verwaltung, dc=domain, dc=net
		''' </summary>
		WorkGroups = 12
		''' <summary>
		''' Liefert den DistinguishedName der Abrechnungsgruppen-OU.
		''' Bsp: ou=Abrechnungsgruppen, ou=verwaltung, dc=domain, dc=net
		''' </summary>
		AccountingGroups = 13
		''' <summary>
		''' Liefert den DistinguishedName der Gruppenmanager-OU.
		''' Bsp: ou=Gruppenmanager, ou=verwaltung, dc=domain, dc=net
		''' </summary>
		GroupManagers = 14
		''' <summary>
		''' Liefert den DistinguishedName der Mappings.
		''' Bsp: ou=Mappings, ou=verwaltung, dc=domain, dc=net
		''' </summary>
		Mappings = 15
	End Enum

End Namespace
