Option Explicit On
Option Infer On
Option Strict On

Namespace Grants.Enums

	Public Enum GrantTypes
		''' <summary>Gruppe ist keine Rolle</summary>
		NoRole = 0
		'''<summary>Allgemeine Rolle</summary>
		CommonRole = 1
		'''<summary>Abteilungsrolle</summary>
		DepartmentRole = 2
		'''<summary>Applikationsrolle</summary>
		ApplicationRole = 3
		'''<summary>Basisrolle</summary>
		BaseRole = 4
		'''<summary>Extrarolle</summary>
		ExtraRole = 5
		'''<summary>Reamrolle</summary>
		TeamRole = 6
		'''<summary>Verweigerungsrolle</summary>
		DenialRole = 7
		'''<summary>Gruppenmanager</summary>
		GroupManager = 8
		'''<summary>Mapping</summary>
		Mapping = 9
	End Enum

End Namespace

