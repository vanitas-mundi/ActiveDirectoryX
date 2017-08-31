Option Explicit On
Option Infer On
Option Strict On

Namespace Grants.Enums

	Public Enum TreeBuilderRootLineTypes
		'''<summary>RootLines sind nicht vorhanden.</summary>
		None = 0
		'''<summary>RootLines werden mit Standardzeichen dargestellt.</summary>
		Regular = 1
		'''<summary>RootLines werden mit Tabulatoren dargestellt.</summary>
		Tabs = 2
		'''<summary>RootLines werden benutzerspezifisch dargestellt.</summary>
		Custom = 3
	End Enum

End Namespace

