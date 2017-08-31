Option Explicit On
Option Infer On
Option Strict On

Namespace Grants.Interfaces

	Public Interface IGrantTables

		ReadOnly Property UserName As String
		ReadOnly Property Item(ByVal index As Int32) As GrantTable
		ReadOnly Property Item(ByVal appName As String) As GrantTable
		Sub Fill()

	End Interface

End Namespace
