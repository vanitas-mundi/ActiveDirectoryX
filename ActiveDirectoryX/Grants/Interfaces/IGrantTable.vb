Option Explicit On
Option Infer On
Option Strict On

Namespace Grants.Interfaces

	Public Interface IGrantTable

		ReadOnly Property Item(ByVal index As Int32) As Grant
		ReadOnly Property Item(ByVal grantName As String) As Grant
		ReadOnly Property GrantNames As String()
		ReadOnly Property AppName As String
		ReadOnly Property UserName As String

    Sub FillAll()
    Sub FillByGrantTree(ByVal gt As GrantTree)

    Function ToGrantString(ByVal grantDelimiter As String) As String
		Function ToGrantString() As String
	End Interface

End Namespace
