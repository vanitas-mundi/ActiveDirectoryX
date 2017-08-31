Imports SSP.ActiveDirectoryX.Grants

Public Enum Grants
    Execute = 0
End Enum

Public Class GrantsManager

    Inherits GrantsManagerBase(Of Grants)

	Public Sub New(ByVal grantsAppName As String, ByVal personId As Int64)
		MyBase.New(grantsAppName, personId)
	End Sub
End Class
