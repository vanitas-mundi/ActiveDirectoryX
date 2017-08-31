Option Explicit On
Option Infer On
Option Strict On

Namespace Core.Enums
	Public Enum LoginResults
		''' <summary>Login erfolgreich</summary>
		Successful = 0
		''' <summary>AppName unbekannt</summary>
		InvalidAppName = 1
		''' <summary>Benutzer unbekannt.</summary>
		InvalidUserName = 2
		''' <summary>Kennwort invalide.</summary>
		InvalidPwd = 3
		''' <summary>Account wurde deaktiviert</summary>
		AccountDeactivated = 4
		''' <summary>Account abgelaufen</summary>
		AccountExpired = 5
	End Enum
End Namespace