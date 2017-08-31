Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "

Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core


#End Region

Namespace Grants.Administration

	Public Class GrantsAdministration

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private Shared _instance As GrantsAdministration
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Shared Sub New()
			_instance = New GrantsAdministration
		End Sub

		Private Sub New()
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		Public Shared ReadOnly Property Instance As GrantsAdministration
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

#Region " --> CreateGrant "
	''' <summary>
	''' Legt eine neue Berechtigung zu einer bestehenden Applikation an.
	''' </summary>
	Public Function CreateGrant(ByVal appName As String, ByVal grantName As String) As AdManipulationResults

		Return CreateGrant(DistinguishedName.GetByOu(appName), grantName)
	End Function

	''' <summary>
	''' Legt eine neue Berechtigung zu einer bestehenden Applikation an.
	''' </summary>
	Public Function CreateGrant(ByVal appDn As DistinguishedName, ByVal grantName As String) As AdManipulationResults

		Return Administrations.Instance.CreateGroup _
		(appDn, String.Concat(appDn.Name, ".", grantName), SpecialDistinguishedNameKeys.Grants)
	End Function
#End Region

#Region " --> DeleteGrant "
	''' <summary>
	''' Löscht eine Berechtigung zu einer bestehenden Applikation.
	''' </summary>
	Public Function DeleteGrant(ByVal appName As String, ByVal grantName As String) As AdManipulationResults

		Return DeleteGrant(DistinguishedName.GetByGroupName(String.Concat(appName, ".", grantName)))
	End Function

	''' <summary>
	''' Löscht eine Berechtigung zu einer bestehenden Applikation.
	''' </summary>
	Public Function DeleteGrant(ByVal grantDn As DistinguishedName) As AdManipulationResults

		Return Administrations.Instance.DeleteGroup(grantDn, SpecialDistinguishedNameKeys.Grants)
	End Function
#End Region

#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace
