Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core.Manipulation
Imports SSP.ActiveDirectoryX.Core
#End Region

Namespace Grants.Administration

	Public Class GrantTablesAdministration

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private Shared _instance As GrantTablesAdministration
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Shared Sub New()
			_instance = New GrantTablesAdministration
		End Sub

		Private Sub New()
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		Public Shared ReadOnly Property Instance As GrantTablesAdministration
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

#Region " --> CreateApp "
	''' <summary>
	''' Legt eine neue Applikation an.
	''' </summary>
	Public Function CreateApp(ByVal appName As String, ByVal description As String) As AdManipulationResults

		Select Case True
		Case GrantTables.AppNameExists(appName)
			Return AdManipulationResults.MemberAlreadyExist
		Case Else
			Try
				AdOrganizationalUnits.CreateOrganizationalUnit(SpecialDistinguishedNames.Item _
				(SpecialDistinguishedNameKeys.Grants), appName.ToLower, description, True)
				Return AdManipulationResults.Successful

			Catch ex As UnauthorizedAccessException
				Return AdManipulationResults.AccesDenied
			Catch ex As System.Exception
				Return AdManipulationResults.UnknownError
			End Try
		End Select
	End Function
#End Region

#Region " --> DeleteApp "
	''' <summary>
	''' Löscht die angegebene Applikation.
	''' </summary>
	Public Function DeleteApp(ByVal appName As String) As AdManipulationResults

		 Return DeleteApp(DistinguishedName.GetByGroupName(appName.ToLower))
	End Function

	''' <summary>
	''' Löscht die angegebene Applikation.
	''' </summary>
	Public Function DeleteApp(ByVal appDn As DistinguishedName) As AdManipulationResults

		Select Case True
		Case (appDn Is Nothing) OrElse (Not appDn.ContainsDn _
		(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Grants)))
			Return AdManipulationResults.IsNotApplication
		Case Else
			Try
				AdOrganizationalUnits.DeleteOrganizationalUnit(appDn, True)
				Return AdManipulationResults.Successful
			Catch ex As NullReferenceException
				Return AdManipulationResults.MemberNotExist
			Catch ex As UnauthorizedAccessException
				Return AdManipulationResults.AccesDenied
			Catch ex As System.Exception
				Return AdManipulationResults.UnknownError
			End Try
		End Select
	End Function
#End Region

#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace


