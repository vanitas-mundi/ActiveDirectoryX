Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports System.DirectoryServices.AccountManagement
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core.Manipulation
Imports SSP.ActiveDirectoryX.Core

#End Region

Namespace Grants.Administration

	Public Class GrantUsersAdministration

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private Shared _instance As GrantUsersAdministration
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Shared Sub New()
			_instance = New GrantUsersAdministration
		End Sub

		Private Sub New()
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		Public Shared ReadOnly Property Instance As GrantUsersAdministration
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

#Region " --> CreateDomainUser "
	Public Function CreateDomainUser(ByVal info As AdUserInfo) As AdManipulationResults
		Return AdUsers.CreateDomainUser(info, True)
	End Function
#End Region

#Region " --> DeleteDomainUser "
	''' <summary>
	''' Löscht den angegebenen User aus dem Active Directory.
	''' </summary>
	Public Function DeleteDomainUser(ByVal userName As String) As AdManipulationResults
		 Return DeleteDomainUser(DistinguishedName.GetByUserName(userName))
	End Function

	''' <summary>
	''' Löscht den angegebenen User aus dem Active Directory.
	''' </summary>
	Public Function DeleteDomainUser(ByVal personId As Int64) As AdManipulationResults
		 Return DeleteDomainUser(DistinguishedName.GetByPersonId(personId))
	End Function

	''' <summary>
	''' Löscht den angegebenen User aus dem Active Directory.
	''' </summary>
	Public Function DeleteDomainUser(ByVal userDn As DistinguishedName) As AdManipulationResults
		Try
			AdUsers.DeleteDomainUser(userDn, True)
			Return AdManipulationResults.Successful
		Catch ex As NullReferenceException
			Return AdManipulationResults.MemberNotExist
		Catch ex As UnauthorizedAccessException
			Return AdManipulationResults.AccesDenied
		Catch ex As System.Exception
			Return AdManipulationResults.UnknownError
		End Try
	End Function
#End Region

#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace
