Option Explicit On
Option Infer On
Option Strict On


#Region " --------------->> Imports "

Imports System.DirectoryServices
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core.Manipulation
Imports SSP.ActiveDirectoryX.Core

#End Region

Namespace Grants.Administration

	Public Class OrganizationalUnitAdministration

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private _organizationalUnit As OrganizationalUnit
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Friend Sub New(ByVal organizationalUnit As OrganizationalUnit)
			_organizationalUnit = organizationalUnit
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
#End Region	'{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region	'{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region	'{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "

#Region " --> CreateOrganizationalUnitChild "
	''' <summary>
	''' Erstellt die untergeordnete Organisationseinheit organizationalUnitName.
	''' </summary>
	Public Function CreateOrganizationalUnitChild(ByVal organizationalUnitName As String) As AdManipulationResults

		Try
			AdOrganizationalUnits.CreateOrganizationalUnit _
			(_organizationalUnit.OrganizationalUnitDn, organizationalUnitName, True)
			Return AdManipulationResults.Successful
		Catch ex As UnauthorizedAccessException
			Return AdManipulationResults.AccesDenied
		Catch ea As DirectoryServicesCOMException
			Return AdManipulationResults.MemberAlreadyExist
		Catch ex As Exception
			Return AdManipulationResults.UnknownError
		End Try
	End Function

#End Region

#Region " --> CreateOrganizationalUnitChild "
	''' <summary>
	''' Löscht die untergeordnete Organisationseinheit organizationalUnitName.
	''' </summary>
	Public Function DeleteOrganizationalUnitChild(ByVal organizationalUnitName As String) As AdManipulationResults

		Return DeleteOrganizationalUnitChild(DistinguishedName.GetByOu(organizationalUnitName))
	End Function

	''' <summary>
	''' Löscht die untergeordnete Organisationseinheit organizationalUnitDn.
	''' </summary>
	Public Function DeleteOrganizationalUnitChild(ByVal organizationalUnitDn As DistinguishedName) As AdManipulationResults

		Select Case True
		Case Not AdministrationTypeResolver.Instance.IsOrganizationalUnit(organizationalUnitDn)
			Return AdManipulationResults.GroupIsNotOrganizationUnit
		Case Not organizationalUnitDn.ContainsDn(_organizationalUnit.OrganizationalUnitDn)
			Return AdManipulationResults.MemberNotExist
		Case Else
			Try
				AdOrganizationalUnits.DeleteOrganizationalUnit(organizationalUnitDn, True)
				Return AdManipulationResults.Successful
			Catch ex As UnauthorizedAccessException
				Return AdManipulationResults.AccesDenied
			Catch ex As Exception
				Return AdManipulationResults.UnknownError
			End Try
		End Select
	End Function

#End Region

#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace



