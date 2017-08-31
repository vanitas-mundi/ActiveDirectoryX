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

	Public Class OrganizationalUnitsAdministration

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
Private Shared _instance As OrganizationalUnitsAdministration
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Shared Sub New()
			_instance = New OrganizationalUnitsAdministration
		End Sub

		Private Sub New()
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		Public Shared ReadOnly Property Instance As OrganizationalUnitsAdministration
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


#Region " --> CreateOrganizationalUnitChild "
	''' <summary>
	''' Erstellt die Organisationseinheit organizationalUnitName unter der Organisationseinheit parentOrganizationalUnitName.
	''' </summary>
	Public Function CreateOrganizationalUnit(ByVal parentOrganizationalUnitName As String _
	, ByVal organizationalUnitName As String) As AdManipulationResults

		Return CreateOrganizationalUnit(DistinguishedName.GetByOu _
		(parentOrganizationalUnitName), organizationalUnitName)
	End Function

	''' <summary>
	''' Erstellt die Organisationseinheit organizationalUnitName unter der Organisationseinheit parentOrganizationalUnitDn.
	''' </summary>
	Public Function CreateOrganizationalUnit(ByVal parentOrganizationalUnitDn As DistinguishedName _
	, ByVal organizationalUnitName As String) As AdManipulationResults

		If Not AdministrationTypeResolver.Instance.IsOrganizationalUnit(parentOrganizationalUnitDn) Then
			Return AdManipulationResults.GroupIsNotOrganizationUnit
		End If

		Try
			AdOrganizationalUnits.CreateOrganizationalUnit(parentOrganizationalUnitDn, organizationalUnitName, True)
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

#Region " --> DeleteOrganizationalUnitChild "
	''' <summary>
	''' Löscht die Organisationseinheit organizationalUnitName.
	''' </summary>
	Public Function DeleteOrganizationalUnitChild(ByVal organizationalUnitName As String) As AdManipulationResults

		Return DeleteOrganizationalUnitChild(DistinguishedName.GetByOu(organizationalUnitName))
	End Function

	''' <summary>
	''' Löscht die Organisationseinheit organizationalUnitDn.
	''' </summary>
	Public Function DeleteOrganizationalUnitChild(ByVal organizationalUnitDn As DistinguishedName) As AdManipulationResults

		If Not AdministrationTypeResolver.Instance.IsOrganizationalUnit(organizationalUnitDn) Then
			Return AdManipulationResults.GroupIsNotOrganizationUnit
		End If

		Try
			AdOrganizationalUnits.DeleteOrganizationalUnit(organizationalUnitDn, True)
			Return AdManipulationResults.Successful
		Catch ex As UnauthorizedAccessException
			Return AdManipulationResults.AccesDenied
		Catch ex As Exception
			Return AdManipulationResults.UnknownError
		End Try
	End Function
#End Region
'#End Region

#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace



