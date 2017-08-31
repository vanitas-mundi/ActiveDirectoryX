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

	Public Class AdministrationGroupsAdministration

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
Private Shared _instance As AdministrationGroupsAdministration
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Shared Sub New()
			_instance = New AdministrationGroupsAdministration
		End Sub

		Private Sub New()
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		Public Shared ReadOnly Property Instance As AdministrationGroupsAdministration
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

#Region " --> CreateAdministrationGroup "
	Public Function CreateAdministrationGroup(ByVal parentOrganizationalUnitName As String _
	, ByVal groupName As String) As AdManipulationResults

		Return CreateAdministrationGroup(DistinguishedName.GetByOu(parentOrganizationalUnitName), groupName)
	End Function

	Public Function CreateAdministrationGroup(ByVal parentOrganizationalUnitDn As DistinguishedName _
	, ByVal groupName As String) As AdManipulationResults

		Return Administrations.Instance.CreateGroup(parentOrganizationalUnitDn _
		, groupName, SpecialDistinguishedNameKeys.Administration)
	End Function
#End Region

#Region " --> DeleteAdministrationGroup "
	Public Function DeleteAdministrationGroup(ByVal groupName As String) As AdManipulationResults

		Return DeleteAdministrationGroup(DistinguishedName.GetByGroupName(groupName))
	End Function

	Public Function DeleteAdministrationGroup(ByVal grouptDn As DistinguishedName) As AdManipulationResults

		Return Administrations.Instance.DeleteGroup(grouptDn, SpecialDistinguishedNameKeys.Administration)
	End Function
#End Region

#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace



