Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core
#End Region

Namespace Grants.Administration

	Public Class GroupManagersAdministration

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private Shared _instance As GroupManagersAdministration
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Shared Sub New()
			_instance = New GroupManagersAdministration
		End Sub

		Private Sub New()
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		Public Shared ReadOnly Property Instance As GroupManagersAdministration
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

#Region " --> CreateGroupManager "
	''' <summary>
	''' Erstellt ein neues Gruppenmanager-Objekt.
	''' </summary>
	Public Function CreateGroupManager(ByVal groupManagerName As String) As AdManipulationResults

		Try
			groupManagerName = groupManagerName.ToLower

        If Not groupManagerName.StartsWith(My.Settings.GroupManagerPrefix) Then
          Return AdManipulationResults.InvalidGroupManagerName
        End If

        Dim groupManagersOuDn = SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.GroupManagers)
			Return Administrations.Instance.CreateGroup(groupManagersOuDn, groupManagerName, SpecialDistinguishedNameKeys.GroupManagers)
		Catch ex As UnauthorizedAccessException
			Return AdManipulationResults.AccesDenied
		Catch ex As System.Exception
			Return AdManipulationResults.UnknownError
		End Try
	End Function
#End Region

#Region " --> DeleteGroupManager "
	''' <summary>
	''' Löscht die angegebene Managergruppe aus dem Active Directory.
	''' </summary>
	Public Function DeleteGroupManager(ByVal groupManagerName As String) As AdManipulationResults
		 Return DeleteGroupManager(DistinguishedName.GetByGroupName(groupManagerName))
	End Function

	''' <summary>
	''' Löscht die angegebene Managergruppe aus dem Active Directory.
	''' </summary>
	Public Function DeleteGroupManager(ByVal groupManagerDn As DistinguishedName) As AdManipulationResults

		Return Administrations.Instance.DeleteGroup(groupManagerDn, SpecialDistinguishedNameKeys.GroupManagers)
	End Function
#End Region

#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace


