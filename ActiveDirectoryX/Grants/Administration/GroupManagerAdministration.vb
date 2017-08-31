Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core
#End Region

Namespace Grants.Administration

	Public Class GroupManagerAdministration

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private _groupManager As GroupManager
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Friend Sub New(ByVal groupManager As GroupManager)
			_groupManager = groupManager
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
#End Region	'{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region	'{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region	'{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "

#Region " --> SetManager "
	''' <summary>
	''' Setzt den Manager der Managergruppe.
	''' </summary>
	Public Function SetManager(ByVal personId As Int64) As AdManipulationResults
		Return SetManager(DistinguishedName.GetByPersonId(personId))
	End Function

	''' <summary>
	''' Setzt den Manager der Managergruppe.
	''' </summary>
	Public Function SetManager(ByVal userName As String) As AdManipulationResults
		Return SetManager(DistinguishedName.GetByUserName(userName))
	End Function

	''' <summary>
	''' Setzt den Manager der Managergruppe. Wird für UserDn Null übergeben
	''' wird der Manager zurückgesetzt.
	''' </summary>
	Public Function SetManager(ByVal userDn As DistinguishedName) As AdManipulationResults
		Try

			If userDn Is Nothing Then
				Dim g = New AdministrationGroup(_groupManager.ManagerGroupDn)
				Return g.Administration.SetGroupManager(Nothing)
			Else
				If Not userDn.IsUser Then Return AdManipulationResults.IsNotUser

				Dim entry = _groupManager.ManagerGroupDn.ToDirectoryEntry(True)
				entry.InvokeSet(AdProperties.managedBy.ToString, New Object() {userDn.Value})
				entry.CommitChanges()
				Return AdManipulationResults.Successful
			End If
		Catch ex As UnauthorizedAccessException
			Return AdManipulationResults.AccesDenied
		Catch ex As System.Exception
			Return AdManipulationResults.UnknownError
		End Try
	End Function
#End Region

#Region " --> AddDeputy "
	''' <summary>
	''' Fügt den übergebenen Stellvertreter der Managergruppe hinzu.
	''' </summary>
	Public Function AddDeputy(ByVal personId As Int64) As AdManipulationResults
		Return AddDeputy(DistinguishedName.GetByPersonId(personId))
	End Function

	''' <summary>
	''' Fügt den übergebenen Stellvertreter der Managergruppe hinzu.
	''' </summary>
	Public Function AddDeputy(ByVal userName As String) As AdManipulationResults
		Return AddDeputy(DistinguishedName.GetByUserName(userName))
	End Function

	''' <summary>
	''' Fügt den übergebenen Stellvertreter der Managergruppe hinzu.
	''' </summary>
	Public Function AddDeputy(ByVal userDn As DistinguishedName) As AdManipulationResults
		Try
			If (userDn Is Nothing) OrElse (Not userDn.IsUser) Then Return AdManipulationResults.IsNotUser
			Dim g = New AdministrationGroup(_groupManager.ManagerGroupDn)
			Return g.Administration.AddUser(userDn)
		Catch ex As UnauthorizedAccessException
			Return AdManipulationResults.AccesDenied
		Catch ex As System.Exception
			Return AdManipulationResults.UnknownError
		End Try
	End Function
#End Region

#Region " --> RemoveDeputy "
	''' <summary>
	''' Entfernt den übergebenen Stellvertreter aus der Managergruppe.
	''' </summary>
	Public Function RemoveDeputy(ByVal personId As Int64) As AdManipulationResults
		Return RemoveDeputy(DistinguishedName.GetByPersonId(personId))
	End Function

	''' <summary>
	''' Entfernt den übergebenen Stellvertreter aus der Managergruppe.
	''' </summary>
	Public Function RemoveDeputy(ByVal userName As String) As AdManipulationResults
		Return RemoveDeputy(DistinguishedName.GetByUserName(userName))
	End Function

	''' <summary>
	''' Entfernt den übergebenen Stellvertreter aus der Managergruppe.
	''' </summary>
	Public Function RemoveDeputy(ByVal userDn As DistinguishedName) As AdManipulationResults
		Try
			If (userDn Is Nothing) OrElse (Not userDn.IsUser) Then Return AdManipulationResults.IsNotUser
			Dim g = New AdministrationGroup(_groupManager.ManagerGroupDn)
			Return g.Administration.RemoveUser(userDn)
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


