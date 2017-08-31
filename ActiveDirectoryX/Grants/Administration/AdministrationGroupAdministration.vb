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

	Public Class AdministrationGroupAdministration

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private _group As AdministrationGroup
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Friend Sub New(ByVal group As AdministrationGroup)
			_group = group
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
#End Region	'{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region	'{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region	'{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "

#Region " --> AddUser "
	''' <summary>
	''' Fügt dem Gruppen-Objekt einen User hinzu.
	''' </summary>
	Public Function AddUser(ByVal personId As Int64) As AdManipulationResults
		Return AddUser(DistinguishedName.GetByPersonId(personId))
	End Function

	''' <summary>
	''' Fügt dem Gruppen-Objekt einen User hinzu.
	''' </summary>
	Public Function AddUser(ByVal userName As String) As AdManipulationResults
		Return AddUser(DistinguishedName.GetByUserName(userName))
	End Function

	''' <summary>
	''' Fügt dem Gruppen-Objekt einen User hinzu.
	''' </summary>
	Public Function AddUser(ByVal userDn As DistinguishedName) As AdManipulationResults
		Try
			Select Case True
			Case Not userDn.IsUser
				Return AdManipulationResults.IsNotUser
			Case _group.Members.Where(Function(dn) dn.IsEqualTo(userDn)).Any
				Return AdManipulationResults.MemberAlreadyExist
			Case Else
				Dim userPrincipal = AdPrincipals.GetUserPrincipal(userDn.GetProperty(AdProperties.sAMAccountName.ToString).ToString)
				AdGroups.AddToGroup(_group.Name, userPrincipal, True)
				Return AdManipulationResults.Successful
			End Select
		Catch ex As PrincipalExistsException
			Return AdManipulationResults.MemberAlreadyExist
		Catch ex As ArgumentNullException
			Return AdManipulationResults.UnknownUserPrincipal
		Catch ex As NullReferenceException
			Return AdManipulationResults.UnknownGroupPrincipal
		Catch ex As UnauthorizedAccessException
			Return AdManipulationResults.AccesDenied
		Catch ex As System.Exception
			Return AdManipulationResults.UnknownError
		End Try
	End Function
#End Region

#Region " --> RemoveUser "
	''' <summary>
	''' Entfernt ein User-Objekt aus der Gruppe.
	''' </summary>
	Public Function RemoveUser(ByVal personId As Int64) As AdManipulationResults
		Return RemoveUser(DistinguishedName.GetByPersonId(personId))
	End Function

	''' <summary>
	''' Entfernt ein User-Objekt aus der Gruppe.
	''' </summary>
	Public Function RemoveUser(ByVal userName As String) As AdManipulationResults
		Return RemoveUser(DistinguishedName.GetByUserName(userName))
	End Function

	''' <summary>
	''' Entfernt ein User-Objekt aus der Gruppe.
	''' </summary>
	Public Function RemoveUser(ByVal userDn As DistinguishedName) As AdManipulationResults
		Try
			Select Case True
			Case Not userDn.IsUser
				Return AdManipulationResults.IsNotUser
			Case Not _group.Members.Where(Function(dn) dn.IsEqualTo(userDn)).Any
				Return AdManipulationResults.MemberNotExist
			Case Else
				Dim userPrincipal = AdPrincipals.GetUserPrincipal _
				(userDn.GetProperty(AdProperties.sAMAccountName.ToString).ToString)

				AdGroups.RemoveFromGroup(_group.Name, userPrincipal, True)
				Return AdManipulationResults.Successful
			End Select
		Catch ex As ArgumentNullException
			Return AdManipulationResults.UnknownUserPrincipal
		Catch ex As NullReferenceException
			Return AdManipulationResults.UnknownGroupPrincipal
		Catch ex As UnauthorizedAccessException
			Return AdManipulationResults.AccesDenied
		Catch ex As System.Exception
			Return AdManipulationResults.UnknownError
		End Try
	End Function
#End Region

#Region " --> AddGroup "
	''' <summary>
	''' Fügt dem Gruppen-Objekt eine Gruppe hinzu.
	''' </summary>
	Public Function AddGroup(ByVal groupName As String) As AdManipulationResults
		Return AddGroup(DistinguishedName.GetByGroupName(groupName))
	End Function

	''' <summary>
	''' Fügt dem Gruppen-Objekt eine Gruppe hinzu.
	''' </summary>
	Public Function AddGroup(ByVal groupDn As DistinguishedName) As AdManipulationResults
		Try
			Select Case True
			Case Not groupDn.IsGroup
				Return AdManipulationResults.IsNotGroup
			Case _group.Members.Where(Function(dn) dn.IsEqualTo(groupDn)).Any
				Return AdManipulationResults.MemberAlreadyExist
			Case Else
				Dim groupPrincipal = AdPrincipals.GetGroupPrincipal(groupDn.Name)
				AdGroups.AddToGroup(_group.Name, groupPrincipal, True)
				Return AdManipulationResults.Successful
			End Select
		Catch ex As PrincipalExistsException
			Return AdManipulationResults.MemberAlreadyExist
		Catch ex As ArgumentNullException
			Return AdManipulationResults.UnknownGroupPrincipal
		Catch ex As NullReferenceException
			Return AdManipulationResults.UnknownGroupPrincipal
		Catch ex As UnauthorizedAccessException
			Return AdManipulationResults.AccesDenied
		Catch ex As System.Exception
			Return AdManipulationResults.UnknownError
		End Try
	End Function
#End Region

#Region " --> RemoveGroup "
	''' <summary>
	''' Entfernt ein Gruppen-Objekt aus der Gruppe.
	''' </summary>
	Public Function RemoveGroup(ByVal groupName As String) As AdManipulationResults
		Return RemoveGroup(DistinguishedName.GetByGroupName(groupName))
	End Function

	''' <summary>
	''' Entfernt ein Gruppen-Objekt aus der Gruppe.
	''' </summary>
	Public Function RemoveGroup(ByVal groupDn As DistinguishedName) As AdManipulationResults
		Try
			Select Case True
			Case Not groupDn.IsGroup
				Return AdManipulationResults.IsNotGroup
			Case Not _group.Members.Where(Function(dn) dn.IsEqualTo(groupDn)).Any
				Return AdManipulationResults.MemberNotExist
			Case Else
				Dim groupPrincipal = AdPrincipals.GetUserPrincipal(groupDn.Name)
				AdGroups.RemoveFromGroup(_group.Name, groupPrincipal, True)
				Return AdManipulationResults.Successful
			End Select
		Catch ex As ArgumentNullException
			Return AdManipulationResults.UnknownUserPrincipal
		Catch ex As NullReferenceException
			Return AdManipulationResults.UnknownGroupPrincipal
		Catch ex As UnauthorizedAccessException
			Return AdManipulationResults.AccesDenied
		Catch ex As System.Exception
			Return AdManipulationResults.UnknownError
		End Try
	End Function
#End Region

#Region " --> AddToGrant "
	''' <summary>
	''' Fügt einer Berechtigung eine Rolle hinzu.
	''' </summary>
	Public Function AddToGrant(ByVal appName As String, ByVal grantName As String) As AdManipulationResults
		Return AddToGrant(DistinguishedName.GetByGrant(appName, grantName))
	End Function

	''' <summary>
	''' Fügt einer Berechtigung eine Rolle hinzu.
	''' </summary>
	Public Function AddToGrant(ByVal grantDn As DistinguishedName) As AdManipulationResults

		Dim roleDn = _group.GroupDistinguishedName
		Select Case True
		Case (grantDn Is Nothing) OrElse (Not grantDn.ContainsDn _
		(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Grants)))
			Return AdManipulationResults.GroupIsNotGrant
		Case (roleDn Is Nothing) OrElse (Not AdministrationTypeResolver.Instance.IsRole(roleDn))
			Return AdManipulationResults.GroupIsNotRole
		Case Else
			Dim grant = New AdministrationGroup(grantDn)
			Return grant.Administration.AddGroup(roleDn)
		End Select
	End Function
#End Region

#Region " --> RemoveRoleFromGrant "
	''' <summary>
	''' Entfernt eine Rolle aus einer Berechtigung.
	''' </summary>
	Public Function RemoveFromGrant(ByVal appName As String, ByVal grantName As String) As AdManipulationResults
		Return RemoveFromGrant(DistinguishedName.GetByGrant(appName, grantName))
	End Function

	''' <summary>
	''' Entfernt eine Rolle aus einer Berechtigung.
	''' </summary>
	Public Function RemoveFromGrant(ByVal grantDn As DistinguishedName) As AdManipulationResults

		Dim roleDn = _group.GroupDistinguishedName

		Select Case True
		Case (grantDn Is Nothing) OrElse (Not grantDn.ContainsDn _
		(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Grants)))
			Return AdManipulationResults.GroupIsNotGrant
		Case (roleDn Is Nothing) OrElse (Not AdministrationTypeResolver.Instance.IsRole(roleDn))
			Return AdManipulationResults.GroupIsNotRole
		Case Else
			Dim grant = New AdministrationGroup(grantDn)
			Return grant.Administration.RemoveGroup(roleDn)
		End Select
	End Function
#End Region

#Region " --> SetGroupManager "
	''' <summary>
	''' Fügt dem Gruppen-Objekt ein GroupManager-Objekt hinzu.
	''' </summary>
	Public Function SetGroupManager(ByVal groupManager As GroupManager) As AdManipulationResults
		Return SetGroupManager(_group, groupManager)
	End Function

	''' <summary>
	''' Fügt dem Gruppen-Objekt ein GroupManager-Objekt hinzu. Wird für groupManager Nothing übergeben
	''' so wird der Gruppenmanager zurückgesetzt.
	''' </summary>
	Public Function SetGroupManager(ByVal group As AdministrationGroup, ByVal groupManager As GroupManager) As AdManipulationResults

		Try
			Dim entry = group.GroupDistinguishedName.ToDirectoryEntry(True)

			If groupManager Is Nothing Then
				entry.Properties.Item(AdProperties.managedBy.ToString).Value = Nothing
			Else
				entry.InvokeSet(AdProperties.managedBy.ToString, New Object() {groupManager.ManagerGroupDn.Value})
			End If
			entry.CommitChanges()
			Return AdManipulationResults.Successful
		Catch ex As UnauthorizedAccessException
			Return AdManipulationResults.AccesDenied
		Catch ex As System.Exception
			Return AdManipulationResults.UnknownError
		End Try
	End Function
#End Region

#Region " --> AddToMapping "
		''' <summary>
		''' Fügt den User einem Mapping hinzu. Mappings sollten besser an Rollen gebunden werden!
		''' </summary>
		Public Function AddToMapping(ByVal mappingName As String) As AdManipulationResults
			Return AddToMapping(DistinguishedName.GetByGroupName(mappingName))
		End Function

		''' <summary>
		''' Fügt die Gruppe einem Mapping hinzu.
		''' </summary>
		Public Function AddToMapping(ByVal mappingDn As DistinguishedName) As AdManipulationResults

			Select Case True
			Case (mappingDn Is Nothing) OrElse (Not AdministrationTypeResolver.Instance.IsMapping(mappingDn))
				Return AdManipulationResults.GroupIsNotMapping
			Case Else
				Return Me.AddGroup(mappingDn.Name)
			End Select
		End Function
#End Region

#Region " --> RemoveFromMapping "
		''' <summary>
		''' Entzieht die Gruppe das angegebene Mapping.
		''' </summary>
		Public Function RemoveFromMapping(ByVal mappingName As String) As AdManipulationResults
			Return RemoveFromMapping(DistinguishedName.GetByGroupName(mappingName))
		End Function

		''' <summary>
		''' Entzieht einem User das angegebene Mapping.
		''' </summary>
		Public Function RemoveFromMapping(ByVal mappingDn As DistinguishedName) As AdManipulationResults

			Select Case True
			Case (mappingDn Is Nothing) OrElse (Not AdministrationTypeResolver.Instance.IsMapping(mappingDn))
				Return AdManipulationResults.GroupIsNotMapping
			Case Else
				Return Me.RemoveGroup(mappingDn.Name)
			End Select
		End Function
#End Region

#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace


