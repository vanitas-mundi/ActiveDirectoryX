Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core
Imports SSP.ActiveDirectoryX.Core.Manipulation
#End Region

Namespace Grants.Administration

	Public Class GrantUserAdministration

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private _user As GrantUser
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Friend Sub New(ByVal user As GrantUser)
			_user = user
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
#End Region	'{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region	'{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
		''' <summary>
		''' Fügt dem User eine Gruppe hinzu mit einem Check des Gruppentyps.
		''' </summary>
		Private Function AddToGroup(ByVal groupDn As DistinguishedName, ByVal groupTypDn As DistinguishedName) As AdManipulationResults
			Try
				Select Case True
				Case (groupDn Is Nothing) OrElse (Not groupDn.IsGroup)
					Return AdManipulationResults.IsNotGroup
				Case Not groupDn.ContainsDn(groupTypDn)
					Select Case True
					Case groupTypDn.Equals(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Roles))
						Return AdManipulationResults.GroupIsNotRole
					Case groupTypDn.Equals(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.OrganizationGroups))
						Return AdManipulationResults.GroupIsNotOrganizationGroup
					Case groupTypDn.Equals(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Administration))
						Return AdManipulationResults.GroupIsNotAdminstrationGroup
					Case Else
						Return AdManipulationResults.GroupIsNotDomainGroup
					End Select
				Case Else

					If Not groupDn.GetMembers.Where(Function(dn) dn.IsEqualTo(_user.UserDistinguishedName)).Any Then
						Try
							Dim userPrincipal = AdPrincipals.GetUserPrincipal _
							(_user.UserDistinguishedName.GetProperty(AdProperties.sAMAccountName.ToString).ToString)
							AdGroups.AddToGroup(groupDn.Name, userPrincipal, True)
							Return AdManipulationResults.Successful
						Catch ex As Exception
							Return AdManipulationResults.UnknownError
						End Try
					Else
						Return AdManipulationResults.MemberAlreadyExist
					End If
				End Select
			Catch ex As UnauthorizedAccessException
				Return AdManipulationResults.AccesDenied
			Catch ex As System.Exception
				Return AdManipulationResults.UnknownError
			End Try
		End Function

		''' <summary>
		''' Entfernt einen User aus einer Gruppe mit einem Check des Gruppentyps.
		''' </summary>
		Private Function RemoveFromGroup(ByVal groupDn As DistinguishedName, ByVal groupTypDn As DistinguishedName) As AdManipulationResults
			Try
				Select Case True
				Case (groupDn Is Nothing) OrElse (Not groupDn.IsGroup)
					Return AdManipulationResults.IsNotGroup
				Case Not groupDn.ContainsDn(groupTypDn)
					Select Case True
					Case groupTypDn.Equals(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Roles))
						Return AdManipulationResults.GroupIsNotRole
					Case groupTypDn.Equals(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.OrganizationGroups))
						Return AdManipulationResults.GroupIsNotOrganizationGroup
					Case groupTypDn.Equals(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Administration))
						Return AdManipulationResults.GroupIsNotAdminstrationGroup
					Case Else
						Return AdManipulationResults.GroupIsNotDomainGroup
					End Select
				Case Else

					If groupDn.GetMembers.Where(Function(dn) dn.IsEqualTo(_user.UserDistinguishedName)).Any Then
						Try
							Dim userPrincipal = AdPrincipals.GetUserPrincipal _
							(_user.UserDistinguishedName.GetProperty(AdProperties.sAMAccountName.ToString).ToString)
							AdGroups.RemoveFromGroup(groupDn.Name, userPrincipal, True)
							Return AdManipulationResults.Successful
						Catch ex As Exception
							Return AdManipulationResults.UnknownError
						End Try
					Else
						Return AdManipulationResults.MemberNotExist
					End If
				End Select
			Catch ex As UnauthorizedAccessException
				Return AdManipulationResults.AccesDenied
			Catch ex As System.Exception
				Return AdManipulationResults.UnknownError
			End Try
		End Function
#End Region	'{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "

#Region " --> SetProperty "
		''' <summary>
		''' Setzt eine AD-Eigenschaft des Benutzers.
		''' </summary>
		Public Function SetProperty(ByVal propertyName As AdProperties, ByVal values As Object()) As AdManipulationResults
			Try
				Using entry = _user.UserDistinguishedName.ToDirectoryEntry(True)
					entry.InvokeSet(propertyName.ToString, values)
					entry.CommitChanges()
					Return AdManipulationResults.Successful
				End Using
			Catch ex As UnauthorizedAccessException
				Return AdManipulationResults.AccesDenied
			Catch ex As System.Exception
				Return AdManipulationResults.UnknownError
			End Try
		End Function
#End Region

#Region " --> AddToRole "
		''' <summary>
		''' Fügt dem User eine Rolle hinzu.
		''' </summary>
		Public Function AddToRole(ByVal roleName As String) As AdManipulationResults
			Return AddToRole(DistinguishedName.GetByGroupName(roleName))
		End Function

		''' <summary>
		''' Fügt dem User eine Rolle hinzu.
		''' </summary>
		Public Function AddToRole(ByVal roleDn As DistinguishedName) As AdManipulationResults
			Try
				Return AddToGroup(roleDn, SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Roles))
			Catch ex As UnauthorizedAccessException
				Return AdManipulationResults.AccesDenied
			Catch ex As System.Exception
				Return AdManipulationResults.UnknownError
			End Try
		End Function
#End Region

#Region " --> RemoveFromRole "
		''' <summary>
		''' Entfernt einen User aus einer Rolle.
		''' </summary>
		Public Function RemoveFromRole(ByVal roleName As String) As AdManipulationResults
			Return RemoveFromRole(DistinguishedName.GetByGroupName(roleName))
		End Function

		''' <summary>
		''' Entfernt einen User aus einer Rolle.
		''' </summary>
		Public Function RemoveFromRole(ByVal roleDn As DistinguishedName) As AdManipulationResults
			Try
				Return RemoveFromGroup(roleDn, SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Roles))
			Catch ex As UnauthorizedAccessException
				Return AdManipulationResults.AccesDenied
			Catch ex As System.Exception
				Return AdManipulationResults.UnknownError
			End Try
		End Function
#End Region

#Region " --> AddToOrganizationGroup "
		''' <summary>
		''' Fügt dem User eine Organisations-Gruppe hinzu.
		''' </summary>
		Public Function AddToOrganizationGroup(ByVal organizationGroupName As String) As AdManipulationResults
			Return AddToOrganizationGroup(DistinguishedName.GetByGroupName(organizationGroupName))
		End Function

		''' <summary>
		''' Fügt dem User eine Organisations-Gruppe hinzu.
		''' </summary>
		Public Function AddToOrganizationGroup(ByVal organizationGroupDn As DistinguishedName) As AdManipulationResults
			Return AddToGroup(organizationGroupDn, SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.OrganizationGroups))
		End Function
#End Region

#Region " --> RemoveFromOrganizationGroup "
		''' <summary>
		''' Entfernt einen User aus einer Organisations-Gruppe.
		''' </summary>
		Public Function RemoveFromOrganizationGroup(ByVal organizationGroupName As String) As AdManipulationResults
			Return RemoveFromOrganizationGroup(DistinguishedName.GetByGroupName(organizationGroupName))
		End Function

		''' <summary>
		''' Entfernt einen User aus einer Organisations-Gruppe.
		''' </summary>
		Public Function RemoveFromOrganizationGroup(ByVal organizationGroupDn As DistinguishedName) As AdManipulationResults
			Return RemoveFromGroup(organizationGroupDn, SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.OrganizationGroups))
		End Function
#End Region

#Region " --> AddToAdministrationGroup "
		''' <summary>
		''' Fügt dem User eine Verwaltungs-Gruppe hinzu.
		''' </summary>
		Public Function AddToAdministrationGroup(ByVal administrationGroupName As String) As AdManipulationResults
			Return AddToAdministrationGroup(DistinguishedName.GetByGroupName(administrationGroupName))
		End Function

		''' <summary>
		''' Fügt dem User eine Verwaltungs-Gruppe hinzu.
		''' </summary>
		Public Function AddToAdministrationGroup(ByVal administrationGroupDn As DistinguishedName) As AdManipulationResults
			Return AddToGroup(administrationGroupDn, SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Administration))
		End Function
#End Region

#Region " --> RemoveFromAdministrationGroup "
		''' <summary>
		''' Entfernt den User aus einer Verwaltungs-Gruppe.
		''' </summary>
		Public Function RemoveFromAdministrationGroup(ByVal administrationGroupName As String) As AdManipulationResults
			Return RemoveFromAdministrationGroup(DistinguishedName.GetByGroupName(administrationGroupName))
		End Function

		''' <summary>
		''' Entfernt den User aus einer Verwaltungs-Gruppe.
		''' </summary>
		Public Function RemoveFromAdministrationGroup(ByVal administrationGroupDn As DistinguishedName) As AdManipulationResults
			Return RemoveFromGroup(administrationGroupDn, SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Administration))
		End Function
#End Region

#Region " --> AddToGroup "
		''' <summary>
		''' Fügt dem User eine allgemeine Domänen-Gruppe hinzu.
		''' </summary>
		Public Function AddToGroup(ByVal groupName As String) As AdManipulationResults
			Return AddToGroup(DistinguishedName.GetByGroupName(groupName))
		End Function

		''' <summary>
		''' Fügt dem User eine allgemeine Domänen-Gruppe hinzu.
		''' </summary>
		Public Function AddToGroup(ByVal groupDn As DistinguishedName) As AdManipulationResults
			Try
				Return AddToGroup(groupDn, SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Domain))
			Catch ex As UnauthorizedAccessException
				Return AdManipulationResults.AccesDenied
			Catch ex As System.Exception
				Return AdManipulationResults.UnknownError
			End Try
		End Function
#End Region

#Region " --> RemoveFromGroup "
		''' <summary>
		''' Entfernt dem User aus einer allgemeinen Domänen-Gruppe.
		''' </summary>
		Public Function RemoveFromGroup(ByVal groupName As String) As AdManipulationResults
			Return RemoveFromGroup(DistinguishedName.GetByGroupName(groupName))
		End Function

		''' <summary>
		''' Entfernt dem User aus einer allgemeinen Domänen-Gruppe.
		''' </summary>
		Public Function RemoveFromGroup(ByVal groupDn As DistinguishedName) As AdManipulationResults
			Try
				Return RemoveFromGroup(groupDn, SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Domain))
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
		''' Fügt den User einem Mapping hinzu. Mappings sollten besser an Rollen gebunden werden!
		''' </summary>
		Public Function AddToMapping(ByVal mappingDn As DistinguishedName) As AdManipulationResults

			Select Case True
			Case (mappingDn Is Nothing) OrElse (Not AdministrationTypeResolver.Instance.IsMapping(mappingDn))
				Return AdManipulationResults.GroupIsNotMapping
			Case Else
				Return Me.AddToGroup(mappingDn.Name)
			End Select
		End Function
#End Region

#Region " --> RemoveFromMapping "
		''' <summary>
		''' Entzieht einem User das angegebene Mapping.
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
				Return Me.RemoveFromGroup(mappingDn.Name)
			End Select
		End Function
#End Region

#Region " --> AddToGrant - ACHTUNG!!! Es sollte AddToRole verwendet werden, da Individualberechtigungen entfallen sollen."
		''' <summary>
		''' ACHTUNG!!! Es sollte AddToRole verwendet werden, da Individualberechtigungen entfallen sollen.
		''' Gewährt die übergebene Individualberechtigung.
		''' </summary>
		<Obsolete("ACHTUNG!!! Es sollte AddToRole verwendet werden, da Individualberechtigungen entfallen sollen.")> _
		Public Function AddToGrant(ByVal appName As String, ByVal grantName As String) As AdManipulationResults
			Return AddToGrant(DistinguishedName.GetByGrant(appName, grantName))
		End Function

		''' <summary>
		''' ACHTUNG!!! Es sollte AddToRole verwendet werden, da Individualberechtigungen entfallen sollen.
		''' Gewährt die übergebene Individualberechtigung.
		''' </summary>
		<Obsolete("ACHTUNG!!! Es sollte AddToRole verwendet werden, da Individualberechtigungen entfallen sollen.")> _
		Public Function AddToGrant(ByVal grantDn As DistinguishedName) As AdManipulationResults

			Select Case True
			Case (grantDn Is Nothing) OrElse (Not grantDn.ContainsDn _
			(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Grants)))
				Return AdManipulationResults.GroupIsNotGrant
			Case Else
				Return Me.AddToGroup(grantDn.Name)
			End Select
		End Function
#End Region

#Region " --> RemoveFromGrant "
		''' <summary>
		''' Entzieht einem User die angegebne Individualberechtigung.
		''' </summary>
		Public Function RemoveFromGrant(ByVal appName As String, ByVal grantName As String) As AdManipulationResults
			Return RemoveFromGrant(DistinguishedName.GetByGrant(appName, grantName))
		End Function

		''' <summary>
		''' Entzieht einem User die angegebne Individualberechtigung.
		''' </summary>
		Public Function RemoveFromGrant(ByVal grantDn As DistinguishedName) As AdManipulationResults

			Select Case True
			Case (grantDn Is Nothing) OrElse (Not grantDn.ContainsDn _
			(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Grants)))
				Return AdManipulationResults.GroupIsNotGrant
			Case Else
				Return Me.RemoveFromGroup(grantDn.Name)
			End Select
		End Function
#End Region

#Region " --> RemoveFromAllGroups "
		''' <summary>
		''' Entzieht einem User alle Gruppenzuweisungen.
		''' </summary>
		Public Function RemoveFromAllGroups() As AdManipulationResultsErrorsValue

			Try

				Dim groups = _user.UserDistinguishedName.GetMemberOf.Where _
				(Function(group) Not group.Name.ToLower.Contains("grvdi")).ToList

				Dim errors = New List(Of AdManipulationResultsErrorValue)

				For Each group In groups
					Try
						Dim result = RemoveFromGroup(group)
						If Not result = AdManipulationResults.Successful Then
							errors.Add(New AdManipulationResultsErrorValue(group, New Exception(result.ToString)))
						End If
					Catch ex As Exception
						errors.Add(New AdManipulationResultsErrorValue(group, ex))
					End Try
				Next group

				Return If(errors.Any, New AdManipulationResultsErrorsValue(AdManipulationResults.SuccessfulWithSomeErrors, errors) _
				, New AdManipulationResultsErrorsValue(AdManipulationResults.Successful))

			Catch ex As UnauthorizedAccessException
				Return New AdManipulationResultsErrorsValue(AdManipulationResults.AccesDenied)
			Catch ex As System.Exception
				Return New AdManipulationResultsErrorsValue(AdManipulationResults.UnknownError)
			End Try
		End Function
#End Region

#Region " --> CloneAllGroupsFromUser "
		''' <summary>
		''' Klont alle Gruppen des Users fromUserName auf User.
		''' </summary>
		Public Function CloneAllGroupsFromUser(ByVal fromUserName As String) As AdManipulationResultsErrorsValue

			Return CloneAllGroupsFromUser(DistinguishedName.GetByUserName(fromUserName))
		End Function

		''' <summary>
		''' Klont alle Gruppen des Users fromUserDn auf User.
		''' </summary>
		Public Function CloneAllGroupsFromUser(ByVal fromUserDn As DistinguishedName) As AdManipulationResultsErrorsValue

			Try
				Dim groups = fromUserDn.GetMemberOf.Where(Function(group) Not group.Name.ToLower.Contains("grvdi")).ToList
				Dim errors = New List(Of AdManipulationResultsErrorValue)

				For Each group In groups
					Try
						Dim result = AddToGroup(group)
						If Not result = AdManipulationResults.Successful Then
							errors.Add(New AdManipulationResultsErrorValue(group, New Exception(result.ToString)))
						End If
					Catch ex As Exception
						errors.Add(New AdManipulationResultsErrorValue(group, ex))
					End Try
				Next group

				Return If(errors.Any, New AdManipulationResultsErrorsValue(AdManipulationResults.SuccessfulWithSomeErrors, errors) _
				, New AdManipulationResultsErrorsValue(AdManipulationResults.Successful))

			Catch ex As UnauthorizedAccessException
				Return New AdManipulationResultsErrorsValue(AdManipulationResults.AccesDenied)
			Catch ex As System.Exception
				Return New AdManipulationResultsErrorsValue(AdManipulationResults.UnknownError)
			End Try
		End Function

#End Region

#Region " --> CloneAllGroupsToUser "
		''' <summary>
		''' Klont alle Gruppen des Users auf toUserName.
		''' </summary>
		Public Function CloneAllGroupsToUser(ByVal toUserName As String) As AdManipulationResultsErrorsValue
			Return CloneAllGroupsToUser(DistinguishedName.GetByUserName(toUserName))
		End Function

		''' <summary>
		''' Klont alle Gruppen des Users auf toUserDn.
		''' </summary>
		Public Function CloneAllGroupsToUser(ByVal toUserDn As DistinguishedName) As AdManipulationResultsErrorsValue

			Try
				Dim toGrantUser = New GrantUser(toUserDn)
				Dim groups = _user.UserDistinguishedName.GetMemberOf.Where(Function(group) Not group.Name.ToLower.Contains("grvdi")).ToList
				Dim errors = New List(Of AdManipulationResultsErrorValue)

				For Each group In groups
					Try
						Dim result = toGrantUser.Administration.AddToGroup(group)
						If Not result = AdManipulationResults.Successful Then
							errors.Add(New AdManipulationResultsErrorValue(group, New Exception(result.ToString)))
						End If
					Catch ex As Exception
						errors.Add(New AdManipulationResultsErrorValue(group, ex))
					End Try
				Next group

				Return If(errors.Any, New AdManipulationResultsErrorsValue(AdManipulationResults.SuccessfulWithSomeErrors, errors) _
				, New AdManipulationResultsErrorsValue(AdManipulationResults.Successful))

			Catch ex As UnauthorizedAccessException
				Return New AdManipulationResultsErrorsValue(AdManipulationResults.AccesDenied)
			Catch ex As System.Exception
				Return New AdManipulationResultsErrorsValue(AdManipulationResults.UnknownError)
			End Try
		End Function
#End Region

#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace



