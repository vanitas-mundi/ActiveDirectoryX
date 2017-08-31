Option Explicit On
Option Strict On
Option Infer On

#Region " --------------->> Imports/ usings "

Imports System.DirectoryServices.AccountManagement
Imports SSP.ActiveDirectoryX.Core.Enums

#End Region

Namespace Core.Manipulation

	Public Class AdGroups

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
#End Region	'{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region	'{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region	'{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
		''' <summary>
		''' Fügt den angegebenen principal der Gruppe groupName, als Member hinzu.
		''' </summary>
		Public Shared Sub AddToGroup(ByVal groupDn As DistinguishedName _
		, ByVal principal As Principal, ByVal useManipulationUser As Boolean)

			AddToGroup(groupDn.Name, principal, useManipulationUser)
		End Sub

		''' <summary>
		''' Fügt den angegebenen principal der Gruppe groupName, als Member hinzu.
		''' </summary>
		Public Shared Sub AddToGroup(ByVal groupName As String, ByVal principal As Principal, ByVal useManipulationUser As Boolean)

			Using group = AdPrincipals.GetGroupPrincipal(groupName, useManipulationUser)
				group.Members.Add(principal)
				group.Save()
			End Using
		End Sub

		''' <summary>
		''' Entfernt den principal aus den Gruppenmembers.
		''' </summary>
		Public Shared Sub RemoveFromGroup(ByVal groupDn As DistinguishedName _
		, ByVal principal As Principal, ByVal useManipulationUser As Boolean)

			RemoveFromGroup(groupDn.Name, principal, useManipulationUser)
		End Sub

		''' <summary>
		''' Entfernt den principal aus den Gruppenmembers.
		''' </summary>
		Public Shared Sub RemoveFromGroup(ByVal groupName As String _
		, ByVal principal As Principal, ByVal useManipulationUser As Boolean)

			Using group = AdPrincipals.GetGroupPrincipal(groupName, useManipulationUser)
				group.Members.Remove(principal)
				group.Save()
			End Using
		End Sub

		''' <summary>
		''' Legt die Gruppe in der angegebenen OU an.
		''' </summary>
		Public Shared Sub CreateGroup(ByVal parentOrganizationalUnitName As String _
		, ByVal groupName As String, ByVal useManipulationUser As Boolean)

			CreateGroup(DistinguishedName.GetByOu(parentOrganizationalUnitName), groupName, useManipulationUser)
		End Sub

		''' <summary>
		''' Legt die Gruppe in der angegebenen OU an.
		''' </summary>
		Public Shared Sub CreateGroup(ByVal parentOrganizationalUnitDn As DistinguishedName _
		, ByVal groupName As String, ByVal useManipulationUser As Boolean)

			Dim newGroupEntry = parentOrganizationalUnitDn.ToDirectoryEntry(useManipulationUser).Children.Add _
			("CN=" & groupName, ObjectClasses.group.ToString)
			newGroupEntry.Properties(AdProperties.sAMAccountName.ToString).Value = groupName
			newGroupEntry.CommitChanges()
		End Sub

		''' <summary>
		''' Löscht die angegebene Gruppe.
		''' </summary>
		Public Shared Sub DeleteGroup(ByVal groupName As String, ByVal useManipulationUser As Boolean)

			DeleteGroup(DistinguishedName.GetByGroupName(groupName), useManipulationUser)
		End Sub

		''' <summary>
		''' Löscht die angegebene Gruppe.
		''' </summary>
		Public Shared Sub DeleteGroup(ByVal groupDn As DistinguishedName, ByVal useManipulationUser As Boolean)

			Dim parentEntry = groupDn.Parent.ToDirectoryEntry(useManipulationUser)
			Dim groupEntry = groupDn.ToDirectoryEntry(useManipulationUser)
			parentEntry.Children.Remove(groupEntry)
			groupEntry.CommitChanges()
			parentEntry.CommitChanges()
		End Sub
#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace
