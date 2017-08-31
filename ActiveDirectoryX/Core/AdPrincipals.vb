Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "

Imports System.DirectoryServices.AccountManagement
Imports SSP.ActiveDirectoryX.Core.Enums

#End Region

Namespace Core

	Public Class AdPrincipals

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
		Public Shared Function GetPrincipalContext() As PrincipalContext

			Return GetPrincipalContext(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Domain).Value)
		End Function

		Public Shared Function GetPrincipalContext _
		(ByVal useManipulationUser As Boolean) As PrincipalContext

			If useManipulationUser Then
				Return New PrincipalContext(ContextType.Domain, Settings.Instance.DomainName _
				, Settings.Instance.ManipulationUserName, Settings.Instance.ManipulationUserPassword)
			Else
				Return AdPrincipals.GetPrincipalContext
			End If
		End Function

		Public Shared Function GetPrincipalContext _
		(ByVal distinguishedName As String) As PrincipalContext

      Return New PrincipalContext(ContextType.Domain, Settings.Instance.DomainName, distinguishedName)
    End Function

		Public Shared Function GetPrincipalContext _
		(ByVal distinguishedName As String, ByVal useManipulationUser As Boolean) As PrincipalContext

			If useManipulationUser Then
				Return New PrincipalContext(ContextType.Domain, Settings.Instance.DomainName _
				, distinguishedName, Settings.Instance.ManipulationUserName, Settings.Instance.ManipulationUserPassword)
			Else
				Return GetPrincipalContext(distinguishedName)
			End If
		End Function

		Public Shared Function GetPrincipalContext _
		(ByVal distinguishedName As String, ByVal userName As String, ByVal password As String) As PrincipalContext

			Return New PrincipalContext(ContextType.Domain, Settings.Instance.DomainName, distinguishedName, userName, password)
		End Function

		Public Shared Function GetPrincipalContext _
		(ByVal dn As DistinguishedName, ByVal useManipulationUser As Boolean) As PrincipalContext

			If useManipulationUser Then
				Return New PrincipalContext(ContextType.Domain, Settings.Instance.DomainName _
				, dn.Value, Settings.Instance.ManipulationUserName, Settings.Instance.ManipulationUserPassword)
			Else
				Return GetPrincipalContext(dn.Value)
			End If
		End Function

		Public Shared Function GetPrincipalContext _
		(ByVal dn As DistinguishedName, ByVal userName As String, ByVal password As String) As PrincipalContext

			Return New PrincipalContext(ContextType.Domain, Settings.Instance.DomainName, dn.Value, userName, password)
		End Function

		Public Shared Function GetPrincipalContext _
		(ByVal userName As String, ByVal password As String) As PrincipalContext

			Return New PrincipalContext(ContextType.Domain, Settings.Instance.DomainName, userName, password)
		End Function

		Public Shared Function ExistsUserPrincipal(ByVal userName As String) As Boolean

			Using u = New UserPrincipal(GetPrincipalContext)
				u.SamAccountName = userName
				Return DirectCast(New PrincipalSearcher(u).FindOne, UserPrincipal) IsNot Nothing
			End Using
		End Function

		Public Shared Function ExistsAdsGroupPrincipal(ByVal groupName As String) As Boolean

			Using g = New GroupPrincipal(GetPrincipalContext)
				g.SamAccountName = groupName
				Return DirectCast(New PrincipalSearcher(g).FindOne, GroupPrincipal) IsNot Nothing
			End Using
		End Function

		Public Shared Function GetUserPrincipal _
		(ByVal username As String, ByVal useManipulationUser As Boolean) As UserPrincipal

			If useManipulationUser Then
				Return GetUserPrincipal(username, Settings.Instance.ManipulationUserName, Settings.Instance.ManipulationUserPassword)
			Else
				Return GetUserPrincipal(username)
			End If
		End Function

    Public Shared Function GetUserPrincipal(ByVal personId As Int64) As UserPrincipal

      Using u = New UserPrincipal(GetPrincipalContext())
        u.EmployeeId = personId.ToString
        Return DirectCast(New PrincipalSearcher(u).FindOne, UserPrincipal)
      End Using
    End Function

    Public Shared Function GetUserPrincipal(ByVal userName As String) As UserPrincipal

			Using u = New UserPrincipal(GetPrincipalContext())
        u.SamAccountName = userName
        Return DirectCast(New PrincipalSearcher(u).FindOne, UserPrincipal)
      End Using
		End Function

		Public Shared Function GetUserPrincipal _
		(ByVal userName As String, ByVal userManipulationName As String _
		, ByVal password As String) As UserPrincipal

			Using u = New UserPrincipal(GetPrincipalContext(userManipulationName, password))
				u.SamAccountName = userName
				Return DirectCast(New PrincipalSearcher(u).FindOne, UserPrincipal)
			End Using
		End Function

		Public Shared Function GetGroupPrincipal _
		(ByVal groupName As String, ByVal useManipulationUser As Boolean) As GroupPrincipal

			If useManipulationUser Then
				Return GetGroupPrincipal(groupName, Settings.Instance.ManipulationUserName, Settings.Instance.ManipulationUserPassword)
			Else
				Return GetGroupPrincipal(groupName)
			End If
		End Function

    Public Shared Function GetGroupPrincipal(ByVal groupName As String) As GroupPrincipal
      Using g = New GroupPrincipal(GetPrincipalContext)
        g.SamAccountName = groupName
        Return DirectCast(New PrincipalSearcher(g).FindOne, GroupPrincipal)
      End Using
    End Function

		Public Shared Function GetGroupPrincipal _
		(ByVal groupName As String, ByVal userManipulationName As String _
		, ByVal password As String) As GroupPrincipal

			Using g = New GroupPrincipal(GetPrincipalContext(userManipulationName, password))
				g.SamAccountName = groupName
				Return DirectCast(New PrincipalSearcher(g).FindOne, GroupPrincipal)
			End Using

		End Function
#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace
