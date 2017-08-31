Option Explicit On
Option Strict On
Option Infer On

#Region " --------------->> Imports/ usings "

Imports System.DirectoryServices
Imports System.DirectoryServices.AccountManagement
Imports System.Management.Automation
Imports System.Management.Automation.Runspaces
Imports System.Security
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core.Exceptions

#End Region

Namespace Core.Manipulation

	Public Class AdUsers

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private Shared _rnd As New Random
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
#End Region	'{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region	'{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
		''' <summary>
		''' Liefert einen zufälligen Exchange-Datenbank-Namen zum Lastausgleich.
		''' </summary>
		Private Shared Function GetRandomExchangeDatabaseName() As String
			Dim index = _rnd.Next(My.Settings.ExchangeDatabaseNames.Count)
			Return My.Settings.ExchangeDatabaseNames.Item(index)
		End Function

		''' <summary>
		''' Inkrementiert die aktuelle UnixId um eins und gibt diese zurück.
		''' </summary>
		Private Shared Function GetUnixId() As Int32

      Dim de = DistinguishedName.GetByDistinguishedName _
      (Settings.Instance.UnixIdDistinguishedName).ToDirectoryEntry(True)

      Dim unixId = Convert.ToInt32(de.InvokeGet(AdProperties.msSFU30MaxUidNumber.ToString))
      Dim nextUnixId = unixId + 1
      de.InvokeSet(AdProperties.msSFU30MaxUidNumber.ToString, nextUnixId)
      de.CommitChanges()
      Return unixId
    End Function

    ''' <summary>
    ''' Setzt das Unix-Attribut für den Linux-Zugriff.
    ''' </summary>
    Private Shared Sub SetUnixAttributes(ByVal de As DirectoryEntry, ByVal samAccountName As String)

      de.Properties.Item(AdProperties.msSFU30NisDomain.ToString).Value = Settings.Instance.SecondLevelDomainName
      de.Properties.Item(AdProperties.msSFU30Name.ToString).Value = samAccountName
      de.Properties.Item(AdProperties.uidNumber.ToString).Value = GetUnixId()
      de.Properties.Item(AdProperties.loginShell.ToString).Value = "/bin/false"
      de.Properties.Item(AdProperties.unixHomeDirectory.ToString).Value = "/no/home"
      de.Properties.Item(AdProperties.gidNumber.ToString).Value = Settings.Instance.DomainUserGroupId
      de.CommitChanges()
    End Sub

    ''' <summary>
    ''' Aktiviert eine Exchange-Mailbox für den übergebenen User.
    ''' </summary>
    ''' <remarks>technet Enable-Mailbox: https://technet.microsoft.com/de-de/library/aa998251(v=exchg.150).aspx</remarks>
    Private Shared Sub GenerateMailbox(ByVal userInfo As AdUserInfo)

      Const enableMailboxCommand = "Enable-Mailbox"
      Dim passwordSecureString = New SecureString()
      Dim exchangePowershellUri = New Uri(My.Settings.ExchangePowershellUrl)

      Settings.Instance.ManipulationUserPassword.ToList.ForEach(Sub(c) passwordSecureString.AppendChar(c))

      Dim credential = New PSCredential(Settings.Instance.ManipulationUserName, passwordSecureString)

      Dim connectionInfo = New WSManConnectionInfo _
      (exchangePowershellUri, My.Settings.ExchangePowershellSchema, credential)

      connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic

      Using runspace = RunspaceFactory.CreateRunspace(connectionInfo)

        runspace.Open()

        Using pipeline = runspace.CreatePipeline()
          Dim command = New Command(enableMailboxCommand)
          command.Parameters.Add(ExchangeParameters.Identity.ToString, userInfo.SamAccountName)
          'command.Parameters.Add(ExchangeParameters.Alias.ToString, userName)
          command.Parameters.Add(ExchangeParameters.Database.ToString, GetRandomExchangeDatabaseName)

          pipeline.Commands.Add(command)
          pipeline.Invoke()

          If (pipeline.Error IsNot Nothing) AndAlso (pipeline.Error.Count > 0) Then
            Throw New GenerateMailBoxException
          End If
        End Using
      End Using
    End Sub
#End Region  '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
		'''<summary>Löscht den Benutzer mit der angegebenen Personen-Id aus dem AD.</summary>
		Public Shared Sub DeleteDomainUser(ByVal personId As Int64, ByVal useManipulationUser As Boolean)

      DeleteDomainUser(DistinguishedName.GetByPersonId(personId), useManipulationUser)
    End Sub

    '''<summary>Löscht den Benutzer mit dem angegebenen userName aus dem AD.</summary>
    Public Shared Sub DeleteDomainUser(ByVal userName As String, ByVal useManipulationUser As Boolean)

      DeleteDomainUser(DistinguishedName.GetByUserName(userName), useManipulationUser)
    End Sub

    ''' <summary>
    ''' Löscht den Benutzer mit dem angegebenen DistinguishedName aus dem AD.
    ''' </summary>
    Public Shared Sub DeleteDomainUser(ByVal userDn As DistinguishedName, ByVal useManipulationUser As Boolean)

      Dim userName = userDn.GetProperty(AdProperties.sAMAccountName.ToString).ToString
      Using u = AdPrincipals.GetUserPrincipal(userName, useManipulationUser)
        u.Delete()
      End Using
    End Sub

    ''' <summary>
    ''' Legt einen neuen Benutzer im AD an.
    ''' </summary>
    Public Shared Function CreateDomainUser(ByVal userInfo As AdUserInfo _
    , ByVal useManipulationUser As Boolean) As AdManipulationResults

      If AdPrincipals.ExistsUserPrincipal(userInfo.SamAccountName) Then
        Return AdManipulationResults.MemberAlreadyExist
      End If

      Using newUser = New UserPrincipal(AdPrincipals.GetPrincipalContext(userInfo.OuDistinguishedName.Value, useManipulationUser))

        With newUser
          Try
            .SamAccountName = userInfo.SamAccountName
            .UserPrincipalName = userInfo.UserPrincipalName
            .Surname = userInfo.LastName
            .GivenName = userInfo.FirstName
            .DisplayName = userInfo.DisplayName
            .EmployeeId = userInfo.EmployeeId.ToString
            .Description = userInfo.Description
            .VoiceTelephoneNumber = userInfo.PhoneNumber
            .Save()
          Catch ex As Exception
            Return AdManipulationResults.SetDefaultAttributesError
          End Try

          Try
            .SetPassword(userInfo.Pwd)
            .Save()
          Catch ex As Exception
            DeleteDomainUser(userInfo.SamAccountName, useManipulationUser)
            Return AdManipulationResults.SetPasswordError
          End Try

          Try
            .ExpirePasswordNow()
            .Save()
          Catch ex As Exception
            DeleteDomainUser(userInfo.SamAccountName, useManipulationUser)
            Return AdManipulationResults.GenerateMailboxError
          End Try

          If userInfo.ActivateAccount Then
            Try
              .Enabled = True
              .Save()
            Catch ex As Exception
              DeleteDomainUser(userInfo.SamAccountName, useManipulationUser)
              Return AdManipulationResults.SetUserEnabledError
            End Try
          End If

          Try
            SetUnixAttributes(DistinguishedName.GetByDistinguishedName _
            (.DistinguishedName).ToDirectoryEntry(True), .SamAccountName)
          Catch ex As Exception
            DeleteDomainUser(userInfo.SamAccountName, useManipulationUser)
            Return AdManipulationResults.SetUnixAttributesError
          End Try

          If userInfo.CreateMailBox Then
            Try
              GenerateMailbox(userInfo)
            Catch ex As Exception
              DeleteDomainUser(userInfo.SamAccountName, useManipulationUser)
              Return AdManipulationResults.GenerateMailboxError
            End Try
          End If
        End With

        Return AdManipulationResults.Successful
      End Using
    End Function

    ''' <summary>
    ''' Setzt das Kennwort des Benutzers zurück.
    ''' </summary>
    Public Shared Function SetPassword(ByVal personId As Int64 _
    , ByVal newPassword As String, ByVal useManipulationUser As Boolean) As AdManipulationResults

      If AdInformation.ExistsPersonId(personId) Then
        Dim dn = DistinguishedName.GetByPersonId(personId)
        Dim userName = dn.GetProperty(AdProperties.sAMAccountName).ToString
        Return SetPassword(userName, newPassword, useManipulationUser)
      Else
        Return AdManipulationResults.UserNotExists
      End If
    End Function

    ''' <summary>
    ''' Setzt das Kennwort des Benutzers zurück.
    ''' </summary>
    Public Shared Function SetPassword(ByVal userName As String _
    , ByVal newPassword As String, ByVal useManipulationUser As Boolean) As AdManipulationResults

      If AdPrincipals.ExistsUserPrincipal(userName) Then
        Using user = AdPrincipals.GetUserPrincipal(userName, useManipulationUser)
          Try
            user.SetPassword(newPassword)
            user.Save()
          Catch ex As Exception
            Return AdManipulationResults.SetPasswordError
          End Try
          Return AdManipulationResults.Successful
        End Using
      Else
        Return AdManipulationResults.UserNotExists
      End If
    End Function

    ''' <summary>
    ''' Aktiviert das Konto im AD.
    ''' </summary>
    Public Shared Function ActivateUser(ByVal personId As Int64 _
    , ByVal useManipulationUser As Boolean) As AdManipulationResults

      If AdInformation.ExistsPersonId(personId) Then
        Dim dn = DistinguishedName.GetByPersonId(personId)
        Dim userName = dn.GetProperty(AdProperties.sAMAccountName).ToString
        Return ActivateUser(userName, useManipulationUser)
      Else
        Return AdManipulationResults.UserNotExists
      End If
    End Function

    ''' <summary>
    ''' Aktiviert das Konto im AD.
    ''' </summary>
    Public Shared Function ActivateUser(ByVal userName As String _
    , ByVal useManipulationUser As Boolean) As AdManipulationResults

      If AdPrincipals.ExistsUserPrincipal(userName) Then

        Using user = AdPrincipals.GetUserPrincipal(userName, useManipulationUser)
          Try
            user.Enabled = True
            user.Save()
          Catch ex As Exception
            Return AdManipulationResults.SetUserEnabledError
          End Try
          Return AdManipulationResults.Successful
        End Using
      Else
        Return AdManipulationResults.UserNotExists
      End If
    End Function

    ''' <summary>
    ''' Deaktiviert das Konto im AD.
    ''' </summary>
    Public Shared Function DeactivateUser(ByVal personId As Int64 _
    , ByVal useManipulationUser As Boolean) As AdManipulationResults

      If AdInformation.ExistsPersonId(personId) Then
        Dim dn = DistinguishedName.GetByPersonId(personId)
        Dim userName = dn.GetProperty(AdProperties.sAMAccountName).ToString
				Return DeactivateUser(userName, useManipulationUser)
			Else
				Return AdManipulationResults.UserNotExists
			End If
		End Function

		''' <summary>
		''' Deaktiviert das Konto im AD.
		''' </summary>
		Public Shared Function DeactivateUser(ByVal userName As String _
		, ByVal useManipulationUser As Boolean) As AdManipulationResults

			If AdPrincipals.ExistsUserPrincipal(userName) Then
				Using user = AdPrincipals.GetUserPrincipal(userName, useManipulationUser)
					Try
						user.Enabled = True
						user.Save()
					Catch ex As Exception
						Return AdManipulationResults.SetUserEnabledError
					End Try
					Return AdManipulationResults.Successful
				End Using
			Else
				Return AdManipulationResults.UserNotExists
			End If
		End Function

#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace
