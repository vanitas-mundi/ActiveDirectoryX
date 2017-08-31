Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports System.Text
Imports SSP.ActiveDirectoryX.Core
Imports SSP.ActiveDirectoryX.Grants
Imports SSP.WebServices.GrantServiceLibrary.Core.Interfaces
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Grants.Enums
#End Region

Namespace Core


  Public Class GrantService

    Implements IGrantService

#Region " --------------->> Enumerationen der Klasse "
#End Region '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
#End Region '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
#End Region '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
    Private Function GetUserName(ByVal personId As Int32) As String
      Dim dn = DistinguishedName.GetByPersonId(personId)
      Return If(dn Is Nothing, "", dn.GetProperty("sAMAccountName").ToString)
    End Function

    Private Function GenerateGrantString _
    (ByVal singleSignOn As Boolean _
    , ByVal appName As String, ByVal userName As String _
    , ByVal pwd As String) As String

      Dim loginResult = GetLoginResult(singleSignOn, appName, userName, pwd)
      Dim loginOk = GetLoginOk(loginResult)

      Dim sb = New StringBuilder
      sb.Append("LoginOk:" & loginOk)
      sb.Append(",LoginResult:" & loginResult)

      If loginOk = "Y" Then
        Dim gt = New GrantTree(userName)
        gt.Generate()

        Dim grantTable = New GrantTable(appName)
        grantTable.FillByGrantTree(gt)

        sb.Append("," & grantTable.ToGrantString)
      Else
        sb.Append(",Execute:N")
      End If

      Return sb.ToString
    End Function

    Private Function GetLoginResult _
    (ByVal singleSignOn As Boolean _
    , ByVal appName As String _
    , ByVal userName As String _
    , ByVal pwd As String) As String

      If (GrantTables.AppNameExists(appName)) AndAlso (userName <> "") Then

        Dim result = If(singleSignOn, AdInformation.IsLoginSingleSignOnSuccessful(userName) _
        , AdInformation.IsLoginSuccessful(userName, pwd))
        Return result.ToString
      Else
        Return If(userName = "", "InvalidUserName", "InvalidAppName")
      End If

    End Function

    Private Function GetLoginOk(ByVal loginResult As String) As String
      Return If(loginResult = "Successful", "Y", "N")
    End Function
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
    ''' <summary>
    ''' Liefert einen Berechtigungsstring unter Angabe von appName, personId, pwd.
    ''' </summary>
    Public Function GetGrantStringByPersonId _
    (ByVal appName As String, ByVal personId As Integer, ByVal pwd As String) As String _
    Implements IGrantService.GetGrantStringByPersonId

      Return GetGrantString(appName, GetUserName(personId), pwd)
    End Function

    ''' <summary>
    ''' Liefert einen Berechtigungsstring unter Angabe von appName, userName, pwd.
    ''' </summary>
    Public Function GetGrantString(ByVal appName As String, ByVal userName As String, ByVal pwd As String) As String _
    Implements IGrantService.GetGrantString

      Return GenerateGrantString(False, appName, userName, pwd)
    End Function

    ''' <summary>
    ''' Liefert einen Berechtigungsstring unter Angabe von appName, personId.
    ''' </summary>
    Public Function GetGrantStringSingleSignOnByPersonId _
    (ByVal appName As String, ByVal personId As Integer) As String _
    Implements IGrantService.GetGrantStringSingleSignOnByPersonId

      Return GetGrantStringSingleSignOn(appName, GetUserName(personId))
    End Function

    ''' <summary>
    ''' Liefert einen Berechtigungsstring unter Angabe von appName, userName.
    ''' </summary>
    Public Function GetGrantStringSingleSignOn(ByVal appName As String, ByVal userName As String) As String _
    Implements IGrantService.GetGrantStringSingleSignOn

      Return GenerateGrantString(True, appName, userName, Nothing)
    End Function

    ''' <summary>
    ''' Liefert alle AppNames.
    ''' </summary>
    Public Function GetAppNames() As String() Implements IGrantService.GetAppNames

      Return GrantTables.GetAppNames.ToArray
    End Function

    ''' <summary>
    ''' Liefert aus dem AD den Wert der Eigenschaft propertyname des Users von userName.
    ''' </summary>
    Public Function GetUserProperty(userName As String, propertyName As String) As String _
    Implements IGrantService.GetUserProperty

      Try
        Return DistinguishedName.GetByUserName(userName).GetProperty(propertyName).ToString
      Catch ex As NullReferenceException
        Throw New Exception(String.Format("AD-Property '{0}' doesn't exist or property value is not set.", propertyName), ex)
      End Try
    End Function

    ''' <summary>
    ''' Liefert aus dem AD den Wert der Eigenschaft propertyname des Users von personId.
    ''' </summary>
    Public Function GetUserPropertyByPersonId(personId As Integer, propertyName As String) As String _
    Implements IGrantService.GetUserPropertyByPersonId

      Try
        Return DistinguishedName.GetByPersonId(personId).GetProperty(propertyName).ToString
      Catch ex As NullReferenceException
        Throw New Exception(String.Format("AD-Property {0} doesn't exist or Property not set.", propertyName), ex)
      End Try
    End Function

    Public Function GetHolidayGroupNamesFromUser(personId As Integer) As String _
    Implements IGrantService.GetHolidayGroupNamesFromUser

      Dim grantUser = New GrantUser(personId)
      Dim groupNames = grantUser.OrganizationGroups.HolidayGroups.Select _
      (Function(x) $"{x.Name}:{x.Description}").ToArray
      Return String.Join(", ", groupNames)
    End Function

    Public Function IsUserInHolidayGroup(ByVal personId As Int32, ByVal holidayGroupName As String) As Boolean _
    Implements IGrantService.IsUserInHolidayGroup

      Dim grantUser = New GrantUser(personId)
      Return grantUser.OrganizationGroups.HolidayGroups.Any _
      (Function(hg) String.Compare(hg.Name, holidayGroupName, True) = 0)
    End Function

    Public Function GetUsersOfHolidayGroup(holidayGroupName As String) As String _
    Implements IGrantService.GetUsersOfHolidayGroup

      Dim dn = DistinguishedName.GetByGroupName(holidayGroupName)
      Return String.Join(", ", dn.GetMembersRecursive(GetMembersRecursiveTypes.UsersOnly).Select _
      (Function(userDn) userDn.BaseProperties.PersonId.ToString))
    End Function

    Public Function GetHolidayGroupManagerPersonIds(holidayGroupName As String) As String _
    Implements IGrantService.GetHolidayGroupManagerPersonIds

      Dim dn = DistinguishedName.GetByGroupName(holidayGroupName)
      Dim hg = New AdministrationGroup(dn)

      Dim manager = hg.GroupManager.ManagerPersonId.ToString
      Dim deputies = String.Join(", ", hg.GroupManager.DeputyPersonIds)

      Return If(deputies = "", manager, String.Format("{0}, {1}", manager, deputies))
    End Function

    Public Function GetGroupNamesByManagerPersonId(ByVal personId As Int32) As String _
    Implements IGrantService.GetGroupNamesByManagerPersonId

      Return String.Join(", ", GrantsBaseRoutines.Instance.GetAssignedAdministrationGroupsByManagerPersonId _
      (personId, True).Select(Function(ag) ag.Name).ToArray)
    End Function

    Public Function GetAllHolidayGroups() As String() Implements IGrantService.GetAllHolidayGroups

      Dim result = New List(Of String)
      Dim organizationGroups = New OrganizationGroups
      Dim holidayGroups = organizationGroups.OrganizationGroups _
      (OrganizationGroupTypes.HolidayGroup)

      holidayGroups.ToList.ForEach(Sub(x) result.Add($"{x.Name}:{x.Description}"))
      Return result.ToArray
    End Function

#End Region  '{Öffentliche Methoden der Klasse}

  End Class

End Namespace
