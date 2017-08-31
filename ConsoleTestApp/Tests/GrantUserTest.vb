Option Explicit On
Option Infer On
Option Strict On

Imports SSP.ActiveDirectoryX.Grants

Public Class GrantUserTest

  Private _grantUser As GrantUser

  Public Sub ShowMenu()

    Dim key As ConsoleKey = Nothing

    Do
      Console.Clear()
      Console.WriteLine("<A>   Erzeuge GrantUser " & If(_grantUser Is Nothing, "", "(" & _grantUser.UserName & ")"))
      If _grantUser IsNot Nothing Then
        Console.WriteLine("<B>   Berechtigungstabelle listen")
        Console.WriteLine("<C>   Mappings listen")
        Console.WriteLine("<D>   Organisationsgruppen listen")
        Console.WriteLine("<E>   Berechtigungsgruppen listen")
        Console.WriteLine("<F>   Rollen listen")
        Console.WriteLine("<G>   Bewilliger listen")
      End If
      Console.WriteLine("<ESC> Beenden")
      Console.WriteLine("")
      Console.WriteLine("Auswahl> ")

      key = Console.ReadKey(True).Key

      Select Case key
        Case ConsoleKey.A
          CreateGrantUser()
        Case ConsoleKey.B
          ListGrantTable()
        Case ConsoleKey.C
          ListMappings()
        Case ConsoleKey.D
          ListOrganizationGroups()
        Case ConsoleKey.E
          ListGrantGroups()
        Case ConsoleKey.F
          ListRoles()
        Case ConsoleKey.G
          ListGroupManager
        Case ConsoleKey.H
      End Select

    Loop Until key = ConsoleKey.Escape
  End Sub

  Private Sub CreateGrantUser()

    Console.WriteLine("Bitte Personen-Id. eingeben: ")
    Dim value = Console.ReadLine

    If IsNumeric(value) Then
      ConsoleStopWatch.Instance.Reset()
      ConsoleStopWatch.Instance.Start()
      _grantUser = New GrantUser(Convert.ToInt64(value))
      Console.WriteLine("<OK> " & ConsoleStopWatch.Instance.PassedSeconds & " Sekunden!")
    Else
      Console.WriteLine("Keine gültige Personen-Id.!")
    End If

    ConsoleStopWatch.Instance.Pause()
    ConsoleHelper.Instance.PressAnyKey()
  End Sub

  Private Sub ListGrantTable()

    If _grantUser Is Nothing Then Exit Sub

    Console.WriteLine("Bitte GrantTable-Namen eingeben: ")
    Dim value = Console.ReadLine

    Try
      ConsoleStopWatch.Instance.Reset()
      ConsoleStopWatch.Instance.Start()
      _grantUser.GrantTables.GrantTable(value).ToList.ForEach(Sub(g) Console.WriteLine(g.NameValueString))
      Console.WriteLine("")
      Console.WriteLine("<OK> " & ConsoleStopWatch.Instance.PassedSeconds & " Sekunden!")
    Catch ex As Exception
      Console.WriteLine(ex.Message)
    End Try

    ConsoleStopWatch.Instance.Pause()
    ConsoleHelper.Instance.PressAnyKey()
  End Sub

  Private Sub ListMappings()
    If _grantUser Is Nothing Then Exit Sub

    Try
      ConsoleStopWatch.Instance.Reset()
      ConsoleStopWatch.Instance.Start()
      _grantUser.Mappings.ToList.ForEach(Sub(x) Console.WriteLine(x.DriveLetter & " | " & x.ToString))
      Console.WriteLine("")
      Console.WriteLine("<OK> " & ConsoleStopWatch.Instance.PassedSeconds & " Sekunden!")
    Catch ex As Exception
      Console.WriteLine(ex.Message)
    End Try

    ConsoleStopWatch.Instance.Pause()
    ConsoleHelper.Instance.PressAnyKey()
  End Sub

  Private Sub ListOrganizationGroups()
    If _grantUser Is Nothing Then Exit Sub

    Try
      ConsoleStopWatch.Instance.Reset()
      ConsoleStopWatch.Instance.Start()
      _grantUser.OrganizationGroups.ToList.ForEach(Sub(x) Console.WriteLine(x.ToString))
      Console.WriteLine("")
      Console.WriteLine("<OK> " & ConsoleStopWatch.Instance.PassedSeconds & " Sekunden!")
    Catch ex As Exception
      Console.WriteLine(ex.Message)
    End Try

    ConsoleStopWatch.Instance.Pause()
    ConsoleHelper.Instance.PressAnyKey()
  End Sub

  Private Sub ListGrantGroups()
    If _grantUser Is Nothing Then Exit Sub

    Try
      ConsoleStopWatch.Instance.Reset()
      ConsoleStopWatch.Instance.Start()
      _grantUser.GrantGroups.ToList.ForEach(Sub(x) Console.WriteLine(x.ToString))
      Console.WriteLine("")
      Console.WriteLine("<OK> " & ConsoleStopWatch.Instance.PassedSeconds & " Sekunden!")
    Catch ex As Exception
      Console.WriteLine(ex.Message)
    End Try

    ConsoleStopWatch.Instance.Pause()
    ConsoleHelper.Instance.PressAnyKey()
  End Sub

  Private Sub ListRoles()
    If _grantUser Is Nothing Then Exit Sub

    Try
      ConsoleStopWatch.Instance.Reset()
      ConsoleStopWatch.Instance.Start()
      _grantUser.Roles.ToList.ForEach(Sub(x) Console.WriteLine(x.ToString))
      Console.WriteLine("")
      Console.WriteLine("<OK> " & ConsoleStopWatch.Instance.PassedSeconds & " Sekunden!")
    Catch ex As Exception
      Console.WriteLine(ex.Message)
    End Try

    ConsoleStopWatch.Instance.Pause()
    ConsoleHelper.Instance.PressAnyKey()
  End Sub

  Private Sub ListGroupManager()
    If _grantUser Is Nothing Then Exit Sub

    'Try
    ConsoleStopWatch.Instance.Reset()
      ConsoleStopWatch.Instance.Start()

      Console.WriteLine("Bewilliger:")
      _grantUser.AssignedManagers.ToList.ForEach _
      (Sub(x) Console.WriteLine(x.BaseProperties.Name & "(" & x.BaseProperties.PersonId & ")"))

      Console.WriteLine("")
      Console.WriteLine("")
      Console.WriteLine("Stellvertreter:")
      _grantUser.AssignedDeputies.ToList.ForEach _
      (Sub(x) Console.WriteLine(x.BaseProperties.Name & "(" & x.BaseProperties.PersonId & ")"))


      Console.WriteLine("")
      Console.WriteLine("")
      Console.WriteLine("Bewilliger und Stellvertreter:")
      _grantUser.AssignedManagersAndDeputies.ToList.ForEach _
      (Sub(x) Console.WriteLine(x.BaseProperties.Name & "(" & x.BaseProperties.PersonId & ")"))


      Console.WriteLine("")
      Console.WriteLine("")
      Console.WriteLine("Bewilliger von (Administrationsgruppen):")
      _grantUser.ManagedGroups.ToList.ForEach(Sub(x) Console.WriteLine _
      (x.Name & If(x.GroupManager.IsDeputy(_grantUser.UserDistinguishedName), " - Stellvertreter", "")))


      Console.WriteLine("")
      Console.WriteLine("")
      Console.WriteLine("Bewilliger von (Benutzern):")
      _grantUser.ManagedUsers.ToList.ForEach(Sub(x) Console.WriteLine(x.Name))


      Console.WriteLine("")
      Console.WriteLine("<OK> " & ConsoleStopWatch.Instance.PassedSeconds & " Sekunden!")
      'Catch ex As Exception
      '  Console.WriteLine(ex.Message)
      'End Try

      ConsoleStopWatch.Instance.Pause()
    ConsoleHelper.Instance.PressAnyKey()
  End Sub

End Class
