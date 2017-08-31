Option Explicit On
Option Infer On
Option Strict On

Imports SSP.ActiveDirectoryX.Grants

Public Class GrantTest

  Private _grantTable As GrantTable

  Public Sub ShowMenu()

    Dim key As ConsoleKey = Nothing

    Do
      Console.Clear()
      Console.WriteLine("<A>   Erzeuge GrantTable " & If(_grantTable Is Nothing, "", "(" & _grantTable.AppName & ")"))
      Console.WriteLine("<B>   Vorhandene Applikationen (GrantTables) listen")
      If _grantTable IsNot Nothing Then
        Console.WriteLine("<C>   Berechtigungen listen")
        Console.WriteLine("<D>   Benutzer mit gewährter Berechtigung listen")
      End If
      Console.WriteLine("<ESC> Beenden")
      Console.WriteLine("")
      Console.WriteLine("Auswahl> ")

      key = Console.ReadKey(True).Key

      Select Case key
        Case ConsoleKey.A
          CreateGrantTable()
        Case ConsoleKey.B
          ListAppications
        Case ConsoleKey.C
          ListGrants()
        Case ConsoleKey.D
          ListAssignedUsers()
        Case ConsoleKey.E
        Case ConsoleKey.F
        Case ConsoleKey.G
        Case ConsoleKey.H
      End Select

    Loop Until key = ConsoleKey.Escape
  End Sub

  Private Sub CreateGrantTable()

    Console.WriteLine("Bitte Applikationsnamen eingeben: ")
    Dim value = Console.ReadLine

    If GrantTables.GetAppNames.Contains(value) Then
      ConsoleStopWatch.Instance.Reset()
      ConsoleStopWatch.Instance.Start()
      _grantTable = New GrantTable(value)
      Console.WriteLine("<OK> " & ConsoleStopWatch.Instance.PassedSeconds & " Sekunden!")
    Else
      Console.WriteLine("Kein gültiger Applikationsname!")
    End If

    ConsoleStopWatch.Instance.Pause()
    ConsoleHelper.Instance.PressAnyKey()
  End Sub

  Private Sub ListAppications()
    ConsoleStopWatch.Instance.Reset()
    ConsoleStopWatch.Instance.Start()
    GrantTables.GetAppNames.ForEach(Sub(x) Console.WriteLine(x))
    Console.WriteLine("<OK> " & ConsoleStopWatch.Instance.PassedSeconds & " Sekunden!")
    ConsoleStopWatch.Instance.Pause()
    ConsoleHelper.Instance.PressAnyKey()
  End Sub

  Private Sub ListGrants()
    If _grantTable Is Nothing Then Exit Sub

    Try
      ConsoleStopWatch.Instance.Reset()
      ConsoleStopWatch.Instance.Start()

      Console.WriteLine("Berechtigungen:")
      _grantTable.FillByAppName()
      _grantTable.ToList.ForEach _
      (Sub(x) Console.WriteLine(String.Format("{0}{1} - {2}{1}", x.GrantName, vbCrLf, x.Description)))

      Console.WriteLine("")
      Console.WriteLine("<OK> " & ConsoleStopWatch.Instance.PassedSeconds & " Sekunden!")
    Catch ex As Exception
      Console.WriteLine(ex.Message)
    End Try

    ConsoleStopWatch.Instance.Pause()
    ConsoleHelper.Instance.PressAnyKey()
  End Sub

  Private Sub ListAssignedUsers()

    If _grantTable Is Nothing Then Exit Sub

    Console.WriteLine("Bitte Berechtigungsnamen eingeben: ")
    Dim value = Console.ReadLine

    If _grantTable.GrantNames.Contains(value) Then
      ConsoleStopWatch.Instance.Reset()
      ConsoleStopWatch.Instance.Start()

      _grantTable.Item(value).AssignedUsers.ToList.ForEach _
      (Sub(x) Console.WriteLine(String.Format("{0}({1})", x.BaseProperties.Name, x.BaseProperties.PersonId)))

      Console.WriteLine("<OK> " & ConsoleStopWatch.Instance.PassedSeconds & " Sekunden!")
    Else
      Console.WriteLine("Kein gültiger Berechtigungsname!")
    End If

    ConsoleStopWatch.Instance.Pause()
    ConsoleHelper.Instance.PressAnyKey()
  End Sub
End Class
