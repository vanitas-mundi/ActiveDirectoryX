Option Explicit On
Option Infer On
Option Strict On
Imports SSP.ActiveDirectoryX.Core
Imports SSP.ActiveDirectoryX.Grants

Public Class GroupManagerTest

  Public Sub ShowMenu()

    Dim key As ConsoleKey = Nothing

    Do
      Console.Clear()
      Console.WriteLine("<A>   Gruppenmanager listen")
      Console.WriteLine("<B>   Verantworliche (Bewilliger) listen")
      Console.WriteLine("<C>   Zugeordnete Administrationsgruppen listen")

      Console.WriteLine("<ESC> Beenden")
      Console.WriteLine("")
      Console.WriteLine("Auswahl> ")

      key = Console.ReadKey(True).Key

      Select Case key
        Case ConsoleKey.A
          ListGroupManager()
        Case ConsoleKey.B
          ListManagerAndDeputies()
        Case ConsoleKey.C
          ListAssignedGroups
        Case ConsoleKey.D
        Case ConsoleKey.E
        Case ConsoleKey.F
        Case ConsoleKey.G
        Case ConsoleKey.H
      End Select

    Loop Until key = ConsoleKey.Escape
  End Sub

  Private Sub ListGroupManager()

    ConsoleStopWatch.Instance.Reset()
    ConsoleStopWatch.Instance.Start()
    GroupManagers.Instance.ToList.ForEach(Sub(x) Console.WriteLine(x.ToString))
    Console.WriteLine("<OK> " & ConsoleStopWatch.Instance.PassedSeconds & " Sekunden!")
    ConsoleStopWatch.Instance.Pause()
    ConsoleHelper.Instance.PressAnyKey()
  End Sub

  Private Function GetGroupManagerDnByName() As DistinguishedName
    Console.WriteLine("Bitte Namen des Gruppenmanagers eingeben:")
    Console.WriteLine("Präfix 'gruppenmanager.' kann entfallen!")
    Dim value = Console.ReadLine

    If Not value.StartsWith("gruppenmanager.") Then
      value = "gruppenmanager." & value
      Console.CursorTop -= 1
      Console.CursorLeft = 0
      Console.WriteLine(value)
    End If

    Return DistinguishedName.GetByGroupName(value)
  End Function

  Private Sub ListManagerAndDeputies()

    Dim dn = GetGroupManagerDnByName()

    If dn IsNot Nothing Then
      ConsoleStopWatch.Instance.Reset()
      ConsoleStopWatch.Instance.Start()

      Console.WriteLine("Hauptverantwortlicher:")
      With GroupManagers.Instance.Item(dn).ManagerDn.BaseProperties
        Console.WriteLine(String.Format("{0}({1})", .Name, .PersonId))
      End With
      Console.WriteLine("")

      Console.WriteLine("Stellvertreter:")
      GroupManagers.Instance.Item(dn).Deputies.ToList.ForEach _
      (Sub(x) Console.WriteLine(String.Format("{0}({1})", x.BaseProperties.Name, x.BaseProperties.PersonId)))
      Console.WriteLine("")

      Console.WriteLine("Verwaltete Benutzer:")
      GroupManagers.Instance.Item(dn).AssignedUsers.ToList.ForEach _
      (Sub(x) Console.WriteLine(String.Format("{0}({1})", x.BaseProperties.Name, x.BaseProperties.PersonId)))

      Console.WriteLine("<OK> " & ConsoleStopWatch.Instance.PassedSeconds & " Sekunden!")
    Else
      Console.WriteLine("Kein gültiger Gruppenmanager-Name!")
    End If

    ConsoleStopWatch.Instance.Pause()
    ConsoleHelper.Instance.PressAnyKey()
  End Sub

  Private Sub ListAssignedGroups()
    Dim dn = GetGroupManagerDnByName()

    If dn IsNot Nothing Then
      ConsoleStopWatch.Instance.Reset()
      ConsoleStopWatch.Instance.Start()

      Console.WriteLine("Zugeordnete Administrationsgruppen:")
      GroupManagers.Instance.GetAdministrationGroupsOf(dn).ToList.ForEach(Sub(x) Console.WriteLine(x.Name))

      Console.WriteLine("<OK> " & ConsoleStopWatch.Instance.PassedSeconds & " Sekunden!")
    Else
      Console.WriteLine("Kein gültiger Gruppenmanager-Name!")
    End If

    ConsoleStopWatch.Instance.Pause()
    ConsoleHelper.Instance.PressAnyKey()
  End Sub

End Class
