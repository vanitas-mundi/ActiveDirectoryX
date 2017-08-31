
Imports SSP.ActiveDirectoryX.Grants
Imports SSP.ActiveDirectoryX.Core
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Data

Module MainModule

	Private Sub TempTest()

		ConsoleStopWatch.Instance.Reset()
		ConsoleStopWatch.Instance.Start()

		Dim hg = New AdministrationGroup(DistinguishedName.GetByGroupName("urlaubsgruppe.U.IT"))
		Console.WriteLine(hg.GroupManager.ManagerDn.BaseProperties.Name)

		Console.WriteLine("Fertig " & ConsoleStopWatch.Instance.Pause() & " Sekunden!")

		ConsoleHelper.Instance.PressAnyKey()
	End Sub

	Private Sub ShowMenu()

		Dim key As ConsoleKey = Nothing

		Do
			Console.Clear()
			Console.WriteLine("<A>   GrantUserTest")
			Console.WriteLine("<B>   GrantTest")
			Console.WriteLine("<C>   GroupManagerTest")
			Console.WriteLine("<X>   TempTest")
			Console.WriteLine("<ESC> Beenden")
			Console.WriteLine()
			Console.WriteLine("Auswahl> ")

			key = Console.ReadKey(True).Key

			Select Case key
				Case ConsoleKey.A
					GrantUserTest()
				Case ConsoleKey.B
					GrantTest()
				Case ConsoleKey.C
					GroupManagerTest()
				Case ConsoleKey.X
					TempTest()
			End Select

		Loop Until key = ConsoleKey.Escape

	End Sub

	Private Sub GrantUserTest()
		Dim grantUserTest = New GrantUserTest
		grantUserTest.ShowMenu()
	End Sub

	Private Sub GrantTest()
		Dim grantTest = New GrantTest
		grantTest.ShowMenu()
	End Sub

	Private Sub GroupManagerTest()
		Dim groupManagerTest = New GroupManagerTest
		groupManagerTest.ShowMenu()
	End Sub

	Sub Main()
		Dim sb = AdRepositoryHelper.Instance.CreateDefaultSelectBuilder

		sb.Select.Add(AdProperties.displayName.ToString)
		sb.Where.Add(String.Format("{0} = '{1}'", AdProperties.employeeID.ToString, "27"))

		MsgBox(sb.ToString)
		MsgBox(Settings.Instance.ConnectionString)

		Console.WriteLine(sb.ExecuteScalar.ToString)
		Console.ReadKey()

    'ShowMenu()
  End Sub


End Module
