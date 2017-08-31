Imports NUnit.Framework
Imports SSP.ActiveDirectoryX.Grants

<TestFixture>
Public Class ConsoleToolTest

  <Test>
  Public Sub TrallaTest()
    Dim gu = New GrantUser(27)

    Assert.IsNotNull(gu)
  End Sub


End Class
