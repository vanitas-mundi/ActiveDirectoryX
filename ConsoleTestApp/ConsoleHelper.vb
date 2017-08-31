Public Class ConsoleHelper

  Private Shared _instance As ConsoleHelper

  Shared Sub New()
    _instance = New ConsoleHelper
  End Sub

  Private Sub New()
  End Sub

  Public Shared ReadOnly Property Instance As ConsoleHelper
    Get
      Return _instance
    End Get
  End Property

  Public Sub PressAnyKey()
    Console.WriteLine()
    Console.WriteLine("<Taste, um fortzufahren>")
    Console.ReadKey(True)
  End Sub
End Class
