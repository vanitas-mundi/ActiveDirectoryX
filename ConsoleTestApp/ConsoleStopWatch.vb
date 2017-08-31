Public Class ConsoleStopWatch

  Private Shared _instance As ConsoleStopWatch
  Private _showInTitle As Boolean = True
  Private WithEvents _timer As New Timers.Timer
  Private _passedSeconds As Int64 = 0

  Private Sub OnTimer(ByVal sender As Object, ByVal e As EventArgs) Handles _timer.Elapsed
    _passedSeconds += 1
    If Me.ShowInTitle Then Console.Title = _passedSeconds.ToString
  End Sub

  Shared Sub New()
    _instance = New ConsoleStopWatch
  End Sub

  Private Sub New()
    _timer.Interval = 1000
  End Sub

  Public Shared ReadOnly Property Instance As ConsoleStopWatch
    Get
      Return _instance
    End Get
  End Property

  Public Property ShowInTitle As Boolean
    Get
      Return _showInTitle
    End Get
    Set(value As Boolean)
      _showInTitle = value
    End Set
  End Property


  Public ReadOnly Property PassedSeconds As Int64
    Get
      Return _passedSeconds
    End Get
  End Property

  Public Sub Start()
    _timer.Start()
  End Sub

  Public Function Pause() As Int64
    _timer.Stop()
    Return _passedSeconds
  End Function

  Public Function Reset() As Int64
    If Me.ShowInTitle Then Console.Title = "0"
    Dim temp = _passedSeconds
    _passedSeconds = 0
    Pause()
    Return temp
  End Function

End Class
