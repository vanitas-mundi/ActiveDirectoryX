Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "

Imports SSP.ActiveDirectoryX.Grants.Administration
Imports SSP.ActiveDirectoryX.Grants.Interfaces
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core
Imports SSP.ActiveDirectoryX.Grants.Enums

#End Region

Namespace Grants

  Public Class GrantTables

    Implements IEnumerable
    Implements IGrantTables
    Implements IEnumerable(Of GrantTable)

#Region " --------------->> Enumerationen der Klasse "
#End Region  '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
    Private _fillType As FillTypes = FillTypes.FillAll
    Private _grantTree As GrantTree = Nothing
    Private _grantTables As New List(Of GrantTable)
#End Region  '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
    Public Sub New()
    End Sub
#End Region  '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
    ''' <summary>
    ''' Stellt Verwaltungsfunktionalität zur Verfügung.
    ''' </summary>
    Public ReadOnly Property Administration As GrantTablesAdministration
      Get
        Return GrantTablesAdministration.Instance
      End Get
    End Property

    ''' <summary>
    ''' Liefert die GrantTable des angegebenen indexes.
    ''' </summary>
    Public ReadOnly Property Item(ByVal index As Int32) As GrantTable Implements IGrantTables.Item
      Get
        CheckFilled()
        Return _grantTables.Item(index)
      End Get
    End Property

    ''' <summary>
    ''' Liefert die GrantTable des angegebenen appNames.
    ''' </summary>
    Public ReadOnly Property Item(ByVal appName As String) As GrantTable Implements IGrantTables.Item
      Get
        CheckFilled()
        Return Me.Where(Function(gt) String.Compare(gt.AppName, appName, True) = 0).FirstOrDefault
      End Get
    End Property

    ''' <summary>
    ''' Liefert den zugrunde liegenden UserName.
    ''' </summary>
    Public ReadOnly Property UserName As String Implements IGrantTables.UserName
      Get
        Return If(_grantTree Is Nothing, "", _grantTree.DistinguishedNameContext.BaseProperties.Name)
      End Get
    End Property
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
    Private Sub CheckFilled()

      If _grantTables.Count = 0 Then

        For Each appName In GetAppNames()
          Dim grantTable = New GrantTable(appName)

          Select Case _fillType
            Case FillTypes.FillAll
              grantTable.FillByAppName()
            Case FillTypes.FillByGrantTree
              grantTable.FillByGrantTree(_grantTree)
          End Select
          _grantTables.Add(grantTable)
        Next appName
      End If
    End Sub
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
    Public Function GetEnumerator() As IEnumerator(Of GrantTable) _
    Implements IEnumerable(Of GrantTable).GetEnumerator
      CheckFilled()
      Return _grantTables.GetEnumerator
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
      CheckFilled()
      Return _grantTables.GetEnumerator
    End Function

    Public Function GrantTable(ByVal index As Int32) As GrantTable
      Return Me.Item(index)
    End Function

    Public Function GrantTable(ByVal appName As String) As GrantTable
      Return Me.Item(appName)
    End Function

    ''' <summary>
    ''' Liefert die GrantTable zum AppName.
    ''' </summary>
    Public Function Table(ByVal appName As String) As GrantTable
      Return Me.Item(appName)
    End Function

    ''' <summary>
    ''' Liefert die GrantTable zum Index.
    ''' </summary>
    Public Function Table(ByVal index As Int32) As GrantTable
      Return Me.Item(index)
    End Function
    ''' <summary>
    ''' Ermittelt die Namen aller Programme mit hinterlegten Berechtigungen.
    ''' </summary>
    Public Shared Function GetAppNames() As List(Of String)

      Dim grantDn = SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Grants)

      Return grantDn.Children.Where _
      (Function(dn) dn.ToRelativeDistinguishedName.ToLower.StartsWith("ou=")).Select _
      (Function(dn) dn.ToRelativeName).ToList
    End Function

    ''' <summary>
    ''' Füllt das Objekt mit allen vorhandenen GrantTables aller AppNames.
    ''' </summary>
    Public Sub FillAll() Implements IGrantTables.Fill
      _grantTables.Clear()
      _grantTree = Nothing
      _fillType = FillTypes.FillAll
    End Sub

    ''' <summary>
    ''' Füllt das Objekt mit allen vorhandenen GrantTables.
    ''' </summary>
    Public Sub FillByGrantTree(ByVal gt As GrantTree)
      _grantTables.Clear()
      _grantTree = gt
      _fillType = FillTypes.FillByGrantTree
    End Sub

    ''' <summary>
    ''' Prüft, ob Berechtigungen zur Anwendung appName existieren.
    ''' </summary>
    Public Shared Function AppNameExists(ByVal appName As String) As Boolean
      Return GetAppNames.Contains(appName)
    End Function
#End Region '{Öffentliche Methoden der Klasse}

  End Class

End Namespace
