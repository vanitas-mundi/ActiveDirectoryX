Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports System.Text
Imports SSP.ActiveDirectoryX.Grants.Administration
Imports SSP.ActiveDirectoryX.Grants.Enums
Imports SSP.ActiveDirectoryX.Grants.Interfaces
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core
#End Region

Namespace Grants

	Public Class GrantTable

		Implements IEnumerable(Of Grant)
    Implements IGrantTable
    Implements IEnumerable

#Region " --------------->> Enumerationen der Klasse "
#End Region '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
    Private _fillType As FillTypes = FillTypes.FillAll
    Private _grantTree As GrantTree = Nothing
    Private _grants As New List(Of Grant)
    Private _appName As String
#End Region  '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
    Public Sub New(ByVal appName As String)
      _appName = appName
    End Sub
#End Region  '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
    ''' <summary>
    '''Stellt Funktionen zur Administration von Berechtigungen zur Verfügung.
    ''' </summary>
    Public ReadOnly Property Administration() As GrantsAdministration
      Get
        Return GrantsAdministration.Instance
      End Get
    End Property

    ''' <summary>
    ''' Liefert die Berechtigung zum Index.
    ''' </summary>
    Public ReadOnly Property Item(index As Integer) As Grant Implements IGrantTable.Item
      Get
        CheckFilled()
        Return _grants.Item(index)
      End Get
    End Property

    ''' <summary>
    ''' Liefert die Berechtigung zum GrantName.
    ''' </summary>
    Public ReadOnly Property Item(grantName As String) As Grant Implements IGrantTable.Item
      Get
        CheckFilled()
        Return _grants.Where(Function(g) g.GrantName = grantName).FirstOrDefault
      End Get
    End Property

    Public ReadOnly Property GrantNames As String() Implements IGrantTable.GrantNames
      Get
        Return Me.Select(Function(g) g.GrantName).ToArray
      End Get
    End Property

    Public ReadOnly Property AppName As String Implements IGrantTable.AppName
      Get
        Return _appName
      End Get
    End Property

    Public ReadOnly Property UserName As String Implements IGrantTable.UserName
      Get
        Return If(_grantTree Is Nothing, "", _grantTree.DistinguishedNameContext.BaseProperties.Name)
      End Get
    End Property
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
    Private Sub CheckFilled()
      If _grants.Count = 0 Then

        'FillByAppName
        _grants.AddRange(GetGrantDistinguishedNamesByAppName(Me.AppName).Select _
        (Function(x) New Grant(Me, x.ToDeEscapedRelativeName.Split("."c).LastOrDefault _
        , GrantValues.N, x.BaseProperties.Descripton)))

        Select Case _fillType
          Case FillTypes.FillAll
            'Falls benötigt hier Code eintragen
            Return
          Case FillTypes.FillByGrantTree
            Dim appNames = _grantTree.GetGrantsByAppName(_appName)
            appNames.ToList.ForEach(Sub(s) Me.Item(s).Value = GrantValues.Y)
        End Select
      End If
    End Sub
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
    Public Function GetEnumerator() As IEnumerator(Of Grant) Implements IEnumerable(Of Grant).GetEnumerator
      CheckFilled()
      Return _grants.GetEnumerator
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
      CheckFilled()
      Return _grants.GetEnumerator
    End Function

    ''' <summary>
    ''' Füllt das Objekt mit allen hinterlegten Berechtigungen der Applikation.
    ''' </summary>
    Public Sub FillByAppName() Implements IGrantTable.FillAll
      _grants.Clear()
      _grantTree = Nothing
      _fillType = FillTypes.FillAll
    End Sub

    ''' <summary>
    ''' Füllt das Objekt mit allen hinterlegten Berechtigungen der Applikation zu einem User.
    ''' </summary>
    Public Sub FillByGrantTree(ByVal gt As GrantTree) Implements IGrantTable.FillByGrantTree
      _grants.Clear()
      _grantTree = gt
      _fillType = FillTypes.FillByGrantTree
    End Sub

    ''' <summary>
    ''' Liefert die Berechtigung zum Index.
    ''' </summary>
    Public Function Grant(ByVal index As Int32) As Grant
			Return _grants.Item(index)
		End Function

    ''' <summary>
    ''' Liefert die Berechtigung zum GrantName.
    ''' </summary>
    Public Function Grant(ByVal grantName As String) As Grant
      Return Me.Item(grantName)
    End Function

    ''' <summary>
    ''' Ermittelt alle Berechtigungsnamen einer Applikation.
    ''' </summary>
    Public Shared Function GetGrantNamesByAppName(ByVal appName As String) As IEnumerable(Of String)

      With SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Grants)
        Return .GetDistinguishedNameFromChild(appName).Children.Select _
        (Function(dn) dn.ToDeEscapedRelativeName.Split("."c).LastOrDefault)
      End With
    End Function

    ''' <summary>
    ''' Ermittelt alle Berechtigungsnamen einer Applikation.
    ''' </summary>
    Public Shared Function GetGrantDistinguishedNamesByAppName(ByVal appName As String) As IEnumerable(Of DistinguishedName)

      With SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Grants)
        Return .GetDistinguishedNameFromChild(appName).Children
      End With
    End Function

    ''' <summary>
    ''' Liefert einen String  mit allen Berechtigungen.
    ''' </summary>
    Public Function ToGrantString(ByVal grantDelimiter As String) As String Implements IGrantTable.ToGrantString

      Return String.Join(",", Me.Select(Function(g) g.NameValueString(grantDelimiter)).ToArray)
    End Function

    ''' <summary>
    ''' Liefert einen String mit allen Berechtigungen.
    ''' </summary>
    Public Function ToGrantString() As String Implements IGrantTable.ToGrantString

      Return ToGrantString(":")
    End Function

    ''' <summary>
    ''' Liefert alle DistinguishedNames der Users, welchen die Berechtigung grantName gewährt wurde.
    ''' </summary>
    Public Function GetAssignedUsers(ByVal grantName As String) As DistinguishedName()
      Dim gt = New GroupTree(String.Format("{0}.{1}", Me.AppName, grantName))
      gt.Generate()
      Return gt.Items.ToArray
    End Function

    ''' <summary>
    ''' Liefert alle DistinguishedNames der Users, welchen die Berechtigung grant gewährt wurde.
    ''' </summary>
    Public Function GetAssignedUsers(ByVal grant As Grant) As DistinguishedName()

      Return GetAssignedUsers(grant.GrantName)
    End Function

    ''' <summary>
    ''' Liefert alle Personen-Ids der Benutzer, welchen die Berechtigung grantName gewährt wurde.
    ''' </summary>
    Public Function GetAssignedUsersPersonIds(ByVal grantName As String) As Int64()
      Return Me.GetAssignedUsers(grantName).Select(Function(x) x.BaseProperties.PersonId).ToArray
    End Function

    ''' <summary>
    ''' Liefert alle Personen-Ids der Benutzer, welchen die Berechtigung grant gewährt wurde.
    ''' </summary>
    Public Function GetAssignedUsersPersonIds(ByVal grant As Grant) As Int64()
      Return GetAssignedUsersPersonIds(grant.GrantName)
    End Function

#End Region  '{Öffentliche Methoden der Klasse}

  End Class

End Namespace
