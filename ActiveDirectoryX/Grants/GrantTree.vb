Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core
Imports System.Collections.ObjectModel
#End Region

Namespace Grants

	Public Class GrantTree

		Inherits AdTree

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Public Sub New(ByVal personId As Int64)
			Me.New(DistinguishedName.GetByPersonId(personId))
		End Sub

		Public Sub New(ByVal userName As String)
			Me.New(DistinguishedName.GetByUserName(userName))
		End Sub

		Public Sub New(ByVal distinguishedName As DistinguishedName)
			MyBase.New(distinguishedName)
		End Sub
#End Region  '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
#End Region  '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region  '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
    Private Function ContainsTilde(ByVal dn As DistinguishedName) As Boolean

      If dn.Value.Contains(Settings.Instance.DenialChar) Then
        Return True
      Else
        Dim parent = dn.Parent
        Return If(parent IsNot Nothing, ContainsTilde(parent), False)
      End If
    End Function

    Private Function FindAppGrants(ByVal distinguishedNames As IEnumerable(Of DistinguishedName) _
    , ByVal appName As String) As List(Of String)

      appName = appName.ToLower
      Dim temp = distinguishedNames.Where(Function(dn) dn.Value.ToLower.Contains("=" & appName & "."))
      Return temp.Select(Function(dn) If(ContainsTilde(dn), Settings.Instance.DenialChar, "") & dn.ToRelativeName).Distinct.ToList
    End Function
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
    Public Function ToGrantUser() As GrantUser
			Return New GrantUser(Me)
		End Function

		''' <summary>
		''' Erzeugt den GrantTree
		''' </summary>
		Public Overloads Sub Generate()
      MyBase.GenerateTree(AdProperties.memberOf)
    End Sub

    '''<summary>Liefert eine generische Liste mit den gewährten Berechtigungen eines appNames.</summary>
    Public Function GetGrantsByAppName(ByVal appName As String) As ReadOnlyCollection(Of String)

      Dim result = FindAppGrants(Me.Items, appName)

      Dim denieds = result.Where(Function(s) s.StartsWith(Settings.Instance.DenialChar)).Select _
      (Function(s) s.Replace(Settings.Instance.DenialChar, String.Empty))

      Dim grants = result.Where(Function(s) Not s.StartsWith(Settings.Instance.DenialChar)).ToList
      denieds.ToList.ForEach(Sub(s) grants.Remove(s))

			Return grants.Select(Function(s) s.Split("."c).Last).ToList.AsReadOnly
		End Function
#End Region '{Öffentliche Methoden der Klasse}

  End Class

End Namespace
