Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.ActiveDirectoryX.Grants.Administration
Imports SSP.ActiveDirectoryX.Core
#End Region

Namespace Grants

	Public Class GroupManager

#Region " --------------->> Enumerationen der Klasse "
#End Region  '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
    Private _baseProperties As GroupManagerBaseProperties
    Private _administration As GroupManagerAdministration
#End Region '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
    Friend Sub New(ByVal baseProperties As GroupManagerBaseProperties)

      Initialize(baseProperties)
    End Sub
#End Region  '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
    ''' <summary>
    ''' Stellt Funktionen zur Verwaltung zur Verfügung.
    ''' </summary>
    Public ReadOnly Property Administration As GroupManagerAdministration
      Get
        Return _administration
      End Get
    End Property

    ''' <summary>
    ''' Liefert den Namen des Gruppenmanagers.
    ''' </summary>
    Public ReadOnly Property Name As String
      Get
        Return _baseProperties.GroupManagerDn.Name
      End Get
    End Property

    ''' <summary>
    ''' Liefert den Distinguished-Name zum Gruppenmanager-Objekt.
    ''' </summary>
    Public ReadOnly Property ManagerGroupDn As DistinguishedName
      Get
        Return _baseProperties.GroupManagerDn
      End Get
    End Property

    ''' <summary>
    ''' Liefert die Personen-Id. des hinterlegten Managers.
    ''' </summary>
    Public ReadOnly Property ManagerPersonId As Int64
      Get
        Return _baseProperties.ManagerDn.BaseProperties.PersonId
      End Get
    End Property

    ''' <summary>
    ''' Liefert den DistinguishedName des hinterlegten Managers.
    ''' </summary>
    Public ReadOnly Property ManagerDn As DistinguishedName
      Get
        Return _baseProperties.ManagerDn
      End Get
    End Property

    ''' <summary>
    ''' Liefert die Personen-Ids der hinterlegten Stellvertreter.
    ''' </summary>
    Public ReadOnly Property DeputyPersonIds As Int64()
      Get
        Return _baseProperties.DeputiesDn.Select(Function(x) x.BaseProperties.PersonId).ToArray
      End Get
    End Property

    ''' <summary>
    ''' Liefert ein Array mit allen hinterlegten Stellvertretern.
    ''' </summary>
    Public ReadOnly Property Deputies As DistinguishedName()
      Get
        Return _baseProperties.DeputiesDn.ToArray
      End Get
    End Property

    ''' <summary>
    ''' Liefert alle Benutzer, welche vom Gruppenmanager verwaltet werden.
    ''' </summary>
    Public ReadOnly Property AssignedUsers As DistinguishedName()
      Get
        Dim result = New List(Of DistinguishedName)
        GroupManagers.Instance.GetAdministrationGroupsOf(Me).ToList.ForEach _
        (Sub(x) result.AddRange(x.GroupDistinguishedName.GetMembers))
        Return result.Where(Function(x) x.BaseProperties.PersonId > 0).GroupBy _
        (Function(x) x.BaseProperties.PersonId).Select(Function(x) x.First).ToArray
      End Get
    End Property

    ''' <summary>
    ''' Liefert ein Array mit allen hinterlegten Stellvertretern und dem Manager.
    ''' </summary>
    Public ReadOnly Property ManagerAndDeputies As DistinguishedName()
      Get
        Dim result = New List(Of DistinguishedName)
        result.Add(Me.ManagerDn)
        result.AddRange(Me.Deputies.ToList)

        Return result.ToArray
      End Get
    End Property

#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
    Private Sub Initialize(ByVal baseProperties As GroupManagerBaseProperties)

      _baseProperties = baseProperties
      _administration = New GroupManagerAdministration(Me)
    End Sub
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
    Public Overrides Function ToString() As String
      Return Me.Name
    End Function

    ''' <summary>
    ''' Prüft, ob der übergebene DistinguishedName Manager der Managergruppe ist.
    ''' </summary>
    Public Function IsManager(ByVal userDn As DistinguishedName) As Boolean
      Return IsManager(userDn.BaseProperties.PersonId)
    End Function

    ''' <summary>
    ''' Prüft, ob der übergebene DistinguishedName Manager der Managergruppe ist.
    ''' </summary>
    Public Function IsManager(ByVal personId As Int64) As Boolean
      Return _baseProperties.ManagerDn.BaseProperties.PersonId = personId
    End Function

    ''' <summary>
    ''' Prüft, ob der übergebene DistinguishedName ein Stellvertreter der Managergruppe ist.
    ''' </summary>
    Public Function IsDeputy(ByVal userDn As DistinguishedName) As Boolean
      Return IsDeputy(userDn.BaseProperties.PersonId)
    End Function

    ''' <summary>
    ''' Prüft, ob der übergebene DistinguishedName ein Stellvertreter der Managergruppe ist.
    ''' </summary>
    Public Function IsDeputy(ByVal personId As Int64) As Boolean
      Return Me.DeputyPersonIds.Contains(personId)
    End Function

    ''' <summary>
    ''' Prüft, ob der übergebene DistinguishedName ein Stellvertreter oder Manager der Managergruppe ist.
    ''' </summary>
    Public Function IsManagerOrDeputy(ByVal userDn As DistinguishedName) As Boolean
      Return IsManagerOrDeputy(userDn.BaseProperties.PersonId)
    End Function

    ''' <summary>
    ''' Prüft, ob der übergebene DistinguishedName ein Stellvertreter oder Manager der Managergruppe ist.
    ''' </summary>
    Public Function IsManagerOrDeputy(ByVal personId As Int64) As Boolean
      Return Me.IsManager(personId) OrElse Me.IsDeputy(personId)
    End Function
#End Region '{Öffentliche Methoden der Klasse}

  End Class

End Namespace


