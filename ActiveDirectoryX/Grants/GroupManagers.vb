Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.ActiveDirectoryX.Grants.Administration
Imports SSP.ActiveDirectoryX.Core
Imports System.DirectoryServices
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Data.Repositories
Imports SSP.ActiveDirectoryX.Grants.Enums
#End Region

Namespace Grants

  Public Class GroupManagers

    Implements IEnumerable(Of GroupManager)

#Region " --------------->> Enumerationen der Klasse "
#End Region  '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
    Private Shared _instance As GroupManagers
    Private _fillType As FillTypes = FillTypes.FillAll

    Private _groupManagerOfAdministrationGroupDictionary As New Dictionary(Of String, String)
    Private _administratorGroupsOfGroupManagerDictionary As New Dictionary(Of String, List(Of String))
    Private _groupManagerBasePropertiesDictionary As New Dictionary(Of String, GroupManagerBaseProperties)

    Private _groupManagers As New List(Of GroupManager)
#End Region  '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
    Shared Sub New()
      _instance = New GroupManagers
    End Sub

    Private Sub New()
    End Sub
#End Region  '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
    Public Shared ReadOnly Property Instance As GroupManagers
      Get
        Return _instance
      End Get
    End Property

    ''' <summary>
    ''' Stellt Funktionalität zum Administrieren von Gruppenmanagern zur Verfügung.
    ''' </summary>
    Public ReadOnly Property Administration As GroupManagersAdministration
      Get
        Return GroupManagersAdministration.Instance
      End Get
    End Property
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
    Private Sub CheckFilled()
      If _groupManagers.Count = 0 Then
        FillGroupManagers()
      End If
    End Sub

    Private Sub CheckAssignmentsFilled()
      CheckFilled()

      If _groupManagerOfAdministrationGroupDictionary.Count = 0 Then
        FillGroupManagerAssignments()
      End If
    End Sub

    Private Sub SetGroupManagerDnInBasePropertiesDictionary()
      With DistinguishedNameRepository.Instance
        Dim distinguishedNames = .GetByDistinguishedNames _
        (_groupManagerBasePropertiesDictionary.Keys).ToList

        distinguishedNames.ForEach _
        (Sub(x) _groupManagerBasePropertiesDictionary.Item(x.Value).GroupManagerDn = x)
      End With
    End Sub

    Private Sub SetManagerDnInBasePropertiesDictionary()
      With DistinguishedNameRepository.Instance
        Dim managerDnStrings = _groupManagerBasePropertiesDictionary.Values.Select _
        (Function(x) x.ManagerDnString).Distinct

        Dim distinguishedNames = .GetByDistinguishedNames(managerDnStrings).ToList
        distinguishedNames.ForEach _
        (Sub(x) _groupManagerBasePropertiesDictionary.Values.Where _
        (Function(y) y.ManagerDnString = x.Value).ToList.ForEach(Sub(y) y.ManagerDn = x))
      End With
    End Sub

    Private Sub SetDeputiesDnInBasePropertiesDictionary()
      With DistinguishedNameRepository.Instance
        Dim deputiesDnStringList = _groupManagerBasePropertiesDictionary.Values.Where _
        (Function(x) x.DeputiesDnString IsNot Nothing).Select _
        (Function(x) x.DeputiesDnString).ToList

        Dim deputiesDnString = New List(Of String)
        deputiesDnStringList.ForEach(Sub(x) deputiesDnString.AddRange(x))

        Dim distinguishedNames = .GetByDistinguishedNames(deputiesDnString.Distinct).ToList
        distinguishedNames.ForEach _
        (Sub(x) _groupManagerBasePropertiesDictionary.Values.Where _
        (Function(y) y.DeputiesDnString.Contains(x.Value)).ToList.ForEach _
        (Sub(y) y.AddToDeputies(x)))
      End With
    End Sub

#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
    Public Function GetEnumerator() As IEnumerator(Of GroupManager) Implements IEnumerable(Of GroupManager).GetEnumerator
      CheckFilled()
      Return _groupManagers.GetEnumerator
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
      CheckFilled()
      Return _groupManagers.GetEnumerator
    End Function

    Public Function Item(ByVal index As Int32) As GroupManager
      CheckFilled()
      Return _groupManagers.Item(index)
    End Function

    Public Function Item(ByVal dn As DistinguishedName) As GroupManager
      CheckFilled()
      Return _groupManagers.Where(Function(m) m.ManagerGroupDn.Value = dn.Value).FirstOrDefault
    End Function

    Public Function GroupManager(ByVal index As Int32) As GroupManager
      Return _groupManagers.Item(index)
    End Function

    Public Function GroupManager(ByVal dn As DistinguishedName) As GroupManager
      Return Me.Item(dn)
    End Function

    ''' <summary>
    ''' Füllt das Gruppenmanager-Objekt mit allen vorhandenen Gruppenmanager-Objekten.
    ''' </summary>
    Public Sub FillAll()
      _groupManagers.Clear()
      _groupManagerOfAdministrationGroupDictionary = New Dictionary(Of String, String)
      _AdministratorGroupsOfGroupManagerDictionary = New Dictionary(Of String, List(Of String))
      _fillType = FillTypes.FillByGrantTree
    End Sub
#End Region  '{Öffentliche Methoden der Klasse}

    Public Sub FillGroupManagerAssignments()

      If _groupManagerOfAdministrationGroupDictionary.Count > 0 Then Return

      Dim entry = SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Administration).ToDirectoryEntry(False)

      Using searcher = New DirectorySearcher(My.Settings.GroupManagersOfAdminGroupsLdapQueryString)

        searcher.SearchRoot = entry
        searcher.PropertiesToLoad.Add(AdProperties.distinguishedName.ToString)
        searcher.PropertiesToLoad.Add(AdProperties.managedBy.ToString)

        For Each item As SearchResult In searcher.FindAll
          Dim groupManagerDn = item.Properties.Item(AdProperties.managedBy.ToString).Item(0).ToString

          If groupManagerDn.StartsWith(My.Settings.GroupManagerPrefix, StringComparison.CurrentCultureIgnoreCase) Then
            Dim administrationGroupDn = item.Properties.Item(AdProperties.distinguishedName.ToString).Item(0).ToString

            _groupManagerOfAdministrationGroupDictionary.Add(administrationGroupDn, groupManagerDn)
            If Not _AdministratorGroupsOfGroupManagerDictionary.Keys.Contains(groupManagerDn) Then
              _AdministratorGroupsOfGroupManagerDictionary.Add(groupManagerDn, New List(Of String))
            End If

            _AdministratorGroupsOfGroupManagerDictionary.Item(groupManagerDn).Add(administrationGroupDn)
          End If
        Next item
      End Using
    End Sub

    Public Function GetGroupManagerOf(ByVal administrationGroup As AdministrationGroup) As GroupManager
      Return GetGroupManagerOf(administrationGroup.GroupDistinguishedName.Value)
    End Function

    Public Function GetGroupManagerOf(ByVal administrationGroupDn As DistinguishedName) As GroupManager
      Return GetGroupManagerOf(administrationGroupDn.Value)
    End Function

    Public Function GetGroupManagerOf(ByVal administrationGroupDnString As String) As GroupManager
      CheckAssignmentsFilled()

      If Not _groupManagerOfAdministrationGroupDictionary.Keys.Contains(administrationGroupDnString) Then
        Return Nothing
      Else
        Dim groupManagerDnString = _groupManagerOfAdministrationGroupDictionary.Item(administrationGroupDnString)
        Return _groupManagers.FirstOrDefault(Function(x) x.ManagerGroupDn.Value = groupManagerDnString)
      End If
    End Function

    Public Function GetAdministrationGroupsOf(ByVal groupManager As GroupManager) As AdministrationGroup()
      Return GetAdministrationGroupsOf(groupManager.ManagerGroupDn.Value)
    End Function

    Public Function GetAdministrationGroupsOf(ByVal groupManagerDn As DistinguishedName) As AdministrationGroup()
      Return GetAdministrationGroupsOf(groupManagerDn.Value)
    End Function

    Public Function GetAdministrationGroupsOf(ByVal groupManagerDnString As String) As AdministrationGroup()
      CheckAssignmentsFilled()

      Dim result = New List(Of AdministrationGroup)

      If _administratorGroupsOfGroupManagerDictionary.Keys.Contains(groupManagerDnString) Then
        Dim administrationGroupDnStringArray = _administratorGroupsOfGroupManagerDictionary.Item(groupManagerDnString)
        result = DistinguishedNameRepository.Instance.GetByDistinguishedNames(administrationGroupDnStringArray).Select _
        (Function(x) New AdministrationGroup(x)).ToList
      End If

      Return result.ToArray
    End Function

    ''' <summary>
    ''' Liefert die von einem Manager (Bewilliger) verwalteten (zu bewilligenden) Administrationsgruppen.
    ''' Der Manager kann alternativ ein Stellvertreter sein.
    ''' </summary>
    Public Function GetManagedGroupsOf(ByVal managerPersonId As Int64) As AdministrationGroup()
      Return GetManagedGroupsOf(managerPersonId, True)
    End Function

    ''' <summary>
    ''' Liefert die von einem Manager (Bewilliger) verwalteten (zu bewilligenden) Administrationsgruppen.
    ''' Der Manager kann alternativ ein Stellvertreter sein.
    ''' </summary>
    Public Function GetManagedGroupsOf(ByVal managerDn As DistinguishedName) As AdministrationGroup()
      Return GetManagedGroupsOf(managerDn, True)
    End Function

    ''' <summary>
    ''' Liefert die von einem Manager (Bewilliger) verwalteten (zu bewilligenden) Administrationsgruppen.
    ''' Der Manager kann ebenfalls ein Stellvertreter sein, wenn includeDeputies 'true'
    ''' </summary>
    Public Function GetManagedGroupsOf(ByVal managerDn As DistinguishedName _
    , ByVal includeDeputies As Boolean) As AdministrationGroup()

      Return GetManagedGroupsOf(managerDn.BaseProperties.PersonId, includeDeputies)
    End Function

    ''' <summary>
    ''' Liefert die von einem Manager (Bewilliger) verwalteten (zu bewilligenden) Administrationsgruppen.
    ''' Der Manager kann ebenfalls ein Stellvertreter sein, wenn includeDeputies 'true'
    ''' </summary>
    Public Function GetManagedGroupsOf(ByVal managerPersonId As Int64 _
    , ByVal includeDeputies As Boolean) As AdministrationGroup()

      CheckAssignmentsFilled()

      Dim result = New List(Of AdministrationGroup)

      _groupManagers.Where(Function(x) (x.ManagerDn.BaseProperties.PersonId = managerPersonId) _
      OrElse (If(includeDeputies, x.DeputyPersonIds.Contains(managerPersonId), False))).ToList.ForEach _
      (Sub(x) result.AddRange(GetAdministrationGroupsOf(x)))

      Return result.Distinct.ToArray
    End Function

    ''' <summary>
    ''' Liefert die von einem Manager (Bewilliger) verwalteten (zu bewilligenden) Benutzer.
    ''' Der Manager kann alternativ ein Stellvertreter sein.
    ''' </summary>
    Public Function GetManagedUsersOf(ByVal managerDn As DistinguishedName) As DistinguishedName()
      Return GetManagedUsersOf(managerDn.BaseProperties.PersonId, True)
    End Function

    ''' <summary>
    ''' Liefert die von einem Manager (Bewilliger) verwalteten (zu bewilligenden) Benutzer.
    ''' Der Manager kann alternativ ein Stellvertreter sein.
    ''' </summary>
    Public Function GetManagedUsersOf(ByVal managerPersonId As Int64) As DistinguishedName()
      Return GetManagedUsersOf(managerPersonId, True)
    End Function

    ''' <summary>
    ''' Liefert die von einem Manager (Bewilliger) verwalteten (zu bewilligenden) Benutzer.
    ''' Der Manager kann ebenfalls ein Stellvertreter sein, wenn includeDeputies 'true'
    ''' </summary>
    Public Function GetManagedUsersOf(ByVal managerDn As DistinguishedName _
    , ByVal includeDeputies As Boolean) As DistinguishedName()

      Return GetManagedUsersOf(managerDn.BaseProperties.PersonId, includeDeputies)
    End Function

    ''' <summary>
    ''' Liefert die von einem Manager (Bewilliger) verwalteten (zu bewilligenden) Benutzer.
    ''' Der Manager kann ebenfalls ein Stellvertreter sein, wenn includeDeputies 'true'
    ''' </summary>
    Public Function GetManagedUsersOf(ByVal managerPersonId As Int64 _
    , ByVal includeDeputies As Boolean) As DistinguishedName()

      CheckAssignmentsFilled()

      Dim result = New List(Of DistinguishedName)
      GetManagedGroupsOf(managerPersonId, includeDeputies).ToList.ForEach(Sub(x) result.AddRange(x.Members))

      Return result.GroupBy(Function(x) x.BaseProperties.PersonId).Select _
      (Function(x) x.First).Where(Function(x) x.BaseProperties.PersonId > 0).ToArray
    End Function

    Private Sub FillGroupManagers()

      _groupManagerBasePropertiesDictionary = New Dictionary(Of String, GroupManagerBaseProperties)
      Dim entry = SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.GroupManagers).ToDirectoryEntry(False)

      Using searcher = New DirectorySearcher(My.Settings.GroupManagersLdapQueryString)

        searcher.SearchRoot = entry
        searcher.PropertiesToLoad.Add(AdProperties.distinguishedName.ToString)
        searcher.PropertiesToLoad.Add(AdProperties.managedBy.ToString)
        searcher.PropertiesToLoad.Add(AdProperties.member.ToString)

        For Each item As SearchResult In searcher.FindAll

          Dim bp = New GroupManagerBaseProperties _
          (item.Properties.Item(AdProperties.distinguishedName.ToString) _
          , item.Properties.Item(AdProperties.managedBy.ToString) _
          , item.Properties.Item(AdProperties.member.ToString))

          _groupManagerBasePropertiesDictionary.Add(bp.GroupManagerDnString, bp)
        Next item
      End Using

      SetGroupManagerDnInBasePropertiesDictionary()
      SetManagerDnInBasePropertiesDictionary()
      SetDeputiesDnInBasePropertiesDictionary()

      _groupManagers = New List(Of GroupManager)
      _groupManagers = _groupManagerBasePropertiesDictionary.Values.Select _
      (Function(x) New GroupManager(x)).ToList
    End Sub

  End Class

End Namespace
