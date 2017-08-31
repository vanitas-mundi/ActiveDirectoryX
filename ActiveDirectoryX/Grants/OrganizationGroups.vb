Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.ActiveDirectoryX.Grants.Administration
Imports SSP.ActiveDirectoryX.Grants.Enums
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core
Imports System.Collections.ObjectModel

#End Region

Namespace Grants

  Public Class OrganizationGroups

    Implements IEnumerable(Of AdministrationGroup)

#Region " --------------->> Enumerationen der Klasse "
#End Region '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
    Private _fillType As FillTypes = FillTypes.FillAll
    Private _grantTree As GrantTree = Nothing
    Private _organizationGroups As New List(Of AdministrationGroup)
#End Region  '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
#End Region  '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
    ''' <summary>
    ''' Stellt Funktionalität zum Administrieren von Organisationsgruppen zur Verfügung.
    ''' </summary>
    Public ReadOnly Property Administration As OrganizationGroupsAdministration
      Get
        Return OrganizationGroupsAdministration.Instance
      End Get
    End Property

    ''' <summary>
    ''' Liefert eine Liste mit allen Organisationsgruppen.
    ''' </summary>
    Public ReadOnly Property OrganizationGroups As ReadOnlyCollection(Of AdministrationGroup)
      Get
        Return _organizationGroups.AsReadOnly
      End Get
    End Property

    ''' <summary>
    ''' Liefert eine Liste mit allen Organisationsgruppen gefiltert nach organizationGroupType.
    ''' </summary>
    Public ReadOnly Property OrganizationGroups(ByVal organizationGroupType As OrganizationGroupTypes) _
    As ReadOnlyCollection(Of AdministrationGroup)
      Get
        Return Me.Where(Function(r) r.OrganizationGroupType = organizationGroupType).ToList.AsReadOnly
      End Get
    End Property

    ''' <summary>
    ''' Liefert eine Liste mit Gruppen, welche keine Organisationsgruppen sind.
    ''' </summary>
    Public ReadOnly Property NoOrganizationGroups As ReadOnlyCollection(Of AdministrationGroup)
      Get
        Return Me.OrganizationGroups(OrganizationGroupTypes.NoOrganizationGroup).ToList.AsReadOnly
      End Get
    End Property

    ''' <summary>
    ''' Liefert eine Liste mit allen allgemeinen Organisationsgruppen.
    ''' </summary>
    Public ReadOnly Property CommonOrganizationGroups As ReadOnlyCollection(Of AdministrationGroup)
      Get
        Return Me.OrganizationGroups(OrganizationGroupTypes.CommonOrganizationGroup).ToList.AsReadOnly
      End Get
    End Property

    ''' <summary>
    ''' Liefert eine Liste mit allen Urlaubsgruppen.
    ''' </summary>
    Public ReadOnly Property HolidayGroups As ReadOnlyCollection(Of AdministrationGroup)
      Get
        Return Me.OrganizationGroups(OrganizationGroupTypes.HolidayGroup).ToList.AsReadOnly
      End Get
    End Property

    ''' <summary>
    ''' Liefert eine Liste mit allen Arbeitsgruppen.
    ''' </summary>
    Public ReadOnly Property WorkGroups As ReadOnlyCollection(Of AdministrationGroup)
      Get
        Return Me.OrganizationGroups(OrganizationGroupTypes.WorkGroup).ToList.AsReadOnly
      End Get
    End Property

    ''' <summary>
    ''' Liefert eine Liste mit allen Abrechnungsgruppen.
    ''' </summary>
    Public ReadOnly Property AccountingGroups As ReadOnlyCollection(Of AdministrationGroup)
      Get
        Return Me.OrganizationGroups(OrganizationGroupTypes.AccountingGroup).ToList.AsReadOnly
      End Get
    End Property

    ''' <summary>
    ''' Liefert eine Liste mit allen Organisationsgruppentypen.
    ''' </summary>
    Public ReadOnly Property OrganizationGroupTypesArray As OrganizationGroupTypes()
      Get
        Return CType(System.Enum.GetValues(GetType(OrganizationGroupTypes)), OrganizationGroupTypes())
      End Get
    End Property

    ''' <summary>
    ''' Liefert eine Liste mit allen Organisationsgruppentypnamen.
    ''' </summary>
    Public ReadOnly Property OrganizationGroupTypeNames As ReadOnlyCollection(Of String)
      Get
        Return GetOrganizationGroupTypeNames()
      End Get
    End Property

    ''' <summary>
    ''' Liefert eine Liste mit allen Organisationsgruppentyp-DistinguishedNames.
    ''' </summary>
    Public Shared ReadOnly Property OrganizationGroupTypeDistinguishedNames As ReadOnlyCollection(Of DistinguishedName)
      Get
        Return GetOrganizationGroupTypeDistinguishedNames()
      End Get
    End Property
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
    Private Sub CheckFilled()
      If _organizationGroups.Count = 0 Then
        Select Case _fillType
          Case FillTypes.FillAll
            Dim adminGroup = New AdministrationGroups
            adminGroup.FillAllOrganizationGroups()
            _organizationGroups.AddRange(adminGroup)
          Case FillTypes.FillByGrantTree
            Dim organizationGroupsDn = SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.OrganizationGroups)
            _organizationGroups.AddRange(_grantTree.Items.Where _
            (Function(dn) (dn.ContainsDn(organizationGroupsDn))).Select _
            (Function(dn) New AdministrationGroup(dn)).ToList)
        End Select
      End If
    End Sub
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
    Public Function GetEnumerator() As IEnumerator(Of AdministrationGroup) _
    Implements IEnumerable(Of AdministrationGroup).GetEnumerator
      CheckFilled()
      Return _organizationGroups.GetEnumerator
    End Function

    Public Function GetEnumerator1() As IEnumerator _
    Implements IEnumerable.GetEnumerator
      CheckFilled()
      Return _organizationGroups.GetEnumerator
    End Function

    Public Function Item(ByVal index As Int32) As AdministrationGroup
      CheckFilled()
      Return _organizationGroups.Item(index)
    End Function

    Public Function Item(ByVal dn As DistinguishedName) As AdministrationGroup
      CheckFilled()
      Return _organizationGroups.Where(Function(g) g.GroupDistinguishedName.Value = dn.Value).FirstOrDefault
    End Function

    Public Function OrganizationGroup(ByVal index As Int32) As AdministrationGroup
      Return _organizationGroups.Item(index)
    End Function

    Public Function OrganizationGroup(ByVal dn As DistinguishedName) As AdministrationGroup
      Return Me.Item(dn)
    End Function

    ''' <summary>
    ''' Liefert eine Liste mit allen Organisationsgruppentypnamen
    ''' </summary>
    Public Shared Function GetOrganizationGroupTypeNames() As ReadOnlyCollection(Of String)
      Return SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.OrganizationGroups).Children.Select _
      (Function(dn) dn.ToRelativeName).ToList.AsReadOnly
    End Function

    ''' <summary>
    ''' Liefert eine Liste mit allen Organisationsgruppentyp-DistinguishedNames
    ''' </summary>
    Public Shared Function GetOrganizationGroupTypeDistinguishedNames() As ReadOnlyCollection(Of DistinguishedName)
      Return SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.OrganizationGroups).Children.ToList.AsReadOnly
    End Function

    ''' <summary>
    ''' Füllt das Organisationsgruppen-Objekt mit allen vorhandenen Organisationsgruppen-Objekten.
    ''' </summary>
    Public Sub FillAll()
      _organizationGroups.Clear()
      _grantTree = Nothing
      _fillType = FillTypes.FillAll
    End Sub

    ''' <summary>
    ''' Liefert eine Liste mit allen Organisationsgruppen eines User anhand eines GrantTrees.
    ''' </summary>
    Public Sub FillByGrantTree(ByVal gt As GrantTree)
      _organizationGroups.Clear()
      _grantTree = gt
      _fillType = FillTypes.FillByGrantTree
    End Sub
#End Region  '{Öffentliche Methoden der Klasse}

  End Class

End Namespace

