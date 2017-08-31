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

	Public Class GrantGroups

		Implements IEnumerable(Of AdministrationGroup)

#Region " --------------->> Enumerationen der Klasse "
#End Region '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
    Private _fillType As FillTypes = FillTypes.FillAll
    Private _grantTree As GrantTree = Nothing
    Private _grantGroups As New List(Of AdministrationGroup)
		Private _roles As Roles
		Private _mappings As Mappings
    Private _grantTables As GrantTables
#End Region  '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
#End Region  '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
    ''' <summary>
    ''' Stellt Funktionalität zum Administrieren von Berechtigungsgruppen zur Verfügung.
    ''' </summary>
    Public ReadOnly Property Administration As GrantGroupsAdministration
		Get
			Return GrantGroupsAdministration.Instance
		End Get
		End Property

		''' <summary>
		''' Liefert eine Liste mit allen Berechtigungsgruppen gefiltert nach grantType.
		''' </summary>
		Public ReadOnly Property GrantGroups(ByVal grantType As GrantTypes) _
		As ReadOnlyCollection(Of AdministrationGroup)
		Get
			Return Me.Where(Function(g) g.GrantType = grantType).ToList.AsReadOnly
		End Get
		End Property


		''' <summary>
		''' Liefert eine Liste mit allen Rollen.
		''' </summary>
		Public ReadOnly Property Roles As Roles
		Get
			Return _roles
		End Get
		End Property

		''' <summary>
		''' Liefert hinterlegte Mappings
		''' </summary>
		Public ReadOnly Property Mappings As Mappings
		Get
			Return _mappings
		End Get
		End Property

		''' <summary>
		''' Liefert eine Liste mit allen ManagerGroups.
		''' </summary>
		Public ReadOnly Property GroupManagers As GroupManagers
		Get
        Return GroupManagers.Instance
      End Get
		End Property

		''' <summary>
		''' Liefert hinterlegte GrantTables (Applikationen).
		''' </summary>
		Public ReadOnly Property GrantTables As GrantTables
		Get
			Return _grantTables
		End Get
		End Property

		 ''' <summary>
		 ''' Liefert eine Liste mit allen Berechtigungstypen.
		 ''' </summary>
		Public ReadOnly Property GrantTypesArray As GrantTypes()
		Get
			Return CType(System.Enum.GetValues(GetType(GrantTypes)), GrantTypes())
		End Get
		End Property

		 ''' <summary>
		 ''' Liefert eine Liste mit allen Berechtigungstypnamen.
		 ''' </summary>
		Public ReadOnly Property GrantTypeNames As ReadOnlyCollection(Of String)
		Get
			Return GetGrantTypeNames()
		End Get
		End Property

		 ''' <summary>
		 ''' Liefert eine Liste mit allen Berechtigungstyp-DistinguishedNames.
		 ''' </summary>
		Public Shared ReadOnly Property GrantTypeDistinguishedNames As ReadOnlyCollection(Of DistinguishedName)
		Get
			Return GetGrantTypeDistinguishedNames()
		End Get
		End Property
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
    Private Sub CheckFilled()

      If _grantGroups.Count = 0 Then

        Select Case _fillType
          Case FillTypes.FillAll
            Dim adminGroup = New AdministrationGroups
            adminGroup.FillAllGrantGroups()
            _grantGroups.AddRange(adminGroup)
          Case FillTypes.FillByGrantTree
            Dim grantGroupsDn = SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Grants)
            _grantGroups.AddRange(_grantTree.Items.Where _
            (Function(dn) (dn.ContainsDn(grantGroupsDn))).Select _
            (Function(dn) New AdministrationGroup(dn)).ToList)
        End Select
      End If
    End Sub
#End Region  '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
    Public Function GetEnumerator() As IEnumerator(Of AdministrationGroup) _
		Implements IEnumerable(Of AdministrationGroup).GetEnumerator
      CheckFilled
      Return _grantGroups.GetEnumerator
		End Function

		Public Function GetEnumerator1() As IEnumerator _
		Implements IEnumerable.GetEnumerator
      CheckFilled
      Return _grantGroups.GetEnumerator
		End Function

		Public Function Item(ByVal index As Int32) As AdministrationGroup
      CheckFilled
      Return _grantGroups.Item(index)
    End Function

		Public Function Item(ByVal dn As DistinguishedName) As AdministrationGroup
      CheckFilled
      Return _grantGroups.Where(Function(g) g.GroupDistinguishedName.Value = dn.Value).FirstOrDefault
    End Function

		Public Function GrantGroup(ByVal index As Int32) As AdministrationGroup
			Return _grantGroups.Item(index)
		End Function

		Public Function GrantGroup(ByVal dn As DistinguishedName) As AdministrationGroup
			Return Me.Item(dn)
		End Function

		 ''' <summary>
		 ''' Liefert eine Liste mit allen Berechtigungstypnamen
		 ''' </summary>
		Public Shared Function GetGrantTypeNames() As ReadOnlyCollection(Of String)
			Return SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Grants).Children.Select _
			(Function(dn) dn.ToRelativeName).ToList.AsReadOnly
		End Function

		 ''' <summary>
		 ''' Liefert eine Liste mit allen Berechtigungstyp-DistinguishedNames
		 ''' </summary>
		Public Shared Function GetGrantTypeDistinguishedNames() As ReadOnlyCollection(Of DistinguishedName)
			Return SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Grants).Children.ToList.AsReadOnly
		End Function

    ''' <summary>
    ''' Füllt das Berechtigungsgruppen-Objekt mit allen vorhandenen Berechtigungsgruppen-Objekten.
    ''' </summary>
    Public Sub FillAll()
      _fillType = FillTypes.FillAll
      _grantTree = Nothing

      _roles = New Roles
      _grantTables = New GrantTables
      _mappings = New Mappings
      _grantGroups = New List(Of AdministrationGroup)

      _roles.FillAll()
      _grantTables.FillAll()
      _mappings.FillAll()
    End Sub

    ''' <summary>
    ''' Liefert eine Liste mit allen Berechtigungsgruppen eines User anhand eines GrantTrees.
    ''' </summary>
    Public Sub FillByGrantTree(ByVal gt As GrantTree)
      _fillType = FillTypes.FillByGrantTree
      _grantTree = gt

      _roles = New Roles
      _grantTables = New GrantTables
      _mappings = New Mappings
      _grantGroups = New List(Of AdministrationGroup)

      _roles.FillByGrantTree(_grantTree)
      _grantTables.FillByGrantTree(_grantTree)
      _mappings.FillByGrantTree(_grantTree)
    End Sub
#End Region  '{Öffentliche Methoden der Klasse}

  End Class

End Namespace

