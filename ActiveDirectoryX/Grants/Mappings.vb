Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "

Imports SSP.ActiveDirectoryX.Grants.Administration
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core
Imports SSP.ActiveDirectoryX.Grants.Enums
#End Region

Namespace Grants

  Public Class Mappings

    Implements IEnumerable(Of Mapping)

#Region " --------------->> Enumerationen der Klasse "
#End Region '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
    Private _fillType As FillTypes = FillTypes.FillAll
    Private _grantTree As GrantTree = Nothing
    Private _mappings As New List(Of Mapping)
#End Region  '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
#End Region  '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
    ''' <summary>
    ''' Stellt Funktionalität zum Administrieren von Mappings zur Verfügung.
    ''' </summary>
    Public ReadOnly Property Administration As MappingsAdministration
      Get
        Return MappingsAdministration.Instance
      End Get
    End Property
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
    Private Sub CheckFilled()

      If _mappings.Count = 0 Then
        Select Case _fillType
          Case FillTypes.FillAll
            Dim adminGroup = New AdministrationGroups
            adminGroup.FillAllMappings()
            _mappings.AddRange(adminGroup.Select(Function(g) New Mapping(g.GroupDistinguishedName)).OrderBy(Function(m) m.Name))
          Case FillTypes.FillByGrantTree
            _mappings.AddRange(_grantTree.Items.Where _
            (Function(dn) dn.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Mappings))).Select _
            (Function(dn) New Mapping(dn)).OrderBy(Function(m) m.Name).ToList)
        End Select
      End If
    End Sub
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
    Public Function GetEnumerator() As IEnumerator(Of Mapping) Implements IEnumerable(Of Mapping).GetEnumerator
      CheckFilled
      Return _mappings.GetEnumerator
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
      CheckFilled
      Return _mappings.GetEnumerator
    End Function

    Public Function Item(ByVal index As Int32) As Mapping
      CheckFilled
      Return _mappings.Item(index)
    End Function

    Public Function Item(ByVal dn As DistinguishedName) As Mapping
      CheckFilled
      Return _mappings.Where(Function(m) m.GroupDistinguishedName.Value = dn.Value).FirstOrDefault
    End Function

    Public Function Mapping(ByVal index As Int32) As Mapping
      Return _mappings.Item(index)
    End Function

    Public Function Mapping(ByVal dn As DistinguishedName) As Mapping
      Return Me.Item(dn)
    End Function

    ''' <summary>
    ''' Füllt das Mappings-Objekt mit allen vorhandenen Mapping-Objekten.
    ''' </summary>
    Public Sub FillAll()

      _mappings.Clear()
      _fillType = FillTypes.FillAll
      _grantTree = Nothing
    End Sub

    ''' <summary>
    ''' Füllt User-Mappings anhand eines GrantTrees.
    ''' </summary>
    Public Sub FillByGrantTree(ByVal gt As GrantTree)

      _mappings.Clear()
      _fillType = FillTypes.FillByGrantTree
      _grantTree = gt
    End Sub
#End Region  '{Öffentliche Methoden der Klasse}

  End Class

End Namespace
