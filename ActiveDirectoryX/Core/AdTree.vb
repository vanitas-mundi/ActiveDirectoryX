Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.ActiveDirectoryX.Core.Enums
Imports System.Collections.ObjectModel
Imports System.DirectoryServices.AccountManagement
#End Region

Namespace Core

  Public Class AdTree

#Region " --------------->> Enumerationen der Klasse "
#End Region  '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
    Private _distinguishedNameContext As DistinguishedName
    Private _items As  List(Of DistinguishedName)
    Private _isGenerated As Boolean
#End Region  '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
    Public Sub New(ByVal distinguishedName As DistinguishedName)
      _distinguishedNameContext = distinguishedName
      _items = New List(Of DistinguishedName)
      _items.Add(_distinguishedNameContext)
    End Sub
#End Region  '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
    ''' <summary>
    ''' Liefert den DistinguishedName-Kontext.
    ''' </summary>
    Public ReadOnly Property DistinguishedNameContext As DistinguishedName
      Get
        Return _distinguishedNameContext
      End Get
    End Property

    ''' <summary>
    ''' Liefert true, wenn GenerateTree aufgerufen wurde.
    ''' </summary>
    Public ReadOnly Property IsGenerated As Boolean
      Get
        Return _isGenerated
      End Get

    End Property

    ''' <summary>
    ''' Liefert alle Items des Trees.
    ''' </summary>
    Public ReadOnly Property Items As ReadOnlyCollection(Of DistinguishedName)
      Get
        Return _items.AsReadOnly
      End Get
    End Property
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
    Private Sub GenerateItems(ByVal propertyName As AdProperties)

      _items = New List(Of DistinguishedName)

      Dim basePropertiesList = New List(Of DistinguishedNameBaseProperties)

      Select Case propertyName
        Case AdProperties.member
          Using group = AdPrincipals.GetGroupPrincipal(_distinguishedNameContext.Name)
            basePropertiesList = (group.GetMembers(True).ToList.Where _
            (Function(x) x.Guid.HasValue).Cast(Of UserPrincipal).Select _
            (Function(x) New DistinguishedNameBaseProperties _
            (x.DistinguishedName, x.SamAccountName, x.Description, x.Guid.Value, x.EmployeeId))).ToList
          End Using
        Case AdProperties.memberOf
          Using user = AdPrincipals.GetUserPrincipal(_distinguishedNameContext.BaseProperties.PersonId)
            basePropertiesList = (user.GetAuthorizationGroups.ToList.Where _
            (Function(x) x.Guid.HasValue).Select _
            (Function(x) New DistinguishedNameBaseProperties _
            (x.DistinguishedName, x.SamAccountName, x.Description, x.Guid.Value, ""))).ToList
          End Using
        Case Else
          Return
      End Select

      Dim distinguishedNames = basePropertiesList.Select(Function(x) New DistinguishedName(x))
      _items.AddRange(distinguishedNames)
    End Sub

    '  ''' <summary>
    '  ''' Liefert rekursiv alle Nodes ab CurrentNode
    '  ''' </summary>
    '  Private Function GetNodesRecursive(ByVal currentNode As AdNode) As List(Of AdNode)

    '	Dim nodes As New List(Of AdNode)
    '	nodes.Add(currentNode)
    '	For Each child In currentNode.Nodes
    '		nodes.AddRange(GetNodesRecursive(child))
    '	Next child
    '	Return nodes
    'End Function
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
    ''' <summary>
    ''' Erzeugt den Tree anhand der übergebenen Property.
    ''' </summary>
    Public Sub GenerateTree(ByVal propertyName As AdProperties)

      GenerateItems(propertyName)
      _isGenerated = True
    End Sub

    ''' <summary>
    ''' Prüft, ob der DistinguishedName-Kontext Mitglied der übergebenen Gruppe ist
    ''' </summary>
    Public Function IsMemberOf(ByVal groupDistinguishedName As DistinguishedName) As Boolean
      Return _items.Where(Function(dn) dn.IsEqualTo(groupDistinguishedName)).Any
    End Function
#End Region '{Öffentliche Methoden der Klasse}

  End Class

End Namespace
