Option Explicit On
Option Infer On
 Option Strict On

#Region " --------------->> Imports/ usings "
Imports System.Text
Imports SSP.ActiveDirectoryX.Core
Imports SSP.ActiveDirectoryX.Grants.Enums
Imports SSP.ActiveDirectoryX.Grants.Exceptions
#End Region

Namespace Grants

  Friend Class TreeBuilder

#Region " --------------->> Enumerationen der Klasse "
#End Region '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
    Private _rootlines As TreeBuilderRootLines
    Private _rootLineType As TreeBuilderRootLineTypes
#End Region  '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
    Public Sub New()
      Me.New(TreeBuilderRootLineTypes.Regular)
    End Sub

    Public Sub New(ByVal rootlines As TreeBuilderRootLineTypes)
      _rootLineType = rootlines
      _rootlines = New TreeBuilderRootLines(rootlines)
    End Sub
#End Region  '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
    Public ReadOnly Property RootLineType As TreeBuilderRootLineTypes
      Get
        Return _rootLineType
      End Get
    End Property

    Public ReadOnly Property Rootlines As TreeBuilderRootLines
      Get
        Return _rootlines
      End Get
    End Property
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
    Private Function GetRootLineString(ByVal indent As String, lastElement As Boolean) As String

      With Rootlines
        If Not Me.RootLineType = TreeBuilderRootLineTypes.None Then
          Dim temp = indent.Replace(.RegularElement, .ParentLine).Replace(.EndElement, .ParentLine)
          Dim pos = temp.LastIndexOf(.ParentLine, StringComparison.Ordinal)

          If (lastElement) AndAlso (pos > -1) Then
            Return String.Format("{0}{1}{2}", temp.Substring(0, pos), .WhiteSpace, temp.Substring(pos + .ParentLine.Length))
          Else
            Return temp
          End If
        Else
          Return ""
        End If
      End With
    End Function

    Private Sub BuildTree(ByVal tree As StringBuilder, ByVal ou As OrganizationalUnit, ByVal indent As String, ByVal lastElement As Boolean)

      tree.AppendLine(indent & ou.Name)

      Dim ous = ou.ChildrenOrganizationalUnits
      For i = 0 To ous.Count - 1

        With Rootlines
          Dim item = ous(i)
          Dim atLastElement = (i + 1) >= ous.Count
          Dim indentEnd = If(atLastElement, .EndElement, .RegularElement)
          BuildTree(tree, item, GetRootLineString(indent, lastElement) & indentEnd, atLastElement)
        End With
      Next i
    End Sub
#End Region  '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
    Public Function GetTreeStringBuilder(ByVal organizationalUnitName As String) As StringBuilder
      Return GetTreeStringBuilder(New OrganizationalUnit _
      (DistinguishedName.GetByOu(organizationalUnitName)))
    End Function

    Public Function GetTreeStringBuilder(ByVal organizationalUnitDn As DistinguishedName) As StringBuilder

      If AdministrationTypeResolver.Instance.IsOrganizationalUnit(organizationalUnitDn) Then
        Return GetTreeStringBuilder(New OrganizationalUnit(organizationalUnitDn))
      Else
        Throw New WrongAdministrationTypeException
      End If
    End Function

    Public Function GetTreeStringBuilder(ByVal organizationalUnit As OrganizationalUnit) As StringBuilder
      Dim tree = New StringBuilder
      BuildTree(tree, organizationalUnit, "", False)
      Return tree
    End Function

    Public Function GetTreeString(ByVal organizationalUnitName As String) As String
      Return GetTreeStringBuilder(organizationalUnitName).ToString
    End Function

    Public Function GetTreeString(ByVal organizationalUnitDn As DistinguishedName) As String
      Return GetTreeStringBuilder(organizationalUnitDn).ToString
    End Function

    Public Function GetTreeString(ByVal organizationalUnit As OrganizationalUnit) As String
      Return GetTreeStringBuilder(organizationalUnit).ToString
    End Function
#End Region '{Öffentliche Methoden der Klasse}

  End Class

End Namespace
