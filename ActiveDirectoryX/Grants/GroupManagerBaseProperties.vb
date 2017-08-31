Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports System.DirectoryServices
Imports System.DirectoryServices.AccountManagement
Imports SSP.ActiveDirectoryX.Grants.Administration
Imports SSP.ActiveDirectoryX.Core
Imports SSP.ActiveDirectoryX.Grants.Exceptions
Imports SSP.ActiveDirectoryX.Data
Imports SSP.ActiveDirectoryX.Data.Repositories
#End Region

Namespace Grants

  Public Class GroupManagerBaseProperties

#Region " --------------->> Enumerationen der Klasse "
#End Region '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
    Private _groupManagerDn As DistinguishedName
    Private _groupManagerDnString As String
    Private _managerDn As DistinguishedName
    Private _managerDnString As String
    Private _deputiesDn As List(Of DistinguishedName)
    Private _deputiesDnString As String()
#End Region '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
    Public Sub New _
    (ByVal groupManagerDnString As String _
    , ByVal managerDnString As String _
    , ByVal deputiesDnString As String())

      Initialize(groupManagerDnString, managerDnString, deputiesDnString)
    End Sub

    Public Sub New _
    (ByVal groupManagerDnString As ResultPropertyValueCollection _
    , ByVal managerDnString As ResultPropertyValueCollection _
    , ByVal deputiesDnString As ResultPropertyValueCollection)

      Me.New(groupManagerDnString.Item(0).ToString, managerDnString.Item(0).ToString, deputiesDnString.Cast(Of String).ToArray)
    End Sub
#End Region '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
    Public Property GroupManagerDn As DistinguishedName
      Get
        If _groupManagerDn Is Nothing Then
          _groupManagerDn = DistinguishedName.GetByDistinguishedName(Me.GroupManagerDnString)
        End If
        Return _groupManagerDn
      End Get
      Set(value As DistinguishedName)
        _groupManagerDn = value
      End Set
    End Property

    Friend ReadOnly Property GroupManagerDnString As String
      Get
        Return _groupManagerDnString
      End Get

    End Property

    Public Property ManagerDn As DistinguishedName
      Get
        If _managerDn Is Nothing Then
          _managerDn = DistinguishedName.GetByDistinguishedName(Me.ManagerDnString)
        End If
        Return _managerDn
      End Get
      Set(value As DistinguishedName)
        _managerDn = value
      End Set
    End Property

    Friend ReadOnly Property ManagerDnString As String
      Get
        Return _managerDnString
      End Get
    End Property

    Public ReadOnly Property DeputiesDn As List(Of DistinguishedName)
      Get
        If _deputiesDn Is Nothing Then
          _deputiesDn = DistinguishedNameRepository.Instance.GetByDistinguishedNames(_deputiesDnString).ToList
        End If
        Return _deputiesDn
      End Get
    End Property

    Friend ReadOnly Property DeputiesDnString As String()
      Get
        Return _deputiesDnString
      End Get
    End Property
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
    Private Sub Initialize _
    (ByVal groupManagerDnString As String _
    , ByVal managerDnString As String _
    , ByVal deputiesDnString As String())

      _groupManagerDnString = groupManagerDnString
      _managerDnString = managerDnString
      _deputiesDnString = deputiesDnString
    End Sub
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
    Friend Sub AddToDeputies(ByVal deputyDn As DistinguishedName)
      If _deputiesDn Is Nothing Then
        _deputiesDn = New List(Of DistinguishedName)
      End If
      _deputiesDn.Add(deputyDn)
    End Sub

    Public Overrides Function ToString() As String
      Return Me._groupManagerDn.Name & ":" & Me._groupManagerDn.BaseProperties.Descripton & ":" & Me._managerDn.Name & "(" & _managerDn.BaseProperties.PersonId & ")" _
      & If(_deputiesDn Is Nothing, "", ":" & String.Join(",", _deputiesDn.ToList.Select(Function(x) x.Name & "(" & x.BaseProperties.EmployeeId & ")").ToArray))
    End Function
#End Region '{Öffentliche Methoden der Klasse}
  End Class
End Namespace

