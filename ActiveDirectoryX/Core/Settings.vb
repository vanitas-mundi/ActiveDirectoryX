Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports System.DirectoryServices.AccountManagement
Imports System.DirectoryServices.ActiveDirectory
Imports SSP.Base
Imports SSP.Data.StatementBuildersAD.Core
#End Region

Namespace Core

  Public Class Settings

#Region " --------------->> Enumerationen der Klasse "
#End Region  '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
    Private Shared _instance As Settings
    Private _connectionStringBuilder As ConnectionStringBuilderAD
    Private _domainName As String
    Private _domainController As String
#End Region  '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
    Private Sub New()
      _domainName = Domain.GetCurrentDomain.Name '"bcw-intern.local"
      Dim context = New DirectoryContext(DirectoryContextType.Domain, DomainName)
      _domainController = Domain.GetDomain(context).FindDomainController.Name '"DC02.bcw-intern.local"
    End Sub

    Shared Sub New()
      _instance = New Settings
      DbResultAD.Initialize(Settings.Instance.ConnectionString)
    End Sub
#End Region  '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
    Public Shared ReadOnly Property Instance As Settings
      Get
        Return _instance
      End Get
    End Property

    ''' <summary>
    ''' Liefert das Kennzeichen für Verweigerungsgruppen.
    ''' </summary>
    Public ReadOnly Property DenialChar As Char
      Get
        Return "~"c
      End Get
    End Property

    ''' <summary>
    ''' Liefert den DN zum Lesen oder Setzen der UnixId.
    ''' </summary>
    Public ReadOnly Property UnixIdDistinguishedName As String
      Get
        Return My.Settings.UnixIdDistinguishedName
      End Get
    End Property

    ''' <summary>
    ''' Liefert den Namen des aktuellen Domain-Controllers.
    ''' </summary>
    Public ReadOnly Property DomainController() As String
      Get
        Return _domainController
        'If String.IsNullOrWhiteSpace(DomainName) Then
        '	Return Domain.GetCurrentDomain.FindDomainController.Name
        'Else
        '	Dim context = New DirectoryContext(DirectoryContextType.Domain, DomainName)
        '	Return Domain.GetDomain(context).FindDomainController.Name
        'End If
      End Get
    End Property

    ''' <summary>
    ''' Liefert den Domänennamen oder legt diesen fest.
    ''' </summary>
    Public Property DomainName As String
      Get
        Return _domainName
      End Get
      Set(value As String)
        _domainName = value
      End Set
    End Property

    ''' <summary>
    ''' Liefert den Top-Level-Domänennamen.
    ''' </summary>
    Public ReadOnly Property TopLevelDomainName As String
      Get
        Return _domainName.Substring(_domainName.LastIndexOf(".") + 1)
      End Get
    End Property

    ''' <summary>
    ''' Liefert den Second-Level-Domänennamen.
    ''' </summary>
    Public ReadOnly Property SecondLevelDomainName As String
      Get
        Return _domainName.Substring(0, _domainName.LastIndexOf("."))
      End Get
    End Property

    ''' <summary>
    ''' Liefert einen Builder zum Festlegen der Verbindungszeichenfolge für den SQL-LIKE-Zugriff.
    ''' </summary>
    Public ReadOnly Property ConnectionStringBuilder As ConnectionStringBuilderAD
      Get
        If _connectionStringBuilder Is Nothing Then
          _connectionStringBuilder = New ConnectionStringBuilderAD _
          (Settings.Instance.DomainName, Settings.Instance.DomainController)
        End If
        Return _connectionStringBuilder
      End Get
    End Property

    ''' <summary>
    ''' Liefert die Verbindungszeichenfolge für den SQL-LIKE-Zugriff oder legt diese fest.
    ''' </summary>
    Public Property ConnectionString As String
      Get
        Return ConnectionStringBuilder.ConnectionString
      End Get
      Set(value As String)
        _connectionStringBuilder = New ConnectionStringBuilderAD(value)
      End Set
    End Property

    ''' <summary>
    ''' Liefert die Gruppen-Id der Domänen-Admin-Gruppe.
    ''' </summary>
    Public ReadOnly Property DomainAdminGroupId As Int32
      Get
        Return My.Settings.DomainAdminGroupId
      End Get
    End Property

    ''' <summary>
    ''' Liefert die Gruppen-Id der Domänen-Benutzer-Gruppe.
    ''' </summary>
    Public ReadOnly Property DomainUserGroupId As Int32
      Get
        Return My.Settings.DomainAdminGroupId
      End Get
    End Property

    ''' <summary>
    ''' Liefert die primäre Mail-Domain.
    ''' </summary>
    Public ReadOnly Property PrimaryMailDomain As String
      Get
        Return My.Settings.PrimaryMailDomain
      End Get
    End Property

    ''' <summary>
    ''' Liefert den UserName welcher für AD-Manupulationen verwendet werden soll.
    ''' Enthält diese Eigenschaft einen Leestring, wird der angemeldete Windows-Users verwendet.
    ''' </summary>
    Public ReadOnly Property ManipulationUserName As String
      Get
        Return My.Settings.ManipulationUserName
      End Get
    End Property

    ''' <summary>
    ''' Liefert das UserPassword welches für AD-Manupulationen verwendet werden soll.
    ''' Enthält diese Eigenschaft einen Leestring, wird der angemeldete Windows-Users verwendet.
    ''' </summary>
    Public ReadOnly Property ManipulationUserPassword As String
      Get
        Return Helper.Crypt.DecryptString(My.Settings.ManipulationUserPassword)
      End Get
    End Property

    ''' <summary>
    ''' Liefert den ManipulationPrincipalContext welcher für AD-Manupulationen verwendet werden soll.
    ''' </summary>
    Public ReadOnly Property ManipulationPrincipalContext As PrincipalContext
      Get
        Return AdPrincipals.GetPrincipalContext(Settings.Instance.ManipulationUserName, Settings.Instance.ManipulationUserPassword)
      End Get
    End Property
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
#End Region '{Öffentliche Methoden der Klasse}

  End Class

End Namespace
