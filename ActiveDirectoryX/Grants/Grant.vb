Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports System.Runtime.Serialization
Imports SSP.ActiveDirectoryX.Core
Imports SSP.ActiveDirectoryX.Grants.Enums
#End Region

Namespace Grants

  <DataContract>
  Public Class Grant

#Region " --------------->> Enumerationen der Klasse "
#End Region '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
    Private _grantTable As GrantTable
    Private _grantName As String
    Private _value As GrantValues
    Private _description As String
#End Region  '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
    Public Sub New _
    (ByVal grantTable As GrantTable, ByVal grantName As String _
    , ByVal value As GrantValues, ByVal description As String)

      _grantTable = grantTable
      _grantName = grantName
      _value = value
      _description = description
    End Sub

    Public Sub New _
    (ByVal grantTable As GrantTable, ByVal grantName As String _
    , ByVal granted As Boolean, ByVal description As String)

      Me.New(grantTable, grantName, If(granted, GrantValues.Y, GrantValues.N), description)
    End Sub
#End Region  '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
    ''' <summary>
    ''' Name der Berechtigung.
    ''' </summary>
    Public ReadOnly Property GrantName As String
      Get
        Return _grantName
      End Get
    End Property

    ''' <summary>
    ''' Wert der Berechtigung.
    ''' </summary>
    Public Property Value As GrantValues
      Get
        Return _value
      End Get
      Set(value As GrantValues)
        _value = value
      End Set
    End Property

    ''' <summary>
    ''' Liefert true, wenn Berechtigung gewährt wurde.
    ''' </summary>
    Public Property IsGranted As Boolean
      Get
        Return _value = GrantValues.Y
      End Get
      Set(value As Boolean)
        _value = If(value, GrantValues.Y, GrantValues.N)
      End Set
    End Property

    ''' <summary>
    ''' Liefert den Wert der Berechtigung als String.
    ''' </summary>
    Public ReadOnly Property ValueString As String
      Get
        Return Me.Value.ToString
      End Get
    End Property

    ''' <summary>
    ''' Liefert die Berechtigungsbeschreibung.
    ''' </summary>
    Public ReadOnly Property Description As String
      Get
        Return _description
      End Get
    End Property

    ''' <summary>
    ''' Liefert alle Benutzer, denen die Berechtigung gewährt wurde.
    ''' </summary>
    Public ReadOnly Property AssignedUsers As DistinguishedName()
      Get
        Return Me.GrantTable.GetAssignedUsers(Me.GrantName)
      End Get
    End Property

    ''' <summary>
    ''' Liefert die Personen-Ids aller Benutzer, denen die Berechtigung gewährt wurde.
    ''' </summary>
    Public ReadOnly Property AssignedUsersPersonIds As Int64()
      Get
        Return Me.GrantTable.GetAssignedUsersPersonIds(Me.GrantName)
      End Get
    End Property

    ''' <summary>
    ''' Liefert die zugrunde liegende GrantTbale.
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property GrantTable As GrantTable
      Get
        Return _grantTable
      End Get
    End Property

#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
    Public Overrides Function ToString() As String
      Return Me.ValueString
    End Function

    ''' <summary>
    ''' Liefert den Berechtigungsnamen und den Berechtigungswert getrennt durch delimiter als String.
    ''' </summary>
    Public Function NameValueString(ByVal delimiter As String) As String
      Return Me.GrantName & delimiter & Me.ValueString
    End Function

    ''' <summary>
    ''' Liefert den Berechtigungsnamen und den Berechtigungswert getrennt durch ': '.
    ''' </summary>
    Public Function NameValueString() As String
      Return Me.NameValueString(": ")
    End Function
#End Region '{Öffentliche Methoden der Klasse}

  End Class

End Namespace
