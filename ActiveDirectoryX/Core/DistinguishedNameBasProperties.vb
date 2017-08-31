Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports/ usings "
#End Region

Namespace Core

  Public Class DistinguishedNameBaseProperties

#Region " --------------->> Enumerationen der Klasse "
#End Region '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
    Private _distinguishedName As String
    Private _name As String
    Private _descripton As String
    Private _employeeId As String
    Private _personId As Int64
    Private _objectGuid As Guid
#End Region '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
    Public Sub New(ByVal distinguishedName As String, ByVal name As String _
    , ByVal description As String, ByVal objectGuidString As String, ByVal employeeId As String)

      Me.New(distinguishedName, name, description, If(String.IsNullOrEmpty(objectGuidString), New Guid, New Guid(objectGuidString)), employeeId)
    End Sub

    Public Sub New(ByVal distinguishedName As String, ByVal name As String _
    , ByVal description As String, ByVal objectGuid As Guid, ByVal employeeId As String)

      _distinguishedName = distinguishedName
      _name = name
      _descripton = description
      _objectGuid = objectGuid
      _employeeId = employeeId
      Int64.TryParse(_employeeId, _personId)
    End Sub

    Public Sub New(ByVal distinguishedName As String, ByVal name As String _
    , ByVal description As String, ByVal objectGuid As Guid, ByVal employeeId As Object)

      _distinguishedName = distinguishedName
      _name = name
      _descripton = description
      _objectGuid = objectGuid
      _employeeId = If(Convert.IsDBNull(employeeId), "", employeeId.ToString)
      Int64.TryParse(_employeeId, _personId)
    End Sub

    Public Sub New(ByVal propertyArray As Object())
      _distinguishedName = propertyArray(0).ToString
      _name = propertyArray(1).ToString
      _descripton = If(Convert.IsDBNull(propertyArray(2)), "", CType(propertyArray(2), Object())(0).ToString)
      _objectGuid = New Guid(CType(propertyArray(3), Byte()))
      _employeeId = propertyArray(4).ToString
      Int64.TryParse(_employeeId, _personId)
    End Sub
#End Region '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
    ''' <summary>
    ''' Liefert den DistinguishedName.
    ''' Bsp: ou=test, dc=domain, dc=net
    ''' </summary>
    Public ReadOnly Property DistinguishedName As String
      Get
        Return _distinguishedName
      End Get
    End Property

    ''' <summary>
    ''' Liefert den Namen des AD-Objektes.
    ''' Bsp: ou=test, dc=domain, dc=net
    ''' </summary>
    Public ReadOnly Property Name As String
      Get
        Return _name
      End Get
    End Property

    '''<summary>Liefert die Beschreibung des AD-Objektes.</summary>
    Public ReadOnly Property Descripton As String
      Get
        Return _descripton
      End Get
    End Property

    '''<summary>Liefert die Employee-Id (Peronen-Id) aus dem AD.</summary>
    Public ReadOnly Property EmployeeId As String
      Get
        Return _employeeId
      End Get
    End Property

    '''<summary>Liefert die Employee-Id (Peronen-Id) aus dem AD.</summary>
    Public ReadOnly Property PersonId As Int64
      Get
        Return _personId
      End Get
    End Property

    ''' <summary>
    ''' Liefert die ObjectGuid des DistinguishedNames.
    ''' Bsp: ou=test, dc=domain, dc=net
    ''' </summary>
    Public ReadOnly Property ObjectGuid As Guid
      Get
        Return _objectGuid
      End Get
    End Property

    ''' <summary>
    ''' Liefert die ObjectGuid des DistinguishedNames als Stirng.
    ''' Bsp: ou=test, dc=domain, dc=net
    ''' </summary>
    Public ReadOnly Property ObjectGuidString As String
      Get
        Return _objectGuid.ToString
      End Get
    End Property

#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
    Public Shared Function GetEmptyProperties() As DistinguishedNameBaseProperties
      Return New DistinguishedNameBaseProperties("", "", "", New Guid, "")
    End Function
#End Region '{Öffentliche Methoden der Klasse}

  End Class

End Namespace

