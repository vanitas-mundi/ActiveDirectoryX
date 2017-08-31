Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.Base.Generators
Imports SSP.Base.Generators.Interfaces
#End Region

Namespace Core.Manipulation

  Public Class AdUserInfo

#Region " --------------->> Enumerationen der Klasse "
#End Region  '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
    Private _ouDistinguishedName As DistinguishedName
    Private _samAccountName As String
    Private _lastName As String
    Private _firstName As String
    Private _description As String
    Private _employeeId As Int64
    Private _phoneNumber As String
    Private _pwd As String
    Private _createMailBox As Boolean = True
    Private _activateAccount As Boolean = True
#End Region '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
    Public Sub New()
      _ouDistinguishedName = DistinguishedName.GetByDistinguishedName(My.Settings.UsersContainerDistinguishedName)
    End Sub
#End Region  '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
    ''' <summary>
    ''' Liefert den DistinguishedName der OU, in welcher der Benutzer angelegt wird oder legt diesen fest.
    ''' </summary>
    Public Property OuDistinguishedName As DistinguishedName
      Get
        Return _ouDistinguishedName
      End Get
      Set(value As DistinguishedName)
        _ouDistinguishedName = value
      End Set
    End Property

    ''' <summary>
    ''' Liefert den Username/ SamAccountName oder legt diesen fest.
    ''' </summary>
    Public Property SamAccountName As String
      Get
        Return _samAccountName
      End Get
      Set(value As String)
        _samAccountName = value
      End Set
    End Property

    ''' <summary>
    ''' Liefert den Username/ SamAccountName oder legt diesen fest.
    ''' </summary>
    Public Property UserName As String
      Get
        Return _samAccountName
      End Get
      Set(value As String)
        _samAccountName = value
      End Set
    End Property

    ''' <summary>
    ''' Liefert den DisplayName oder legt diesen fest.
    ''' </summary>
    Public ReadOnly Property DisplayName As String
      Get
        Return String.Concat(Me.LastName, ", ", Me.FirstName)
      End Get
    End Property

    ''' <summary>
    ''' Liefert die Personenbeschreibung oder legt diese fest.
    ''' </summary>
    Public Property Description As String
      Get
        Return _description
      End Get
      Set(value As String)
        _description = value
      End Set
    End Property

    ''' <summary>
    ''' Liefert die EmployeeId (Personen-Id) oder legt diese fest.
    ''' </summary>
    Public Property EmployeeId As Int64
      Get
        Return _employeeId
      End Get
      Set(value As Int64)
        _employeeId = value
      End Set
    End Property

    ''' <summary>
    ''' Liefert die EmployeeId (Personen-Id) oder legt diese fest.
    ''' </summary>
    Public Property PersonId As Int64
      Get
        Return _employeeId
      End Get
      Set(value As Int64)
        _employeeId = value
      End Set
    End Property

    ''' <summary>
    ''' Liefert den Vornamen oder legt diesen fest.
    ''' </summary>
    Public Property FirstName As String
      Get
        Return _firstName
      End Get
      Set(value As String)
        _firstName = value
      End Set
    End Property

    ''' <summary>
    ''' Liefert den Nachname oder legt diesen fest.
    ''' </summary>
    Public Property LastName As String
      Get
        Return _lastName
      End Get
      Set(value As String)
        _lastName = value
      End Set
    End Property

    ''' <summary>
    ''' Liefert den UserPrincipalName oder legt diesen fest (Bsp.: hans.meier@bcw-intern.local).
    ''' </summary>
    Public ReadOnly Property UserPrincipalName As String
      Get
        Return String.Concat(Me.SamAccountName, "@", Settings.Instance.DomainName).ToLower
      End Get
    End Property

    ''' <summary>
    ''' Liefert die Telefonnummer oder legt diesen fest.
    ''' </summary>
    Public Property PhoneNumber As String
      Get
        Return _phoneNumber
      End Get
      Set(value As String)
        _phoneNumber = value
      End Set
    End Property

    ''' <summary>
    ''' Liefert das Kennwort oder legt diese fest.
    ''' </summary>
    Public Property Pwd As String
      Get
        Return _pwd
      End Get
      Set(value As String)
        _pwd = value
      End Set
    End Property

    Public ReadOnly Property UserDistinguishedName As DistinguishedName
      Get
        Try
          Return DistinguishedName.GetByPersonId(Convert.ToInt64(Me.EmployeeId))
        Catch ex As Exception
          Return Nothing
        End Try
      End Get
    End Property

    Public Property CreateMailBox As Boolean
      Get
        Return _createMailBox
      End Get
      Set(value As Boolean)
        _createMailBox = value
      End Set
    End Property

    Public Property ActivateAccount As Boolean
      Get
        Return _activateAccount
      End Get
      Set(value As Boolean)
        _activateAccount = value
      End Set
    End Property

#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
    ''' <summary>
    ''' Erzeugt anhand der Eigenschaften firstName, lastName, personId und dem angegebenen Format einen Username
    ''' und speichert ihn als samAccountName
    ''' </summary>
    Public Sub GenerateUserName(ByVal format As UserNameFormats)

			If (String.IsNullOrEmpty(Me.FirstName)) _
			OrElse (String.IsNullOrEmpty(Me.LastName)) _
			OrElse (Me.PersonId = 0) Then
				Return
			End If

			Dim ug = New UserNameGenerator(Me.FirstName, Me.LastName, Me.PersonId) With {
				.Format = format
			}

			Me.SamAccountName = ug.Generate
		End Sub

    ''' <summary>
    ''' Erzeugt ein zufälliges komplexes Kennwort und speichert es in der Eigenschaft Pwd.
    ''' </summary>
    Public Sub GeneratePassord()

      Me.Pwd = PasswordGenerator.Generate(5, 1, 1, 1)
    End Sub
#End Region  '{Öffentliche Methoden der Klasse}

  End Class

End Namespace
