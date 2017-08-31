Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.ActiveDirectoryX.Grants.Administration
Imports SSP.ActiveDirectoryX.Grants.Enums
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core
Imports SSP.ActiveDirectoryX.Grants.Exceptions
#End Region

Namespace Grants

	Public Class AdministrationGroup

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private _groupDistinguishedName As DistinguishedName
    Private _administration As AdministrationGroupAdministration
#End Region '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
    Public Sub New(ByVal distinguishedName As DistinguishedName)
      If Not AdministrationTypeResolver.Instance.IsAdministrationGroup(distinguishedName) Then
        Throw New WrongAdministrationTypeException
      End If

      _groupDistinguishedName = distinguishedName
      _administration = New AdministrationGroupAdministration(Me)
    End Sub
#End Region  '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
    ''' <summary>
    ''' Guid des AD-Objektes.
    ''' </summary>
    Public ReadOnly Property ObjectGuid As Guid
      Get
        Return _groupDistinguishedName.BaseProperties.ObjectGuid
      End Get
    End Property

    ''' <summary>
    ''' String der Guid des AD-Objektes.
    ''' </summary>
    Public ReadOnly Property ObjectGuidString As String
      Get
        Return Me.ObjectGuid.ToString
      End Get
    End Property

    ''' <summary>
    ''' Stellt Funktionalität zum Administrieren einer Gruppe zur Verfügung.
    ''' </summary>
    Public ReadOnly Property Administration As AdministrationGroupAdministration
		Get
			Return _administration
		End Get
		End Property

		''' <summary>
		''' Liefert den Namen der Gruppe.
		''' </summary>
		Public ReadOnly Property Name As String
		Get
			Return _groupDistinguishedName.ToRelativeName
		End Get
		End Property

		''' <summary>
		''' Liefert den Namen der Gruppe ohne führendem Gruppentypnamen.
		''' </summary>
		Public ReadOnly Property NameSimple As String
		Get
			Return Me.Name.Split("."c).LastOrDefault
		End Get
		End Property

		''' <summary>
		''' Liefert die Beschreibung der Gruppe.
		''' </summary>
		Public ReadOnly Property Description As String
		Get
			Dim obj = Me.GroupDistinguishedName.GetProperty(AdProperties.description.ToString)
			Return If(obj Is Nothing, "", obj.ToString)
		End Get
		End Property

		''' <summary>
		''' Liefert den DistinguishedName der Gruppe.
		''' </summary>
		Public ReadOnly Property GroupDistinguishedName As DistinguishedName
		Get
			Return _groupDistinguishedName
		End Get
		End Property

		''' <summary>
		''' Liefert die Mitglieder als DistinguishedName-Array.
		''' </summary>
		Public ReadOnly Property Members As DistinguishedName()
      Get
        Return _groupDistinguishedName.GetMembers
      End Get
    End Property

    ''' <summary>
    ''' Liefert ein DistinguishedName-Array der AD-Objekte, welche die Rolle als Mitglied beinhalten.
    ''' </summary>
    Public ReadOnly Property MembersOf As DistinguishedName()
		Get
			Return _groupDistinguishedName.GetMemberOf
		End Get
		End Property

		''' <summary>
		''' Liefert den Organisationsgruppentyp der Gruppe.
		''' </summary>
		Public ReadOnly Property OrganizationGroupType As OrganizationGroupTypes
		Get
			Select Case True
			Case Me.GroupDistinguishedName.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.HolidayGroups))
				Return OrganizationGroupTypes.HolidayGroup
			Case Me.GroupDistinguishedName.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.AccountingGroups))
				Return OrganizationGroupTypes.AccountingGroup
			Case Me.GroupDistinguishedName.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.WorkGroups))
				Return OrganizationGroupTypes.WorkGroup
			Case Me.GroupDistinguishedName.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.OrganizationGroups))
				Return OrganizationGroupTypes.CommonOrganizationGroup
			Case Else
				Return OrganizationGroupTypes.NoOrganizationGroup
			End Select
		End Get
		End Property

		''' <summary>
		''' Prüft, ob Gruppe eine Organisationsgruppe ist.
		''' </summary>
		Public ReadOnly Property IsOrganizationGroup As Boolean
		Get
			Return Not Me.OrganizationGroupType = OrganizationGroupTypes.NoOrganizationGroup
		End Get
		End Property

		''' <summary>
		''' Liefert den Berechtigungstyp der Gruppe.
		''' </summary>
		Public ReadOnly Property GrantType As GrantTypes
		Get
			With Me.GroupDistinguishedName
				Select Case True
				Case .ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.DepartmentRoles))
					Return GrantTypes.DepartmentRole
				Case .ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.ApplicationRoles))
					Return GrantTypes.ApplicationRole
				Case .ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.BaseRoles))
					Return GrantTypes.BaseRole
				Case .ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.DenialRoles))
					Return GrantTypes.DenialRole
				Case .ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.ExtraRoles))
					Return GrantTypes.ExtraRole
				Case .ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.TeamRoles))
					Return GrantTypes.TeamRole
				Case Me.GroupDistinguishedName.ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.GroupManagers))
					Return GrantTypes.GroupManager
				Case .ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Mappings))
					Return GrantTypes.Mapping
				Case .ContainsDn(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Roles))
					Return GrantTypes.CommonRole
				Case Else
					Return GrantTypes.NoRole
				End Select
			End With
		End Get
		End Property

		''' <summary>
		''' Prüft, ob Gruppe eine Berechtigungsrolle ist.
		''' </summary>
		Public ReadOnly Property IsGrantGroup As Boolean
		Get
			Return Not Me.GrantType = GrantTypes.NoRole
		End Get
		End Property

		''' <summary>
		''' Liefert das GroupManager-Objekt der Gruppe
		''' </summary>
		Public ReadOnly Property GroupManager As GroupManager
      Get
        Return GroupManagers.Instance.GetGroupManagerOf(Me)
      End Get
    End Property
#End Region	'{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region	'{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region	'{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
		Public Overrides Function ToString() As String
			Return Me.GroupDistinguishedName.Name
		End Function
#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace


