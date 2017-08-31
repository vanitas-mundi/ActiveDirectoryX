Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "

Imports SSP.ActiveDirectoryX.Grants.Administration
Imports SSP.ActiveDirectoryX.Grants.Enums
Imports SSP.ActiveDirectoryX.Data.Repositories
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core
Imports SSP.Data.StatementBuildersAD.Core
Imports SSP.ActiveDirectoryX.Data
#End Region

Namespace Grants

	Public Class AdministrationGroups

		Implements IEnumerable(Of AdministrationGroup)

#Region " --------------->> Enumerationen der Klasse "
#End Region '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private _administrationGroups As New List(Of AdministrationGroup)
#End Region  '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
#End Region  '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		''' <summary>
		''' Stellt Funktionalität zum Administrieren von Gruppen zur Verfügung.
		''' </summary>
		Public ReadOnly Property Administration As AdministrationGroupsAdministration
			Get
				Return AdministrationGroupsAdministration.Instance
			End Get
		End Property
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
		Private Function GetBaseStatement(ByVal specialDistinguishedName As SpecialDistinguishedNameKeys) As SelectBuilderAD

			_administrationGroups.Clear()
			Dim sb = AdRepositoryHelper.Instance.CreateDefaultSelectBuilder(specialDistinguishedName)
			AdRepositoryHelper.Instance.SelectAddAdProperty(sb, AdProperties.distinguishedName)

			AdRepositoryHelper.Instance.SetSingleWhereAdPropertyCondition _
			(sb, AdProperties.objectClass, ObjectClasses.group.ToString)
			Return sb
		End Function

		Private Sub FillBase(ByVal dnName As SpecialDistinguishedNameKeys)

			Dim sb = GetBaseStatement(dnName)
			Dim distinguishedNames = DistinguishedNameRepository.Instance.GetByDistinguishedNames(sb)
			_administrationGroups.AddRange(distinguishedNames.Select(Function(dn) New AdministrationGroup(dn)))
		End Sub
#End Region  '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
		Public Function Item(ByVal index As Int32) As AdministrationGroup

			Return _administrationGroups.Item(index)
		End Function

		Public Function AdministrationGroup(ByVal index As Int32) As AdministrationGroup

			Return _administrationGroups.Item(index)
		End Function

		Public Function Item(ByVal dn As DistinguishedName) As AdministrationGroup

			Return _administrationGroups.Where(Function(g) g.GroupDistinguishedName.Value = dn.Value).FirstOrDefault
		End Function

		Public Function AdministrationGroup(ByVal dn As DistinguishedName) As AdministrationGroup

			Return Me.Item(dn)
		End Function

		Public Function GetEnumerator() As IEnumerator(Of AdministrationGroup) _
		Implements IEnumerable(Of AdministrationGroup).GetEnumerator

			Return _administrationGroups.GetEnumerator
		End Function

		Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator

			Return _administrationGroups.GetEnumerator
		End Function

    ''' <summary>
    ''' Füllt das Objekt mit allen Administrationsgruppen.
    ''' </summary>
    Public Sub FillAll()

			Dim sb = GetBaseStatement(SpecialDistinguishedNameKeys.Administration)
			Dim distinguishedNames = sb.GetFieldList(Of String)(AdProperties.distinguishedName.ToString)

			Dim temp = DistinguishedNameRepository.Instance.GetByDistinguishedNames(distinguishedNames)
			_administrationGroups.AddRange(temp.Select(Function(dn) New AdministrationGroup(dn)))
		End Sub

    ''' <summary>
    ''' Füllt das Objekt mit allen Berechtigungsgruppen.
    ''' </summary>
    Public Sub FillAllGrantGroups()
			FillBase(SpecialDistinguishedNameKeys.Grants)
		End Sub

		''' <summary>
		''' Füllt das Objekt mit Rollen einen Rollentyps.
		''' </summary>
		Public Sub FillByGrantType(ByVal grantType As GrantTypes)

			Me.FillAllGrantGroups()
			Dim result = Me.Where(Function(g) g.GrantType = grantType).ToArray()
			_administrationGroups.Clear()
			_administrationGroups.AddRange(result)
		End Sub

		''' <summary>
		''' Füllt das Objekt mit allen Organisationsgruppen.
		''' </summary>
		Public Sub FillAllOrganizationGroups()
			FillBase(SpecialDistinguishedNameKeys.OrganizationGroups)
		End Sub

		''' <summary>
		''' Füllt das Objekt mit Organisationsgruppen eines Organisationsgruppentyps.
		''' </summary>
		Public Sub FillByOrganizationGroupType(ByVal organizationGroupType As OrganizationGroupTypes)

			Me.FillAllOrganizationGroups()
			Dim result = Me.Where(Function(g) g.OrganizationGroupType = organizationGroupType).ToArray()
			_administrationGroups.Clear()
			_administrationGroups.AddRange(result)
		End Sub

		''' <summary>
		''' Füllt das Objekt mit allen Organisationsgruppen.
		''' </summary>
		Public Sub FillAllMappings()
			FillBase(SpecialDistinguishedNameKeys.Mappings)
		End Sub

		''' <summary>
		''' Füllt das Objekt mit allen Gruppenmanagern.
		''' </summary>
		Public Sub FillAllGroupManagers()
			FillBase(SpecialDistinguishedNameKeys.GroupManagers)
		End Sub
#End Region  '{Öffentliche Methoden der Klasse}

	End Class

End Namespace

