Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.ActiveDirectoryX.Grants.Administration
Imports SSP.ActiveDirectoryX.Data
Imports SSP.ActiveDirectoryX.Data.Repositories
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core
Imports SSP.Data.StatementBuildersAD.Core
#End Region

Namespace Grants

	Public Class OrganizationalUnits

		Implements IEnumerable(Of OrganizationalUnit)

#Region " --------------->> Enumerationen der Klasse "
#End Region  '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private _organizationalUnits As New List(Of OrganizationalUnit)
#End Region  '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
#End Region  '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		''' <summary>
		''' Stellt Funktionalität zum Administrieren von Administartions-Organisationseinheiten zur Verfügung.
		''' </summary>
		Public ReadOnly Property Administration As OrganizationalUnitsAdministration
			Get
				Return OrganizationalUnitsAdministration.Instance
			End Get
		End Property
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
		Private Function GetBaseStatement(ByVal fromDn As DistinguishedName) As SelectBuilderAD

			_organizationalUnits.Clear()
			Dim sb = AdRepositoryHelper.Instance.CreateDefaultSelectBuilder(fromDn)
			AdRepositoryHelper.Instance.SelectAddAdProperty(sb, AdProperties.distinguishedName)

			AdRepositoryHelper.Instance.SetSingleWhereAdPropertyCondition _
			(sb, AdProperties.objectClass, ObjectClasses.organizationalUnit.ToString)
			Return sb
		End Function
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
		Public Function Item(ByVal index As Int32) As OrganizationalUnit

			Return _organizationalUnits.Item(index)
		End Function

		Public Function OrganizationalUnit(ByVal index As Int32) As OrganizationalUnit

			Return _organizationalUnits.Item(index)
		End Function

		Public Function Item(ByVal dn As DistinguishedName) As OrganizationalUnit

			Return _organizationalUnits.Where(Function(ou) ou.OrganizationalUnitDn.Value = dn.Value).FirstOrDefault
		End Function

		Public Function OrganizationalUnit(ByVal dn As DistinguishedName) As OrganizationalUnit

			Return Me.Item(dn)
		End Function

		Public Function GetEnumerator() As IEnumerator(Of OrganizationalUnit) _
		Implements IEnumerable(Of OrganizationalUnit).GetEnumerator

			Return _organizationalUnits.GetEnumerator
		End Function

		Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator

			Return _organizationalUnits.GetEnumerator
		End Function

		''' <summary>
		''' Füllt das Objekt mit allen Administrations-Organisationseinheiten.
		''' </summary>
		Public Sub FillAll()

			Dim sb = GetBaseStatement(SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Administration))
			Dim distinguishedNames = sb.GetFieldList(Of String)(AdProperties.distinguishedName.ToString)

			Dim temp = Repositories.DistinguishedNameRepository.Instance.GetByDistinguishedNames(distinguishedNames)
			_organizationalUnits.AddRange(temp.Select(Function(dn) New OrganizationalUnit(dn)))
		End Sub

		''' <summary>
		''' Füllt das Objekt mit allen Administrations-Organisationseinheiten einer Administrations-Organisationseinheit.
		''' </summary>
		Public Sub FillByParentOu(ByVal parentOrganizationUnitDn As DistinguishedName)

			Dim sb = GetBaseStatement(parentOrganizationUnitDn)
			Dim distinguishedNames = sb.GetFieldList(Of String)(AdProperties.distinguishedName.ToString)

			Dim temp = DistinguishedNameRepository.Instance.GetByDistinguishedNames(distinguishedNames)
			_organizationalUnits.AddRange(temp.Select(Function(dn) New OrganizationalUnit(dn)))
		End Sub

#End Region  '{Öffentliche Methoden der Klasse}

	End Class

End Namespace

