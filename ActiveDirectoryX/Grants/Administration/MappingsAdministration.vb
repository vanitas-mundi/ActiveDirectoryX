Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Core
#End Region

Namespace Grants.Administration

	Public Class MappingsAdministration

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private Shared _instance As MappingsAdministration
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Shared Sub New()
			_instance = New MappingsAdministration
		End Sub

		Private Sub New()
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		Public Shared ReadOnly Property Instance As MappingsAdministration
		Get
			Return _instance
		End Get
		End Property
#End Region	'{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region	'{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region	'{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "

#Region " --> CreateMapping "

	''' <summary>
	''' Erstellt ein neues Mapping.
	''' </summary>
	Public Function CreateMapping(ByVal mappingName As String _
	, ByVal unc As String, ByVal driveLetter As String) As AdManipulationResults

		Try
			mappingName = mappingName.ToLower

			If Not mappingName.StartsWith("mapping.") Then Return AdManipulationResults.InvalidMappingName

			Dim mappingsOuDn = SpecialDistinguishedNames.Item(SpecialDistinguishedNameKeys.Mappings)
			Dim result = Administrations.Instance.CreateGroup _
			(mappingsOuDn, mappingName, SpecialDistinguishedNameKeys.Mappings)

			If result = AdManipulationResults.Successful Then
				Dim dn = DistinguishedName.GetByGroupName(mappingName)
				Dim entry = dn.ToDirectoryEntry(True)
				Dim value = String.Format("{0} | {1}:\", unc, driveLetter.Replace(":", "".Replace("\", "")))
				entry.InvokeSet(AdProperties.description.ToString, New Object() {value})
				entry.CommitChanges()
				Return AdManipulationResults.Successful
			Else
				Return result
			End If
		Catch ex As UnauthorizedAccessException
			Return AdManipulationResults.AccesDenied
		Catch ex As System.Exception
			Return AdManipulationResults.UnknownError
		End Try
	End Function
#End Region

#Region " --> DeleteMapping "
	''' <summary>
	''' Löscht das angegebene Mapping.
	''' </summary>
	Public Function DeleteMapping(ByVal mappingName As String) As AdManipulationResults
		 Return DeleteMapping(DistinguishedName.GetByGroupName(mappingName))
	End Function

	''' <summary>
	''' Löscht das angegebene Mapping.
	''' </summary>
	Public Function DeleteMapping(ByVal mappingDn As DistinguishedName) As AdManipulationResults

		Return Administrations.Instance.DeleteGroup(mappingDn, SpecialDistinguishedNameKeys.Mappings)
	End Function
#End Region

#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace


