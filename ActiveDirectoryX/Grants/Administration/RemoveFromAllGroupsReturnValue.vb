Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports/ usings "
Imports SSP.ActiveDirectoryX.Core.Enums
Imports System.Collections.ObjectModel
Imports SSP.ActiveDirectoryX.Core
Imports System.ComponentModel

#End Region

Namespace Grants.Administration

	<DefaultProperty("AdManipulationResult")>
	Public Class AdManipulationResultsErrorsValue

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private _adManipulationResult As AdManipulationResults
		Private _errors As New List(Of AdManipulationResultsErrorValue)
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Friend Sub New(ByVal adManipulationResult As AdManipulationResults, _
		 ByVal groupsErrorValue As IEnumerable(Of AdManipulationResultsErrorValue))
			_adManipulationResult = adManipulationResult
			_errors.AddRange(groupsErrorValue)
		End Sub

		Friend Sub New(ByVal adManipulationResult As AdManipulationResults)
			_adManipulationResult = adManipulationResult
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		Public ReadOnly Property AdManipulationResult As AdManipulationResults
		Get
			Return _adManipulationResult
		End Get
		End Property

		Public ReadOnly Property Errors As ReadOnlyCollection(Of AdManipulationResultsErrorValue)
		Get
			Return _errors.AsReadOnly
		End Get
		End Property
#End Region	'{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region	'{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region	'{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
		Public Overrides Function ToString() As String
			Return Me.AdManipulationResult.ToString
		End Function
#End Region	'{Öffentliche Methoden der Klasse}

	End Class

End Namespace
