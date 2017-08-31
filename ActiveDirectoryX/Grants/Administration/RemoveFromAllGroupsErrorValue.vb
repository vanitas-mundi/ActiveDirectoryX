Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports/ usings "
Imports SSP.ActiveDirectoryX.Core
#End Region

Namespace Grants.Administration

	Public Class AdManipulationResultsErrorValue

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
		Private _groupDn As DistinguishedName
		Private _exception As System.Exception
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
		Friend Sub New(ByVal groupDn As DistinguishedName, ByVal exception As System.Exception)
			_groupDn = groupDn
			_exception = exception
		End Sub
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
		Public ReadOnly Property GroupDn As DistinguishedName
		Get
			Return _groupDn
		End Get
		End Property

		Public ReadOnly Property Exception As System.Exception
		Get
			Return _exception
		End Get
		End Property
#End Region	'{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region	'{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region	'{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
		Public Overrides Function ToString() As String
			Return Me.GroupDn.Name & " - " & Me.Exception.Message
		End Function
#End Region	'{Öffentliche Methoden der Klasse}

	End Class
End Namespace
