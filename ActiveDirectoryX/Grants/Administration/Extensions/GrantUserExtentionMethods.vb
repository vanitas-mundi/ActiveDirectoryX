Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports/ usings "
Imports System.Runtime.CompilerServices
Imports SSP.ActiveDirectoryX.Core.Enums
#End Region

Namespace Grants.Administration.Extensions

	Module GrantUserExtentionMethods

#Region " --------------->> Enumerationen der Klasse "
#End Region	'{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
#End Region	'{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
#End Region	'{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
#End Region	'{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region	'{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region	'{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
		''' <summary>
		''' Aktualisiert die übergebene AD-Eigenschaft.
		''' </summary>
		<Extension>
		Public Function SetProperty _
		(ByVal grantUser As GrantUser _
		, ByVal propertyName As AdProperties _
		, ByVal values As Object()) As AdManipulationResults

			Return grantUser.Administration.SetProperty(propertyName, values)
		End Function
#End Region	'{Öffentliche Methoden der Klasse}

	End Module

End Namespace