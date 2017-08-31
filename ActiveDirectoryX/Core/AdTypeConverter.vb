Option Explicit On
Option Infer On
Option Strict On

#Region " --------------->> Imports "

Imports System.Text
Imports SSP.ActiveDirectoryX.Core.Enums
Imports SSP.ActiveDirectoryX.Data
Imports SSP.Data.StatementBuildersAD.Core
'Imports SSP.Data.StatementBuildersAD.Core

#End Region

Namespace Core

	Public Class AdTypeConverter

#Region " --------------->> Enumerationen der Klasse "
#End Region '{Enumerationen der Klasse}

#Region " --------------->> Eigenschaften der Klasse "
#End Region '{Eigenschaften der Klasse}

#Region " --------------->> Konstruktor und Destruktor der Klasse "
#End Region '{Konstruktor und Destruktor der Klasse}

#Region " --------------->> Zugriffsmethoden der Klasse "
#End Region '{Zugriffsmethoden der Klasse}

#Region " --------------->> Ereignismethoden Methoden der Klasse "
#End Region '{Ereignismethoden der Klasse}

#Region " --------------->> Private Methoden der Klasse "
#End Region '{Private Methoden der Klasse}

#Region " --------------->> Öffentliche Methoden der Klasse "
		''' <summary>
		''' Liefert den LowPart eines LargeInteger(Integer8) AD-Types
		''' </summary>
		Public Shared Function ObjectToLargeIntegerLowPart(ByVal value As Object) As Int64
			Return Convert.ToInt64(value.GetType.InvokeMember _
				("LowPart", Reflection.BindingFlags.GetProperty, Nothing, value, Nothing))
		End Function

		''' <summary>
		''' Liefert den HighPart eines LargeInteger(Integer8) AD-Types
		''' </summary>
		Public Shared Function ObjectToLargeIntegerHighPart(ByVal value As Object) As Int64
			Return Convert.ToInt64(value.GetType.InvokeMember _
				("HighPart", Reflection.BindingFlags.GetProperty, Nothing, value, Nothing))
		End Function

		''' <summary>
		''' Wandelt einen LargeInteger(Integer8) AD-Type in ein Datum um.
		''' </summary>
		Public Shared Function LargeIntegerToDate(ByVal value As Object) As Date?
			Try
				Dim low = ObjectToLargeIntegerLowPart(value)
				Dim high = ObjectToLargeIntegerHighPart(value)

				If (high << 32 = -1) AndAlso (low = -1) Then
					Return Nothing
				Else
					Dim fileTime = Convert.ToInt64((high << 32) + low)
					Dim tempDate = DateTime.FromFileTime(fileTime)
					If tempDate.Year = 1601 Then
						Return Nothing
					Else
						Return tempDate
					End If
				End If

			Catch ex As Exception
				Return Nothing
			End Try
		End Function

		''' <summary>
		''' Prüft, ob ein Account userName im AD existiert.
		''' </summary>
		Public Shared Function UserNameExists(ByVal userName As String) As Boolean
			Dim sb = AdRepositoryHelper.Instance.CreateDefaultSelectBuilder
			sb.Select.Add(AdProperties.distinguishedName.ToString)
			sb.Where.Add(String.Format("{0}='{1}'", AdProperties.sAMAccountName.ToString, userName))
			Return sb.ExecuteScalar IsNot Nothing
		End Function

		''' <summary>
		''' Liefert die NativeGuid einer Guid.
		''' </summary>
		Public Shared Function GuidToNativeGuid(ByVal guid As Guid) As String
			Return GuidToNativeGuid(guid.ToString)
		End Function

		''' <summary>
		''' Liefert die NativeGuid einer Guid.
		''' </summary>
		Public Shared Function GuidToNativeGuid(ByVal guid As String) As String
			'Eingabe: 9D405236-1D3A-4332-BB1B-E242A3321846
			'Ausgabe: 3652409d3a1d3243bb1be242a3321846

			'Trenne die Guid zwischen den Bindestrichen und erzeuge daraus ein Array.
			Dim parts = guid.Split("-"c)
			Dim sb As StringBuilder

			'Durchlaufe die ersten 3 Array-Elemente
			For i = 0 To 2
				Dim part = parts(i)
				sb = New StringBuilder

				'Durchlaufe das Array-Element in Zwei-Zeichen-Blöcken, angefangen vom letzen zum ertsen Block
				For j = (part.Length \ 2) - 1 To 0 Step -1
					'Vertausche die Zweierblöcke
					sb.Append(part.Substring(j * 2, 2))
				Next j

				parts(i) = sb.ToString
			Next i

			'Gib die umgewandelte Guid zurück
			Return String.Join("", parts)
		End Function

		''' <summary>
		''' Liefert die maskierte NativeGuid einer Guid.
		''' </summary>
		Public Shared Function GuidToEscapedNativeGuid(ByVal guid As Guid) As String

			Return GuidToEscapedNativeGuid(guid)
		End Function

		''' <summary>
		''' Liefert die maskierte NativeGuid einer Guid.
		''' </summary>
		Public Shared Function GuidToEscapedNativeGuid(ByVal guid As String) As String
			'Eingabe: 9D405236-1D3A-4332-BB1B-E242A3321846
			'Ausgabe: \36\52\40\9d\3a\1d\32\43\bb\1b\e2\42\a3\32\18\46

			'Erzeuge NativeGuid
			Dim nativeGuidString = GuidToNativeGuid(guid)
			Dim sb = New StringBuilder

			'Maskiere die NativeGuid mit Backslashes
			For i = 0 To (nativeGuidString.Length \ 2) - 1
				sb.Append("\" & nativeGuidString.Substring(i * 2, 2))
			Next i

			Return sb.ToString
		End Function
#End Region '{Öffentliche Methoden der Klasse}

	End Class

End Namespace




