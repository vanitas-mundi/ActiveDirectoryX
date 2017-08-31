Option Explicit On
Option Infer On
Option Strict On

Namespace Grants.Enums

  Public Enum FillTypes
    ''' <summary>Füllt die Granttable anhand von den hinterlegten Berechtigungen der Applikation.</summary>
    FillAll
    ''' <summary>
    ''' Füllt die Granttable anhand von den hinterlegten Berechtigungen der Applikation
    ''' und lädt die benutzerspezifischen Einstellungen aus dem GrantTree.</summary>
    FillByGrantTree
  End Enum

End Namespace

