Option Explicit On
Option Infer On
Option Strict On

Namespace Core.Enums

	Public Enum AdProperties
		distinguishedName = 0
		sAMAccountName = 1
		employeeID = 2
		description = 3
		member = 4
		memberOf = 5
		managedBy = 6
		accountExpires = 7
		userAccountControl = 8
		objectClass = 9
		group = 10
		user = 11
		objectGUID = 12
		cn = 13
		ou = 14
		displayName = 15
		name = 16
		msSFU30MaxUidNumber = 17
		msSFU30NisDomain = 18
		msSFU30Name = 19
		uidNumber = 20
		loginShell = 21
		unixHomeDirectory = 22
		gidNumber = 23
		mail = 24
	End Enum

End Namespace
