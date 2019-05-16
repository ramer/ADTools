
Imports System.Collections.ObjectModel

Public Module mdlVariables

    Public Const ADS_UF_SCRIPT = 1 '0x1
    Public Const ADS_UF_ACCOUNTDISABLE = 2 '0x2
    Public Const ADS_UF_HOMEDIR_REQUIRED = 8 '0x8
    Public Const ADS_UF_LOCKOUT = 16 '0x10
    Public Const ADS_UF_PASSWD_NOTREQD = 32 '0x20
    Public Const ADS_UF_PASSWD_CANT_CHANGE = 64 '0x40
    Public Const ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = 128 '0x80
    Public Const ADS_UF_TEMP_DUPLICATE_ACCOUNT = 256 '0x100
    Public Const ADS_UF_NORMAL_ACCOUNT = 512 '0x200
    Public Const ADS_UF_INTERDOMAIN_TRUST_ACCOUNT = 2048 '0x800
    Public Const ADS_UF_WORKSTATION_TRUST_ACCOUNT = 4096 '0x1000
    Public Const ADS_UF_SERVER_TRUST_ACCOUNT = 8192 '0x2000
    Public Const ADS_UF_DONT_EXPIRE_PASSWD = 65536 '0x10000
    Public Const ADS_UF_MNS_LOGON_ACCOUNT = 131072 '0x20000
    Public Const ADS_UF_SMARTCARD_REQUIRED = 262144 '0x40000
    Public Const ADS_UF_TRUSTED_FOR_DELEGATION = 524288 '0x80000
    Public Const ADS_UF_NOT_DELEGATED = 1048576 '0x100000
    Public Const ADS_UF_USE_DES_KEY_ONLY = 2097152 '0x200000
    Public Const ADS_UF_DONT_REQUIRE_PREAUTH = 4194304 '0x400000
    Public Const ADS_UF_PASSWORD_EXPIRED = 8388608 '0x800000
    Public Const ADS_UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION = 16777216 '0x1000000

    Public Const ADS_GROUP_TYPE_GLOBAL_GROUP = 2 '0x00000002
    Public Const ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP = 4 '0x00000004
    Public Const ADS_GROUP_TYPE_UNIVERSAL_GROUP = 8 '0x00000008
    Public Const ADS_GROUP_TYPE_SECURITY_ENABLED = -2147483648 '0x80000000

    Public Enum enmSearchMode
        [Default] = 0
        Advanced = 1
    End Enum

    Public Enum enmClipboardAction
        Copy = 0
        Cut = 1
    End Enum

    Public Enum enmADSType
        ADSTYPE_INVALID = 0
        ADSTYPE_DN_STRING = 1
        ADSTYPE_CASE_EXACT_STRING = 2
        ADSTYPE_CASE_IGNORE_STRING = 3
        ADSTYPE_PRINTABLE_STRING = 4
        ADSTYPE_NUMERIC_STRING = 5
        ADSTYPE_BOOLEAN = 6
        ADSTYPE_INTEGER = 7
        ADSTYPE_OCTET_STRING = 8
        ADSTYPE_UTC_TIME = 9
        ADSTYPE_LARGE_INTEGER = 10
        ADSTYPE_PROV_SPECIFIC = 11
        ADSTYPE_OBJECT_CLASS = 12
        ADSTYPE_CASEIGNORE_LIST = 13
        ADSTYPE_OCTET_LIST = 14
        ADSTYPE_PATH = 15
        ADSTYPE_POSTALADDRESS = 16
        ADSTYPE_TIMESTAMP = 17
        ADSTYPE_BACKLINK = 18
        ADSTYPE_TYPEDNAME = 19
        ADSTYPE_HOLD = 20
        ADSTYPE_NETADDRESS = 21
        ADSTYPE_REPLICAPOINTER = 22
        ADSTYPE_FAXNUMBER = 23
        ADSTYPE_EMAIL = 24
        ADSTYPE_NT_SECURITY_DESCRIPTOR = 25
        ADSTYPE_UNKNOWN = 26
        ADSTYPE_DN_WITH_BINARY = 27
        ADSTYPE_DN_WITH_STRING = 28
    End Enum

    Public Enum enmDirectoryObjectSchemaClass
        User
        Contact
        Computer
        Group
        OrganizationalUnit
        Container
        DomainDNS
        UnknownContainer
        Unknown
    End Enum

    Public Enum enmDirectoryObjectStatus
        Normal
        Expired
        Blocked
    End Enum

    ''' <summary>
    ''' Useless procedure to set brake points
    ''' </summary>
    Public Sub shit()

    End Sub

    Public AttributeSchemeType As New Dictionary(Of String, Integer) From {
        {"2.5.5.1", enmADSType.ADSTYPE_DN_STRING},
        {"2.5.5.2", enmADSType.ADSTYPE_CASE_IGNORE_STRING},
        {"2.5.5.4", enmADSType.ADSTYPE_CASE_IGNORE_STRING},
        {"2.5.5.5", enmADSType.ADSTYPE_PRINTABLE_STRING},
        {"2.5.5.6", enmADSType.ADSTYPE_NUMERIC_STRING},
        {"2.5.5.7", enmADSType.ADSTYPE_DN_WITH_BINARY},
        {"2.5.5.8", enmADSType.ADSTYPE_BOOLEAN},
        {"2.5.5.9", enmADSType.ADSTYPE_INTEGER},
        {"2.5.5.10", enmADSType.ADSTYPE_OCTET_STRING},
        {"2.5.5.11", enmADSType.ADSTYPE_UTC_TIME},
        {"2.5.5.12", enmADSType.ADSTYPE_CASE_IGNORE_STRING},
        {"2.5.5.13", enmADSType.ADSTYPE_CASE_IGNORE_STRING},
        {"2.5.5.14", enmADSType.ADSTYPE_DN_WITH_STRING},
        {"2.5.5.15", enmADSType.ADSTYPE_NT_SECURITY_DESCRIPTOR},
        {"2.5.5.16", enmADSType.ADSTYPE_LARGE_INTEGER},
        {"2.5.5.17", enmADSType.ADSTYPE_OCTET_STRING}
    }

    Public attributesDefaultNames As String() = {
    "accountExpires",
    "company",
    "department",
    "description",
    "displayName",
    "distinguishedName",
    "givenName",
    "initials",
    "lastLogon",
    "location",
    "logonCount",
    "mail",
    "manager",
    "managedBy",
    "name",
    "objectGUID",
    "objectSid",
    "physicalDeliveryOfficeName",
    "pwdLastSet",
    "sAMAccountName",
    "sn",
    "telephoneNumber",
    "title",
    "userPrincipalName",
    "whenCreated"
    }

    Public attributesForSearchDefault As New ObservableCollection(Of clsAttributeSchema) From {
        New clsAttributeSchema("RDN", True, 13, "2.5.5.12", "name"),
        New clsAttributeSchema("Display-Name", True, 5, "2.5.5.12", "displayName"),
        New clsAttributeSchema("SAM-Account-Name", True, 13, "2.5.5.12", "sAMAccountName"),
        New clsAttributeSchema("User-Principal-Name", True, 1, "2.5.5.12", "userPrincipalName")}

    Public attributesForSearchExchangePermissionTarget As New ObservableCollection(Of clsAttributeSchema) From {
        New clsAttributeSchema("RDN", True, 13, "2.5.5.12", "name"),
        New clsAttributeSchema("Display-Name", True, 5, "2.5.5.12", "displayName"),
        New clsAttributeSchema("User-Principal-Name", True, 1, "2.5.5.12", "userPrincipalName")}

    Public attributesForSearchExchangePermissionFullAccess As New ObservableCollection(Of clsAttributeSchema) From {
        New clsAttributeSchema("SAM-Account-Name", True, 13, "2.5.5.12", "sAMAccountName")}

    Public attributesForSearchExchangePermissionSendAs As New ObservableCollection(Of clsAttributeSchema) From {
        New clsAttributeSchema("SAM-Account-Name", True, 13, "2.5.5.12", "sAMAccountName")}

    Public attributesForSearchExchangePermissionSendOnBehalf As New ObservableCollection(Of clsAttributeSchema) From {
        New clsAttributeSchema("RDN", True, 13, "2.5.5.12", "name")}

End Module
