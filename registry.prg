* Copyright (c) 1995,1996 Sierra Systems, Microsoft Corporation
*
* Written by Randy Brown
* Contributions from Matt Oshry, Calvin Hsia
*
* The Registry class provides a complete library of API
* calls to access the Windows Registry. Support is provided
* for Windows 32S, Windows NT amd Windows 95. Included for
* backward compatibility with older applications which still
* use INI files are several routines which access INI sections
* and details. Finally, several valuable routines are included
* for accessing ODBC drivers and data sources.
*


* Operating System codes
#Define  OS_W32S        1
#Define  OS_NT        2
#Define  OS_WIN95      3
#Define  OS_MAC        4
#Define  OS_DOS        5
#Define  OS_UNIX        6

* DLL Paths for various operating systems
#Define DLLPATH_32S      "\SYSTEM\"    &&used for ODBC only
#Define DLLPATH_NT      "\SYSTEM32\"
#Define DLLPATH_WIN95    "\SYSTEM\"

* DLL files used to read INI files
#Define  DLL_KERNEL_W32S    "W32SCOMB.DLL"
#Define  DLL_KERNEL_NT    "KERNEL32.DLL"
#Define  DLL_KERNEL_WIN95  "KERNEL32.DLL"

* DLL files used to read registry
#Define  DLL_ADVAPI_W32S    "W32SCOMB.DLL"
#Define  DLL_ADVAPI_NT    "ADVAPI32.DLL"
#Define  DLL_ADVAPI_WIN95  "ADVAPI32.DLL"

* DLL files used to read ODBC info
#Define DLL_ODBC_W32S    "ODBC32.DLL"
#Define DLL_ODBC_NT      "ODBC32.DLL"
#Define DLL_ODBC_WIN95    "ODBC32.DLL"

* Registry roots
#Define HKEY_CLASSES_ROOT           -2147483648  && BITSET(0,31)
#Define HKEY_CURRENT_USER           -2147483647  && BITSET(0,31)+1
#Define HKEY_LOCAL_MACHINE          -2147483646  && BITSET(0,31)+2
#Define HKEY_USERS                  -2147483645  && BITSET(0,31)+3

* Misc
#Define APP_PATH_KEY    "\Shell\Open\Command"
#Define OLE_PATH_KEY    "\Protocol\StdFileEditing\Server"
#Define VFP_OPTIONS_KEY    "Software\Microsoft\VisualFoxPro\6.0\Options"
#Define VFP_OPT32S_KEY    "VisualFoxPro\6.0\Options"
#Define CURVER_KEY      "\CurVer"
#Define ODBC_DATA_KEY    "Software\ODBC\ODBC.INI\"
#Define ODBC_DRVRS_KEY    "Software\ODBC\ODBCINST.INI\"
#Define SQL_FETCH_NEXT    1
#Define SQL_NO_DATA      100
#Define VFP_OPTIONS_KEY1  "Software\Microsoft\VisualFoxPro\"
#Define VFP_OPTIONS_KEY2  "\Options"


* Error Codes
#Define ERROR_SUCCESS    0  && OK
#Define ERROR_EOF       259 && no more entries in key

* Note these next error codes are specific to this Class, not DLL
#Define ERROR_NOAPIFILE    -101  && DLL file to check registry not found
#Define ERROR_KEYNOREG    -102  && key not registered
#Define ERROR_BADPARM    -103  && bad parameter passed
#Define ERROR_NOENTRY    -104  && entry not found
#Define  ERROR_BADKEY    -105  && bad key passed
#Define  ERROR_NONSTR_DATA  -106  && data type for value is not a data string
#Define ERROR_BADPLAT    -107  && platform not supported
#Define ERROR_NOINIFILE    -108  && DLL file to check INI not found
#Define ERROR_NOINIENTRY  -109  && No entry in INI file
#Define ERROR_FAILINI    -110  && failed to get INI entry
#Define ERROR_NOPLAT    -111  && call not supported on this platform
#Define ERROR_NOODBCFILE  -112  && DLL file to check ODBC not found
#Define ERROR_ODBCFAIL    -113  && failed to get ODBC environment

* Data types for keys
#Define REG_SZ         1  && Data string
#Define REG_BINARY       3  && Binary data in any form.
#Define REG_DWORD       4  && A 32-bit number.

* Data types labels
#Define REG_BINARY_LOC    "*Binary*"      && Binary data in any form.
#Define REG_DWORD_LOC     "*Dword*"      && A 32-bit number.
#Define REG_UNKNOWN_LOC    "*Unknown type*"  && unknown type

* FoxPro ODBC drivers
#Define FOXODBC_25      "FoxPro Files (*.dbf)"
#Define FOXODBC_26      "Microsoft FoxPro Driver (*.dbf)"
#Define FOXODBC_30      "Microsoft Visual FoxPro Driver"

Define Class registry As Custom

    nUserKey = HKEY_CURRENT_USER
    cVFPOptPath = ""
    cRegDLLFile = ""
    cINIDLLFile = ""
    cODBCDLLFile = ""
    nCurrentOS = 0
    nCurrentKey = 0
    lLoadedDLLs = .F.
    lLoadedINIs = .F.
    lLoadedODBCs = .F.
    cAppPathKey = ""
    lCreateKey = .F.
    lhaderror = .F.

    Procedure Init
    This.cVFPOptPath = VFP_OPTIONS_KEY1 + _vfp.Version + VFP_OPTIONS_KEY2
    Do Case
    Case _Dos Or _Unix Or _Mac
        Return .F.
    Case Atc("Windows 3",Os(1)) # 0
        This.nCurrentOS = OS_W32S
        This.cRegDLLFile = DLL_ADVAPI_W32S
        This.cINIDLLFile = DLL_KERNEL_W32S
        This.cODBCDLLFile = DLL_ODBC_W32S
        This.nUserKey = HKEY_CLASSES_ROOT
    Case Atc("Windows NT",Os(1)) # 0
        This.nCurrentOS = OS_NT
        This.cRegDLLFile = DLL_ADVAPI_NT
        This.cINIDLLFile = DLL_KERNEL_NT
        This.cODBCDLLFile = DLL_ODBC_NT
    Otherwise
* Windows 95
        This.nCurrentOS = OS_WIN95
        This.cRegDLLFile = DLL_ADVAPI_WIN95
        This.cINIDLLFile = DLL_KERNEL_WIN95
        This.cODBCDLLFile = DLL_ODBC_WIN95
    Endcase
    Endproc

    Procedure Error
    Lparameters nError, cMethod, nLine
    This.lhaderror = .T.
    =Messagebox(Message())
    Endproc

    Procedure LoadRegFuncs
* Loads funtions needed for Registry
    Local nHKey,cSubKey,nResult
    Local hKey,iValue,lpszValue,lpcchValue,lpdwType,lpbData,lpcbData
    Local lpcStr,lpszVal,nLen,lpdwReserved
    Local lpszValueName,dwReserved,fdwType
    Local iSubKey,lpszName,cchName

    If This.lLoadedDLLs
        Return ERROR_SUCCESS
    Endif

    Declare Integer RegOpenKey In Win32API ;
        Integer nHKey, String @cSubKey, Integer @nResult

    If This.lhaderror && error loading library
        Return -1
    Endif

    Declare Integer RegCreateKey In Win32API ;
        Integer nHKey, String @cSubKey, Integer @nResult

    Declare Integer RegDeleteKey In Win32API ;
        Integer nHKey, String @cSubKey

    Declare Integer RegDeleteValue In Win32API ;
        Integer nHKey, String cSubKey

    Declare Integer RegCloseKey In Win32API ;
        Integer nHKey

    Declare Integer RegSetValueEx In Win32API ;
        Integer hKey, String lpszValueName, Integer dwReserved,;
        Integer fdwType, String lpbData, Integer cbData

    Declare Integer RegQueryValueEx In Win32API ;
        Integer nHKey, String lpszValueName, Integer dwReserved,;
        Integer @lpdwType, String @lpbData, Integer @lpcbData

    Declare Integer RegEnumKey In Win32API ;
        Integer nHKey,Integer iSubKey, String @lpszName, Integer @cchName

    Declare Integer RegEnumKeyEx In Win32API ;
        Integer nHKey,Integer iSubKey, String @lpszName, Integer @cchName,;
        Integer dwReserved,String @lpszName, Integer @cchName,String @cchName

    Declare Integer RegEnumValue In Win32API ;
        Integer hKey, Integer iValue, String @lpszValue, ;
        Integer @lpcchValue, Integer lpdwReserved, Integer @lpdwType, ;
        String @lpbData, Integer @lpcbData

    This.lLoadedDLLs = .T.

* Need error check here
    Return ERROR_SUCCESS
    Endproc

    Procedure OpenKey
* Opens a registry key
    Lparameter cLookUpKey,nRegKey,lCreateKey

    Local nSubKey,nErrCode,nPCount,lSaveCreateKey
    nSubKey = 0
    nPCount = Parameters()

    If Type("m.nRegKey") # "N" Or Empty(m.nRegKey)
        m.nRegKey = HKEY_CLASSES_ROOT
    Endif

* Load API functions
    nErrCode = This.LoadRegFuncs()
    If m.nErrCode # ERROR_SUCCESS
        Return m.nErrCode
    Endif

    lSaveCreateKey = This.lCreateKey
    If m.nPCount>2 And Type("m.lCreateKey") = "L"
        This.lCreateKey = m.lCreateKey
    Endif

    If This.lCreateKey
* Try to open or create registry key
        nErrCode = RegCreateKey(m.nRegKey,m.cLookUpKey,@nSubKey)
    Else
* Try to open registry key
        nErrCode = RegOpenKey(m.nRegKey,m.cLookUpKey,@nSubKey)
    Endif

    This.lCreateKey = m.lSaveCreateKey

    If nErrCode # ERROR_SUCCESS
        Return m.nErrCode
    Endif

    This.nCurrentKey = m.nSubKey
    Return ERROR_SUCCESS
    Endproc

    Procedure CloseKey
* Closes a registry key
    =RegCloseKey(This.nCurrentKey)
    This.nCurrentKey =0
    Endproc

    Procedure SetRegKey
* This routine sets a registry key setting
* ex. THIS.SetRegKey("ResWidth","640",;
*    "Software\Microsoft\VisualFoxPro\4.0\Options",;
*    HKEY_CURRENT_USER)
    Lparameter cOptName,cOptVal,cKeyPath,nUserKey
    Local iPos,cOptKey,cOption,nErrNum
    iPos = 0
    cOption = ""
    nErrNum = ERROR_SUCCESS

* Open registry key
    m.nErrNum = This.OpenKey(m.cKeyPath,m.nUserKey)
    If m.nErrNum # ERROR_SUCCESS
        Return m.nErrNum
    Endif

* Set Key value
    nErrNum = This.SetKeyValue(m.cOptName,m.cOptVal)

* Close registry key
    This.CloseKey()    &&close key
    Return m.nErrNum
    Endproc

    Procedure GetRegKey
* This routine gets a registry key setting
* ex. THIS.GetRegKey("ResWidth",@cValue,;
*    "Software\Microsoft\VisualFoxPro\4.0\Options",;
*    HKEY_CURRENT_USER)
    Lparameter cOptName,cOptVal,cKeyPath,nUserKey
    Local iPos,cOptKey,cOption,nErrNum
    iPos = 0
    cOption = ""
    nErrNum = ERROR_SUCCESS

* Open registry key
    m.nErrNum = This.OpenKey(m.cKeyPath,m.nUserKey)
    If m.nErrNum # ERROR_SUCCESS
        Return m.nErrNum
    Endif

* Get the key value
    nErrNum = This.GetKeyValue(cOptName,@cOptVal)

* Close registry key
    This.CloseKey()    &&close key
    Return m.nErrNum
    Endproc

    Procedure GetKeyValue
* Obtains a value from a registry key
* Note: this routine only handles Data strings (REG_SZ)
    Lparameter cValueName,cKeyValue

    Local lpdwReserved,lpdwType,lpbData,lpcbData,nErrCode
    Store 0 To lpdwReserved,lpdwType
    Store Space(256) To lpbData
    Store Len(m.lpbData) To m.lpcbData

    Do Case
    Case Type("THIS.nCurrentKey")#'N' Or This.nCurrentKey = 0
        Return ERROR_BADKEY
    Case Type("m.cValueName") #"C"
        Return ERROR_BADPARM
    Endcase

    m.nErrCode=RegQueryValueEx(This.nCurrentKey,m.cValueName,;
        m.lpdwReserved,@lpdwType,@lpbData,@lpcbData)

* Check for error
    If m.nErrCode # ERROR_SUCCESS
        Return m.nErrCode
    Endif

* Make sure we have a data string data type
    If lpdwType # REG_SZ
        Return ERROR_NONSTR_DATA
    Endif

    m.cKeyValue = Left(m.lpbData,m.lpcbData-1)
    Return ERROR_SUCCESS
    Endproc

    Procedure SetKeyValue
* This routine sets a key value
* Note: this routine only handles data strings (REG_SZ)
    Lparameter cValueName,cValue
    Local nValueSize,nErrCode

    Do Case
    Case Type("THIS.nCurrentKey")#'N' Or This.nCurrentKey = 0
        Return ERROR_BADKEY
    Case Type("m.cValueName") #"C" Or Type("m.cValue")#"C"
        Return ERROR_BADPARM
    Case Empty(m.cValueName) Or Empty(m.cValue)
        Return ERROR_BADPARM
    Endcase

* Make sure we null terminate this guy
    cValue = m.cValue+Chr(0)
    nValueSize = Len(m.cValue)

* Set the key value here
    m.nErrCode = RegSetValueEx(This.nCurrentKey,m.cValueName,0,;
        REG_SZ,m.cValue,m.nValueSize)

* Check for error
    If m.nErrCode # ERROR_SUCCESS
        Return m.nErrCode
    Endif

    Return ERROR_SUCCESS
    Endproc

    Procedure DeleteKey
* This routine deletes a Registry Key
    Lparameter nUserKey,cKeyPath
    Local nErrNum
    nErrNum = ERROR_SUCCESS

* Delete key
    m.nErrNum = RegDeleteKey(m.nUserKey,m.cKeyPath)
    Return m.nErrNum
    Endproc

    Procedure EnumOptions
* Enumerates through all entries for a key and populates array
    Lparameter aRegOpts,cOptPath,nUserKey,lEnumKeys
    Local iPos,cOptKey,cOption,nErrNum
    iPos = 0
    cOption = ""
    nErrNum = ERROR_SUCCESS

    If Parameters()<4 Or Type("m.lEnumKeys") # "L"
        lEnumKeys = .F.
    Endif

* Open key
    m.nErrNum = This.OpenKey(m.cOptPath,m.nUserKey)
    If m.nErrNum # ERROR_SUCCESS
        Return m.nErrNum
    Endif

* Enumerate through keys
    If m.lEnumKeys
* Enumerate and get key names
        nErrNum = This.EnumKeys(@aRegOpts)
    Else
* Enumerate and get all key values
        nErrNum = This.EnumKeyValues(@aRegOpts)
    Endif

* Close key
    This.CloseKey()    &&close key
    Return m.nErrNum
    Endproc

    Function IsKey
* Checks to see if a key exists
    Lparameter cKeyName,nRegKey

* Open extension key
    nErrNum = This.OpenKey(m.cKeyName,m.nRegKey)
    If m.nErrNum  = ERROR_SUCCESS
* Close extension key
        This.CloseKey()
    Endif

    Return m.nErrNum = ERROR_SUCCESS
    Endfunc

    Procedure EnumKeys
    Parameter aKeyNames
    Local nKeyEntry,cNewKey,cNewSize,cbuf,nbuflen,cRetTime
    nKeyEntry = 0
    Dimension aKeyNames[1]
    Do While .T.
        nKeySize = 0
        cNewKey = Space(100)
        nKeySize = Len(m.cNewKey)
        cbuf=Space(100)
        nbuflen=Len(m.cbuf)
        cRetTime=Space(100)

        m.nErrCode = RegEnumKeyEx(This.nCurrentKey,m.nKeyEntry,@cNewKey,@nKeySize,0,@cbuf,@nbuflen,@cRetTime)

        Do Case
        Case m.nErrCode = ERROR_EOF
            Exit
        Case m.nErrCode # ERROR_SUCCESS
            Exit
        Endcase

        cNewKey = Alltrim(m.cNewKey)
        cNewKey = Left(m.cNewKey,Len(m.cNewKey)-1)
        If !Empty(aKeyNames[1])
            Dimension aKeyNames[ALEN(aKeyNames)+1]
        Endif
        aKeyNames[ALEN(aKeyNames)] = m.cNewKey
        nKeyEntry = m.nKeyEntry + 1
    Enddo

    If m.nErrCode = ERROR_EOF And m.nKeyEntry # 0
        m.nErrCode = ERROR_SUCCESS
    Endif
    Return m.nErrCode
    Endproc

    Procedure EnumKeyValues
* Enumerates through values of a registry key
    Lparameter aKeyValues

    Local lpszValue,lpcchValue,lpdwReserved
    Local lpdwType,lpbData,lpcbData
    Local nErrCode,nKeyEntry,lArrayPassed

    Store 0 To nKeyEntry

    If Type("THIS.nCurrentKey")#'N' Or This.nCurrentKey = 0
        Return ERROR_BADKEY
    Endif

* Sorry, Win32s does not support this one!
    If This.nCurrentOS = OS_W32S
        Return ERROR_BADPLAT
    Endif

    Do While .T.

        Store 0 To lpdwReserved,lpdwType,nErrCode
        Store Space(256) To lpbData, lpszValue
        Store Len(lpbData) To m.lpcchValue
        Store Len(lpszValue) To m.lpcbData

        nErrCode=RegEnumValue(This.nCurrentKey,m.nKeyEntry,@lpszValue,;
            @lpcchValue,m.lpdwReserved,@lpdwType,@lpbData,@lpcbData)

        Do Case
        Case m.nErrCode = ERROR_EOF
            Exit
        Case m.nErrCode # ERROR_SUCCESS
            Exit
        Endcase

        nKeyEntry = m.nKeyEntry + 1

* Set array values
        Dimension aKeyValues[m.nKeyEntry,2]
        aKeyValues[m.nKeyEntry,1] = Left(m.lpszValue,m.lpcchValue)
        Do Case
        Case lpdwType = REG_SZ
            aKeyValues[m.nKeyEntry,2] = Left(m.lpbData,m.lpcbData-1)
        Case lpdwType = REG_BINARY
* Don't support binary
            aKeyValues[m.nKeyEntry,2] = REG_BINARY_LOC
        Case lpdwType = REG_DWORD
* You will need to use ASC() to check values here.
            aKeyValues[m.nKeyEntry,2] = Left(m.lpbData,m.lpcbData-1)
        Otherwise
            aKeyValues[m.nKeyEntry,2] = REG_UNKNOWN_LOC
        Endcase
    Enddo

    If m.nErrCode = ERROR_EOF And m.nKeyEntry # 0
        m.nErrCode = ERROR_SUCCESS
    Endif
    Return m.nErrCode
    Endproc

Enddefine


Define Class oldinireg As registry

    Procedure GetINISection
    Parameters aSections,cSection,cINIFile
    Local cINIValue, nTotEntries, i, nLastPos
    cINIValue = ""
    If Type("m.cINIFile") # "C"
        cINIFile = ""
    Endif

    If This.GetINIEntry(@cINIValue,cSection,0,m.cINIFile) # ERROR_SUCCESS
        Return ERROR_FAILINI
    Endif

    nTotEntries=Occurs(Chr(0),m.cINIValue)
    Dimension aSections[m.nTotEntries]
    nLastPos = 1
    For i = 1 To m.nTotEntries
        nTmpPos = At(Chr(0),m.cINIValue,m.i)
        aSections[m.i] = Substr(m.cINIValue,m.nLastPos,m.nTmpPos-m.nLastPos)
        nLastPos = m.nTmpPos+1
    Endfor

    Return ERROR_SUCCESS
    Endproc

    Procedure GetINIEntry
    Lparameter cValue,cSection,cEntry,cINIFile

* Get entry from INI file
    Local cBuffer,nBufSize,nErrNum,nTotParms
    nTotParms = Parameters()

* Load API functions
    nErrNum= This.LoadINIFuncs()
    If m.nErrNum # ERROR_SUCCESS
        Return m.nErrNum
    Endif

* Parameter checks here
    If m.nTotParms < 3
        m.cEntry = 0
    Endif

    m.cBuffer=Space(2000)

    If Empty(m.cINIFile)
* WIN.INI file
        m.nBufSize = GetWinINI(m.cSection,m.cEntry,"",@cBuffer,Len(m.cBuffer))
    Else
* Private INI file
        m.nBufSize = GetPrivateINI(m.cSection,m.cEntry,"",@cBuffer,Len(m.cBuffer),m.cINIFile)
    Endif

    If m.nBufSize = 0 &&could not find entry in INI file
        Return ERROR_NOINIENTRY
    Endif

    m.cValue=Left(m.cBuffer,m.nBufSize)

** All is well
    Return ERROR_SUCCESS
    Endproc

    Procedure WriteINIEntry
    Lparameter cValue,cSection,cEntry,cINIFile

* Get entry from INI file
    Local nErrNum

* Load API functions
    nErrNum = This.LoadINIFuncs()
    If m.nErrNum # ERROR_SUCCESS
        Return m.nErrNum
    Endif

    If Empty(m.cINIFile)
* WIN.INI file
        nErrNum = WriteWinINI(m.cSection,m.cEntry,m.cValue)
    Else
* Private INI file
        nErrNum = WritePrivateINI(m.cSection,m.cEntry,m.cValue,m.cINIFile)
    Endif

** All is well
    Return Iif(m.nErrNum=1,ERROR_SUCCESS,m.nErrNum)
    Endproc

    Procedure LoadINIFuncs
* Loads funtions needed for reading INI files
    If This.lLoadedINIs
        Return ERROR_SUCCESS
    Endif

    Declare Integer GetPrivateProfileString In Win32API ;
        AS GetPrivateINI String,String,String,String,Integer,String

    If This.lhaderror && error loading library
        Return -1
    Endif

    Declare Integer GetProfileString In Win32API ;
        AS GetWinINI String,String,String,String,Integer

    Declare Integer WriteProfileString In Win32API ;
        AS WriteWinINI String,String,String

    Declare Integer WritePrivateProfileString In Win32API ;
        AS WritePrivateINI String,String,String,String

    This.lLoadedINIs = .T.

* Need error check here
    Return ERROR_SUCCESS
    Endproc

Enddefine

Define Class foxreg As registry

    Procedure SetFoxOption
    Lparameter cOptName,cOptVal
    Return This.SetRegKey(cOptName,cOptVal,This.cVFPOptPath,This.nUserKey)
    Endproc

    Procedure GetFoxOption
    Lparameter cOptName,cOptVal
    Return This.GetRegKey(cOptName,@cOptVal,This.cVFPOptPath,This.nUserKey)
    Endproc

    Procedure EnumFoxOptions
    Lparameter aFoxOpts
    Return This.EnumOptions(@aFoxOpts,This.cVFPOptPath,This.nUserKey,.F.)
    Endproc

Enddefine

Define Class odbcreg As registry

    Procedure LoadODBCFuncs
    If This.lLoadedODBCs
        Return ERROR_SUCCESS
    Endif

* Check API file containing functions

    If Empty(This.cODBCDLLFile)
        Return ERROR_NOODBCFILE
    Endif

    Local henv,fDirection,szDriverDesc,cbDriverDescMax
    Local pcbDriverDesc,szDriverAttributes,cbDrvrAttrMax,pcbDrvrAttr
    Local szDSN,cbDSNMax,pcbDSN,szDescription,cbDescriptionMax,pcbDescription

    Declare Short SQLDrivers In (This.cODBCDLLFile) ;
        Integer henv, Integer fDirection, ;
        String @ szDriverDesc, Integer cbDriverDescMax, Integer pcbDriverDesc, ;
        String @ szDriverAttributes, Integer cbDrvrAttrMax, Integer pcbDrvrAttr

    If This.lhaderror && error loading library
        Return -1
    Endif

    Declare Short SQLDataSources In (This.cODBCDLLFile) ;
        Integer henv, Integer fDirection, ;
        String @ szDSN, Integer cbDSNMax, Integer @ pcbDSN, ;
        String @ szDescription, Integer cbDescriptionMax,Integer pcbDescription

    This.lLoadedODBCs = .T.

    Return ERROR_SUCCESS
    Endproc

    Procedure GetODBCDrvrs
    Parameter aDrvrs,lDataSources
    Local nODBCEnv,nRetVal,dsn,dsndesc,mdsn,mdesc

    lDataSources = Iif(Type("m.lDataSources")="L",m.lDataSources,.F.)

* Load API functions
    nRetVal = This.LoadODBCFuncs()
    If m.nRetVal # ERROR_SUCCESS
        Return m.nRetVal
    Endif

* Get ODBC environment handle
    nODBCEnv=Val(Sys(3053))

* -- Possible error messages
* 527 "cannot load odbc library"
* 528 "odbc entry point missing"
* 182 "not enough memory"

    If Inlist(nODBCEnv,527,528,182)
* Failed
        Return ERROR_ODBCFAIL
    Endif

    Dimension aDrvrs[1,IIF(m.lDataSources,2,1)]
    aDrvrs[1] = ""

    Do While .T.
        dsn=Space(100)
        dsndesc=Space(100)
        mdsn=0
        mdesc=0

* Return drivers or data sources
        If m.lDataSources
            nRetVal = SQLDataSources(m.nODBCEnv,SQL_FETCH_NEXT,@dsn,100,@mdsn,@dsndesc,255,@mdesc)
        Else
            nRetVal = SQLDrivers(m.nODBCEnv,SQL_FETCH_NEXT,@dsn,100,@mdsn,@dsndesc,100,@mdesc)
        Endif

        Do Case
        Case m.nRetVal = SQL_NO_DATA
            nRetVal = ERROR_SUCCESS
            Exit
        Case m.nRetVal # ERROR_SUCCESS And m.nRetVal # 1
            Exit
        Otherwise
            If !Empty(aDrvrs[1])
                If m.lDataSources
                    Dimension aDrvrs[ALEN(aDrvrs,1)+1,2]
                Else
                    Dimension aDrvrs[ALEN(aDrvrs,1)+1,1]
                Endif
            Endif
            dsn = Alltrim(m.dsn)
            aDrvrs[ALEN(aDrvrs,1),1] = Left(m.dsn,Len(m.dsn)-1)

            aDrvrs[ALEN(aDrvrs,1),1] = Chrtran( aDrvrs[ALEN(aDrvrs,1),1] , Chr(10)+Chr(13),'')

            If m.lDataSources
                dsndesc = Alltrim(m.dsndesc)
                aDrvrs[ALEN(aDrvrs,1),2] = Left(m.dsndesc,Len(m.dsndesc)-1)
            Endif
        Endcase
    Enddo
    Return nRetVal
    Endproc

    Procedure EnumODBCDrvrs
    Lparameter aDrvrOpts,cODBCDriver
    Local cSourceKey
    cSourceKey = ODBC_DRVRS_KEY+m.cODBCDriver
    Return This.EnumOptions(@aDrvrOpts,m.cSourceKey,HKEY_LOCAL_MACHINE,.F.)
    Endproc

    Procedure EnumODBCData
    Lparameter aDrvrOpts,cDataSource
    Local cSourceKey
    cSourceKey = ODBC_DATA_KEY+cDataSource
    Return This.EnumOptions(@aDrvrOpts,m.cSourceKey,HKEY_CURRENT_USER,.F.)
    Endproc

Enddefine

Define Class filereg As registry

    Procedure GetAppPath
* Checks and returns path of application
* associated with a particular extension (e.g., XLS, DOC).
    Lparameter cExtension,cExtnKey,cAppKey,lServer
    Local nErrNum,cOptName
    cOptName = ""

* Check Extension parameter
    If Type("m.cExtension") # "C" Or Len(m.cExtension) > 3
        Return ERROR_BADPARM
    Endif
    m.cExtension = "."+m.cExtension

* Open extension key
    nErrNum = This.OpenKey(m.cExtension)
    If m.nErrNum  # ERROR_SUCCESS
        Return m.nErrNum
    Endif

* Get key value for file extension
    nErrNum = This.GetKeyValue(cOptName,@cExtnKey)

* Close extension key
    This.CloseKey()

    If m.nErrNum  # ERROR_SUCCESS
        Return m.nErrNum
    Endif
    Return This.GetApplication(cExtnKey,@cAppKey,lServer)
    Endproc

    Procedure GetLatestVersion
* Checks and returns path of application
* associated with a particular extension (e.g., XLS, DOC).
    Lparameter cClass,cExtnKey,cAppKey,lServer

    Local nErrNum,cOptName
    cOptName = ""

* Open class key (e.g., Excel.Sheet)
    nErrNum = This.OpenKey(m.cClass+CURVER_KEY)
    If m.nErrNum  # ERROR_SUCCESS
        Return m.nErrNum
    Endif

* Get key value for file extension
    nErrNum = This.GetKeyValue(cOptName,@cExtnKey)

* Close extension key
    This.CloseKey()

    If m.nErrNum  # ERROR_SUCCESS
        Return m.nErrNum
    Endif
    Return This.GetApplication(cExtnKey,@cAppKey,lServer)
    Endproc

    Procedure GetApplication
    Parameter cExtnKey,cAppKey,lServer

    Local nErrNum,cOptName
    cOptName = ""

* lServer - checking for OLE server.
    If Type("m.lServer") = "L" And m.lServer
        This.cAppPathKey = OLE_PATH_KEY
    Else
        This.cAppPathKey = APP_PATH_KEY
    Endif

* Open extension app key
    m.nErrNum = This.OpenKey(m.cExtnKey+This.cAppPathKey)
    If m.nErrNum  # ERROR_SUCCESS
        Return m.nErrNum
    Endif

* Get application path
    nErrNum = This.GetKeyValue(cOptName,@cAppKey)

* Close application path key
    This.CloseKey()

    Return m.nErrNum
    Endproc

Enddefine
