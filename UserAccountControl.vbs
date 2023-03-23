'Author: Rob Lawton.
'Ver: 1.1.
'Date: 01.Mar.2023.
'Note: VBS version.
'Usage: Supply me an integer response from a UserAccountControl attribute -
'       - and magic will show the MS aligned text response, with much -
'       - displayed working out supporting the determined outcome.
'
' Addition: If you do find this script useful or any sections of code within it, -
'       - feel free to distribute.  However all I ask is that you keep the author -
'       - details present.
'
' Version 	v1.0 - Initial release.
'	v1.1 - Add undocumented-(ish) bit #14. 
'--------------------------------------
option explicit
dim args,intVar,binVar,x,newbinVar,useraccountcontrolVar,useraccountcontrolArray,useraccountControlFlagVar
dim useraccountControlFlagArray,countupVar,arrayposVar,bitVar,resultVar
if wscript.arguments.count =0 then
    wscript.echo "Supply a useraccountcontrol integer!"
else
    Set args = WScript.Arguments
    intVar=args(0)
    if not isnumeric(intVar) then
        wscript.echo "Supply an INTEGER! -  (1,2,3,4...730801....)"
    else
        binVar=fcnDecimalToBinary(intVar)
        
        wscript.echo "Signed integer 4 byte (32-bit) two's complement binary:" & binVar
        wscript.echo "Calculated start at offset (0-31):" & 31-len(binVar)
        for x=0 to     31-len(binVar)
            newbinVar=newbinVar & 0
        next
        wscript.echo "Padding bits:" & newbinVar
        wscript.echo "MSB order:" &  newbinVar & binVar
        binVar = newbinVar & binVar
        binVar=strreverse(binVar)
        wscript.echo "LSB order:" & binVar & vbcrlf
        useraccountcontrolVar="UNUSED_MUST_BE_ZERO-IGNORED,UNUSED_MUST_BE_ZERO-IGNORED,UNUSED_MUST_BE_ZERO-IGNORED,UNUSED_MUST_BE_ZERO-IGNORED,UNUSED_MUST_BE_ZERO-IGNORED,ADS_UF_PARTIAL_SECRETS_ACCOUNT,ADS_UF_NO_AUTH_DATA_REQUIRED,ADS_UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION,ADS_UF_PASSWORD_EXPIRED,ADS_UF_DONT_REQUIRE_PREAUTH,ADS_UF_USE_DES_KEY_ONLY,ADS_UF_NOT_DELEGATED,ADS_UF_TRUSTED_FOR_DELEGATION,ADS_UF_SMARTCARD_REQUIRED,UF_MNS_LOGON_ACCOUNT (apparently not UNUSED),ADS_UF_DONT_EXPIRE_PASSWD,UNUSED_MUST_BE_ZERO-IGNORED,UNUSED_MUST_BE_ZERO-IGNORED,ADS_UF_SERVER_TRUST_ACCOUNT,ADS_UF_WORKSTATION_TRUST_ACCOUNT,ADS_UF_INTERDOMAIN_TRUST_ACCOUNT,UNUSED_MUST_BE_ZERO-IGNORED,ADS_UF_NORMAL_ACCOUNT,UNUSED_MUST_BE_ZERO-IGNORED,ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED,ADS_UF_PASSWD_CANT_CHANGE,ADS_UF_PASSWD_NOTREQD,ADS_UF_LOCKOUT,ADS_UF_HOMEDIR_REQUIRED,UNUSED_MUST_BE_ZERO-IGNORED,ADS_UF_ACCOUNT_DISABLE,UNUSED_MUST_BE_ZERO-IGNORED"
        useraccountcontrolArray=split(useraccountcontrolVar,",")
        useraccountControlFlagVar= "X_,X_,X_,X_,X_,PS,NA,TA,PE,DR,DK,ND,TD,SR,X?,DP,X_,X_,ST,WT,ID,X_,N_,X_,ET,CC,NR,L_,HR,X_,D_,X_"
        useraccountControlFlagArray=split(useraccountControlFlagVar,",")
        arrayposVar=0
        for x = len(binVar) to 1 step -1
            bitVar = "0" & arrayposVar & " - "
            if (arrayposVar) > 9 then bitVar = arrayposVar & " - "
            if mid(binVar,x,1) = 1 then
                resultVar = resultVar & vbcrlf & useraccountcontrolArray(arrayposVar)
                wscript.echo "* bit " & bitVar & "[" & mid(binVar,x,1) & "] - " & useraccountControlFlagArray(arrayposVar) & " - " & useraccountcontrolArray(arrayposVar)
            else
                wscript.echo "  bit " & bitVar & "[" & mid(binVar,x,1) & "] - " & useraccountControlFlagArray(arrayposVar) & " - " & useraccountcontrolArray(arrayposVar)
            end if
            arrayposVar = arrayposVar + 1
        next
        if arrayposVar < 31 then
            for x = arrayposVar to 31
                bitVar = "  bit 0" & x & " - "
                if len(x) > 1 then bitVar = "  bit " & x & " - "  'or if x+1 > 9 then ...
                wscript.echo bitvar & "[0] - " & useraccountControlFlagArray(x) & " - " & useraccountcontrolArray(x)
            next
        end if
    end if
wscript.echo vbcrlf & "Result: " & resultVar
end if
'--------------------------------------
Function fcnDecimalToBinary(intDecimal)
    wscript.echo "Integer:" & intDecimal    
    Dim strBinary, lngNumber1, lngNumber2, strDigit
    strBinary = ""
    intDecimal = cDbl(intDecimal)
        While (intDecimal > 1)
                lngNumber1 = intDecimal / 2
                lngNumber2 = Fix(lngNumber1)
                If (lngNumber1 > lngNumber2) Then
                        strDigit = "1"
                Else
                        strDigit = "0"
                End If
                strBinary = strDigit & strBinary
                intDecimal = Fix(intDecimal / 2)
        Wend
        strBinary = "1" & strBinary
        fcnDecimalToBinary = strBinary
End Function