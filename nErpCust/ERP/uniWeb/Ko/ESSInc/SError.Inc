<%
'======================================================================================================================	
' Name : DisplaySQLErrorMessageConn()
' Desc : Display SQL-DB error message
'======================================================================================================================	
Function CheckSQLError(objConn,pBool)
    Dim objError
    On Error Resume Next														'☜: 

    CheckSQLError = False

    If objConn.Errors.Count > 0 Then
       For Each objError in objConn.Errors
           If ShowODBCErrorCode(objError.Number,objError.Description) = False Then
              If pBool = True Then
                 Call SvrMsgBox(objError.Description &  vbCrLf & "Error Code : " & objError.Number , vbCritical, I_MKSCRIPT)    
              End If   
           End If
           Exit For
       Next
       objConn.Errors.Clear 
       CheckSQLError = True
    End If
End Function

'============================================================================================================
' Name : CheckComproxyError
' Desc : This function check comproxy dll
'============================================================================================================
Function CheckComproxyError(pComDll,pBool)
    On Error Resume Next														'☜: 

    CheckComproxyError = False
    
    If Trim(pComDll.OperationStatusMessage) = "" Then
       Exit Function
    End If
    
	If Not (pComDll.OperationStatusMessage = MSG_OK_STR) Then
		If (pComDll.OperationStatusMessage = "900011" Or pComDll.OperationStatusMessage = "900012") Then
           If pBool = True Then
              Call DisplayMsgBox(pComDll.OperationStatusMessage, vbInformation, "", "", I_MKSCRIPT)	
           End If 
		Else
            Select Case pComDll.OperationStatusMessage
                    Case MSG_DEADLOCK_STR
                         If pBool = True Then
                            Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
                         End If   
                    Case MSG_DBERROR_STR
                         If pBool = True Then
                            Call DisplayMsgBox2(pComDll.ExportErrEabSqlCodeSqlcode, pComDll.ExportErrEabSqlCodeSeverity, pComDll.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
                         End If   
                    Case Else
                         If pBool = True Then
                            Call DisplayMsgBox(pComDll.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
                         End If   
            End Select
		End If
        CheckComproxyError = True
	End If

End Function   


'======================================================================================================================
' Name : CheckSYSTEMError2()
' Desc : Display SYSTEM error message
' ZR00 : When not found
' ZU00 : When not unique
' ZC00 : When already exists
' ZU01 : When permitted value violation
' ZC01 : When permitted value violation
' ZU02 : When database deadlock or timeout
' ZC02 : When database deadlock or timeout
' ZD00 : When database deadlock or timeout
' ZU03 : When database error
' ZC03 : When database error
' ZD01 : When database error
'======================================================================================================================
Function CheckSYSTEMError2(objError, ByVal pBool, ByVal pArg1, ByVal pArg2, ByVal Opt1, ByVal Opt2, ByVal Opt3)
    Dim iDesc0
    Dim iDesc1
    Dim iRet

    CheckSYSTEMError2 = False
    
    If objError.Number = 0 Then
       Exit Function
    End If
    
    CheckSYSTEMError2 = True
    
    If InStr(UCase(objError.Description), "B_CASE") > 0 Then
       If HandleBCaseError(objError.Number, objError.Description, pArg1, pArg2) = True Then
          Exit Function
       End If
    End If
    
    If objError.Number = vbObjectError Then
       If InStr(UCase(objError.Description), "B_MESSAGE") > 0 Then
          If HandleBMessageError(objError.Number, objError.Description, pArg1, pArg2) = True Then
             Exit Function
          End If
       End If
    End If
    
    CheckSYSTEMError2 = CheckSYSTEMError(objError, pBool)
    
End Function
'======================================================================================================================
' Name : DisplaySYSErrorMessage()
' Desc : Display SYSTEM error message
'======================================================================================================================
Function CheckSYSTEMError(objError, pBool)

    Dim iDesc

    CheckSYSTEMError = False
    
    If objError.Number = 0 Then
       Exit Function
    End If
    
    CheckSYSTEMError = True
    
    If objError.Number = vbObjectError Then
      If InStr(UCase(objError.Description), "B_MESSAGE") > 0 Then
         If HandleBMessageError(objError.Number, objError.Description, "", "") = True Then
            Exit Function
         End If
      End If
    End If

    If InStr(UCase(objError.Description), "B_CASE") > 0 Then
       If HandleBCaseError(objError.Number, objError.Description, "", "") = True Then
          Exit Function
       End If
    End If
    
       
    If ShowODBCErrorCode(objError.Number, objError.Description) = False Then
       If pBool = True Then
          Call SvrMsgBox(objError.Description & vbCrLf & "Error Code : " & objError.Number, vbCritical, I_MKSCRIPT)
       End If
    End If
    
    objError.Clear
    
End Function

'======================================================================================================================
' Name : HandleBMessageError
' Desc :
'======================================================================================================================
Function HandleBMessageError(pErrorCode, pErrorDescription, pArg1, pArg2)
    Dim iDesc
    
    HandleBMessageError = False
    
    If pErrorCode <> vbObjectError Then
       Exit Function
    End If
    
    If Not (InStr(pErrorDescription, Chr(11)) > 1) Then
       Exit Function
    End If
    
    iDesc = Split(pErrorDescription, Chr(11))
       
    If InStr(UCase(iDesc(0)), "B_MESSAGE") > 0 Then
       Call DisplayMsgBox(iDesc(1), vbOKOnly, pArg1, pArg2, I_MKSCRIPT)
       HandleBMessageError = True
       Exit Function
    End If

End Function
'======================================================================================================================
' Name : HandleBCaseError
' Desc :
'======================================================================================================================
Function HandleBCaseError(pErrorCode, pErrorDescription, pArg1, pArg2)
    Dim iDesc0, iDesc1, iRet
    Dim Tmp
    
    HandleBCaseError = False
    
    If Not (InStr(pErrorDescription, Chr(11)) > 1) Then
       Exit Function
    End If
    
    iDesc0 = Split(pErrorDescription, Chr(11))
       
    iDesc1 = Split(iDesc0(1), "::")
       
    Select Case pErrorCode
       Case -2147217913, -2147217833 '  Data type cannot be converted.'  A literal value in the command overflowed the range of the type of the associated column.
                        iRet = FindDataCode(iDesc1, "ZU01")
                        If iRet > "" Then
                            Call DisplayMsgBox(iRet, vbOKOnly, pArg1, pArg2, I_MKSCRIPT)
                            HandleBCaseError = True
                            Exit Function
                         End If
                             
                         iRet = FindDataCode(iDesc1, "ZC01")
                               
                         If iRet > "" Then
                            Call DisplayMsgBox(iRet, vbOKOnly, pArg1, pArg2, I_MKSCRIPT)
                            HandleBCaseError = True
                            Exit Function
                         End If
                                
       Case -2147217873  '  A specified value violated the integrity constraints for a column or table
                         iRet = FindDataCode(iDesc1, "ZU00")
                         If iRet > "" Then
                            Call DisplayMsgBox(iRet, vbOKOnly, pArg1, pArg2, I_MKSCRIPT)
                            HandleBCaseError = True
                            Exit Function
                         End If
                             
                         iRet = FindDataCode(iDesc1, "ZC00")
                         If iRet > "" Then
                            Call DisplayMsgBox(iRet, vbOKOnly, pArg1, pArg2, I_MKSCRIPT)
                            HandleBCaseError = True
                            Exit Function
                         End If
                         
                         iRet = FindDataCode(iDesc1, "ZD01")
                         If iRet > "" Then
                            Call DisplayMsgBox("900020", vbOKOnly, pArg1, pArg2, I_MKSCRIPT)
                            HandleBCaseError = True
                            Exit Function
                         End If
                         
       Case -2147217871   ' Timeout expired
                         Tmp = "Error Code : " & pErrorCode & vbCrLf & vbCrLf & "Error Desc : Timeout " & vbCrLf & vbCrLf & "Detail Error Desc : " & pErrorDescription
                         Call SvrMsgBox(Tmp, vbCritical, I_MKSCRIPT)  
    End Select
    
End Function


'======================================================================================================================
' Name : FindDataCode
' Desc : Find Error Code
'======================================================================================================================
Function FindDataCode(pDesc, pCode)
   Dim ii
   Dim iTemp
   
   FindDataCode = ""
   
   For ii = 0 To UBound(pDesc)
   
      If InStr(pDesc(ii), pCode) > 0 Then
         iTemp = Mid(pDesc(ii), 5, 6)
         If Trim(iTemp) > "" Then
            FindDataCode = iTemp
            Exit Function
         End If
      End If
      
   Next

End Function

'======================================================================================================================	
' Name : ShowODBCErrorCode
' Desc : Display ODBC error message
'======================================================================================================================	
Function ShowODBCErrorCode(pErrorCode,pErrorMsg)
  Dim Tmp
  Dim ODBCMsg
  													'☜: 
  ODBCMsg           = ""
  ShowODBCErrorCode = True

  Select Case pErrorCode
  
         Case  -2147217920   :  ODBCMsg =   "Invalid accessor"
         Case  -2147217919   :  ODBCMsg =   "Creating another row would have exceeded the total number of active rows supported by the rowset"
         Case  -2147217918   :  ODBCMsg =   "Unable to write with a read-only accessor"
         Case  -2147217917   :  ODBCMsg =   "Given values violate the database schema"
         Case  -2147217916   :  ODBCMsg =   "Invalid row handle"
         Case  -2147217915   :  ODBCMsg =   "An object was open"
         Case  -2147217914   :  ODBCMsg =   "Invalid chapter"
         Case  -2147217913   :  ODBCMsg =   "A literal value in the command could not be" & vbCrLf &"converted to the correct type due to a reason" & vbCrLf &"other than data overflow"
         Case  -2147217912   :  ODBCMsg =   "Invalid binding info"
         Case  -2147217911   :  ODBCMsg =   "Permission denied"
         Case  -2147217910   :  ODBCMsg =   "Specified column does not contain bookmarks or chapters"
         Case  -2147217909   :  ODBCMsg =   "Some cost limits were rejected"
         Case  -2147217908   :  ODBCMsg =   "No command has been set for the command object"
         Case  -2147217907   :  ODBCMsg =   "Unable to find a query plan within the given cost limit"
         Case  -2147217906   :  ODBCMsg =   "Invalid bookmark"
         Case  -2147217905   :  ODBCMsg =   "Invalid lock mode"
         Case  -2147217904   :  ODBCMsg =   "No value given for one or more required parameters"
         Case  -2147217903   :  ODBCMsg =   "Invalid column ID"
         Case  -2147217902   :  ODBCMsg =   "Invalid ratio"
         Case  -2147217901   :  ODBCMsg =   "Invalid value"
         Case  -2147217900   :  ODBCMsg =   "The command contained one or more errors"
         Case  -2147217899   :  ODBCMsg =   "The executing command cannot be canceled"
         Case  -2147217898   :  ODBCMsg =   "The provider does not support the specified dialect"
         Case  -2147217897   :  ODBCMsg =   "A data source with the specified name already exists"
         Case  -2147217896   :  ODBCMsg =   "The rowset was built over a live data feed and cannot be restarted"
         Case  -2147217895   :  ODBCMsg =   "No key matching the described characteristics could be found within the current range"
         Case  -2147217894   :  ODBCMsg =   "Ownership of this tree has been given to the provider"
         Case  -2147217893   :  ODBCMsg =   "The provider is unable to determine identity for newly inserted rows"
         Case  -2147217892   :  ODBCMsg =   "No nonzero weights specified for any goals" & vbCrLf &"supported, so goal was rejected; current goal" & vbCrLf &"was not changed"
         Case  -2147217891   :  ODBCMsg =   "Requested conversion is not supported"
         Case  -2147217890   :  ODBCMsg =   "lRowsOffset would position you past either end" & vbCrLf &"of the rowset, regardless of the cRows value" & vbCrLf &"specified; cRowsObtained is 0"
         Case  -2147217889   :  ODBCMsg =   "Information was requested for a query, and the query was not set"
         Case  -2147217888   :  ODBCMsg =   "Provider called a method from IRowsetNotify in the consumer and NT"
         Case  -2147217887   :  ODBCMsg =   "Errors occurred"
         Case  -2147217886   :  ODBCMsg =   "A non-NULL controlling IUnknown was specified" & vbCrLf &"and the object being created does not support aggregation"
         Case  -2147217885   :  ODBCMsg =   "A given HROW referred to a hard- or soft- deleted row"
         Case  -2147217884   :  ODBCMsg =   "The rowset does not support fetching backwards"
         Case  -2147217883   :  ODBCMsg =   "All HROWs must be released before new ones can be obtained"
         Case  -2147217882   :  ODBCMsg =   "One of the specified storage flags was not supported"
         Case  -2147217881   :  ODBCMsg =   "The comparision operator was invalid"
         Case  -2147217880   :  ODBCMsg =   "The specified status flag was neither" & vbCrLf &"DBCOLUMNSTATUS_OK nor DBCOLUMNSTATUS_ISNULL"
         Case  -2147217879   :  ODBCMsg =   "The rowset cannot scroll backwards"
         Case  -2147217878   :  ODBCMsg =   "Invalid region handle"
         Case  -2147217877   :  ODBCMsg =   "The specified set of rows was not contiguous to" & vbCrLf &"or overlapping the rows in the specified watch region"
         Case  -2147217876   :  ODBCMsg =   "A transition from ALL* to MOVE* or EXTEND* was specified"
         Case  -2147217875   :  ODBCMsg =   "The specified region is not a proper subregion" & vbCrLf &"of the region identified by the given watch region handle"
         Case  -2147217874   :  ODBCMsg =   "The provider does not support multi-statement commands"
         Case  -2147217873   :  
                               If UCase(gLang) = "KO" Then
                                  ODBCMsg = "동일 데이타가 이미 존재 합니다. 데이터를 확인해 보십시요."
                               Else
                                  ODBCMsg = "A specified value violated the integrity constraints for a column or table"
                               End If

         Case  -2147217872   :  ODBCMsg =   "The given type name was unrecognized"
         Case  -2147217871   :  ODBCMsg =   "Execution aborted because a resource limit has been reached; no results have been returned"
         Case  -2147217870   :  ODBCMsg =   "Cannot clone a command object whose command tree contains a rowset or rowsets"
         Case  -2147217869   :  ODBCMsg =   "Cannot represent the current tree as text"
         Case  -2147217868   :  ODBCMsg =   "The specified index already exists"
         Case  -2147217867   :  ODBCMsg =   "The specified index does not exist"
         Case  -2147217866   :  ODBCMsg =   "The specified index was in use"
         Case  -2147217865   :  ODBCMsg =   "The specified table does not exist"
         Case  -2147217864   :  ODBCMsg =   "The rowset was using optimistic concurrency and" & vbCrLf &"the value of a column has been changed since" & vbCrLf &"it was last read"
         Case  -2147217863   :  ODBCMsg =   "Errors were detected during the copy"
         Case  -2147217862   :  ODBCMsg =   "A specified precision was invalid"
         Case  -2147217861   :  ODBCMsg =   "A specified scale was invalid"
         Case  -2147217860   :  ODBCMsg =   "Invalid table ID"
         Case  -2147217859   :  ODBCMsg =   "A specified type was invalid"
         Case  -2147217858   :  ODBCMsg =   "A column ID was occurred more than once in the specification"
         Case  -2147217857   :  ODBCMsg =   "The specified table already exists"
         Case  -2147217856   :  ODBCMsg =   "The specified table was in use"
         Case  -2147217855   :  ODBCMsg =   "The specified locale ID was not supported"
         Case  -2147217854   :  ODBCMsg =   "The specified record number is invalid"
         Case  -2147217853   :  ODBCMsg =   "Although the bookmark was validly formed, no row could be found to match it"
         Case  -2147217852   :  ODBCMsg =   "The value of a property was invalid"
         Case  -2147217851   :  ODBCMsg =   "The rowset was not chaptered"
         Case  -2147217850   :  ODBCMsg =   "Invalid accessor"
         Case  -2147217849   :  ODBCMsg =   "Invalid storage flags"
         Case  -2147217848   :  ODBCMsg =   "By-ref accessors are not supported by this provider"
         Case  -2147217847   :  ODBCMsg =   "Null accessors are not supported by this provider"
         Case  -2147217846   :  ODBCMsg =   "The command was not prepared"
         Case  -2147217845   :  ODBCMsg =   "The specified accessor was not a parameter accessor"
         Case  -2147217844   :  ODBCMsg =   "The given accessor was write-only"
         Case  -2147217843   :  ODBCMsg =   "Authentication failed"
         Case  -2147217842   :  ODBCMsg =   "The change was canceled during notification; no columns are changed"
         Case  -2147217841   :  ODBCMsg =   "The rowset was single-chaptered and the chapter was not released"
         Case  -2147217840   :  ODBCMsg =   "Invalid source handle"
         Case  -2147217839   :  ODBCMsg =   "The provider cannot derive parameter info and SetParameterInfo has not been called"
         Case  -2147217838   :  ODBCMsg =   "The data source object is already initialized"
         Case  -2147217837   :  ODBCMsg =   "The provider does not support this method"
         Case  -2147217836   :  ODBCMsg =   "The number of rows with pending changes has exceeded the set limit"
         Case  -2147217835   :  ODBCMsg =   "The specified column did not exist"
         Case  -2147217834   :  ODBCMsg =   "There are pending changes on a row with a reference count of zero"
         Case  -2147217833   :
                                If UCase(gLang) = "KO" Then
                                   ODBCMsg = "제한된 숫자 입력 범위를 초과하였습니다."
                                Else
                                   ODBCMsg = "A literal value in the command overflowed the range of the type of the associated column"
                                End If
         
         Case  -2147217832   :  ODBCMsg =   "The supplied HRESULT was invalid"
         Case  -2147217831   :  ODBCMsg =   "The supplied LookupID was invalid"
         Case  -2147217830   :  ODBCMsg =   "The supplied DynamicErrorID was invalid"
         Case  -2147217829   :  ODBCMsg =   "Unable to get visible data for a newly-inserted row that has not yet been updated"
         Case  -2147217828   :  ODBCMsg =   "Invalid conversion flag"
         Case  -2147217827   :  ODBCMsg =   "The given parameter name was unrecognized"
         Case  -2147217826   :  ODBCMsg =   "Multiple storage objects can not be open simultaneously"
         Case  -2147217825   :  ODBCMsg =   "The requested filter could not be opened"
         Case  -2147217824   :  ODBCMsg =   "The requested order could not be opened"
         Case  -2147217823   :  ODBCMsg =   "Bad tuple"
         Case  -2147217822   :  ODBCMsg =   "Bad coordinate"
         Case  -2147217821   :  ODBCMsg =   "The given axis was not valid for this Dataset"
         Case  -2147217820   :  ODBCMsg =   "One or more of the given cell ordinals was invalid"
         Case  -2147217819   :  ODBCMsg =   "The supplied columnID was invalid"
         Case  -2147217817   :  ODBCMsg =   "The supplied command does not have a DBID"
         Case  -2147217816   :  ODBCMsg =   "The supplied DBID already exists"
         Case  -2147217815   :  ODBCMsg =   "The maximum number of Sessions supported by the" & vbCrLf &"provider has already been created. The consumer" & vbCrLf &"must release one or more currently held Sessions" & vbCrLf &"before obtaining a new Session object"
         Case  -2147217806   :  ODBCMsg =   "The index ID is invalid"
         Case  -2147217805   :  ODBCMsg =   "The initialization string does not conform to specification"
         Case  -2147217804   :  ODBCMsg =   "The OLE DB root enumerator did not return any" & vbCrLf &"providers that matched an of the SOURCES_TYPEs requested"
         Case  -2147217803   :  ODBCMsg =   "The initialization string specifies a provider which does not match the currently active provider."
         Case  -2147217802   :  ODBCMsg =   "The specified DBID is invalid"
         Case  -2147217814   :  ODBCMsg =   "Invalid trustee value"
         Case  -2147217813   :  ODBCMsg =   "The trustee is not for the current data source"
         Case  -2147217812   :  ODBCMsg =   "The trustee does not support memberships/ collections"
         Case  -2147217811   :  ODBCMsg =   "The object is invalid or unknown to the provider"
         Case  -2147217810   :  ODBCMsg =   "No owner exists for the object"
         Case  -2147217809   :  ODBCMsg =   "The access entry list supplied is invalid"
         Case  -2147217808   :  ODBCMsg =   "The trustee supplied as owner is invalid or unknown to the provider"
         Case  -2147217807   :  ODBCMsg =   "The permission supplied in the access entry list is invalid"
         Case  -2147217801   :  ODBCMsg =   "The ConstraintType was invalid or not supported by the provider."
         Case  -2147217800   :  ODBCMsg =   "The ConstraintType was not CONSTRAINTTYPE_FOREIGNKEY and cForeignKeyColumns was not zero"
         Case  -2147217799   :  ODBCMsg =   "The Deferrability was invalid or the value was not supported by the provider"
         Case  -2147217792   :  ODBCMsg =   "The MatchType was invalid or the value was not supported by the provider"
         Case  -2147217782   :  ODBCMsg =   "The UpdateRule or DeleteRule was invalid or the value was not supported by the provider"
         Case  -2147217781   :  ODBCMsg =   "The pConstraintID did not exist in the data source"
         Case  -2147217780   :  ODBCMsg =   "The dwFlags was invalid"
         Case  -2147217779   :  ODBCMsg =   "The rguidColumnType pointed to a GUID that does not" & vbCrLf &"match the object type of this column or this column" & vbCrLf &"was not set"
         Case  -2147217778   :  ODBCMsg =   "The requested URL was out-of-scope"
         Case  -2147217776   :  ODBCMsg =   "The provider could not drop the object"
         Case  -2147217775   :  ODBCMsg =   "There is no source row"
         Case  -2147217774   :  ODBCMsg =   "The OLE DB object represented by this URL is locked by one or more other processes"
         Case  -2147217773   :  ODBCMsg =   "The client requested an object type that is only valid for a collection"
         Case  -2147217772   :  ODBCMsg =   "The caller requested write access to a read-only object"
         Case  -2147217771   :  ODBCMsg =   "The provider could not connect to the server for this object"
         Case  -2147217770   :  ODBCMsg =   "The provider could not connect to the server for this object"
         Case  -2147217769   :  ODBCMsg =   "The attempt to bind to the object timed out"
         Case  -2147217768   :  ODBCMsg =   "The provider was unable to create an object at" & vbCrLf &"this URL because an object named by this URL already exists"
         Case  -2147217767   :  ODBCMsg =   "The provider could not drop the object"
         Case  -2147217766   :  ODBCMsg =   "The provider was unable to create an object" & vbCrLf &"at this URL because the server was out of physical storage"
         Case       265920   :  ODBCMsg =   "Fetching requested number of rows would have" & vbCrLf &"exceeded total number of active rows supported by the rowset"
         Case       265921   :  ODBCMsg =   "One or more column types are incompatible; conversion errors will occur during copying"
         Case       265922   :  ODBCMsg =   "Parameter type information has been overridden by caller"
         Case       265923   :  ODBCMsg =   "Skipped bookmark for deleted or non-member row"
         Case       265924   :  ODBCMsg =   "Errors found in validating tree"
         Case       265925   :  ODBCMsg =   "There are no more rowsets"
         Case       265926   :  ODBCMsg =   "Reached start or end of rowset or chapter"
         Case       265927   :  ODBCMsg =   "The provider re-executed the command"
         Case       265928   :  ODBCMsg =   "Variable data buffer full"
         Case       265929   :  ODBCMsg =   "There are no more results"
         Case       265930   :  ODBCMsg =   "Server cannot release or downgrade a lock until the end of the transaction"
         Case       265931   :  ODBCMsg =   "Specified weight was not supported or exceeded the supported limit and was set to 0 or the supported limit"
         Case       265932   :  ODBCMsg =   "Consumer is uninterested in receiving further notification calls for this reason"
         Case       265933   :  ODBCMsg =   "Input dialect was ignored and text was returned in different dialect"
         Case       265934   :  ODBCMsg =   "Consumer is uninterested in receiving further notification calls for this phase"
         Case       265935   :  ODBCMsg =   "Consumer is uninterested in receiving further notification calls for this reason"
         Case       265936   :  ODBCMsg =   "The operation is being processed asynchronously"
         Case       265937   :  ODBCMsg =   "In order to reposition to the start of the" & vbCrLf &"rowset, the provider had to reexecute the" & vbCrLf &"query; either the order of the columns changed" & vbCrLf &"or columns were added to or removed from the rowset"
         Case       265938   :  ODBCMsg =   "The method had some errors; errors have been returned in the error array"
         Case       265939   :  ODBCMsg =   "Invalid row handle"
         Case       265940   :  ODBCMsg =   "A given HROW referred to a hard-deleted row"
         Case       265941   :  ODBCMsg =   "The provider was unable to keep track of all" & vbCrLf &"the changes; the client must refetch the data" & vbCrLf &"associated with the watch region using another method"
         Case       265942   :  ODBCMsg =   "Execution stopped because a resource limit has" & vbCrLf &"been reached; results obtained so far have been" & vbCrLf &"returned but execution cannot be resumed"
         Case       265943   :  ODBCMsg =   "The bind failed because the provider was unable" & vbCrLf &"to satisfy all of the bind flags or properties"
         Case       265944   :  ODBCMsg =   "A lock was upgraded from the value specified"
         Case       265945   :  ODBCMsg =   "One or more properties were changed as allowed by provider"
         Case       265946   :  ODBCMsg =   "Errors occurred"
         Case       265947   :  ODBCMsg =   "A specified parameter was invalid"
         Case       265948   :  ODBCMsg =   "Updating this row caused more than one row to be updated in the data source"
         Case       265949   :  ODBCMsg =   "The row has no row-specific columns"
         Case       Else     : ShowODBCErrorCode = False
  End Select                                

  If Trim(pErrorMsg) <> "" Then
'    Tmp = "Error Code : " & pErrorCode & vbCrLf & vbCrLf & "Error Desc : " & ODBCMsg & vbCrLf & vbCrLf & "Detail Error Desc : " & pErrorMsg  '20050823
     Tmp = "Error Desc : " & ODBCMsg
  Else
     Tmp = "Error Code : " & pErrorCode & vbCrLf & "ODBC Err Desc : " & ODBCMsg
  End If
     
  If ShowODBCErrorCode =  True Then
     Call SvrMsgBox(Tmp, vbCritical, I_MKSCRIPT)  
  End If   
                                            
End Function                                
                                            
                                            
%>
