<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim lgStrPrevKey
	Dim ChgSave1,ChgSave2
	Const C_SHEETMAXROWS_D = 100	
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)

    lgCurrentSpd      = Request("lgCurrentSpd")                                      '☜: "M"(Spread #1) "S"(Spread #2)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)         '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    ChgSave1 = Request("ChgSave1")
    ChgSave2 = Request("ChgSave2")
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    Call SubBizQueryMulti()
End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    Dim iDx
    Dim iKey1

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    If lgCurrentSpd = "M" Then
        iKey1 = " ocpt_type = " & FilterVar(UCase(lgKeyStream(0)), "''", "S")
        if Trim(UCase(lgKeyStream(1))) <> "" then
            iKey1 = iKey1 & " AND bas_amt_type >= " & FilterVar(UCase(lgKeyStream(1)), "''", "S")
        end if
    else
        iKey1 = " bas_amt_type = " & FilterVar(UCase(lgKeyStream(0)), "''", "S")
        if Trim(UCase(lgKeyStream(1))) <> "" then
            iKey1 = iKey1 & " AND ocpt_type = " & FilterVar(UCase(lgKeyStream(1)), "''", "S")
        end if    
    end if
    
    Call SubMakeSQLStatements("MR",iKey1,"X",C_EQGT)                                 '☆: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        If lgCurrentSpd = "M" Then
           Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        End If   
        Call SetErrorStatus()
    Else
        
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        

        lgstrData = ""
        
        iDx = 1

        Do While Not lgObjRs.EOF
            if lgCurrentSpd = "M" then
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("bas_amt_type") )                   
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("allow_nm"))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dilig_cd"))
                    lgstrData = lgstrData & Chr(11) & ""                 
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("dilig_nm") )                   
                    lgstrData = lgstrData & Chr(11) & "(1)"
                    lgstrData = lgstrData & Chr(11) & "("
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("compute4_n"),    2,0)
                    lgstrData = lgstrData & Chr(11) & "+"
                    lgstrData = lgstrData & Chr(11) & "수당합계"
                    lgstrData = lgstrData & Chr(11) & ")"
                    lgstrData = lgstrData & Chr(11) & "*"
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("compute1_n"),    2,0)
                    lgstrData = lgstrData & Chr(11) & "/"
		            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("compute2_n"),    2,0)
                    lgstrData = lgstrData & Chr(11) & "*"
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("compute3_n"),    2,0)
                    lgstrData = lgstrData & Chr(11) & "(2)"
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("std_amt"),ggAmtOfMoney.DecPoint,0)
               Else
                    lgstrData = lgstrData & chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("allow_cd"))
               End if      
        	lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

        	iDx =  iDx + 1
		    lgObjRs.MoveNext
               	    
            If iDx > C_SHEETMAXROWS_D Then
            	lgStrPrevKey = lgStrPrevKey + 1				
            Exit Do
            End If   
               
        Loop 
 
    End If
    
    If iDx <= C_SHEETMAXROWS_D Then
	  	lgStrPrevKey = ""
    End If   

    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing

    Err.Clear                                                                        '☜: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

End Sub    

'============================================================================================================
' Name : SubBizSaveMultiCreate
' Desc : Save Multi Data
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgCurrentSpd = Trim(UCase(arrColVal(2)))
    
    if ChgSave1 = "T" then
        lgStrSQL = "INSERT INTO HDA160T( bas_amt_type, ocpt_type, dilig_cd," 
        lgStrSQL = lgStrSQL & " compute4_n, compute1_n, compute2_n, compute3_n, std_amt,"
        lgStrSQL = lgStrSQL & " ISRT_EMP_NO , ISRT_DT  , UPDT_EMP_NO , UPDT_DT      )" 
        lgStrSQL = lgStrSQL & " VALUES(" 
        lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S") & ","
        lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S") & ","
        lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S") & ","
        lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(6),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(7),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(8),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(9),0) & ","
		lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(10),0) & ","
        lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                     & "," 
        lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S") & "," 
        lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
        lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")
        lgStrSQL = lgStrSQL & ")"
    elseif ChgSave2 = "T" then
        lgStrSQL = "INSERT INTO HDA020T( bas_amt_type, ocpt_type, allow_cd," 
        lgStrSQL = lgStrSQL & " ISRT_EMP_NO , ISRT_DT , UPDT_EMP_NO , UPDT_DT      )" 
        lgStrSQL = lgStrSQL & " VALUES(" 
        lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S") & ","
        lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S") & ","
        lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S") & ","
        lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                     & "," 
        lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S") & "," 
        lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
        lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")
        lgStrSQL = lgStrSQL & ")"
    end if

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgCurrentSpd = Trim(UCase(arrColVal(2)))

    if lgCurrentSpd = "M" then
        lgStrSQL = "UPDATE  HDA160T"
        lgStrSQL = lgStrSQL & " SET " 
        lgStrSQL = lgStrSQL & " dilig_cd = " & FilterVar(UCase(arrColVal(5)), "''", "S") & ","
        lgStrSQL = lgStrSQL & " compute4_n = " & UNIConvNum(arrColVal(6),0) & ","
        lgStrSQL = lgStrSQL & " compute1_n = " & UNIConvNum(arrColVal(7),0) & ","
        lgStrSQL = lgStrSQL & " compute2_n = " & UNIConvNum(arrColVal(8),0) & ","
        lgStrSQL = lgStrSQL & " compute3_n = " & UNIConvNum(arrColVal(9),0) & ","
        lgStrSQL = lgStrSQL & " std_amt = " & UNIConvNum(arrColVal(10),0) & ","
        lgStrSQL = lgStrSQL & " updt_emp_no = " & FilterVar(gUsrId, "''", "S") & ","
        lgStrSQL = lgStrSQL & " updt_dt = " & FilterVar(GetSvrDateTime,NULL,"S")
        lgStrSQL = lgStrSQL & " WHERE   "
        lgStrSQL = lgStrSQL & " bas_amt_type = " & FilterVar(UCase(arrColVal(4)), "''", "S")
        lgStrSQL = lgStrSQL & " AND ocpt_type = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    else
    end if    

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
   
End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgCurrentSpd = Trim(UCase(arrColVal(2)))

    if  ChgSave1 = "T" then
        lgStrSQL = "DELETE  HDA160T"
        lgStrSQL = lgStrSQL & " WHERE        "
        lgStrSQL = lgStrSQL & "     bas_amt_type = " & FilterVar(UCase(arrColVal(4)), "''", "S")
        lgStrSQL = lgStrSQL & " AND ocpt_type = " & FilterVar(UCase(arrColVal(3)), "''", "S")

        lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	    Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

        lgStrSQL = "DELETE  HDA020T"
        lgStrSQL = lgStrSQL & " WHERE        "
        lgStrSQL = lgStrSQL & "     bas_amt_type = " & FilterVar(UCase(arrColVal(4)), "''", "S")
        lgStrSQL = lgStrSQL & " AND ocpt_type = " & FilterVar(UCase(arrColVal(3)), "''", "S")

        lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	    Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
       
    elseif  ChgSave2 = "T" then
        lgStrSQL = "DELETE  HDA020T"
        lgStrSQL = lgStrSQL & " WHERE        "
        lgStrSQL = lgStrSQL & " bas_amt_type = " & FilterVar(UCase(arrColVal(4)), "''", "S")
        lgStrSQL = lgStrSQL & " AND ocpt_type = " & FilterVar(UCase(arrColVal(3)), "''", "S")
        lgStrSQL = lgStrSQL & " AND allow_cd = " & FilterVar(UCase(arrColVal(5)), "''", "S")

		lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
		Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

    end if    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Select Case Mid(pDataType,1,1)
        Case "M"
        
			iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
          
           Select Case Mid(pDataType,2,1)
               Case "R"
                       If lgCurrentSpd = "M" Then
                          lgStrSQL = "Select TOP " & iSelCount  & " bas_amt_type, "
                          lgStrSQL = lgStrSQL & "                   dbo.ufn_H_GetCodeName(" & FilterVar("HDA010T", "''", "S") & ",BAS_AMT_TYPE,'') ALLOW_nm, "
                          lgStrSQL = lgStrSQL & "                   dilig_cd, "
                          lgStrSQL = lgStrSQL & "                   dbo.ufn_H_GetCodeName(" & FilterVar("HCA010T", "''", "S") & ",DILIG_CD,'') DILIG_nm, "                
                          lgStrSQL = lgStrSQL & "                   compute4_n, "
                          lgStrSQL = lgStrSQL & "                   compute1_n, "
                          lgStrSQL = lgStrSQL & "                   compute2_n, " 
                          lgStrSQL = lgStrSQL & "                   compute3_n, "
                          lgStrSQL = lgStrSQL & "                   std_amt "                          
                          lgStrSQL = lgStrSQL & " From  HDA160T "                          
                          lgStrSQL = lgStrSQL & " WHERE " & pCode
                       Else
                          lgStrSQL = "Select TOP " & iSelCount  
                          lgStrSQL = lgStrSQL & "				allow_cd, "                                                    
                          lgStrSQL = lgStrSQL & "				bas_amt_type "
                          lgStrSQL = lgStrSQL & " From  HDA020T "
                          lgStrSQL = lgStrSQL & " WHERE " & pCode
                          
                       End If             
           End Select             
    End Select

End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select
End Sub

%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                If Trim("<%=lgCurrentSpd%>") = "M" Then
                   .ggoSpread.Source     = .frm1.vspdData
                   .lgStrPrevKey    = "<%=lgStrPrevKey%>"
				   .ggoSpread.SSShowData "<%=lgstrData%>"
				   .DBQueryOk        
	            Else
                   .ggoSpread.Source     = .frm1.vspdData1
                   .lgStrPrevKey1    = "<%=lgStrPrevKey%>"
                   .ggoSpread.SSShowData "<%=lgstrData%>"
                   .DBQueryOk2
                End If  
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
                If parent.Frm1.ChgSave1.value="T"  Then
					If parent.Frm1.ChgSave2.value="T"  Then                
						parent.Frm1.ChgSave1.value = "F"
						Parent.DBSave
					else
						Parent.DBSaveOk
					end if
	            Else
					Parent.DBSaveOk
				End If
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
    End Select    
       
</Script>	

