<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Response.Buffer = True%>
<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	dim lgGetSvrDateTime
	
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

	lgGetSvrDateTime = GetSvrDateTime

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case "REFLECT"																 '☜: REFLECT
             Call SubReflect()             
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection
 
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iKey1, iKey2
	Dim amtSum1,amtSum2,amtSum3,amtSum4,amtSum5
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")
    iKey2 = FilterVar(lgKeyStream(1), "''", "S")

    Call SubMakeSQLStatements("MR",iKey1,C_EQ,iKey2,C_EQ)                                 '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
 %>
<Script Language=vbscript>       
	Parent.frm1.txtSum1.value  = "<%=UNINumClientFormat(0, ggAmtOfMoney.DecPoint,0)%>"      
	Parent.frm1.txtSum2.value  = "<%=UNINumClientFormat(0, ggAmtOfMoney.DecPoint,0)%>"      
</Script>       
<%        
    Else
    
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx = 1
        amtSum1 = 0
        amtSum2 = 0 
        amtSum3 = 0 
        amtSum4 = 0 
        amtSum5 = 0 
        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_NAME"))
            lgstrData = lgstrData & Chr(11) & ""            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_REL"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_REL_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_TYPE"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("INSUR_AMT"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUBMIT_FLAG"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUBMIT_FLAGNM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("YEAR_FLAG")) 
                       
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
		
			Select Case lgObjRs("FAMILY_TYPE")
			    Case "3","5","6","7"
			            amtSum1 = amtSum1 + cdbl(lgObjRs("INSUR_AMT")) 			
			    Case else 
			            amtSum2 = amtSum2 + cdbl(lgObjRs("INSUR_AMT")) 
			
			End Select            
			
		    lgObjRs.MoveNext

            iDx =  iDx + 1
'            If iDx > C_SHEETMAXROWS_D Then
'               lgStrPrevKey = lgStrPrevKey + 1
'               Exit Do
'            End If   
               
        Loop 
 %>
<Script Language=vbscript>       
	Parent.frm1.txtSum1.value  = "<%=UNINumClientFormat(amtSum1, ggAmtOfMoney.DecPoint,0)%>"      
	Parent.frm1.txtSum2.value  = "<%=UNINumClientFormat(amtSum2, ggAmtOfMoney.DecPoint,0)%>"      
</Script>       
<%
    End If
    
'    If iDx <= C_SHEETMAXROWS_D Then
'       lgStrPrevKey = ""
'    End If   

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

	lgStrSQL = "INSERT INTO HFA170T("
	lgStrSQL = lgStrSQL & " YEAR_YY           ," 
	lgStrSQL = lgStrSQL & " EMP_NO       ," 
	lgStrSQL = lgStrSQL & " FAMILY_NAME  ," 
	lgStrSQL = lgStrSQL & " FAMILY_REL   ," 
	lgStrSQL = lgStrSQL & " FAMILY_TYPE  ," 
	lgStrSQL = lgStrSQL & " INSUR_AMT      ," 
	lgStrSQL = lgStrSQL & " SUBMIT_FLAG      ," 
	
	lgStrSQL = lgStrSQL & " ISRT_EMP_NO  ," 
	lgStrSQL = lgStrSQL & " ISRT_DT      ," 
	lgStrSQL = lgStrSQL & " UPDT_EMP_NO  ," 
	lgStrSQL = lgStrSQL & " UPDT_DT      )" 
	lgStrSQL = lgStrSQL & " VALUES(" 
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S")   & ","
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(1), "''", "S")   & ","
	lgStrSQL = lgStrSQL & FilterVar(arrColVal(2), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(arrColVal(3), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(arrColVal(4), "''", "S")     & ","
	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(5),0)           & ","
	lgStrSQL = lgStrSQL & FilterVar(arrColVal(6), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
	lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
	lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S")
	lgStrSQL = lgStrSQL & ")"
	
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL =            "UPDATE HFA170T"
	lgStrSQL = lgStrSQL & "   SET FAMILY_TYPE = " & FilterVar(UCase(arrColVal(4)), "''", "S")   & ","
	lgStrSQL = lgStrSQL & "       INSUR_AMT     = " & UNIConvNum(arrColVal(5),0)                   & ","
	lgStrSQL = lgStrSQL & "       SUBMIT_FLAG = " & FilterVar(UCase(arrColVal(6)), "''", "S")   & ","
	lgStrSQL = lgStrSQL & "	      YEAR_FLAG		= 'N',"	
	lgStrSQL = lgStrSQL & "       UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")                      & ","
	lgStrSQL = lgStrSQL & "       UPDT_DT     = " & FilterVar(lgGetSvrDateTime,NULL,"S")
	lgStrSQL = lgStrSQL & " WHERE YEAR_YY          = " & FilterVar(lgKeyStream(0), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND EMP_NO      = " & FilterVar(lgKeyStream(1), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND FAMILY_NAME = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND FAMILY_REL  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
	lgStrSQL = lgStrSQL & "   AND FAMILY_Type = " & FilterVar(UCase(arrColVal(4)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
   
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	lgStrSQL =            "DELETE HFA170T"
	lgStrSQL = lgStrSQL & " WHERE YEAR_YY          = " & FilterVar(lgKeyStream(0), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND EMP_NO      = " & FilterVar(lgKeyStream(1), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND FAMILY_NAME = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND FAMILY_REL  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND FAMILY_TYPE = " & FilterVar(UCase(arrColVal(4)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubReflect()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear    
                                                                        '☜: Clear Error status
'장애자전용보험 
    lgStrSQL =            " UPDATE HFA030T"
	lgStrSQL = lgStrSQL & "    SET DISABLED_SUB_AMT = INSURE.INSURE_SUM,   " ', Disabled_edu_cnt = INSURE.INSURE_CNT , "
	lgStrSQL = lgStrSQL & "        UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")   & ","
	lgStrSQL = lgStrSQL & "        UPDT_DT     = " & FilterVar(lgGetSvrDateTime,NULL,"S")	
	lgStrSQL = lgStrSQL & "   FROM  ( SELECT ISNULL(SUM(INSUR_AMT),0) INSURE_SUM,  count(*) INSURE_CNT"
	lgStrSQL = lgStrSQL & "				FROM HFA170T "
	lgStrSQL = lgStrSQL & " 			WHERE YEAR_yy = " & FilterVar(lgKeyStream(0), "''", "S") 
	lgStrSQL = lgStrSQL & " 				AND emp_no = " & FilterVar(lgKeyStream(1), "''", "S") 
	lgStrSQL = lgStrSQL & " 				AND family_type in ( '3','5','6','7')) INSURE "		
	lgStrSQL = lgStrSQL & " WHERE YY          = " & FilterVar(lgKeyStream(0), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND EMP_NO      = " & FilterVar(lgKeyStream(1), "''", "S") 


    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

'기타보험료 

    lgStrSQL =            " UPDATE HFA030T"
	lgStrSQL = lgStrSQL & "    SET OTHER_INSUR = INSURE.INSURE_SUM ,  "', UNIV_INSURE_CNT = INSURE.INSURE_CNT , "
	lgStrSQL = lgStrSQL & "        UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")   & ","
	lgStrSQL = lgStrSQL & "        UPDT_DT     = " & FilterVar(lgGetSvrDateTime,NULL,"S")	
	lgStrSQL = lgStrSQL & "   FROM  ( SELECT ISNULL(SUM(INSUR_AMT),0) INSURE_SUM,  count(*) INSURE_CNT"
	lgStrSQL = lgStrSQL & "				FROM HFA170T "
	lgStrSQL = lgStrSQL & " 			WHERE YEAR_yy = " & FilterVar(lgKeyStream(0), "''", "S") 
	lgStrSQL = lgStrSQL & " 				AND emp_no = " & FilterVar(lgKeyStream(1), "''", "S") 
	lgStrSQL = lgStrSQL & " 				AND family_type not in ( '3','5','6','7')) INSURE "		
	lgStrSQL = lgStrSQL & " WHERE YY          = " & FilterVar(lgKeyStream(0), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND EMP_NO      = " & FilterVar(lgKeyStream(1), "''", "S") 
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	

    lgStrSQL =            " UPDATE HFA170T "
	lgStrSQL = lgStrSQL & "    SET YEAR_FLAG = 'Y', "
	lgStrSQL = lgStrSQL & "        UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")   & ","
	lgStrSQL = lgStrSQL & "        UPDT_DT     = " & FilterVar(lgGetSvrDateTime,NULL,"S")	
	lgStrSQL = lgStrSQL & " WHERE YEAR_YY = " & FilterVar(lgKeyStream(0), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(1), "''", "S") 
	'lgStrSQL = lgStrSQL & "   AND family_type IN ('1','2','3','4','5','H','M') "


    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)	
	
End Sub
'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode1,pComp1,pCode2,pComp2)
    Dim iSelCount

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Select Case Mid(pDataType,1,1)
        Case "M"
        
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"
                     lgStrSQL = "Select   FAMILY_NAME, FAMILY_REL,dbo.ufn_GetCodeName('H0140',FAMILY_REL) FAMILY_REL_NM, FAMILY_TYPE, INSUR_AMT,SUBMIT_FLAG,CASE WHEN SUBMIT_FLAG='Y' THEN '국세청자료' ELSE '그밖의자료' END SUBMIT_FLAGNM, YEAR_FLAG"
                     lgStrSQL = lgStrSQL & " From  HFA170T "
                     lgStrSQL = lgStrSQL & " WHERE YEAR_YY     " & pComp1 & pCode1
                     lgStrSQL = lgStrSQL & "   AND EMP_NO " & pComp2 & pCode2
                   
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
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .lgStrPrevKey       = "<%=lgStrPrevKey%>"
                .DBQueryOk        
	         End with
          End If   

       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If   

       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
       Case "REFLECT" 
			If Trim("<%=lgErrorStatus%>") = "NO" Then
		            Parent.ReflectOk
			 Else
			        Parent.ReflectNo
			End If             
    End Select    
       
</Script>	
