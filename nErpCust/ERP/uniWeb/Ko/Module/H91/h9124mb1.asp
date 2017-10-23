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
	Dim amtSum1,amtSum2
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = lgKeyStream(0)
    iKey2 = FilterVar(lgKeyStream(1), "''", "S")

    Call SubMakeSQLStatements("MR",iKey1,C_EQ,iKey2,C_EQ)                                 '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
 %>
<Script Language=vbscript>       
	Parent.frm1.txtSum1.value  = 0    
	Parent.frm1.txtSum2.value  = 0  
</Script>       
<%          
        Call SetErrorStatus()
    Else
    
'        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx = 1
        amtSum1 = 0 
        amtSum2 = 0 
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("MED_DT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MED_NAME"))            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MED_RGST_NO"))

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_NM"))
            lgstrData = lgstrData & Chr(11) & "" 

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_REL"))    
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_REL_NM"))     
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_RES_NO"))     

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_TYPE"))       
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FAMILY_TYPE_NM")) 

            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("MED_AMT"), ggAmtOfMoney.DecPoint,0) 
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PROV_CNT"), ggAmtOfMoney.DecPoint,0)            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MED_TEXT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUBMIT_FLAG"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUBMIT_FLAGNM"))
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("YEAR_FLAG"))           
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

			Select Case lgObjRs("FAMILY_TYPE")
			    Case "A","B"
			            amtSum2 = amtSum2 + cdbl(lgObjRs("MED_AMT")) 			
			    Case Else
						If lgObjRs("FAMILY_REL") = "0" Then
							amtSum2 = amtSum2 + cdbl(lgObjRs("MED_AMT")) 		
						Else
							amtSum1 = amtSum1 + cdbl(lgObjRs("MED_AMT")) 
						End If	
			End Select 
			            
		    lgObjRs.MoveNext

'            iDx =  iDx + 1
 '           If iDx > C_SHEETMAXROWS_D Then
  '             lgStrPrevKey = lgStrPrevKey + 1
   '            Exit Do
    '        End If   
               
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

	lgStrSQL = "INSERT INTO HFA130T("
	lgStrSQL = lgStrSQL & " YEAR_YY           ,"  '0
	lgStrSQL = lgStrSQL & " EMP_NO       ,"  '1
	lgStrSQL = lgStrSQL & " MED_DT  ," '2
	lgStrSQL = lgStrSQL & " MED_AMT   ,"  '9
	lgStrSQL = lgStrSQL & " PROV_CNT   ,"  '10
	lgStrSQL = lgStrSQL & " MED_TEXT  ," '11
	lgStrSQL = lgStrSQL & " SUBMIT_FLAG  ," '11
	lgStrSQL = lgStrSQL & " MED_NAME      ," '4 
	lgStrSQL = lgStrSQL & " MED_RGST_NO      ," '3
	lgStrSQL = lgStrSQL & " FAMILY_NM      ,"	'5
	lgStrSQL = lgStrSQL & " FAMILY_REL      ,"   '6
	lgStrSQL = lgStrSQL & " FAMILY_RES_NO      ," '7
	lgStrSQL = lgStrSQL & " FAMILY_TYPE      ," 
	lgStrSQL = lgStrSQL & " ISRT_EMP_NO  ," 
	lgStrSQL = lgStrSQL & " ISRT_DT      ," 
	lgStrSQL = lgStrSQL & " UPDT_EMP_NO  ," 
	lgStrSQL = lgStrSQL & " UPDT_DT      )" 
	lgStrSQL = lgStrSQL & " VALUES(" 
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S")   & ","
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(1), "''", "S")   & ","
	lgStrSQL = lgStrSQL & FilterVar(UNIConvDate(arrColVal(2)),"''","S")  & ","
	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(9),0)           & ","
	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(10),1)           & ","	
	lgStrSQL = lgStrSQL & FilterVar(arrColVal(11), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(arrColVal(12), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(arrColVal(4), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(arrColVal(3), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(arrColVal(5), "''", "S")     & ","	
	lgStrSQL = lgStrSQL & FilterVar(arrColVal(6), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(arrColVal(7), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(arrColVal(8), "''", "S")     & ","
	
	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
	lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
	lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S")
	lgStrSQL = lgStrSQL & ")"
'Response.Write "lgStrSQL:" & lgStrSQL
'Response.End
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

    lgStrSQL =            "UPDATE HFA130T"
	lgStrSQL = lgStrSQL & "   SET MED_NAME  = " & FilterVar(UCase(arrColVal(4)), "''", "S") & ","
	lgStrSQL = lgStrSQL & "       MED_AMT     = " & UNIConvNum(arrColVal(9),0)                   & ","
	lgStrSQL = lgStrSQL & "       PROV_CNT    = " & UNIConvNum(arrColVal(10),1)                   & ","	
	lgStrSQL = lgStrSQL & "       MED_TEXT     = " &  FilterVar(UCase(arrColVal(11)), "''", "S")   & ","
	'lgStrSQL = lgStrSQL & "       SUBMIT_FLAG     = " &  FilterVar(UCase(arrColVal(12)), "''", "S")   & ","
	lgStrSQL = lgStrSQL & "	      YEAR_FLAG		= 'N',"
	lgStrSQL = lgStrSQL & "       UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")                      & ","
	lgStrSQL = lgStrSQL & "       UPDT_DT     = " & FilterVar(lgGetSvrDateTime,NULL,"S")
	lgStrSQL = lgStrSQL & " WHERE YEAR_YY          = " & FilterVar(lgKeyStream(0), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND EMP_NO      = " & FilterVar(lgKeyStream(1), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND MED_DT = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    'lgStrSQL = lgStrSQL & "   AND MED_NAME  = " & FilterVar(UCase(arrColVal(4)), "''", "S")
	lgStrSQL = lgStrSQL & "   AND FAMILY_NM = " & FilterVar(UCase(arrColVal(5)), "''", "S")
	lgStrSQL = lgStrSQL & "   AND FAMILY_REL = " & FilterVar(UCase(arrColVal(6)), "''", "S")	
	lgStrSQL = lgStrSQL & "   AND MED_RGST_NO = " & FilterVar(UCase(arrColVal(3)), "''", "S")	
	lgStrSQL = lgStrSQL & "   AND  SUBMIT_FLAG     = " &  FilterVar(UCase(arrColVal(12)), "''", "S") 

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

	lgStrSQL =            "DELETE HFA130T"
	lgStrSQL = lgStrSQL & " WHERE YEAR_YY          = " & FilterVar(lgKeyStream(0), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND EMP_NO      = " & FilterVar(lgKeyStream(1), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND MED_DT = " & FilterVar(UCase(arrColVal(2)), "''", "S")
	lgStrSQL = lgStrSQL & "   AND MED_RGST_NO = " & FilterVar(UCase(arrColVal(4)), "''", "S")
	lgStrSQL = lgStrSQL & "   AND FAMILY_NM = " & FilterVar(UCase(arrColVal(5)), "''", "S")
	lgStrSQL = lgStrSQL & "   AND FAMILY_REL = " & FilterVar(UCase(arrColVal(6)), "''", "S")
	lgStrSQL = lgStrSQL & "   AND SUBMIT_FLAG = " & FilterVar(UCase(arrColVal(7)), "''", "S")	

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

'본인/경로자/장애인 의료비 

    lgStrSQL =            " UPDATE HFA030T"
	lgStrSQL = lgStrSQL & "    SET SPECI_MED = MED.MED_SUM , "
	lgStrSQL = lgStrSQL & "        UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")   & ","
	lgStrSQL = lgStrSQL & "        UPDT_DT     = " & FilterVar(lgGetSvrDateTime,NULL,"S")	
	lgStrSQL = lgStrSQL & "   FROM  ( SELECT ISNULL(SUM(med_amt),0) MED_SUM "
	lgStrSQL = lgStrSQL & "				FROM HFA130T "
	lgStrSQL = lgStrSQL & " 			WHERE year_yy = " & FilterVar(lgKeyStream(0), "''", "S") 
	lgStrSQL = lgStrSQL & " 				AND emp_no = " & FilterVar(lgKeyStream(1), "''", "S") 
	lgStrSQL = lgStrSQL & " 				AND MED_DT BETWEEN convert(datetime," & FilterVar(lgKeyStream(0), "''", "S") &" +'0101',112) "
	lgStrSQL = lgStrSQL & " 					AND convert(datetime," & FilterVar(lgKeyStream(0), "''", "S") &" +'1231',112) "	
	lgStrSQL = lgStrSQL & " 				AND ( family_rel='0' OR family_type in ('A','B') )) MED "		
	lgStrSQL = lgStrSQL & " WHERE YY          = " & FilterVar(lgKeyStream(0), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND EMP_NO      = " & FilterVar(lgKeyStream(1), "''", "S") 
'Response.Write "lgStrSQL:" & lgStrSQL
'Response.end

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
		                                                                       '☜: Clear Error status
'일반 의료비 
    lgStrSQL =            " UPDATE HFA030T"
	lgStrSQL = lgStrSQL & "    SET TOT_MED = MED.MED_SUM , "
	lgStrSQL = lgStrSQL & "        UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")   & ","
	lgStrSQL = lgStrSQL & "        UPDT_DT     = " & FilterVar(lgGetSvrDateTime,NULL,"S")	
	lgStrSQL = lgStrSQL & "   FROM  ( SELECT ISNULL(SUM(med_amt),0) MED_SUM "
	lgStrSQL = lgStrSQL & "				FROM HFA130T "
	lgStrSQL = lgStrSQL & " 			WHERE year_yy = " & FilterVar(lgKeyStream(0), "''", "S") 
	lgStrSQL = lgStrSQL & " 				AND emp_no = " & FilterVar(lgKeyStream(1), "''", "S") 
	lgStrSQL = lgStrSQL & " 				AND MED_DT BETWEEN convert(datetime," & FilterVar(lgKeyStream(0), "''", "S") &" +'0101',112) "
	lgStrSQL = lgStrSQL & " 					AND convert(datetime," & FilterVar(lgKeyStream(0), "''", "S") &" +'1231',112) "	
	lgStrSQL = lgStrSQL & " 				AND (family_type NOT in ('A','B') OR family_type IS NULL)AND family_rel<>'0') MED "		
	lgStrSQL = lgStrSQL & " WHERE YY          = " & FilterVar(lgKeyStream(0), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND EMP_NO      = " & FilterVar(lgKeyStream(1), "''", "S") 
'Response.Write "lgStrSQL:" & lgStrSQL
'Response.end

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	
'반영falg update
    lgStrSQL =            " UPDATE HFA130T"
	lgStrSQL = lgStrSQL & "    SET YEAR_FLAG = 'Y', "
	lgStrSQL = lgStrSQL & "        UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")   & ","
	lgStrSQL = lgStrSQL & "        UPDT_DT     = " & FilterVar(lgGetSvrDateTime,NULL,"S")	
	lgStrSQL = lgStrSQL & " WHERE year_yy = " & FilterVar(lgKeyStream(0), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(1), "''", "S") 
	lgStrSQL = lgStrSQL & "   AND MED_DT BETWEEN convert(datetime," & FilterVar(lgKeyStream(0), "''", "S") &" +'0101',112) "
	lgStrSQL = lgStrSQL & "       AND convert(datetime," & FilterVar(lgKeyStream(0), "''", "S") &" +'1231',112) "	
	lgStrSQL = lgStrSQL & "   AND (( family_rel='0' OR family_type in ('A','B') ) OR ((family_type NOT in ('A','B') OR family_type IS NULL)AND family_rel<>'0'))"	    
'Response.Write "lgStrSQL:" & lgStrSQL
'Response.end

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
'                     lgStrSQL = "Select TOP " & iSelCount  & " MED_DT, MED_AMT, MED_TEXT, MED_NAME,MED_RGST_NO,FAMILY_NM ,FAMILY_RES_NO"
                     lgStrSQL = "Select  MED_DT, MED_AMT, MED_TEXT, MED_NAME,MED_RGST_NO,FAMILY_NM ,FAMILY_RES_NO"
                     lgStrSQL = lgStrSQL & ", FAMILY_REL,dbo.ufn_GetCodeName('H0140',FAMILY_REL) FAMILY_REL_NM, FAMILY_RES_NO "
                     lgStrSQL = lgStrSQL & ", FAMILY_TYPE, CASE FAMILY_TYPE WHEN 'A' THEN '장애자' WHEN 'B' THEN '경로자' ELSE '' END FAMILY_TYPE_NM ,PROV_CNT,SUBMIT_FLAG,CASE WHEN SUBMIT_FLAG='Y' THEN '국세청자료' ELSE '그밖의자료' END SUBMIT_FLAGNM,YEAR_FLAG"
                     lgStrSQL = lgStrSQL & " From  HFA130T "
                     lgStrSQL = lgStrSQL & " WHERE YEAR_YY     " & pComp1 & FilterVar(pCode1, "''", "S")
                     lgStrSQL = lgStrSQL & "   AND EMP_NO " & pComp2 & pCode2
                     lgStrSQL = lgStrSQL & "   AND MED_DT >= " & FilterVar(pCode1 & "-01-01", "''", "S")

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
