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
	Dim amtSum1,amtSum2,amtSum3,amtSum4,amtSum5,amtSum6
		
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
	Parent.frm1.txtSum3.value  = "<%=UNINumClientFormat(0, ggAmtOfMoney.DecPoint,0)%>"      
	Parent.frm1.txtSum4.value  = "<%=UNINumClientFormat(0, ggAmtOfMoney.DecPoint,0)%>"      
	Parent.frm1.txtSum5.value  = "<%=UNINumClientFormat(0, ggAmtOfMoney.DecPoint,0)%>"       
	Parent.frm1.txtSum6.value  = "<%=UNINumClientFormat(0, ggAmtOfMoney.DecPoint,0)%>" 	
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
        amtSum6 = 0         
        
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("CONTR_DT"))
         
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CONTR_RGST_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CONTR_NAME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CONTR_CODE")) 
            lgstrData = lgstrData & Chr(11) & ""           
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CONTR_TYPE"))     
			lgstrData = lgstrData & Chr(11) & ""                    
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("CONTR_AMT"), ggAmtOfMoney.DecPoint,0)   
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PROV_CNT"), ggAmtOfMoney.DecPoint,0)   
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUBMIT_FLAG"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUBMIT_FLAGNM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("YEAR_FLAG")) 
                                    
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
            
			Select Case lgObjRs("CONTR_CODE")
			    Case "10"
			            amtSum1 = amtSum1 + cdbl(lgObjRs("CONTR_AMT")) 			
			    Case "20"		
						amtSum2 = amtSum2 + cdbl(lgObjRs("CONTR_AMT")) 
			    Case "21"		
						amtSum3 = amtSum3 + cdbl(lgObjRs("CONTR_AMT")) 						
			    Case "30"
						amtSum4 = amtSum4 + cdbl(lgObjRs("CONTR_AMT")) 
			    Case "42"			    
			            amtSum5 = amtSum5 + cdbl(lgObjRs("CONTR_AMT")) 
			    Case "40","41","50"
			            amtSum6 = amtSum6 + cdbl(lgObjRs("CONTR_AMT")) 			            
			End Select  
			            
		    lgObjRs.MoveNext

            iDx =  iDx + 1
'            If iDx > C_SHEETMAXROWS_D Then
 '              lgStrPrevKey = lgStrPrevKey + 1
  '             Exit Do
   '         End If   
               
        Loop 
 %>
<Script Language=vbscript>       
	Parent.frm1.txtSum1.value  = "<%=UNINumClientFormat(amtSum1, ggAmtOfMoney.DecPoint,0)%>"      
	Parent.frm1.txtSum2.value  = "<%=UNINumClientFormat(amtSum2, ggAmtOfMoney.DecPoint,0)%>"      
	Parent.frm1.txtSum3.value  = "<%=UNINumClientFormat(amtSum3, ggAmtOfMoney.DecPoint,0)%>"      
	Parent.frm1.txtSum4.value  = "<%=UNINumClientFormat(amtSum4, ggAmtOfMoney.DecPoint,0)%>"      
	Parent.frm1.txtSum5.value  = "<%=UNINumClientFormat(amtSum5, ggAmtOfMoney.DecPoint,0)%>"       
	Parent.frm1.txtSum6.value  = "<%=UNINumClientFormat(amtSum6, ggAmtOfMoney.DecPoint,0)%>" 	
</Script>       
<%        
    End If
    
'    If iDx <= C_SHEETMAXROWS_D Then
 '      lgStrPrevKey = ""
  '  End If   

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
	lgStrSQL = "INSERT INTO HFA140T("
	lgStrSQL = lgStrSQL & " YEAR_YY           ," 
	lgStrSQL = lgStrSQL & " EMP_NO       ," 
	lgStrSQL = lgStrSQL & " CONTR_DT  ," 
	lgStrSQL = lgStrSQL & " CONTR_AMT   ," '6
	lgStrSQL = lgStrSQL & " PROV_CNT   ," '7	
	lgStrSQL = lgStrSQL & " CONTR_RGST_NO  ," '3
	lgStrSQL = lgStrSQL & " CONTR_NAME  ,"		'4
	lgStrSQL = lgStrSQL & " CONTR_TYPE      ,"	'5
	lgStrSQL = lgStrSQL & " CONTR_CODE      ,"	'6
	lgStrSQL = lgStrSQL & " SUBMIT_FLAG      ," 
	lgStrSQL = lgStrSQL & " ISRT_EMP_NO  ," 
	lgStrSQL = lgStrSQL & " ISRT_DT      ," 
	lgStrSQL = lgStrSQL & " UPDT_EMP_NO  ," 
	lgStrSQL = lgStrSQL & " UPDT_DT      )" 
	lgStrSQL = lgStrSQL & " VALUES(" 
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S")   & ","
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(1), "''", "S")   & ","
	lgStrSQL = lgStrSQL & FilterVar(UNIConvDate(arrColVal(2)),"''","S") & ","
	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(7),0)			 & ","
	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(8),1)			 & ","
	
	
	lgStrSQL = lgStrSQL & FilterVar(arrColVal(3), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(arrColVal(4), "''", "S")     & ","
	lgStrSQL = lgStrSQL & FilterVar(arrColVal(5), "''", "S")     & ","		
	lgStrSQL = lgStrSQL & FilterVar(arrColVal(6), "''", "S")     & ","	
	lgStrSQL = lgStrSQL & FilterVar(arrColVal(9), "''", "S")     & ","
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

    lgStrSQL =            "UPDATE HFA140T"
	lgStrSQL = lgStrSQL & "   SET CONTR_CODE = " & FilterVar(UCase(arrColVal(5)), "''", "S")	& ","
	lgStrSQL = lgStrSQL & "       CONTR_AMT     = " & UNIConvNum(arrColVal(6),0)					& ","
	lgStrSQL = lgStrSQL & "       PROV_CNT     = " & UNIConvNum(arrColVal(7),1)						& ","	
	lgStrSQL = lgStrSQL & "	      CONTR_NAME	= " & FilterVar(UCase(arrColVal(8)), "''", "S")		& ","		
	lgStrSQL = lgStrSQL & "	      SUBMIT_FLAG	= " & FilterVar(UCase(arrColVal(9)), "''", "S")		& ","		
	lgStrSQL = lgStrSQL & "	      YEAR_FLAG		= 'N',"	

	lgStrSQL = lgStrSQL & "       UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")					& ","
	lgStrSQL = lgStrSQL & "       UPDT_DT     = " & FilterVar(lgGetSvrDateTime,NULL,"S")
	lgStrSQL = lgStrSQL & " WHERE YEAR_YY          = " & FilterVar(lgKeyStream(0), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND EMP_NO		= " & FilterVar(lgKeyStream(1), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND CONTR_DT		= " & FilterVar(UNIConvDate(arrColVal(2)),"''","S")
    lgStrSQL = lgStrSQL & "   AND CONTR_RGST_NO = " & FilterVar(UCase(arrColVal(3)), "''", "S")
	lgStrSQL = lgStrSQL & "   AND CONTR_CODE	= " & FilterVar(UCase(arrColVal(5)), "''", "S")

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

	lgStrSQL =            "DELETE HFA140T"
	lgStrSQL = lgStrSQL & " WHERE YEAR_YY          = " & FilterVar(lgKeyStream(0), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND EMP_NO		= " & FilterVar(lgKeyStream(1), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND CONTR_DT		= " & FilterVar(UNIConvDate(arrColVal(2)),"''","S")
    lgStrSQL = lgStrSQL & "   AND CONTR_RGST_NO = " & FilterVar(UCase(arrColVal(3)), "''", "S")
	lgStrSQL = lgStrSQL & "   AND CONTR_CODE	= " & FilterVar(UCase(arrColVal(5)), "''", "S")

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
'법정기부금 : 100 %
    lgStrSQL =            " UPDATE HFA030T"
	lgStrSQL = lgStrSQL & "    SET LEGAL_CONTR = CNTR.CONTR_SUM , "
	lgStrSQL = lgStrSQL & "        UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")   & ","
	lgStrSQL = lgStrSQL & "        UPDT_DT     = " & FilterVar(lgGetSvrDateTime,NULL,"S")	
	lgStrSQL = lgStrSQL & "   FROM  ( SELECT ISNULL(SUM(contr_amt),0) CONTR_SUM "
	lgStrSQL = lgStrSQL & "				FROM HFA140T "
	lgStrSQL = lgStrSQL & " 			WHERE year_yy = " & FilterVar(lgKeyStream(0), "''", "S") 
	lgStrSQL = lgStrSQL & " 				AND emp_no = " & FilterVar(lgKeyStream(1), "''", "S") 
	lgStrSQL = lgStrSQL & " 				AND CONTR_CODE = '10' ) CNTR "		
	lgStrSQL = lgStrSQL & " WHERE YY          = " & FilterVar(lgKeyStream(0), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND EMP_NO      = " & FilterVar(lgKeyStream(1), "''", "S") 
'Response.Write "lgStrSQL:" & lgStrSQL
'Response.end

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
 
'정치자금기부금: 100 %

    lgStrSQL =            " UPDATE HFA030T"
	lgStrSQL = lgStrSQL & "    SET POLI_CONTRA_AMT1 = CNTR.CONTR_SUM , "
	lgStrSQL = lgStrSQL & "        UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")   & ","
	lgStrSQL = lgStrSQL & "        UPDT_DT     = " & FilterVar(lgGetSvrDateTime,NULL,"S")	
	lgStrSQL = lgStrSQL & "   FROM  ( SELECT ISNULL(SUM(contr_amt),0) CONTR_SUM "
	lgStrSQL = lgStrSQL & "				FROM HFA140T "
	lgStrSQL = lgStrSQL & " 			WHERE year_yy = " & FilterVar(lgKeyStream(0), "''", "S") 
	lgStrSQL = lgStrSQL & " 				AND emp_no = " & FilterVar(lgKeyStream(1), "''", "S") 
	lgStrSQL = lgStrSQL & " 				AND CONTR_CODE = '20') CNTR "		
	lgStrSQL = lgStrSQL & " WHERE YY          = " & FilterVar(lgKeyStream(0), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND EMP_NO      = " & FilterVar(lgKeyStream(1), "''", "S") 
'Response.Write "lgStrSQL:" & lgStrSQL
'Response.end

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)


'진흥기금 : 100 %

    lgStrSQL =            " UPDATE HFA030T"
	lgStrSQL = lgStrSQL & "    SET TAXLAW_CONTR_AMT2 = CNTR.CONTR_SUM , "
	lgStrSQL = lgStrSQL & "        UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")   & ","
	lgStrSQL = lgStrSQL & "        UPDT_DT     = " & FilterVar(lgGetSvrDateTime,NULL,"S")	
	lgStrSQL = lgStrSQL & "   FROM  ( SELECT ISNULL(SUM(contr_amt),0) CONTR_SUM "
	lgStrSQL = lgStrSQL & "				FROM HFA140T "
	lgStrSQL = lgStrSQL & " 			WHERE year_yy = " & FilterVar(lgKeyStream(0), "''", "S") 
	lgStrSQL = lgStrSQL & " 				AND emp_no = " & FilterVar(lgKeyStream(1), "''", "S") 
	lgStrSQL = lgStrSQL & " 				AND CONTR_CODE = '21' ) CNTR "		
	lgStrSQL = lgStrSQL & " WHERE YY          = " & FilterVar(lgKeyStream(0), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND EMP_NO      = " & FilterVar(lgKeyStream(1), "''", "S") 
'Response.Write "lgStrSQL:" & lgStrSQL
'Response.end

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

'특례기부 : 50 %
    lgStrSQL =            " UPDATE HFA030T"
	lgStrSQL = lgStrSQL & "    SET TAXLAW_CONTR_AMT = CNTR.CONTR_SUM , "
	lgStrSQL = lgStrSQL & "        UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")   & ","
	lgStrSQL = lgStrSQL & "        UPDT_DT     = " & FilterVar(lgGetSvrDateTime,NULL,"S")	
	lgStrSQL = lgStrSQL & "   FROM  ( SELECT ISNULL(SUM(contr_amt),0) CONTR_SUM "
	lgStrSQL = lgStrSQL & "				FROM HFA140T "
	lgStrSQL = lgStrSQL & " 			WHERE year_yy = " & FilterVar(lgKeyStream(0), "''", "S") 
	lgStrSQL = lgStrSQL & " 				AND emp_no = " & FilterVar(lgKeyStream(1), "''", "S") 
	lgStrSQL = lgStrSQL & " 				AND CONTR_CODE = '30') CNTR "		
	lgStrSQL = lgStrSQL & " WHERE YY          = " & FilterVar(lgKeyStream(0), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND EMP_NO      = " & FilterVar(lgKeyStream(1), "''", "S") 
'Response.Write "lgStrSQL:" & lgStrSQL
'Response.end

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)		
	
'우리사주기부금 : 30%
    lgStrSQL =            " UPDATE HFA030T"
	lgStrSQL = lgStrSQL & "    SET OURSTOCK_CONTRA_AMT = CNTR.CONTR_SUM , "
	lgStrSQL = lgStrSQL & "        UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")   & ","
	lgStrSQL = lgStrSQL & "        UPDT_DT     = " & FilterVar(lgGetSvrDateTime,NULL,"S")	
	lgStrSQL = lgStrSQL & "   FROM  ( SELECT ISNULL(SUM(contr_amt),0) CONTR_SUM "
	lgStrSQL = lgStrSQL & "				FROM HFA140T "
	lgStrSQL = lgStrSQL & " 			WHERE year_yy = " & FilterVar(lgKeyStream(0), "''", "S") 
	lgStrSQL = lgStrSQL & " 				AND emp_no = " & FilterVar(lgKeyStream(1), "''", "S") 
	lgStrSQL = lgStrSQL & " 				AND CONTR_CODE = '42') CNTR "		
	lgStrSQL = lgStrSQL & " WHERE YY          = " & FilterVar(lgKeyStream(0), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND EMP_NO      = " & FilterVar(lgKeyStream(1), "''", "S") 
'Response.Write "lgStrSQL:" & lgStrSQL
'Response.end

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	
'지정기부금  : 10%
    lgStrSQL =            " UPDATE HFA030T"
	lgStrSQL = lgStrSQL & "    SET APP_CONTR = CNTR.CONTR_SUM , "
	lgStrSQL = lgStrSQL & "        UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")   & ","
	lgStrSQL = lgStrSQL & "        UPDT_DT     = " & FilterVar(lgGetSvrDateTime,NULL,"S")	
	lgStrSQL = lgStrSQL & "   FROM  ( SELECT ISNULL(SUM(contr_amt),0) CONTR_SUM "
	lgStrSQL = lgStrSQL & "				FROM HFA140T "
	lgStrSQL = lgStrSQL & " 			WHERE year_yy = " & FilterVar(lgKeyStream(0), "''", "S") 
	lgStrSQL = lgStrSQL & " 				AND emp_no = " & FilterVar(lgKeyStream(1), "''", "S") 
	lgStrSQL = lgStrSQL & " 				AND CONTR_CODE IN ('40','41','50')) CNTR "		
	lgStrSQL = lgStrSQL & " WHERE YY          = " & FilterVar(lgKeyStream(0), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND EMP_NO      = " & FilterVar(lgKeyStream(1), "''", "S") 
'Response.Write "lgStrSQL:" & lgStrSQL
'Response.end

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)	
	
'반영falg update
    lgStrSQL =            " UPDATE HFA140T"
	lgStrSQL = lgStrSQL & "    SET YEAR_FLAG = 'Y', "
	lgStrSQL = lgStrSQL & "        UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S")   & ","
	lgStrSQL = lgStrSQL & "        UPDT_DT     = " & FilterVar(lgGetSvrDateTime,NULL,"S")	
	lgStrSQL = lgStrSQL & " WHERE year_yy = " & FilterVar(lgKeyStream(0), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(1), "''", "S") 
	lgStrSQL = lgStrSQL & "   AND CONTR_CODE in ('10', '20', '21' ,'30' , '42' , '40','41','50') "
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
                     lgStrSQL = "Select CONTR_DT,CONTR_RGST_NO,CONTR_NAME, CONTR_TYPE ,dbo.ufn_GetCodeName('H0125', CONTR_TYPE) CONTR_TYPE_NM, "
                     lgStrSQL = lgStrSQL & " CONTR_CODE ,dbo.ufn_GetCodeName('H0126', CONTR_CODE) CONTR_TYPE_NM, CONTR_AMT ,PROV_CNT ,SUBMIT_FLAG,CASE WHEN SUBMIT_FLAG='Y' THEN '국세청자료' ELSE '그밖의자료' END SUBMIT_FLAGNM,YEAR_FLAG  "
                     lgStrSQL = lgStrSQL & " From  HFA140T "
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
