
<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%    
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

	Call CheckVersion(lgKeyStream(1), lgKeyStream(2))	' 2005-03-11 버전관리기능 추가 
	
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)
             Call SubBizQuery()
        Case CStr(UID_M0002)
       
             Call SubBizSave()
        Case CStr(UID_M0003)
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1
    dim strWhere

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strWhere = " co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
	strWhere = strWhere & "  and fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
	strWhere = strWhere & "  and rep_type =" &  FilterVar(Trim(lgKeyStream(2)),"","S")
	

    Call SubMakeSQLStatements("R",strWhere)                                       '☜ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
          'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
         
				%>
				<Script Language=vbscript>
				     
				     Parent.FNCnew
	 		    
				  </Script>       
				<%        
              Call SetErrorStatus()
    Else 
				%>
				<Script Language=vbscript>
				       With Parent	
				                .Frm1.txtW1_S.TEXT			  = "<%=ConvSPChars(lgObjRs("w1_S"))%>"
				                .Frm1.txtW1_E.TEXT			  = "<%=ConvSPChars(lgObjRs("w1_E"))%>"
				                .Frm1.txtW2_S.TEXT			  = "<%=ConvSPChars(lgObjRs("w2_S"))%>"
				                .Frm1.txtW2_E.TEXT			  = "<%=ConvSPChars(lgObjRs("w2_E"))%>"
				               
				                .Frm1.txtW6.Value			  = "<%=ConvSPChars(lgObjRs("w6"))%>"

				                .Frm1.txtW7.Value			  = "<%=ConvSPChars(lgObjRs("w7"))%>"
				               
				                .Frm1.txtW8.Value			  = "<%=ConvSPChars(lgObjRs("w8"))%>"
				                .Frm1.txtW9.Value			  = "<%=ConvSPChars(lgObjRs("w9"))%>"
				                .Frm1.txtW9_value.Value			  = "<%=ConvSPChars(lgObjRs("w9_Value"))%>"
				                .Frm1.txtW10.Value			  = "<%=ConvSPChars(lgObjRs("w10"))%>"
				                .Frm1.txtW11.Value			  = "<%=ConvSPChars(lgObjRs("w11"))%>"
				                .Frm1.txtW12.Value			  = "<%=ConvSPChars(lgObjRs("w12"))%>"
				                .Frm1.txtW13.Value			  = "<%=ConvSPChars(lgObjRs("w13"))%>"
				                .Frm1.txtW14.Value			  = "<%=ConvSPChars(lgObjRs("w14"))%>"
				                .Frm1.txtW15.Value			  = "<%=ConvSPChars(lgObjRs("w15"))%>"
				                 
				               
				       End With          
				</Script>       
				<%     
    End If
    
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)
    
End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next
    Err.Clear
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)

    Select Case lgIntFlgMode
        Case  OPMD_CMODE        
                                                '☜ : Create
              Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE           
              Call SubBizSaveSingleUpdate()
    End Select
End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next
    Err.Clear

    lgStrSQL = "DELETE From  TB_68"

    lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S") 
    lgStrSQL = lgStrSQL & "		  and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S") 
    lgStrSQL = lgStrSQL & "		  and rep_type =" &FilterVar(Trim(UCase(lgKeyStream(2))),"","S") 

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    On Error Resume Next
    Err.Clear
    
    lgStrSQL = "INSERT INTO TB_68("
    lgStrSQL = lgStrSQL & " co_cd, "
    lgStrSQL = lgStrSQL & " FISC_YEAR, "
    lgStrSQL = lgStrSQL & " rep_type, "
    lgStrSQL = lgStrSQL & " W1_S, "
    lgStrSQL = lgStrSQL & " W1_E, "
    lgStrSQL = lgStrSQL & " W2_S, "
    lgStrSQL = lgStrSQL & " W2_E, "

    lgStrSQL = lgStrSQL & " W6, "
    lgStrSQL = lgStrSQL & " W7, "
    lgStrSQL = lgStrSQL & " W8, "
    lgStrSQL = lgStrSQL & " W9, "
    lgStrSQL = lgStrSQL & " W9_Value, "
    lgStrSQL = lgStrSQL & " W10, "
    lgStrSQL = lgStrSQL & " W11, "
    lgStrSQL = lgStrSQL & " W12, "
    lgStrSQL = lgStrSQL & " W13, "
    lgStrSQL = lgStrSQL & " W14, "             
    lgStrSQL = lgStrSQL & " W15, "                          
    lgStrSQL = lgStrSQL & " INSRT_USER_ID, "
    lgStrSQL = lgStrSQL & " INSRT_DT, "
    lgStrSQL = lgStrSQL & " UPDT_USER_ID, "
    lgStrSQL = lgStrSQL & " UPDT_DT)"
    lgStrSQL = lgStrSQL & " VALUES("
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")     & ","                      
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")     & ","           
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")     & ","             
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW1_S"),"''","S")   & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW1_E"),"''","S")   & ","      
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW2_S"),"''","S")   & ","      
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW2_E"),"''","S")   & ","            

    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW6"),0)     & ","   

    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW7"),0)     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8"),0)     & ","   
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW9"),"''","S")     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW9_Value"),0)     & ","     
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW10"),0)     & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW11"),0)     & ","     
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW12"),0)     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW13"),0)     & ","        
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW14"),0)     & ","    
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW15"),0)     & ","    
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                       
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","                                     
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                     
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S")                                            
    lgStrSQL = lgStrSQL & ")" 


    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords

	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    On Error Resume Next
    Err.Clear

    lgStrSQL = lgStrSQL & " Update TB_68 set"
    lgStrSQL = lgStrSQL & " W1_S  = " & FilterVar(Trim(Request("txtW1_S")),"","S")   & ","
    lgStrSQL = lgStrSQL & " W1_E  = " & FilterVar(Trim(Request("txtW1_E")),"","S")   & ","  
    lgStrSQL = lgStrSQL & " W2_S  = " & FilterVar(Trim(Request("txtW2_S")),"","S")   & ","  
    lgStrSQL = lgStrSQL & " W2_E  = " & FilterVar(Trim(Request("txtW2_E")),"","S")   & ","    
 	
    lgStrSQL = lgStrSQL & " W6		 = " & UNIConvNum(Request("txtW6"),0)   & ","  
    lgStrSQL = lgStrSQL & " W7		 = " & UNIConvNum(Request("txtW7"),0)   & ","  
    lgStrSQL = lgStrSQL & " W8		 = " & UNIConvNum(Request("txtW8"),0)   & ","  
    lgStrSQL = lgStrSQL & " W9		 = " & FilterVar(Trim(Request("txtW9")),"","S")    & ","
    lgStrSQL = lgStrSQL & " W9_Value = " & UNIConvNum(Request("txtW9_value"),0)   & ","    
    lgStrSQL = lgStrSQL & " W10		 = " & UNIConvNum(Request("txtW10"),0)   & ","  
    lgStrSQL = lgStrSQL & " W11		 = " & UNIConvNum(Request("txtW11"),0) & ","  
    lgStrSQL = lgStrSQL & " W12		 = " & UNIConvNum(Request("txtW12"),0)   & ","
    lgStrSQL = lgStrSQL & " W13		 = " & UNIConvNum(Request("txtW13"),0)   & ","
    lgStrSQL = lgStrSQL & " W14		 = " & UNIConvNum(Request("txtW14"),0)   & ","
    lgStrSQL = lgStrSQL & " W15		 = " & UNIConvNum(Request("txtW15"),0)   & ","                                                                        
    lgStrSQL = lgStrSQL & " UPDT_USER_ID   = " & FilterVar(gUsrId,"''","S")   & ","  
    lgStrSQL = lgStrSQL & " UPDT_DT   = " &  FilterVar(GetSvrDateTime,"","S")   
    lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S") 
    lgStrSQL = lgStrSQL & "  and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S") 
    lgStrSQL = lgStrSQL & "  and rep_type =" &FilterVar(Trim(UCase(lgKeyStream(2))),"","S") 


    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords

	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)


    Select Case pMode 
      Case "R"
            lgStrSQL = "    Select TOP 1  "
            lgStrSQL = lgStrSQL & "  W1_S,				W1_E,				W2_S,				W2_E,  "
            lgStrSQL = lgStrSQL & "  W6,				W7,					W8,					W9,	   W9_Value,"
            lgStrSQL = lgStrSQL & "  W10,				W11,				W12,				W13,	   "
            lgStrSQL = lgStrSQL & "  W14,				W15				 "					
            lgStrSQL = lgStrSQL & " FROM  TB_68"
            lgStrSQL = lgStrSQL & " where "
            lgStrSQL = lgStrSQL &   pCode  
 

 
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

    lgErrorStatus     = "YES"
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next
    Err.Clear

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
       Case "<%=UID_M0001%>"

          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBQueryOk        
          End If   
       Case "<%=UID_M0002%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If   
       Case "<%=UID_M0003%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
    End Select    
       
</Script>	
