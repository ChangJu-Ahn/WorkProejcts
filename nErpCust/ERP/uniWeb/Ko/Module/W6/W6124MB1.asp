
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
' Name : SubBizQueryF
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
				                .Frm1.txtW3_1_A.TEXT			  = "<%=ConvSPChars(lgObjRs("w3_1_A"))%>"
				                .Frm1.txtW3_1_B.TEXT			  = "<%=ConvSPChars(lgObjRs("w3_1_B"))%>"
				                .Frm1.txtW3_1_C.TEXT			  = "<%=ConvSPChars(lgObjRs("w3_1_C"))%>"
				                .Frm1.txtW4_1.TEXT			      = "<%=ConvSPChars(lgObjRs("w4_1"))%>"
				               
				                 .Frm1.txtW3_2_A.TEXT			  = "<%=ConvSPChars(lgObjRs("w3_2_A"))%>"
				                .Frm1.txtW3_2_B.TEXT			  = "<%=ConvSPChars(lgObjRs("w3_2_B"))%>"
				                .Frm1.txtW3_2_C.TEXT			  = "<%=ConvSPChars(lgObjRs("w3_2_C"))%>"
				                .Frm1.txtW4_2.TEXT			      = "<%=ConvSPChars(lgObjRs("w4_2"))%>"


								.Frm1.txtW1_3.value				  = "<%=ConvSPChars(lgObjRs("w1_3"))%>"
				                .Frm1.txtW2_3.value				  = "<%=ConvSPChars(lgObjRs("w2_3"))%>"
				                .Frm1.txtW3_3.value			       = "<%=ConvSPChars(lgObjRs("w3_3"))%>"
				                .Frm1.txtW4_3.value			      = "<%=ConvSPChars(lgObjRs("w4_3"))%>"
				                .Frm1.txtW4_Sum.value			      = "<%=ConvSPChars(lgObjRs("w4_Sum"))%>"
				                
				                
								.Frm1.txtW5_1.value				  = "<%=ConvSPChars(lgObjRs("w5_1"))%>"
				                .Frm1.txtW5_2.TEXT				  = "<%=ConvSPChars(lgObjRs("w5_2"))%>"
				                .Frm1.txtW5_3.TEXT			      = "<%=ConvSPChars(lgObjRs("w5_3"))%>"
				                .Frm1.txtW5_4.TEXT			      = "<%=ConvSPChars(lgObjRs("w5_4"))%>"
				                
				                .Frm1.txtW5_4_GB.value			      = "<%=ConvSPChars(lgObjRs("w5_4_GB"))%>"
 				                 
				               
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

    lgStrSQL = "DELETE From  TB_8_1"

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
    
    lgStrSQL = "INSERT INTO TB_8_1("
    lgStrSQL = lgStrSQL & " co_cd, "
    lgStrSQL = lgStrSQL & " FISC_YEAR, "
    lgStrSQL = lgStrSQL & " rep_type, "
    lgStrSQL = lgStrSQL & " W3_1_A, "
    lgStrSQL = lgStrSQL & " W3_1_B, "
    lgStrSQL = lgStrSQL & " W3_1_C, "
    lgStrSQL = lgStrSQL & " W4_1, "

    lgStrSQL = lgStrSQL & " W3_2_A, "
    lgStrSQL = lgStrSQL & " W3_2_B, "
    lgStrSQL = lgStrSQL & " W3_2_C, "
    lgStrSQL = lgStrSQL & " W4_2, "
 
    lgStrSQL = lgStrSQL & " W1_3, "
    lgStrSQL = lgStrSQL & " W2_3, "
    lgStrSQL = lgStrSQL & " W3_3, "
    lgStrSQL = lgStrSQL & " W4_3, "
    lgStrSQL = lgStrSQL & " W4_Sum, "
    
    lgStrSQL = lgStrSQL & " W5_1, "
    lgStrSQL = lgStrSQL & " W5_2, "
    lgStrSQL = lgStrSQL & " W5_3, "
    lgStrSQL = lgStrSQL & " W5_4, "
    lgStrSQL = lgStrSQL & " W5_4_GB, "       
    lgStrSQL = lgStrSQL & " INSRT_USER_ID, "
    lgStrSQL = lgStrSQL & " INSRT_DT, "
    lgStrSQL = lgStrSQL & " UPDT_USER_ID, "
    lgStrSQL = lgStrSQL & " UPDT_DT)"
    lgStrSQL = lgStrSQL & " VALUES("
    
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")     & ","                      
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")     & ","           
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")     & ","     
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW3_1_A"),0)     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW3_1_B"),0)     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW3_1_C"),0)     & ","    
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW4_1"),0)     & "," 
    
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW3_2_A"),0)     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW3_2_B"),0)     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW3_2_C"),0)     & ","    
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW4_2"),0)     & "," 
    
    
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW1_3"),"''","S")   & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW2_3"),"''","S")   & ","      
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW3_3"),"''","S")   & ","      
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW4_3"),0)     & "," 
   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW4_Sum"),0)     & "," 
    
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW5_1"),"''","S")   & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW5_2"),"''","S")   & ","      
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW5_3"),"''","S")   & ","      
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW5_4"),0)     & "," 
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW5_4_GB"),"''","S")   & ","  
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

    lgStrSQL = lgStrSQL & " Update TB_8_1 set"
    lgStrSQL = lgStrSQL & " W3_1_A  = " &  UNIConvNum(Request("txtW3_1_A"),0)   & ","
    lgStrSQL = lgStrSQL & " W3_1_B  = " &  UNIConvNum(Request("txtW3_1_B"),0)   & "," 
    lgStrSQL = lgStrSQL & " W3_1_C  = " &  UNIConvNum(Request("txtW3_1_C"),0)   & ","
    lgStrSQL = lgStrSQL & " W4_1   = " &   UNIConvNum(Request("txtW4_1"),0)   & ","
    
    lgStrSQL = lgStrSQL & " W3_2_A  = " & UNIConvNum(Request("txtW3_2_A"),0)   & ","
    lgStrSQL = lgStrSQL & " W3_2_B  = " & UNIConvNum(Request("txtW3_2_B"),0)   & ","
    lgStrSQL = lgStrSQL & " W3_2_C  = " & UNIConvNum(Request("txtW3_2_C"),0)   & ","
    lgStrSQL = lgStrSQL & " W4_2   = " &  UNIConvNum(Request("txtW4_2"),0)   & ","              
    
    
    lgStrSQL = lgStrSQL & " W1_3  = " & FilterVar(Trim(Request("txtW1_3")),"","S")   & ","
    lgStrSQL = lgStrSQL & " W2_3  = " & FilterVar(Trim(Request("txtW2_3")),"","S")   & ","
    lgStrSQL = lgStrSQL & " W3_3  = " & FilterVar(Trim(Request("txtW3_3")),"","S")   & ","
    lgStrSQL = lgStrSQL & " W4_3   = " & UNIConvNum(Request("txtW4_3"),0)   & ","       
    lgStrSQL = lgStrSQL & " W4_SUM  = " & UNIConvNum(Request("txtW4_SUM"),0)   & ","       
    
    lgStrSQL = lgStrSQL & " W5_1  = " & FilterVar(Trim(Request("txtW5_1")),"","S")   & ","
    lgStrSQL = lgStrSQL & " W5_2  = " & FilterVar(Trim(Request("txtW5_2")),"","S")   & ","  
    lgStrSQL = lgStrSQL & " W5_3  = " & FilterVar(Trim(Request("txtW5_3")),"","S")   & ","  
    lgStrSQL = lgStrSQL & " W5_4  = " & UNIConvNum(Request("txtW5_4"),0)     & "," 
    lgStrSQL = lgStrSQL & " W5_4_GB  = " & FilterVar(Trim(Request("txtw5_4_GB")),"","S")   & ","  
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
            lgStrSQL = lgStrSQL & "  W3_1_A,			W3_1_B,				W3_1_C,				W4_1,  "
            lgStrSQL = lgStrSQL & "  W3_2_A,			W3_2_B,				W3_2_C,				W4_2,  "
            lgStrSQL = lgStrSQL & "  W1_3  ,			W2_3  ,				W3_3,				W4_3,     W4_Sum,  " 
            lgStrSQL = lgStrSQL & "  W5_1  ,			W5_2  ,				W5_3,				W5_4,      W5_4_GB"  
            lgStrSQL = lgStrSQL & " FROM  TB_8_1"
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
