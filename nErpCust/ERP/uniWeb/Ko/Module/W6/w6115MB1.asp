<%@ LANGUAGE="VBScript" CODEPAGE=949 TRANSACTION=Required%>
<% Option Explicit%>
<% session.CodePage=949 %>
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
				                .Frm1.txtW1_1A.Value		  = "<%=ConvSPChars(lgObjRs("w1_1A"))%>"
				                .Frm1.txtW1_1B.Value		  = "<%=ConvSPChars(lgObjRs("w1_1B"))%>"
				                .Frm1.txtW1_2A.Value		  = "<%=ConvSPChars(lgObjRs("w1_2A"))%>"
				                .Frm1.txtW1_2B.Value		  = "<%=ConvSPChars(lgObjRs("w1_2B"))%>"
				                .Frm1.txtW1_3A.Value		  = "<%=ConvSPChars(lgObjRs("w1_3A"))%>"
				                .Frm1.txtW1_3B.Value		  = "<%=ConvSPChars(lgObjRs("w1_3B"))%>"
				                .Frm1.txtW1_SUM.Value		  = "<%=ConvSPChars(lgObjRs("w1_SUM"))%>"
				                
				                IF <%=ConvSPChars(lgObjRs("w2_4"))%> =  "1" THEN
				                    .Frm1.chkW2_4.checked		  = True 
				                Else
				                    .Frm1.chkW2_4.checked		  = False
				                End if   
				                
				                IF <%=ConvSPChars(lgObjRs("w2_5"))%> =  "1" THEN
				                    .Frm1.chkW2_5.checked		  = True 
				                Else
				                    .Frm1.chkW2_5.checked		  = False     
				                End if   
				                
				                IF <%=ConvSPChars(lgObjRs("w2_6"))%> =  "1" THEN
				                    .Frm1.chkW2_6.checked		  = True 
				                Else
				                    .Frm1.chkW2_6.checked		  = False     
				                End if  
				                
				                IF <%=ConvSPChars(lgObjRs("w2_7"))%> =  "1" THEN
				                    .Frm1.chkW2_7.checked		  = True 
				                Else
				                    .Frm1.chkW2_7.checked		  = False     
				                End if    
				                
				                
				                 .Frm1.txtW3_8.Value		  = "<%=ConvSPChars(lgObjRs("w3_8"))%>"
				                 .Frm1.txtW3_9.Value		  = "<%=ConvSPChars(lgObjRs("w3_9"))%>"
				                 .Frm1.txtW3_10.Value		  = "<%=ConvSPChars(lgObjRs("w3_10"))%>"
				                 .Frm1.txtW3_11.Value		  = "<%=ConvSPChars(lgObjRs("w3_11"))%>"
				                 .Frm1.txtW3_12.Value		  = "<%=ConvSPChars(lgObjRs("w3_12"))%>"
				                 .Frm1.txtW3_Sum.Value		  = "<%=ConvSPChars(lgObjRs("w3_Sum"))%>"
				                
				                 
				               
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

    lgStrSQL = "DELETE From  TB_8_5_2"

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
    
    lgStrSQL = "INSERT INTO TB_8_5_2("
    lgStrSQL = lgStrSQL & " co_cd, "
    lgStrSQL = lgStrSQL & " FISC_YEAR, "
    lgStrSQL = lgStrSQL & " rep_type, "
    lgStrSQL = lgStrSQL & " W1_1A, "
    lgStrSQL = lgStrSQL & " W1_1B, "
    lgStrSQL = lgStrSQL & " W1_2A, "
    lgStrSQL = lgStrSQL & " W1_2B, "
    lgStrSQL = lgStrSQL & " W1_3A, "
    lgStrSQL = lgStrSQL & " W1_3B, "
    lgStrSQL = lgStrSQL & " W1_Sum, "
    lgStrSQL = lgStrSQL & " W2_4, "
    lgStrSQL = lgStrSQL & " W2_5, "
    lgStrSQL = lgStrSQL & " W2_6, "
    lgStrSQL = lgStrSQL & " W2_7, "
    lgStrSQL = lgStrSQL & " W3_8, "
    lgStrSQL = lgStrSQL & " W3_9, "
    lgStrSQL = lgStrSQL & " W3_10, "
    lgStrSQL = lgStrSQL & " W3_11, "
    lgStrSQL = lgStrSQL & " W3_12, "
    lgStrSQL = lgStrSQL & " W3_Sum, "             
    lgStrSQL = lgStrSQL & " INSRT_USER_ID, "
    lgStrSQL = lgStrSQL & " INSRT_DT, "
    lgStrSQL = lgStrSQL & " UPDT_USER_ID, "
    lgStrSQL = lgStrSQL & " UPDT_DT)"
    lgStrSQL = lgStrSQL & " VALUES("
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")     & ","                      
    lgStrSQL = lgStrSQL & FilterVar(Trim((lgKeyStream(1))),"","S")     & ","           
    lgStrSQL = lgStrSQL & FilterVar(Trim((lgKeyStream(2))),"","S")     & ","             
    
    lgStrSQL = lgStrSQL & FilterVar(Trim((lgKeyStream(3))),"","S")    & ","
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(4),0)     & ","
	lgStrSQL = lgStrSQL & FilterVar(Trim((lgKeyStream(5))),"","S")    & ","
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(6),0)     & ","                     
    lgStrSQL = lgStrSQL & FilterVar(Trim((lgKeyStream(7))),"","S")    & ","
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(8),0)     & ","                     
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(9),0)     & ","  
    
    
    lgStrSQL = lgStrSQL & FilterVar(Trim((lgKeyStream(10))),"","S")    & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim((lgKeyStream(11))),"","S")    & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim((lgKeyStream(12))),"","S")    & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim((lgKeyStream(13))),"","S")    & ","
        
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(14),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(15),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(16),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(17),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(18),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(19),0)     & ","  
     
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

    lgStrSQL = lgStrSQL & " Update TB_8_5_2 set"
    lgStrSQL = lgStrSQL & " W1_1A  = " & FilterVar(Trim((lgKeyStream(3))),"","S")   & ","  
    lgStrSQL = lgStrSQL & " W1_1B  = " & UNIConvNum(lgKeyStream(4),0)  & ","
    lgStrSQL = lgStrSQL & " W1_2A  = " & FilterVar(Trim((lgKeyStream(5))),"","S")   & ","  
    lgStrSQL = lgStrSQL & " W1_2B  = " & UNIConvNum(lgKeyStream(6),0)  & ","
    lgStrSQL = lgStrSQL & " W1_3A  = " & FilterVar(Trim((lgKeyStream(7))),"","S")   & ","  
    lgStrSQL = lgStrSQL & " W1_3B  = " & UNIConvNum(lgKeyStream(8),0)  & ","   
    lgStrSQL = lgStrSQL & " W1_Sum  = " & UNIConvNum(lgKeyStream(9),0)  & ","        
    
    lgStrSQL = lgStrSQL & " W2_4  = " & FilterVar(Trim((lgKeyStream(10))),"","S")   & ","
    lgStrSQL = lgStrSQL & " W2_5  = " & FilterVar(Trim((lgKeyStream(11))),"","S")   & ","  
    lgStrSQL = lgStrSQL & " W2_6  = " & FilterVar(Trim((lgKeyStream(12))),"","S")   & ","  
    lgStrSQL = lgStrSQL & " W2_7  = " & FilterVar(Trim((lgKeyStream(13))),"","S")   & "," 
    
    lgStrSQL = lgStrSQL & " W3_8   = " & UNIConvNum(lgKeyStream(14),0)  & ","
    lgStrSQL = lgStrSQL & " W3_9   = " & UNIConvNum(lgKeyStream(15),0)  & ","                
    lgStrSQL = lgStrSQL & " W3_10  = " & UNIConvNum(lgKeyStream(16),0)  & ","        
    lgStrSQL = lgStrSQL & " W3_11  = " & UNIConvNum(lgKeyStream(17),0)  & ","        
    lgStrSQL = lgStrSQL & " W3_12  = " & UNIConvNum(lgKeyStream(18),0)  & ","        
    lgStrSQL = lgStrSQL & " W3_Sum = " & UNIConvNum(lgKeyStream(19),0)  & ","           
    
                                    
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
            lgStrSQL = lgStrSQL & "  W1_1A,				W1_1B,				W1_2A,        W1_2B,   "
            lgStrSQL = lgStrSQL & "  W1_3A,				W1_3B,				W1_Sum,			  "
            lgStrSQL = lgStrSQL & "  W2_4,				W2_4,				W2_5,		  W2_6,       "
            lgStrSQL = lgStrSQL & "  W2_7,				W3_8,				W3_9,		  W3_10 ,"			
            lgStrSQL = lgStrSQL & "  W3_11,				W3_12,				W3_Sum"				
            lgStrSQL = lgStrSQL & " FROM  TB_8_5_2"
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
