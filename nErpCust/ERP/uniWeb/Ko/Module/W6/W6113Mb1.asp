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
                .Frm1.txtW1.Value		  = "<%=ConvSPChars(lgObjRs("w1"))%>"
                .Frm1.txtW1_RATE.Value    = "<%=ConvSPChars(lgObjRs("w1_rate_view"))%>"
                .Frm1.txtW1_RATE_NEW.Value   = "<%=ConvSPChars(lgObjRs("w1_rate_value"))%>"
                .Frm1.txtW2.Value			 = "<%=ConvSPChars(lgObjRs("w2"))%>"
                .Frm1.txtW3.Value			 = "<%=ConvSPChars(lgObjRs("w3"))%>"              
                .Frm1.txtW4.Value			 = "<%=ConvSPChars(lgObjRs("w4"))%>"

                .Frm1.txtW5.Value			 = "<%=ConvSPChars(lgObjRs("w5"))%>"
                .Frm1.txtW5_RATE.Value		 = "<%=ConvSPChars(lgObjRs("w5_rate_view"))%>"
                .Frm1.txtW5_RATE_NEW.Value	 = "<%=ConvSPChars(lgObjRs("w5_rate_value"))%>"
                .Frm1.txtW6.Value			 = "<%=ConvSPChars(lgObjRs("w6"))%>"
                .Frm1.txtW7.Value			 = "<%=ConvSPChars(lgObjRs("w7"))%>"
                .Frm1.txtW8.Value			 = "<%=ConvSPChars(lgObjRs("w8"))%>"
                .Frm1.txtW9.Value			 = "<%=ConvSPChars(lgObjRs("w9"))%>"
                .Frm1.txtW10.Value			  = "<%=ConvSPChars(lgObjRs("w10"))%>"
               
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

    lgStrSQL = "DELETE From  TB_JT11_5"

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
    
    lgStrSQL = "INSERT INTO TB_JT11_5("
    lgStrSQL = lgStrSQL & " co_cd, "
    lgStrSQL = lgStrSQL & " FISC_YEAR, "
    lgStrSQL = lgStrSQL & " rep_type, "
    lgStrSQL = lgStrSQL & " W1, "
    lgStrSQL = lgStrSQL & " W1_Rate_Value, "
    lgStrSQL = lgStrSQL & " W1_Rate_VIEW, "
    lgStrSQL = lgStrSQL & " W2, "
    lgStrSQL = lgStrSQL & " W3, "
    lgStrSQL = lgStrSQL & " W4, "
    lgStrSQL = lgStrSQL & " W5, "
    lgStrSQL = lgStrSQL & " W5_Rate_Value, "
    lgStrSQL = lgStrSQL & " W5_Rate_VIEW, "
    lgStrSQL = lgStrSQL & " W6, "
    lgStrSQL = lgStrSQL & " W7, "
    lgStrSQL = lgStrSQL & " W8, "
    lgStrSQL = lgStrSQL & " W9, "
    lgStrSQL = lgStrSQL & " W10, "        
    lgStrSQL = lgStrSQL & " INSRT_USER_ID, "
    lgStrSQL = lgStrSQL & " INSRT_DT, "
    lgStrSQL = lgStrSQL & " UPDT_USER_ID, "
    lgStrSQL = lgStrSQL & " UPDT_DT)"
    lgStrSQL = lgStrSQL & " VALUES("
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")     & ","                      
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")     & ","           
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")     & ","             
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW1"),0)     & ","      
    lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(3)),0)     & ","    
    lgStrSQL = lgStrSQL & FilterVar(Trim((lgKeyStream(4))),"","S")      & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW2"),0)     & ","    
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW3"),0)     & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW4"),0)     & ","     
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW5"),0)     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(5)),0)     & ","    
    lgStrSQL = lgStrSQL & FilterVar(Trim((lgKeyStream(6))),"","S")      & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW6"),0)     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW7"),0)     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8"),0)     & ","    
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW9"),0)     & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW10"),0)     & ","     
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

    lgStrSQL = lgStrSQL & " Update TB_JT11_5 set"
    lgStrSQL = lgStrSQL & " W1  = " & UNIConvNum(Request("txtW1"),0)   & ","  
    lgStrSQL = lgStrSQL & " W1_RATE_VALUE = " &  UNIConvNum(Trim(lgKeyStream(3)),0) & ","  
    lgStrSQL = lgStrSQL & " W1_RATE_VIEW  = " & FilterVar(Trim(UCase(lgKeyStream(4))),"","S") & "," 
    lgStrSQL = lgStrSQL & " W2  = " & UNIConvNum(Request("txtW2"),0)   & ","   
    lgStrSQL = lgStrSQL & " W3  = " & UNIConvNum(Request("txtW3"),0)   & ","  
    lgStrSQL = lgStrSQL & " W4		 = " & UNIConvNum(Request("txtW4"),0) & ","  
    lgStrSQL = lgStrSQL & " W5		 = " & UNIConvNum(Request("txtW5"),0)  & ","  	
    lgStrSQL = lgStrSQL & " W5_RATE_VALUE = " &  UNIConvNum(Trim(lgKeyStream(5)),0) & ","  
    lgStrSQL = lgStrSQL & " W5_RATE_VIEW  = " & FilterVar(Trim(UCase(lgKeyStream(6))),"","S") & ","    
    lgStrSQL = lgStrSQL & " W6		 = " & UNIConvNum(Request("txtW6"),0)   & ","  
    lgStrSQL = lgStrSQL & " W7		 = " & UNIConvNum(Request("txtW7"),0)   & ","  
    lgStrSQL = lgStrSQL & " W8		 = " & UNIConvNum(Request("txtW8"),0)   & ","  
    lgStrSQL = lgStrSQL & " W9		 = " & UNIConvNum(Request("txtW9"),0)   & ","  
    lgStrSQL = lgStrSQL & " W10		 = " & UNIConvNum(Request("txtW10"),0)   & ","                  
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
            lgStrSQL = lgStrSQL & " W1,				W1_RATE_Value,		W1_RATE_VIEW,		W2,"
            lgStrSQL = lgStrSQL & " W3,				W4,					W5,					W5_RATE_Value,"
            lgStrSQL = lgStrSQL & " W5_RATE_VIEW,   W6,					W7,					W8,		      "
            lgStrSQL = lgStrSQL & " W9,				W10"	
            lgStrSQL = lgStrSQL & " FROM  TB_JT11_5 "
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
