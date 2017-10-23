<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
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
    'lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    
'PrintForm


    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

	Call CheckVersion(wgFISC_YEAR, wgREP_TYPE)	' 2005-03-11 버전관리기능 추가 
	
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

    strWhere = " co_cd = " & FilterVar(Trim(UCase(wgCO_CD)),"","S")  
	strWhere = strWhere & "  and fisc_year =" & FilterVar(Trim(wgFISC_YEAR),"","S")
	strWhere = strWhere & "  and rep_type =" &  FilterVar(Trim(wgREP_TYPE),"","S")
	

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
       With parent.Frm1
                .txtW8_A_SUM.Value  = "<%=ConvSPChars(lgObjRs("W8_A_SUM"))%>"
                .txtW8_A_1.Value    = "<%=ConvSPChars(lgObjRs("W8_A_1"))%>"
                .txtW8_A_2.Value    = "<%=ConvSPChars(lgObjRs("W8_A_2"))%>"
                
                .txtW8_B_SUM.Value  = "<%=ConvSPChars(lgObjRs("W8_B_SUM"))%>"
                .txtW8_B_1.Value    = "<%=ConvSPChars(lgObjRs("W8_B_1"))%>"
                .txtW8_B_2.Value    = "<%=ConvSPChars(lgObjRs("W8_B_2"))%>"

                .txtW8_C_SUM.Value  = "<%=ConvSPChars(lgObjRs("W8_C_SUM"))%>"
                .txtW8_C_1.Value    = "<%=ConvSPChars(lgObjRs("W8_C_1"))%>"
                .txtW8_C_2.Value    = "<%=ConvSPChars(lgObjRs("W8_C_2"))%>"

                .txtW8_D_SUM.Value  = "<%=ConvSPChars(lgObjRs("W8_D_SUM"))%>"
                .txtW8_D_1.Value    = "<%=ConvSPChars(lgObjRs("W8_D_1"))%>"
                .txtW8_D_2.Value    = "<%=ConvSPChars(lgObjRs("W8_D_2"))%>"

                .txtW8_E_SUM.Value  = "<%=ConvSPChars(lgObjRs("W8_E_SUM"))%>"
                .txtW8_E_1.Value    = "<%=ConvSPChars(lgObjRs("W8_E_1"))%>"
                .txtW8_E_2.Value    = "<%=ConvSPChars(lgObjRs("W8_E_2"))%>"

                .txtW8_F_SUM.Value  = "<%=ConvSPChars(lgObjRs("W8_F_SUM"))%>"
                .txtW8_F_1.Value    = "<%=ConvSPChars(lgObjRs("W8_F_1"))%>"
                .txtW8_F_2.Value    = "<%=ConvSPChars(lgObjRs("W8_F_2"))%>"

                .txtW8_HAP_SUM.Value  = "<%=ConvSPChars(lgObjRs("W8_HAP_SUM"))%>"
                .txtW8_HAP_1.Value    = "<%=ConvSPChars(lgObjRs("W8_HAP_1"))%>"
                .txtW8_HAP_2.Value    = "<%=ConvSPChars(lgObjRs("W8_HAP_2"))%>"

                .txtW11.Value    = "<%=ConvSPChars(lgObjRs("W11"))%>"

                .txtW12_GA_A.Value		= "<%=ConvSPChars(lgObjRs("W12_GA_A"))%>"
                .txtW12_GA_B_VIEW.Value = "<%=ConvSPChars(lgObjRs("W12_GA_B_VIEW"))%>"
                .txtW12_GA_B_VAL.Value  = "<%=ConvSPChars(lgObjRs("W12_GA_B_VAL"))%>"
                .txtW12_GA_C.Value		= "<%=ConvSPChars(lgObjRs("W12_GA_C"))%>"

                .txtW12_NA_A.Value		= "<%=ConvSPChars(lgObjRs("W12_NA_A"))%>"
                .txtW12_NA_B_VIEW.Value = "<%=ConvSPChars(lgObjRs("W12_NA_B_VIEW"))%>"
                .txtW12_NA_B_VAL.Value  = "<%=ConvSPChars(lgObjRs("W12_NA_B_VAL"))%>"
                .txtW12_NA_C.Value		= "<%=ConvSPChars(lgObjRs("W12_NA_C"))%>"
                
                .txtW12_HAP_C.Value		= "<%=ConvSPChars(lgObjRs("W12_HAP_C"))%>"
                
                .txtW13.Value		= "<%=ConvSPChars(lgObjRs("W13"))%>"
                .txtW14_VIEW.Value		= "<%=ConvSPChars(lgObjRs("W14_VIEW"))%>"
                .txtW14_VAL.Value		= "<%=ConvSPChars(lgObjRs("W14_VAL"))%>"
                .txtW14.Value		= "<%=ConvSPChars(lgObjRs("W14"))%>"
                .txtW15.Value		= "<%=ConvSPChars(lgObjRs("W15"))%>"
                
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

    lgStrSQL = "DELETE From  TB_JT2_2_200603"

    lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(wgCO_CD)),"","S") 
    lgStrSQL = lgStrSQL & "		  and fisc_year =" & FilterVar(Trim(UCase(wgFISC_YEAR)),"","S") 
    lgStrSQL = lgStrSQL & "		  and rep_type =" &FilterVar(Trim(UCase(wgREP_TYPE)),"","S") 

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
    
    lgStrSQL = "INSERT INTO TB_JT2_2_200603 ("
    lgStrSQL = lgStrSQL & " co_cd, "
    lgStrSQL = lgStrSQL & " FISC_YEAR, "
    lgStrSQL = lgStrSQL & " rep_type, "
    lgStrSQL = lgStrSQL & " W8_A_SUM, W8_A_1, W8_A_2, "
    lgStrSQL = lgStrSQL & " W8_B_SUM, W8_B_1, W8_B_2, "
    lgStrSQL = lgStrSQL & " W8_C_SUM, W8_C_1, W8_C_2, "
    lgStrSQL = lgStrSQL & " W8_D_SUM, W8_D_1, W8_D_2, "
    lgStrSQL = lgStrSQL & " W8_E_SUM, W8_E_1, W8_E_2, "
    lgStrSQL = lgStrSQL & " W8_F_SUM, W8_F_1, W8_F_2, "
    lgStrSQL = lgStrSQL & " W8_HAP_SUM, W8_HAP_1, W8_HAP_2, "
    lgStrSQL = lgStrSQL & " W11, "
    lgStrSQL = lgStrSQL & " W12_GA_A, W12_GA_B_VIEW, W12_GA_B_VAL, W12_GA_C, "
    lgStrSQL = lgStrSQL & " W12_NA_A, W12_NA_B_VIEW, W12_NA_B_VAL, W12_NA_C, "
    lgStrSQL = lgStrSQL & " W12_HAP_C, "
    lgStrSQL = lgStrSQL & " W13, "
    lgStrSQL = lgStrSQL & " W14_VIEW, W14_VAL, W14, "
    lgStrSQL = lgStrSQL & " W15, "
    lgStrSQL = lgStrSQL & " INSRT_USER_ID, "
    lgStrSQL = lgStrSQL & " INSRT_DT, "
    lgStrSQL = lgStrSQL & " UPDT_USER_ID, "
    lgStrSQL = lgStrSQL & " UPDT_DT)"
    lgStrSQL = lgStrSQL & " VALUES("
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"","S")     & ","                      
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgFISC_YEAR)),"","S")     & ","           
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgREP_TYPE)),"","S")     & ","             
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_A_SUM"),0)     & ","      
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_A_1"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_A_2"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_B_SUM"),0)     & ","      
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_B_1"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_B_2"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_C_SUM"),0)     & ","      
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_C_1"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_C_2"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_D_SUM"),0)     & ","      
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_D_1"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_D_2"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_E_SUM"),0)     & ","      
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_E_1"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_E_2"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_F_SUM"),0)     & ","      
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_F_1"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_F_2"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_HAP_SUM"),0)     & ","      
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_HAP_1"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8_HAP_2"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW11"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW12_GA_A"),0)     & "," 
    lgStrSQL = lgStrSQL & FilterVar(Trim(Request("txtW12_GA_B_VIEW")),"","S")     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW12_GA_B_VAL"),0)     & ","        
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW12_GA_C"),0)     & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW12_NA_A"),0)     & "," 
    lgStrSQL = lgStrSQL & FilterVar(Trim(Request("txtW12_NA_B_VIEW")),"","S")     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW12_NA_B_VAL"),0)     & ","        
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW12_NA_C"),0)     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW12_HAP_C"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW13"),0)     & "," 
    lgStrSQL = lgStrSQL & FilterVar(Trim(Request("txtW14_VIEW")),"","S")     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW14_VAL"),0)     & ","  
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

    lgStrSQL = lgStrSQL & " Update TB_JT2_2_200603 set"

    lgStrSQL = lgStrSQL & " W8_A_SUM	= " & UNIConvNum(Request("txtW8_A_SUM"),0)   & ","  
    lgStrSQL = lgStrSQL & " W8_A_1		= " & UNIConvNum(Request("txtW8_A_1"),0)   & ","  
    lgStrSQL = lgStrSQL & " W8_A_2		= " & UNIConvNum(Request("txtW8_A_2"),0)   & ","  

    lgStrSQL = lgStrSQL & " W8_B_SUM	= " & UNIConvNum(Request("txtW8_B_SUM"),0)   & ","  
    lgStrSQL = lgStrSQL & " W8_B_1		= " & UNIConvNum(Request("txtW8_B_1"),0)   & ","  
    lgStrSQL = lgStrSQL & " W8_B_2		= " & UNIConvNum(Request("txtW8_B_2"),0)   & ","  

    lgStrSQL = lgStrSQL & " W8_C_SUM	= " & UNIConvNum(Request("txtW8_C_SUM"),0)   & ","  
    lgStrSQL = lgStrSQL & " W8_C_1		= " & UNIConvNum(Request("txtW8_C_1"),0)   & ","  
    lgStrSQL = lgStrSQL & " W8_C_2		= " & UNIConvNum(Request("txtW8_C_2"),0)   & ","  

    lgStrSQL = lgStrSQL & " W8_D_SUM	= " & UNIConvNum(Request("txtW8_D_SUM"),0)   & ","  
    lgStrSQL = lgStrSQL & " W8_D_1		= " & UNIConvNum(Request("txtW8_D_1"),0)   & ","  
    lgStrSQL = lgStrSQL & " W8_D_2		= " & UNIConvNum(Request("txtW8_D_2"),0)   & ","  

    lgStrSQL = lgStrSQL & " W8_E_SUM	= " & UNIConvNum(Request("txtW8_E_SUM"),0)   & ","  
    lgStrSQL = lgStrSQL & " W8_E_1		= " & UNIConvNum(Request("txtW8_E_1"),0)   & ","  
    lgStrSQL = lgStrSQL & " W8_E_2		= " & UNIConvNum(Request("txtW8_E_2"),0)   & ","  

    lgStrSQL = lgStrSQL & " W8_F_SUM	= " & UNIConvNum(Request("txtW8_F_SUM"),0)   & ","  
    lgStrSQL = lgStrSQL & " W8_F_1		= " & UNIConvNum(Request("txtW8_F_1"),0)   & ","  
    lgStrSQL = lgStrSQL & " W8_F_2		= " & UNIConvNum(Request("txtW8_F_2"),0)   & ","  

    lgStrSQL = lgStrSQL & " W8_HAP_SUM	= " & UNIConvNum(Request("txtW8_HAP_SUM"),0)   & ","  
    lgStrSQL = lgStrSQL & " W8_HAP_1	= " & UNIConvNum(Request("txtW8_HAP_1"),0)   & ","  
    lgStrSQL = lgStrSQL & " W8_HAP_2	= " & UNIConvNum(Request("txtW8_HAP_2"),0)   & ","  
    
    lgStrSQL = lgStrSQL & " W11  = " & UNIConvNum(Request("txtW11"),0)   & "," 
    
    lgStrSQL = lgStrSQL & " W12_GA_A  = " & UNIConvNum(Request("txtW12_GA_A"),0)   & ","  
    lgStrSQL = lgStrSQL & " W12_GA_B_VIEW  = " & FilterVar(Trim(Request("txtW12_GA_B_VIEW")),"","S")   & ","
    lgStrSQL = lgStrSQL & " W12_GA_B_VAL  = " & UNIConvNum(Request("txtW12_GA_B_VAL"),0)   & ","    
    lgStrSQL = lgStrSQL & " W12_GA_C  = " & UNIConvNum(Request("txtW12_GA_C"),0)   & ","  
    
    lgStrSQL = lgStrSQL & " W12_NA_A  = " & UNIConvNum(Request("txtW12_NA_A"),0)   & ","  
    lgStrSQL = lgStrSQL & " W12_NA_B_VIEW  = " & FilterVar(Trim(Request("txtW12_NA_B_VIEW")),"","S")   & ","
    lgStrSQL = lgStrSQL & " W12_NA_B_VAL  = " & UNIConvNum(Request("txtW12_NA_B_VAL"),0)   & ","    
    lgStrSQL = lgStrSQL & " W12_NA_C  = " & UNIConvNum(Request("txtW12_NA_C"),0)   & ","
    
    lgStrSQL = lgStrSQL & " W12_HAP_C  = " & UNIConvNum(Request("txtW12_HAP_C"),0)   & ","    
    
    lgStrSQL = lgStrSQL & " W13  = " & UNIConvNum(Request("txtW13"),0)   & ","
    lgStrSQL = lgStrSQL & " W14_VIEW  = " & FilterVar(Trim(Request("txtW14_VIEW")),"","S")   & ","   
    lgStrSQL = lgStrSQL & " W14_VAL  = " & UNIConvNum(Request("txtW14_VAL"),0)   & ","
    lgStrSQL = lgStrSQL & " W14  = " & UNIConvNum(Request("txtW14"),0)   & ","    
    
    lgStrSQL = lgStrSQL & " W15  = " & UNIConvNum(Request("txtW15"),0)   & ","    

    lgStrSQL = lgStrSQL & " UPDT_USER_ID   = " & FilterVar(gUsrId,"''","S")   & ","  
    lgStrSQL = lgStrSQL & " UPDT_DT   = " &  FilterVar(GetSvrDateTime,"","S")   
    lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(wgCO_CD)),"","S") 
    lgStrSQL = lgStrSQL & "  and fisc_year =" & FilterVar(Trim(UCase(wgFISC_YEAR)),"","S") 
    lgStrSQL = lgStrSQL & "  and rep_type =" &FilterVar(Trim(UCase(wgREP_TYPE)),"","S") 


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
            lgStrSQL = lgStrSQL & " *"
            lgStrSQL = lgStrSQL & " FROM  TB_JT2_2_200603 "
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

' 폼 출력 
Public Sub PrintForm()
	Dim item
%>
<table>
<%	
	For Each item In Request.Form
%>
<tr><td bgcolor=yellow><%=item%></td><td><%=Request.Form(item)%></td></tr>
<%	
	Next
%>
</table>
<%	
	Response.End
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
