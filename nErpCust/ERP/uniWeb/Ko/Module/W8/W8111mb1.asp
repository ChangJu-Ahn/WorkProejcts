
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
				
				              .frm1.txtw5_AMT.value			=   "<%=UNIConvNum(lgObjRs("W5_AMT"),0)%>" 
				              .frm1.txtw5_Rate.value		=   "<%=UNIConvNum(lgObjRs("W5_RATE"),0)%>" 
				              .frm1.txtw5_Rate_VAL.value    =   "<%=UNIConvNum(lgObjRs("W5_RATE_VAL"),0)%>" 
				              .frm1.txtw5_Tax.value			=   "<%=UNIConvNum(lgObjRs("W5_TAX"),0)%>" 
				              
				              .frm1.txtw6.value				=   "<%=ConvSPChars(lgObjRs("W6"))%>" 
				              .frm1.txtw6_AMT.value			=   "<%=UNIConvNum(lgObjRs("W6_AMT"),0)%>" 
				              .frm1.txtw6_Rate.value		=   "<%=UNIConvNum(lgObjRs("W6_RATE"),0)%>" 
				              .frm1.txtw6_Rate_VAL.value    =   "<%=UNIConvNum(lgObjRs("W6_RATE_VAL"),0)%>" 
				              .frm1.txtw6_Tax.value			=   "<%=UNIConvNum(lgObjRs("W6_TAX"),0)%>" 
				               
				              .frm1.txtw7.value				=   "<%=ConvSPChars(lgObjRs("W7"))%>" 
				              .frm1.txtw7_AMT.value			=   "<%=UNIConvNum(lgObjRs("W7_AMT"),0)%>" 
				              .frm1.txtw7_Rate.value		=   "<%=UNIConvNum(lgObjRs("W7_RATE"),0)%>" 
				              .frm1.txtw7_Rate_VAL.value    =   "<%=UNIConvNum(lgObjRs("W7_RATE_VAL"),0)%>" 
				              .frm1.txtw7_Tax.value			=   "<%=UNIConvNum(lgObjRs("W7_TAX"),0)%>"  
				              
				
				              .frm1.txtw8_AMT.value			=   "<%=UNIConvNum(lgObjRs("W8_AMT"),0)%>" 
				              .frm1.txtw8_Tax.value			=   "<%=UNIConvNum(lgObjRs("W8_TAX"),0)%>"  
				              
				              
				              .frm1.txtw10_AMT.value		=   "<%=UNIConvNum(lgObjRs("W10_AMT"),0)%>" 
				              .frm1.txtw10_Rate.value		=   "<%=UNIConvNum(lgObjRs("W10_RATE"),0)%>" 
				              .frm1.txtw10_Rate_VAL.value   =   "<%=UNIConvNum(lgObjRs("W10_RATE_VAL"),0)%>" 
				              .frm1.txtw10_Tax.value		=   "<%=UNIConvNum(lgObjRs("W10_TAX"),0)%>"  
				              
				              .frm1.txtw11.value			=   "<%=ConvSPChars(lgObjRs("W11"))%>" 
				              .frm1.txtw11_AMT.value		=   "<%=UNIConvNum(lgObjRs("W11_AMT"),0)%>" 
				              .frm1.txtw11_Rate.value		=   "<%=UNIConvNum(lgObjRs("W11_RATE"),0)%>" 
				              .frm1.txtw11_Rate_VAL.value   =   "<%=UNIConvNum(lgObjRs("W11_RATE_VAL"),0)%>" 
				              .frm1.txtw11_Tax.value		=   "<%=UNIConvNum(lgObjRs("W11_TAX"),0)%>" 
				              
				              .frm1.txtw12_AMT.value		=   "<%=UNIConvNum(lgObjRs("W12_AMT"),0)%>" 
				              .frm1.txtw12_Tax.value		=   "<%=UNIConvNum(lgObjRs("W12_TAX"),0)%>"  
				              
				              
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

    lgStrSQL = "DELETE From  TB_12"

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
    
    lgStrSQL = "INSERT INTO TB_12("
    lgStrSQL = lgStrSQL & " co_cd, "
    lgStrSQL = lgStrSQL & " FISC_YEAR, "
    lgStrSQL = lgStrSQL & " rep_type, "
    lgStrSQL = lgStrSQL & " W5_AMT, "
    lgStrSQL = lgStrSQL & " W5_RATE, "
    lgStrSQL = lgStrSQL & " W5_RATE_VAL, "
    lgStrSQL = lgStrSQL & " W5_Tax, "
    
    lgStrSQL = lgStrSQL & " W6, "
    lgStrSQL = lgStrSQL & " W6_AMT, "
    lgStrSQL = lgStrSQL & " W6_Rate, "
    lgStrSQL = lgStrSQL & " W6_Rate_VAL, "
    lgStrSQL = lgStrSQL & " W6_Tax, "
    
    lgStrSQL = lgStrSQL & " W7, "
    lgStrSQL = lgStrSQL & " W7_AMT, "
    lgStrSQL = lgStrSQL & " W7_Rate, "
    lgStrSQL = lgStrSQL & " W7_Rate_VAL, "
    lgStrSQL = lgStrSQL & " W7_Tax, "
    
    lgStrSQL = lgStrSQL & " W8_AMT, "
    lgStrSQL = lgStrSQL & " W8_TAX, "
    
    lgStrSQL = lgStrSQL & " W10_AMT, "
    lgStrSQL = lgStrSQL & " W10_Rate, "
    lgStrSQL = lgStrSQL & " W10_Rate_VAL, "
    lgStrSQL = lgStrSQL & " W10_Tax, "
    
    lgStrSQL = lgStrSQL & " W11, "
    lgStrSQL = lgStrSQL & " W11_AMT, "
    lgStrSQL = lgStrSQL & " W11_Rate, "
    lgStrSQL = lgStrSQL & " W11_Rate_VAL, "
    lgStrSQL = lgStrSQL & " W11_Tax, "
    
        
    lgStrSQL = lgStrSQL & " W12_AMT, "
    lgStrSQL = lgStrSQL & " W12_TAX, "
    
    lgStrSQL = lgStrSQL & " INSRT_USER_ID, "
    lgStrSQL = lgStrSQL & " INSRT_DT, "
    lgStrSQL = lgStrSQL & " UPDT_USER_ID, "
    lgStrSQL = lgStrSQL & " UPDT_DT)"
    lgStrSQL = lgStrSQL & " VALUES("
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")     & ","                      
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")     & ","           
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(3),0)     & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(4))),"","S")     & ","       
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(5),0)     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(6),0)     & ","
    
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(7))),"","S")     & ","            
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(8),"''","S")     & "," 
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(9))),"","S")     & ","  
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(10),"''","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(11),"''","S")     & ","
    
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(12))),"","S")     & ","            
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(13),"''","S")     & "," 
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(14))),"","S")     & ","  
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(15),"''","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(16),"''","S")     & ","
    
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(17),"''","S")     & ","            
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(18),"''","S")     & ","  
    
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(19),0)     & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(20))),"","S")     & ","       
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(21),0)     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(22),0)     & ","
    
     
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(23))),"","S")     & ","            
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(24),"''","S")     & "," 
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(25))),"","S")     & ","  
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(26),"''","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(27),"''","S")     & ","
    
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(28),"''","S")     & ","            
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(29),"''","S")     & "," 
 
    
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

    lgStrSQL = lgStrSQL & " Update TB_12 set"


    lgStrSQL = lgStrSQL & " W5_AMT            =  "& UNIConvNum(lgKeyStream(3),0)             & ","       
    lgStrSQL = lgStrSQL & " W5_RATE			  =  "& FilterVar(lgKeyStream(4),"''","S")             & ","       
    lgStrSQL = lgStrSQL & " W5_RATE_VAL       =  "& UNIConvNum(lgKeyStream(5),0)            & ","     
    lgStrSQL = lgStrSQL & " W5_TAX			  =  "& UNIConvNum(lgKeyStream(6),0)      & "," 
    
    lgStrSQL = lgStrSQL & " W6				  =  "& FilterVar(lgKeyStream(7),"''","S")       & ","
    lgStrSQL = lgStrSQL & " W6_AMT			  =  "& UNIConvNum(lgKeyStream(8),0)             & ","       
    lgStrSQL = lgStrSQL & " W6_RATE			  =  "& FilterVar(lgKeyStream(9),"''","S")       & ","
    lgStrSQL = lgStrSQL & " W6_RATE_VAL		  =  "& UNIConvNum(lgKeyStream(10),0)       & ","
    lgStrSQL = lgStrSQL & " W6_TAX			  =  "& UNIConvNum(lgKeyStream(11),0)      & ","     
    
    lgStrSQL = lgStrSQL & " W7				  =  "& FilterVar(lgKeyStream(12),"''","S")       & ","
    lgStrSQL = lgStrSQL & " W7_AMT			  =  "& UNIConvNum(lgKeyStream(13),0)             & ","       
    lgStrSQL = lgStrSQL & " W7_RATE			  =  "& FilterVar(lgKeyStream(14),"''","S")       & ","
    lgStrSQL = lgStrSQL & " W7_RATE_VAL		  =  "& UNIConvNum(lgKeyStream(15),0)       & ","
    lgStrSQL = lgStrSQL & " W7_TAX			  =  "& UNIConvNum(lgKeyStream(16),0)      & ","    
    
    lgStrSQL = lgStrSQL & " W8_AMT			  =  "& UNIConvNum(lgKeyStream(17),0)       & ","
    lgStrSQL = lgStrSQL & " W8_TAX			  =  "& UNIConvNum(lgKeyStream(18),0)      & ","    
    
    
    lgStrSQL = lgStrSQL & " W10_AMT           =  "& UNIConvNum(lgKeyStream(19),0)             & ","       
    lgStrSQL = lgStrSQL & " W10_RATE		  =  "& FilterVar(lgKeyStream(20),"''","S")             & ","       
    lgStrSQL = lgStrSQL & " W10_RATE_VAL      =  "& UNIConvNum(lgKeyStream(21),0)            & ","     
    lgStrSQL = lgStrSQL & " W10_TAX			  =  "& UNIConvNum(lgKeyStream(22),0)      & ","     
    
    
    lgStrSQL = lgStrSQL & " W11				  =  "& FilterVar(lgKeyStream(23),"''","S")       & ","
    lgStrSQL = lgStrSQL & " W11_AMT			  =  "& UNIConvNum(lgKeyStream(24),0)             & ","       
    lgStrSQL = lgStrSQL & " W11_RATE		  =  "& FilterVar(lgKeyStream(25),"''","S")       & ","
    lgStrSQL = lgStrSQL & " W11_RATE_VAL	  =  "& UNIConvNum(lgKeyStream(26),0)       & ","
    lgStrSQL = lgStrSQL & " W11_TAX			  =  "& UNIConvNum(lgKeyStream(27),0)      & ","      
    
    lgStrSQL = lgStrSQL & " W12_AMT			  =  "& UNIConvNum(lgKeyStream(28),0)       & ","
    lgStrSQL = lgStrSQL & " W12_TAX			  =  "& UNIConvNum(lgKeyStream(29),0)      & ","    
    
    
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
            lgStrSQL = lgStrSQL & "  W5_AMT,			W5_RATE,		    W5_RATE_VAL,		W5_TAX,		     "
            lgStrSQL = lgStrSQL & "  W6,				W6_AMT,				W6_RATE,			W6_RATE_VAL,	      W6_TAX,"
            lgStrSQL = lgStrSQL & "  W7,				W7_AMT,				W7_RATE,			W7_RATE_VAL,	      W7_TAX,"
            lgStrSQL = lgStrSQL & "  W8_AMT,			W8_TAX,		"
            lgStrSQL = lgStrSQL & "  W10_AMT,			W10_RATE,		    W10_RATE_VAL,		W10_TAX,		     "		
            lgStrSQL = lgStrSQL & "  W11,				W11_AMT,			W11_RATE,			W11_RATE_VAL,	      W11_TAX,"			
            lgStrSQL = lgStrSQL & "  W12_AMT,			W12_TAX	"																								
            lgStrSQL = lgStrSQL & " FROM  TB_12"
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
