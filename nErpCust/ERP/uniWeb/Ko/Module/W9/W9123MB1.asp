<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<!-- #Include file="../wcm/inc_SvrDebug.asp" -->
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
				       

            
            
            
				                .Frm1.cboW1.Value		     = "<%=ConvSPChars(lgObjRs("w1"))%>"
				                .Frm1.cboW2.Value		     = "<%=ConvSPChars(lgObjRs("w2"))%>"
				                .Frm1.txtW2_ETC.Value  		 = "<%=ConvSPChars(lgObjRs("w2_ETC"))%>"
				                .Frm1.cboW3.Value			 = "<%=ConvSPChars(lgObjRs("w3"))%>"   
				                .Frm1.txtW3_ETC.Value  		 = "<%=ConvSPChars(lgObjRs("w3_ETC"))%>"           
				                .Frm1.cboW4.Value			 = "<%=ConvSPChars(lgObjRs("w4"))%>"
								.Frm1.txtW4_ETC.Value  		 = "<%=ConvSPChars(lgObjRs("w4_ETC"))%>"
								.Frm1.txtW4_1.Value  		 = "<%=ConvSPChars(lgObjRs("w4_1"))%>"
				                .Frm1.cboW5.Value			 = "<%=ConvSPChars(lgObjRs("w5"))%>"
				                .Frm1.txtW5_ETC.Value  		 = "<%=ConvSPChars(lgObjRs("w5_ETC"))%>"
				                .Frm1.cboW6.Value			 = "<%=ConvSPChars(lgObjRs("w6"))%>"
				                .Frm1.txtW6_ETC.Value  		 = "<%=ConvSPChars(lgObjRs("w6_ETC"))%>"
				                .Frm1.txtW7_1.Value  		 = "<%=ConvSPChars(lgObjRs("W7_1"))%>"
				                .Frm1.txtW7_2.Value  		 = "<%=ConvSPChars(lgObjRs("W7_2"))%>"
				                .Frm1.cboW8.Value  			 = "<%=ConvSPChars(lgObjRs("W8"))%>"
				       
				                if "<%=ConvSPChars(lgObjRs("W9_1"))%>" = "Y" then
				                    .Frm1.chkW9_1.Checked = True
				                Else
				                     .Frm1.chkW9_1.Checked = False
				                end if
				                
				                if "<%=ConvSPChars(lgObjRs("W9_2"))%>" = "Y" then
				                    .Frm1.chkW9_2.Checked = True
				                Else
				                     .Frm1.chkW9_2.Checked = False
				                end if
				                
				               if "<%=ConvSPChars(lgObjRs("W9_3"))%>" = "Y" then
				                    .Frm1.chkW9_3.Checked = True
				                Else
				                     .Frm1.chkW9_3.Checked = False
				                end if
				                 if "<%=ConvSPChars(lgObjRs("W9_4"))%>" = "Y" then
				                    .Frm1.chkW9_4.Checked = True
				                Else
				                     .Frm1.chkW9_4.Checked = False
				                end if
				                if "<%=ConvSPChars(lgObjRs("W9_5"))%>" = "Y" then
				                    .Frm1.chkW9_5.Checked = True
				                Else
				                     .Frm1.chkW9_5.Checked = False
				                end if
				                
				                 if "<%=ConvSPChars(lgObjRs("W9_6"))%>" = "Y" then
				                    .Frm1.chkW9_6.Checked = True
				                Else
				                     .Frm1.chkW9_6.Checked = False
				                end if
				                 .Frm1.txtW9_6_ETC.Value  		 = "<%=ConvSPChars(lgObjRs("W9_6__ETC"))%>"
				                if "<%=ConvSPChars(lgObjRs("W10_1"))%>" = "Y" then
				                       .Frm1.chkW10_1.Checked = True
				                Else
				                     .Frm1.chkW10_1.Checked = False
				                end if
				                
				                if "<%=ConvSPChars(lgObjRs("W10_2"))%>" = "Y" then
				                    .Frm1.chkW10_2.Checked = True
				                Else
				                     .Frm1.chkW10_2.Checked = False
				                end if
				                
				               if "<%=ConvSPChars(lgObjRs("W10_3"))%>" = "Y" then
				                    .Frm1.chkW10_3.Checked = True
				                Else
				                     .Frm1.chkW10_3.Checked = False
				                end if
				                 if "<%=ConvSPChars(lgObjRs("W10_4"))%>" = "Y" then
				                    .Frm1.chkW10_4.Checked = True
				                Else
				                     .Frm1.chkW10_4.Checked = False
				                end if
				                if "<%=ConvSPChars(lgObjRs("W10_5"))%>" = "Y" then
				                    .Frm1.chkW10_5.Checked = True
				                Else
				                     .Frm1.chkW10_5.Checked = False
				                end if
				                
				                 if "<%=ConvSPChars(lgObjRs("W10_6"))%>" = "Y" then
				                    .Frm1.chkW10_6.Checked = True
				                Else
				                     .Frm1.chkW10_6.Checked = False
				                end if
				                
				                  if "<%=ConvSPChars(lgObjRs("W10_7"))%>" = "Y" then
				                    .Frm1.chkW10_7.Checked = True
				                Else
				                     .Frm1.chkW10_7.Checked = False
				                end if
				                 if "<%=ConvSPChars(lgObjRs("W10_8"))%>" = "Y" then
				                    .Frm1.chkW10_8.Checked = True
				                Else
				                     .Frm1.chkW10_8.Checked = False
				                end if
				                  if "<%=ConvSPChars(lgObjRs("W10_9"))%>" = "Y" then
				                    .Frm1.chkW10_9.Checked = True
				                Else
				                     .Frm1.chkW10_9.Checked = False
				                end if
				                 .Frm1.txtW10_9_ETC.Value  		 = "<%=ConvSPChars(lgObjRs("W10_9__ETC"))%>"
				                
				                
				               
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
    'On Error Resume Next
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

    lgStrSQL = "DELETE From  TB_JS1"

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
    'On Error Resume Next
    Err.Clear

    lgStrSQL = "INSERT INTO TB_JS1("
    lgStrSQL = lgStrSQL & " co_cd, "
    lgStrSQL = lgStrSQL & " FISC_YEAR, "
    lgStrSQL = lgStrSQL & " rep_type, "
    lgStrSQL = lgStrSQL & " W1, "     '3
    lgStrSQL = lgStrSQL & " W2, "	  '4
    lgStrSQL = lgStrSQL & " W2_ETC, "
    lgStrSQL = lgStrSQL & " W3, "
    lgStrSQL = lgStrSQL & " W3_ETC, "
    lgStrSQL = lgStrSQL & " W4, "
    lgStrSQL = lgStrSQL & " W4_ETC, "
    lgStrSQL = lgStrSQL & " W4_1, "
    
    lgStrSQL = lgStrSQL & " W5, "
    lgStrSQL = lgStrSQL & " W5_ETC, "
    lgStrSQL = lgStrSQL & " W6, "
    lgStrSQL = lgStrSQL & " W6_ETC, "
    lgStrSQL = lgStrSQL & " W7_1, "
    lgStrSQL = lgStrSQL & " W7_2, "
    lgStrSQL = lgStrSQL & " W8, "
    lgStrSQL = lgStrSQL & " W9_1, "
    lgStrSQL = lgStrSQL & " W9_2, "
    lgStrSQL = lgStrSQL & " W9_3, "
    
    lgStrSQL = lgStrSQL & " W9_4, "
    lgStrSQL = lgStrSQL & " W9_5, "
    lgStrSQL = lgStrSQL & " W9_6, "
    lgStrSQL = lgStrSQL & " W9_6__ETC, "
    lgStrSQL = lgStrSQL & " W10_1, "
    lgStrSQL = lgStrSQL & " W10_2, "
    lgStrSQL = lgStrSQL & " W10_3, "
    lgStrSQL = lgStrSQL & " W10_4, "
    lgStrSQL = lgStrSQL & " W10_5, "
    lgStrSQL = lgStrSQL & " W10_6, "
    
    lgStrSQL = lgStrSQL & " W10_7, "
    lgStrSQL = lgStrSQL & " W10_8, "
    lgStrSQL = lgStrSQL & " W10_9, "
    lgStrSQL = lgStrSQL & " W10_9__ETC, "  
               
    lgStrSQL = lgStrSQL & " INSRT_USER_ID, "
    lgStrSQL = lgStrSQL & " INSRT_DT, "
    lgStrSQL = lgStrSQL & " UPDT_USER_ID, "
    lgStrSQL = lgStrSQL & " UPDT_DT)"
    lgStrSQL = lgStrSQL & " VALUES("
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")     & ","                      
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")     & ","           
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")     & ","             
    
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(3))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(4))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(5))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(6))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(7))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(8))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(9))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(10))),"","S")     & ","       
    
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(11))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(12))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(13))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(14))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(15))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(16))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(17))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(18))),"","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(19))),"","S")     & ","              
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(20))),"","S")     & ","       
    
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(21))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(22))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(23))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(24))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(25))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(26))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(27))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(28))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(29))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(30))),"","S")     & ","       
    
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(31))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(32))),"","S")     & ","       
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(33))),"","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(34))),"","S")     & ","              

    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                       
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","                                     
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                     
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S")                                            
    lgStrSQL = lgStrSQL & ")" 

	PrintLog lgStrSQL
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

    lgStrSQL = lgStrSQL & " Update TB_JS1 set"
    
	lgStrSQL = lgStrSQL & " co_cd			= " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " FISC_YEAR		= " & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " rep_type		= " & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W1				= " & FilterVar(Trim(UCase(lgKeyStream(3))),"","S")     & ","    	
    lgStrSQL = lgStrSQL & " W2				= " & FilterVar(Trim(UCase(lgKeyStream(4))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W2_ETC			= " & FilterVar(Trim(UCase(lgKeyStream(5))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W3				= " & FilterVar(Trim(UCase(lgKeyStream(6))),"","S")     & ","    	
    lgStrSQL = lgStrSQL & " W3_ETC			= " & FilterVar(Trim(UCase(lgKeyStream(7))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W4				= " & FilterVar(Trim(UCase(lgKeyStream(8))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W4_ETC			= " & FilterVar(Trim(UCase(lgKeyStream(9))),"","S")     & ","    	
    lgStrSQL = lgStrSQL & " W4_1			= " & FilterVar(Trim(UCase(lgKeyStream(10))),"","S")     & ","    
    
    lgStrSQL = lgStrSQL & " W5				= " & FilterVar(Trim(UCase(lgKeyStream(11))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W5_ETC			= " & FilterVar(Trim(UCase(lgKeyStream(12))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W6				= " & FilterVar(Trim(UCase(lgKeyStream(13))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W6_ETC			= " & FilterVar(Trim(UCase(lgKeyStream(14))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W7_1			= " & FilterVar(Trim(UCase(lgKeyStream(15))),"","S")     & ","    	
    lgStrSQL = lgStrSQL & " W7_2			= " & FilterVar(Trim(UCase(lgKeyStream(16))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W8				= " & FilterVar(Trim(UCase(lgKeyStream(17))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W9_1			= " & FilterVar(Trim(UCase(lgKeyStream(18))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W9_2			= " & FilterVar(Trim(UCase(lgKeyStream(19))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W9_3			= " & FilterVar(Trim(UCase(lgKeyStream(20))),"","S")     & ","    
    
    lgStrSQL = lgStrSQL & " W9_4			= " & FilterVar(Trim(UCase(lgKeyStream(21))),"","S")     & ","    	
    lgStrSQL = lgStrSQL & " W9_5			= " & FilterVar(Trim(UCase(lgKeyStream(22))),"","S")     & ","    	
    lgStrSQL = lgStrSQL & " W9_6			= " & FilterVar(Trim(UCase(lgKeyStream(23))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W9_6__ETC		= " & FilterVar(Trim(UCase(lgKeyStream(24))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W10_1			= " & FilterVar(Trim(UCase(lgKeyStream(25))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W10_2			= " & FilterVar(Trim(UCase(lgKeyStream(26))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W10_3			= " & FilterVar(Trim(UCase(lgKeyStream(27))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W10_4			= " & FilterVar(Trim(UCase(lgKeyStream(28))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W10_5			= " & FilterVar(Trim(UCase(lgKeyStream(29))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W10_6			= " & FilterVar(Trim(UCase(lgKeyStream(30))),"","S")     & ","    
    
    lgStrSQL = lgStrSQL & " W10_7			= " & FilterVar(Trim(UCase(lgKeyStream(31))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W10_8			= " & FilterVar(Trim(UCase(lgKeyStream(32))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W10_9			= " & FilterVar(Trim(UCase(lgKeyStream(33))),"","S")     & ","    
    lgStrSQL = lgStrSQL & " W10_9__ETC		= " & FilterVar(Trim(UCase(lgKeyStream(34))),"","S")     & ","    	  
      
    lgStrSQL = lgStrSQL & " UPDT_USER_ID   = " & FilterVar(gUsrId,"''","S")   & ","  
    lgStrSQL = lgStrSQL & " UPDT_DT   = " &  FilterVar(GetSvrDateTime,"","S")   
    lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S") 
    lgStrSQL = lgStrSQL & "  and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S") 
    lgStrSQL = lgStrSQL & "  and rep_type =" &FilterVar(Trim(UCase(lgKeyStream(2))),"","S") 

	PrintLog lgStrSQL
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
            lgStrSQL = lgStrSQL & "  W1,				W2,					W2_ETC,           "
            lgStrSQL = lgStrSQL & "  W3,				W3_ETC,				W4,			  "
            lgStrSQL = lgStrSQL & "  W4_ETC,			W4_1,				W5,		      "
            lgStrSQL = lgStrSQL & "  W5_ETC,			W6,					W6_ETC,		     W7_1 ,"					
            lgStrSQL = lgStrSQL & "  W7_2,				W8,					W9_1,		     W9_2 ,"	
            lgStrSQL = lgStrSQL & "  W9_3,				W9_4,				W9_5,		     W9_6 ,"	
            lgStrSQL = lgStrSQL & "  W9_6__ETC,			W10_1,				W10_2,		     W10_3 , "
            lgStrSQL = lgStrSQL & "  W10_4,				W10_5,				W10_6,		     W10_7 ,"		
            lgStrSQL = lgStrSQL & "  W10_8,				W10_9,				W10_9__ETC"
            lgStrSQL = lgStrSQL & " FROM  TB_JS1"
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
