 <%@ LANGUAGE=VBSCript%>
 <%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<%

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

	Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("Q","M", "NOCOOKIE", "MB")   


	Dim lgStrPrevKey

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)
    

    Const C_SHEETMAXROWS_D  = 100


	'------ Developer Coding part (Start ) ------------------------------------------------------------------
 

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select
                
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
'    On Error Resume Next                                                             '¢Ð: Protect system from crashing
'    Err.Clear                                                                        '¢Ð: Clear Error status
   Call SubBizQueryMulti()
    
End Sub	
'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iKey1
    Dim strWhere, strWhere1

   On Error Resume Next                                                             '¢Ð: Protect system from crashing
   Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    

    strWhere = ""
    strWhere = strWhere & "  and  a.DTI_WDATE >= '" & FilterVar(UNIConvDate(Request("txtIssuedFromDt")),"","SNM") & "'"
    strWhere = strWhere & "  and  a.DTI_WDATE <= '" & FilterVar(UNIConvDate(Request("txtIssuedToDt")),"","SNM") & "'"
    
    		        
    if Trim(Request("cboBillStatus")) <> "" then    
		strWhere = strWhere & " and  F.DTI_STATUS  = '" & FilterVar(Request("cboBillStatus"), "", "SNM") & "'" 			    
    End if
    
    if Trim(Request("txtSupplierNm")) <> "" then
		strWhere = strWhere & " and BYR_COM_REGNO like '" & FilterVar(Request("txtSupplierNm"), "", "SNM") & "%'" 
    End if
    
    if Trim(Request("cboTaxArea")) <> "" then    
		strWhere = strWhere & " and Z.TAX_BIZ_AREA_CD  = '" & FilterVar(Request("cboTaxArea"), "", "SNM") & "'" 			    
    End if
    
     if Trim(Request("cboAmendCode")) <> "" then    
		strWhere = strWhere & " and A.AMEND_CODE  = '" & FilterVar(Request("cboAmendCode"), "", "SNM") & "'" 			    
    End if

            
    strWhere1 = ""
                  
    Call SubMakeSQLStatements("MR",strWhere,"X",strWhere1) 

                                '¡Ù: Make sql statements
   If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
           Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found.  
           Call SetErrorStatus()
    Else

       Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
       lgstrData = ""
       iDx       = 1
       Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ""            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CONVERSATION_ID"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DTI_TYPE"))            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MINOR_NAME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SEQ_ID"))
            lgstrData = lgstrData & Chr(11) & UniDateClientFormat(lgObjRs("DTI_WDATE"))  
            'lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("TOTAL_AMOUNT"), gCurrency, ggAmtOfMoneyNo, "X", "X")												            
            'lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("SUP_AMOUNT"), gCurrency, ggAmtOfMoneyNo, "X", "X")
            'lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("TAX_AMOUNT"), gCurrency, ggAmtOfMoneyNo, "X", "X") 
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("TOTAL_AMOUNT"), ggAmtOfMoney.DecPoint,0)            			            
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("SUP_AMOUNT"), ggAmtOfMoney.DecPoint,0)            			
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("TAX_AMOUNT"), ggAmtOfMoney.DecPoint,0)            			                                 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BYR_COM_REGNO"))                        
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BYR_COM_NAME"))                                    
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BYR_EMAIL"))                                                
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REF_INVOICE_NUM"))             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DTI_STATUS"))             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("S_MINOR_NAME")) 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DISUSE_REASON"))            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("AMEND_CODE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("AMEND_CODE_NM"))            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REMARK"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REMARK2"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REMARK3"))            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUP_COM_REGNO"))
            lgstrData = lgstrData & Chr(11) & ""						
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

    	    lgObjRs.MoveNext
          
            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey = lgStrPrevKey + 1
               Exit Do
            End If   
                       
        Loop 
    End If

      If iDx <= C_SHEETMAXROWS_D Then
         lgStrPrevKey = ""
      End If   

'      If CheckSQLError(lgObjRs.ActiveConnection) = True Then
'         ObjectContext.SetAbort
'      End If
            
      Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
      Call SubCloseRs(lgObjRs)                                            '¢Ð: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    
End Sub    


'============================================================================================================
' Name : SubBizSaveMultiCreate
' Desc : Save Multi Data
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
                                                     '¢Ð: Clear Error status

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
                      
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords 
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
   
End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords 
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
   Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
      
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
           
				Case "R"

				lgStrSQL = " "												
                lgStrSQL = lgStrSQL & vbCrLf & "  Select TOP " & iSelCount  & " A.CONVERSATION_ID,  A.DTI_TYPE,  Y.MINOR_NM AS MINOR_NAME,   A.SEQ_ID,	  A.DTI_WDATE ,	 "										 	 					 		 		 								
		        lgStrSQL = lgStrSQL & vbCrLf & "         A.SUP_AMOUNT + A.TAX_AMOUNT AS TOTAL_AMOUNT,   A.SUP_AMOUNT,   A.TAX_AMOUNT,	"	 										 								
		        lgStrSQL = lgStrSQL & vbCrLf & "         A.BYR_COM_REGNO,	A.BYR_COM_NAME,   A.BYR_EMAIL,	 A.REF_INVOICE_NUM,	"				        										 							
   	            lgStrSQL = lgStrSQL & vbCrLf & "         F.DTI_STATUS,      B.MINOR_NM as s_minor_name,  " 	 									
		        lgStrSQL = lgStrSQL & vbCrLf & "         F.SBDESCRIPTION AS DISUSE_REASON,  "
		        lgStrSQL = lgStrSQL & vbCrLf & "         A.AMEND_CODE ,  O.MINOR_NM AS AMEND_CODE_NM,	"	 
		        lgStrSQL = lgStrSQL & vbCrLf & "        A.REMARK,    A.REMARK2,    A.REMARK3,   A.SUP_COM_REGNO "		 		        
                lgStrSQL = lgStrSQL & vbCrLf & "  FROM	XXSB_DTI_MAIN A (NOLOCK) "
                lgStrSQL = lgStrSQL & vbCrLf & "  INNER	JOIN XXSB_DTI_STATUS F (NOLOCK)  "
                lgStrSQL = lgStrSQL & vbCrLf & "        ON (A.CONVERSATION_ID = F.CONVERSATION_ID  "
                lgStrSQL = lgStrSQL & vbCrLf & "       AND A.SUPBUY_TYPE = F.SUPBUY_TYPE    AND A.DIRECTION = F.DIRECTION) "
                lgStrSQL = lgStrSQL & vbCrLf & " LEFT OUTER JOIN B_TAX_BIZ_AREA Z (NOLOCK)   "
                lgStrSQL = lgStrSQL & vbCrLf & "      ON A.SUP_COM_REGNO = REPLACE(Z.OWN_RGST_NO,'-','') "
                lgStrSQL = lgStrSQL & vbCrLf & " LEFT OUTER JOIN B_MINOR B (NOLOCK)  "
                lgStrSQL = lgStrSQL & vbCrLf & "            ON (B.MAJOR_CD = 'DT409' AND B.MINOR_CD = F.DTI_STATUS) "
                lgStrSQL = lgStrSQL & vbCrLf & " LEFT OUTER JOIN S_TAX_BILL_HDR D (NOLOCK)  "
                lgStrSQL = lgStrSQL & vbCrLf & "            ON (SUBSTRING(A.CONVERSATION_ID,21,12) = D.BUYER_IV_NO ) "
                lgStrSQL = lgStrSQL & vbCrLf & " LEFT OUTER JOIN B_MINOR O (NOLOCK)  "
                lgStrSQL = lgStrSQL & vbCrLf & "            ON (O.MAJOR_CD = 'DT408' AND O.MINOR_CD = A.AMEND_CODE) "
                lgStrSQL = lgStrSQL & vbCrLf & " LEFT OUTER JOIN B_MINOR Y (NOLOCK)  "
                lgStrSQL = lgStrSQL & vbCrLf & "             ON (Y.MAJOR_CD = 'DT403' AND Y.MINOR_CD = A.DTI_TYPE) "
                lgStrSQL = lgStrSQL & vbCrLf & " WHERE	A.SUPBUY_TYPE	= 'AR'	 AND	A.DIRECTION		= '1'	 "				
			    lgStrSQL = lgStrSQL & vbCrLf &   pCode 
           End Select             
    End Select    
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '¢Ð: Protect system from crashing
    Err.Clear                                                                         '¢Ð: Clear Error status

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
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                   .ggoSpread.Source     = .frm1.vspdData
				   .ggoSpread.SSShowData "<%=lgstrData%>"	
                   .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                   .frm1.htxtIssuedFromDt.Value = "<%=ConvSPChars(Request("txtIssuedFromDt"))%>"
                   .frm1.htxtIssuedToDt.Value = "<%=ConvSPChars(Request("txtIssuedToDt"))%>"                   
			       .frm1.hcboBillStatus.Value = "<%=ConvSPChars(Request("cboBillStatus"))%>"			       
			       .frm1.htxtSupplierNm.Value = "<%=ConvSPChars(Request("txtSupplierNm"))%>" 			       
			       .frm1.hcboTaxArea.Value = "<%=ConvSPChars(Request("cboTaxArea"))%>"
			       .frm1.hcboAmendCode.Value = "<%=ConvSPChars(Request("cboAmendCode"))%>"
                   .DBQueryOk()        
	         End with
	      Else
	               parent.DBQueryNotOk()            
          End If   
	   Case "<%=UID_M0002%>"                                                         '¢Ð : Save
    End Select           
</Script>	
