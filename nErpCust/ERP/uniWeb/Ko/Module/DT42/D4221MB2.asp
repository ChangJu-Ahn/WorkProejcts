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
	Dim lgStrPrevKey1

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)

    lgStrPrevKey1 = UNICInt(Trim(Request("lgStrPrevKey1")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Const C_SHEETMAXROWS_D  = 500


	'------ Developer Coding part (Start ) ------------------------------------------------------------------
 

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             'Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             'Call SubBizDelete()
    End Select
                
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
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
    Dim YYYYMM

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
 
    strWhere = ""
    strWhere = strWhere & " and A.CONVERSATION_ID  = '" & FilterVar(Request("txtConvid"), "", "SNM") & "'" 	
    
    		                
    strWhere1 = ""
                  
    Call SubMakeSQLStatements("MR",strWhere,"X",strWhere1) 

                                '¡Ù: Make sql statements
   If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey1 = ""
           'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found.  
           Call SetErrorStatus()
    Else
       Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey1)
       lgstrData = ""
       iDx       = 1
       Do While Not lgObjRs.EOF
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CONVERSATION_ID"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CODE"))                        
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NAME"))            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_SIZE"))            
            'lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("UNIT_PRICE"), gCurrency, ggAmtOfMoneyNo, "X", "X")                                                
            'lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("ITEM_QTY"),   gCurrency, ggAmtOfMoneyNo, "X", "X")            
            'lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("SUP_AMOUNT"), gCurrency, ggAmtOfMoneyNo, "X", "X")
            'lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("TAX_AMOUNT"), gCurrency, ggAmtOfMoneyNo, "X", "X")            
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("UNIT_PRICE"), ggAmtOfMoney.DecPoint,0)            			            
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ITEM_QTY"), ggAmtOfMoney.DecPoint,0)            			            
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("SUP_AMOUNT"), ggAmtOfMoney.DecPoint,0)            			            
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("TAX_AMOUNT"), ggAmtOfMoney.DecPoint,0)            			            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CURRENCY_CODE"))                                                
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REMARK"))            
            lgstrData = lgstrData & Chr(11) & UniDateClientFormat(lgObjRs("ITEM_MD"))                        
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DTI_LINE_NUM"))            
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

    	    lgObjRs.MoveNext
          
            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey1 = lgStrPrevKey1 + 1
               Exit Do
            End If   
                       
        Loop 
    End If
           
      If iDx <= C_SHEETMAXROWS_D Then
         lgStrPrevKey1 = ""
      End If   

'      If CheckSQLError(lgObjRs.ActiveConnection) = True Then
'         ObjectContext.SetAbort
'      End If
            
      Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
      Call SubCloseRs(lgObjRs)                                                              '¢Ð: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '¢Ð: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '¢Ð: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '¢Ð: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '¢Ð: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '¢Ð: Delete
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

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------


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
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------        
                 
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
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
  
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
      
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey1 + 1
           
           Select Case Mid(pDataType,2,1)
           
				Case "R"

                lgStrSQL = " "				
                lgStrSQL = lgStrSQL & vbCrLf & " Select TOP " & iSelCount  & " A.CONVERSATION_ID,   A.ITEM_CODE,     A.ITEM_NAME, "
		        lgStrSQL = lgStrSQL & vbCrLf & "         A.ITEM_SIZE,       A.UNIT_PRICE,    A.ITEM_QTY, "
		        lgStrSQL = lgStrSQL & vbCrLf & "         A.SUP_AMOUNT,      A.TAX_AMOUNT,    A.CURRENCY_CODE, "
		        lgStrSQL = lgStrSQL & vbCrLf & "         A.REMARK,   	    A.ITEM_MD,    	 A.DTI_LINE_NUM "
                lgStrSQL = lgStrSQL & vbCrLf & " FROM	XXSB_DTI_ITEM A (NOLOCK) "
                lgStrSQL = lgStrSQL & vbCrLf & " WHERE	A.SUPBUY_TYPE		= 'AR'	"
                lgStrSQL = lgStrSQL & vbCrLf & "        AND	A.DIRECTION			= '1' "                  
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
                   .ggoSpread.Source     = .frm1.vspdData2
				   .ggoSpread.SSShowData "<%=lgstrData%>"	
                   .lgStrPrevKey1    = "<%=lgStrPrevKey1%>"
                    .DBQueryOk2     
	         End with 
	      else
	             .DBQueryNotOk2           
          End If   
	   Case "<%=UID_M0002%>"														'¢Ð : Save
		  If Trim("<%=lgErrorStatus%>") = "NO" Then
             'Parent.DBSaveOk2
'          Else
'            Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   		   
    End Select    
       
</Script>	
