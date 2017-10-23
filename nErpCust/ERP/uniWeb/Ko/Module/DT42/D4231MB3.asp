<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%> 
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<%
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status


	Const C_SHEETMAXROWS_D = 100

    call LoadBasisGlobalInf()
    
    
	Call loadInfTB19029B( "I", "M","NOCOOKIE","MB") 
	Call LoadBNumericFormatB("I","M", "NOCOOKIE", "MB")  
  
    Call HideStatusWnd 

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
'    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)


    Call SubOpenDB(lgObjConn)           
   

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
'             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             'Call SubBizSaveMulti2()                          
    End Select
    
    Call SubCloseDB(lgObjConn)
    
    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next
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
' Name : SubBizSaveCreate2
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)

    On Error Resume Next
    Err.Clear                   

        lgStrSQL = ""                     
		lgStrSQL = lgStrSQL & " INSERT INTO  DT_VAT_ITEM("							
		lgStrSQL = lgStrSQL & " VAT_NO,  "		
		lgStrSQL = lgStrSQL & " VAT_SEQ,  "		
		lgStrSQL = lgStrSQL & " ITEM_CODE,  "		
		lgStrSQL = lgStrSQL & " ITEM_NAME,  "		
		lgStrSQL = lgStrSQL & " ITEM_SIZE,  "		
		lgStrSQL = lgStrSQL & " ITEM_MD,  "		
		lgStrSQL = lgStrSQL & " UNIT_PRICE,  "		
		lgStrSQL = lgStrSQL & " ITEM_QTY,  "		
		lgStrSQL = lgStrSQL & " SUP_AMOUNT,  "		
		lgStrSQL = lgStrSQL & " TAX_AMOUNT,  "		
		lgStrSQL = lgStrSQL & " REMARK,  "						
		lgStrSQL = lgStrSQL & " INSRT_USER_ID,  "
		lgStrSQL = lgStrSQL & " INSRT_DT,  "
		lgStrSQL = lgStrSQL & " UPDT_USER_ID,  "
		lgStrSQL = lgStrSQL & " UPDT_DT  "
		lgStrSQL = lgStrSQL & " )"																			
		lgStrSQL = lgStrSQL & " VALUES("   		
		lgStrSQL = lgStrSQL & FilterVar(arrColVal(2), "''", "S")     & ","		'세금계산서번호		
		lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(3),0)     & ","				'세금계산서순번		
		lgStrSQL = lgStrSQL & "null,"											'품목코드
		lgStrSQL = lgStrSQL & FilterVar(arrColVal(9), "''", "S")     & ","		'품목명	
		lgStrSQL = lgStrSQL & FilterVar(arrColVal(8), "''", "S")     & ","		'규격					
		lgStrSQL = lgStrSQL & "'" &  UNIConvDate(Trim(arrColVal(11))) & "',"	'발행일자				
		lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(6),0)     & ","				'단가		
		lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(7),0)     & ","				'수량				
		lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(5),0)     & ","				'금액
		lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(4),0)     & ","				'부가세		
		lgStrSQL = lgStrSQL & FilterVar(arrColVal(10), "''", "S")     & ","		'비고								
		lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","
		lgStrSQL = lgStrSQL & "getdate(),"
		lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","
		lgStrSQL = lgStrSQL & "getdate() "
		lgStrSQL = lgStrSQL & ")"	
		
								  		   			 										                   
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
		
End Sub


'============================================================================================================
' Name : SubBizSaveMultiUpdate2
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next
    Err.Clear                   
         
        lgStrSQL = ""                     
		lgStrSQL = lgStrSQL & " UPDATE  DT_VAT_ITEM "	
   	    lgStrSQL = lgStrSQL & " SET "     
   	    lgStrSQL = lgStrSQL & " item_name = " & FilterVar(Trim(arrColVal(9)),"''","S")     & ","    	    
   	    lgStrSQL = lgStrSQL & " item_size = " & FilterVar(Trim(arrColVal(8)),"''","S")     & ","    	    
   	    
   	    if Trim(arrColVal(11)) = "" then
   	       lgStrSQL = lgStrSQL & " item_md = null, " 
   	    else
           lgStrSQL = lgStrSQL & " item_md = '" &  UNIConvDate(Trim(arrColVal(11))) & "',"		'발행일   	          	    
   	    end if
   	    
   	    lgStrSQL = lgStrSQL & " unit_price =   " & UNIConvNum(arrColVal(6),0)   & ","     		   	       	    
		lgStrSQL = lgStrSQL & " item_qty =   " & UNIConvNum(arrColVal(7),0)   & ","     		
		lgStrSQL = lgStrSQL & " sup_amount =   " & UNIConvNum(arrColVal(5),0)   & ","     		
		lgStrSQL = lgStrSQL & " tax_amount =   " & UNIConvNum(arrColVal(4),0)   & ","     		
		lgStrSQL = lgStrSQL & " remark = " & FilterVar(Trim(arrColVal(10)),"''","S")     & ","    	    		
		lgStrSQL = lgStrSQL & " UPDT_USER_ID  = " & FilterVar(gUsrId,"''","S")   & ","
		lgStrSQL = lgStrSQL & " UPDT_DT    =  getdate() "    
		lgStrSQL = lgStrSQL & " WHERE   "
		lgStrSQL = lgStrSQL & " vat_no   =   " & FilterVar(arrColVal(2), "''", "S")  
		lgStrSQL = lgStrSQL & " and vat_seq   =   " & UNIConvNum(arrColVal(3),0)   
		
								  		   			 										                   
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
		
End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete2
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next
    Err.Clear                   
         
        lgStrSQL = ""          
		lgStrSQL = lgStrSQL & " Delete  DT_VAT_ITEM "	
		lgStrSQL = lgStrSQL & " WHERE   "
		lgStrSQL = lgStrSQL & " vat_no   =   " & FilterVar(arrColVal(2), "''", "S")  
		lgStrSQL = lgStrSQL & " and vat_seq   =   " & UNIConvNum(arrColVal(3),0)   
		
											  		   			 										                   
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
		
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
    Response.Write "<BR> Commit Event occur"
End Sub
'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
    Response.Write "<BR> Abort Event occur"
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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else   
          End If   
    End Select       
       
</Script>	
