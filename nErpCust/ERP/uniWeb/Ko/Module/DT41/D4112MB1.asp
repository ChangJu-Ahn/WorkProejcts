<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<% 											'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next														'☜:
Err.Clear


Dim lgStrPrevKey

Const C_SHEETMAXROWS_D = 100

Call LoadBasisGlobalInf()

Call loadInfTB19029B("I", "B","NOCOOKIE","MB")

Call HideStatusWnd 

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    'lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
               
    Call SubOpenDB(lgObjConn)           
   
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
'             Call SubBizDelete()                          
    End Select
    
    Call SubCloseDB(lgObjConn)

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

End Sub	


'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere, strFg
    Dim strWhere1
  
 '   On Error Resume Next    
 '   Err.Clear                                                               '☜: Clear Error status
  
	    strWhere = ""		
    

'		if Request("txtFrStartDt") <> "" then
'		strWhere = strWhere & "  and  a.pu_start_dt >= '" & FilterVar(UNIConvDate(Request("txtFrStartDt")),"","SNM") & "'"
'		end if
		
    
		if Request("txtuserId") <> "" then
		strWhere = strWhere & "  Where  CERT_REGNO = '" & FilterVar(Request("txtuserId"), "", "SNM") & "'"
		end if
    
                 
        strWhere1 = "" 
         
	    Call SubMakeSQLStatements("MR",strWhere,"X",strWhere1)   
    
                 '☆: Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
           Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.  
           Call SetErrorStatus()
    Else
       Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
       lgstrData = ""
       iDx       = 1
       Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CHECK1"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CERT_REGNO"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CERT_COMNAME"))            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("USERDN"))
            lgstrData = lgstrData & Chr(11) & UniDateClientFormat(lgObjRs("EXPIRATION_DATE"))                        
                                    
            'lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ALLOW_UNIT"), ggAmtOfMoney.DecPoint,0)
            
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
      Call SubCloseRs(lgObjRs)    

End Sub    


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

'    On Error Resume Next
'    Err.Clear                                                                        '☜: Clear Error status
    
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
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveCreate(arrColVal)

    On Error Resume Next
    Err.Clear          
        
    
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
	
				
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)

    On Error Resume Next
    Err.Clear                   
            
      
        lgStrSQL = ""
		lgStrSQL = lgStrSQL & " INSERT INTO XXSB_DTI_CERT("		
		lgStrSQL = lgStrSQL & " CERT_REGNO, " 		
		lgStrSQL = lgStrSQL & " CERT_COMNAME, " 									
		lgStrSQL = lgStrSQL & " INSERT_USER_ID ,"
		lgStrSQL = lgStrSQL & " INSERT_DATE	  ,"
		lgStrSQL = lgStrSQL & " UPDATE_USER_ID  ,"
		lgStrSQL = lgStrSQL & " UPDATE_DATE    )"        
		lgStrSQL = lgStrSQL & " VALUES("   		
		lgStrSQL = lgStrSQL & FilterVar(arrColVal(2), "''", "S")     & ","		
		lgStrSQL = lgStrSQL & FilterVar(arrColVal(3), "''", "S")     & ","				
		'lgStrSQL = lgStrSQL & "'" &  UNIConvDate(Trim(arrColVal(4))) & "',"		'PU 부여일		
		'lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(7),0)     & ","				'부여Unit				
		lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","
		lgStrSQL = lgStrSQL & "getdate(),"
		lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","
		lgStrSQL = lgStrSQL & "getdate())" 

                   
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
		
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    
    On Error Resume Next
    Err.Clear      
 
  
            
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
		
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db

'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
		lgStrSQL = ""
		lgStrSQL = " Delete  XXSB_DTI_CERT"
		lgStrSQL = lgStrSQL & " WHERE   "
        lgStrSQL = lgStrSQL & " CERT_REGNO   =   " & FilterVar(arrColVal(2), "''", "S")   

			    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
    
End Sub


'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
	
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
					   lgStrSQL = ""						   
					   lgStrSQL = lgStrSQL & vbCrLf & " Select TOP " & iSelCount & " '0' as check1, CERT_REGNO, CERT_COMNAME, USERDN, USERINFO, EXPIRATION_DATE "						   
					   lgStrSQL = lgStrSQL & vbCrLf & " from  XXSB_DTI_CERT (NOLOCK)  "
					   lgStrSQL = lgStrSQL & vbCrLf &   pCode 	
					   					   					   					   					                           						   												 												  						   						                                                                    				   	                                                                                                   						        				        
          End Select
    End Select
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
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .frm1.hUserId.Value		= "<%=ConvSPChars(Request("txtuserId"))%>"
                
                .DBQueryOk()       
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select       
       
</Script>	
