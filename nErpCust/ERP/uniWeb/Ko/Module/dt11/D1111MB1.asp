<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%    
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    
    
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection

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
    Dim strData
    Dim iLoop

    iKey1 = Request("txtUsrId1")          

    Call SubMakeSQLStatements(iKey1)                                       '¢Ð : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
          Call SetErrorStatus()
    Else
    
          iLoop   = 0 
          strData = ""
          
          Do While  Not (lgObjRs.EOF Or lgObjRs.BOF)
          
			strData = strData & Chr(11) & ConvSPChars(lgObjRs("USER_ID"))
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & ConvSPChars(lgObjRs("USER_NAME"))
			strData = strData & Chr(11) & ConvSPChars(lgObjRs("BP_RGST_NO"))
			strData = strData & Chr(11) & ConvSPChars(lgObjRs("DT_ID"))
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & ConvSPChars(lgObjRs("CERTI_PATH"))
			strData = strData & Chr(11) & ConvSPChars(lgObjRs("CERTI_PW"))
			strData = strData & Chr(11) & ConvSPChars(lgObjRs("CERTI_PATH2"))
			strData = strData & Chr(11) & ConvSPChars(lgObjRs("CERTI_PW2"))
			
			strData = strData & Chr(11) & iLoop
			strData = strData & Chr(11) & Chr(12)

			lgObjRs.MoveNext
			iLoop = iLoop + 1
          Loop
%>    
<Script Language=vbscript>
		parent.ggoSpread.Source = parent.frm1.vspdData
		parent.ggoSpread.SSShowDataByClip "<%=strData%>"
</Script>
    
<%    
    End If
    Call SubCloseRs(lgObjRs)
End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    
End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()


End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate(lgObjRs)

End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pCode)
     
     
     lgStrSQL = ""
            
     lgStrSQL = lgStrSQL & " SELECT A.USER_ID, B.USR_NM USER_NAME,  "
     lgStrSQL = lgStrSQL & " A.DT_ID, A.DT_PW,                      "
     lgStrSQL = lgStrSQL & " A.USER_DN, A.USER_INFO,                "
     lgStrSQL = lgStrSQL & " A.CERTI_PATH, A.CERTI_PW,              "
     lgStrSQL = lgStrSQL & " A.CERTI_PATH2, A.CERTI_PW2,            "
     lgStrSQL = lgStrSQL & " A.BP_RGST_NO                           "
     lgStrSQL = lgStrSQL & " FROM   DT_USER_INFO A (NOLOCK)         "
     lgStrSQL = lgStrSQL & " LEFT OUTER JOIN Z_USR_MAST_REC B (NOLOCK) ON A.USER_ID = B.USR_ID "
            
     if Trim(pCode) > "" then
        lgStrSQL = lgStrSQL & " WHERE A.USER_ID >= '" & pCode & "'"
     end if
            
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
