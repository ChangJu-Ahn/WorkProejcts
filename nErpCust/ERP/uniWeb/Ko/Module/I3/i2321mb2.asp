<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--
'======================================================================================================
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : i2321mb2.asp
'*  4. Program Name         : 표준단가 수정(biz logic)
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2006/05/12
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Lee Seung Wook
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim strPlantCd
    Call HideStatusWnd
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""

    lgLngMaxRow     = Request("txtMaxRows")
    strPlantCd		= Trim(Request("txtPlantCd"))

    Call SubOpenDB(lgObjConn)
    
    Call SubBizSaveMulti()
    
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
    Err.Clear
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)

        
        
        Select Case arrColVal(0)
            Case "U"
				Call SubBizSaveMultiUpdate(arrColVal)
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

End Sub    

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next
    Err.Clear
    'replace(xxx,gComNum1000,"")
    lgStrSQL = "UPDATE  I_MATERIAL_VALUATION " & _
			   "   SET  STD_PRC    = " &  FilterVar(Trim(replace(arrColVal(4),gComNum1000,"")),"","D")   & "," & _
			   "		PREV_STD_PRC   = " &  FilterVar(Trim(replace(arrColVal(5),gComNum1000,"")),"","D")   & "," & _
			   "		UPDT_USER_ID	= " &  FilterVar(gUsrId,"''","S") & "," & _
			   "		UPDT_DT			= " &  FilterVar(GetSvrDateTime,"''","S") & _
			   " WHERE  PLANT_CD		= " &  FilterVar(strPlantCd,"''","S") & _
			   "   AND  ITEM_CD			= " &  FilterVar(Trim(UCase(arrColVal(2))),"''","S") & _	
			   "   AND  TRACKING_NO		= " &  FilterVar(Trim(UCase(arrColVal(3))),"''","S")

    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
    
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
    ObjectContext.SetAbort                                                         '☜: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
	Dim lsMsg
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
    
    If Trim("<%=lgErrorStatus%>") = "NO" Then
       Parent.DBSaveOk
    Else
       Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
    End If   
       
</Script>	