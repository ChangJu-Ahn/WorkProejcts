<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim lgStrPrevKey
	Dim lgStrColorFlag
	Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd                                                               '��: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)							 '��: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '��: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '��: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    Call SubBizQueryMulti()
End Sub    

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

     strWhere = FilterVar(Trim(lgKeyStream(0)),"''","S")
    
    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQGT)                                 '�� : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("spec"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BILL_unit"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("BILL_qty"),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs("BILL_PRICE"), 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs("BILL_AMT_LOC"), 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs("VAT_AMT_LOC"), 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs("tot_amt"), 0)
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

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '��: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '��: Protect system from crashing

    Err.Clear                                                                        '��: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '��: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '��: Split Column data
        CALL SVRMSGBOX(arrColVal ,0,1)
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '��: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '��: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '��: Delete
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
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next
    Err.Clear
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
					lgStrSQL =            "			select	TOP " & iSelCount  & "  a.item_cd, "
					lgStrSQL = lgStrSQL & "					b.item_nm, "
					lgStrSQL = lgStrSQL & "					b.spec, "
					lgStrSQL = lgStrSQL & "					a.BILL_unit, "
					lgStrSQL = lgStrSQL & "					BILL_qty = sum(a.BILL_qty), "
					lgStrSQL = lgStrSQL & "					BILL_price = avg(a.BILL_PRICE), "
					lgStrSQL = lgStrSQL & "					bill_amt_loc = sum(a.BILL_AMT_LOC), "
					lgStrSQL = lgStrSQL & "					vat_AMT_loc = sum(a.vat_AMT_loc), "
					lgStrSQL = lgStrSQL & "					tot_amt = sum(a.BILL_AMT_LOC + A.VAT_AMT_loc) "
					lgStrSQL = lgStrSQL & "			 from	S_BILL_DTL A , B_ITEM B , S_BILL_HDR C "
					lgStrSQL = lgStrSQL & "			where	a.item_cd *= b.item_cd "
					lgStrSQL = lgStrSQL & "           and   a.bill_no = c.bill_no "
					lgStrSQL = lgStrSQL & "           and   c.post_flag = " & FilterVar("Y","''", "S")
					If lgkeystream(0) <> "" Then
						lgStrSQL = lgStrSQL & "             and 	A.BILL_NO = " & FilterVar(lgKeyStream(0),"''", "S")
					End If
					
					lgStrSQL = lgStrSQL & "         group by a.item_cd,b.item_nm,b.spec,a.BILL_unit ,		a.BILL_PRICE "
					lgStrSQL = lgStrSQL & "			order by 		a.item_cd "

           End Select             
                    
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
    lgErrorStatus     = "YES"                                                         '��: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
	Dim lsMsg
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Select Case pOpCode
        Case "MC"
        Case "MD"
        Case "MR"
        Case "MU"
    End Select
End Sub

%>

<Script Language="VBScript">
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '�� : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData2
                .lgStrPrevKey1    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .lgStrColorFlag = "<%=lgStrColorFlag%>"
                .DBDtlQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '�� : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '�� : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
      
</Script>	                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 