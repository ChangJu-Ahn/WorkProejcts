<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/adovbs.inc"  -->
<!-- #Include file="../../inc/incServerAdoDB.asp"  -->

<!-- #Include file="../../inc/lgsvrvariables.inc"  -->
<% 
    Call LoadBasisGlobalInf() 

	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    'Single
    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    
    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)


 '------ Developer Coding part (Start ) ------------------------------------------------------------------


 '------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    Dim iLcNo

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Dim strBizAreaCD2
 Dim strIssueDt3
 Dim strIssueDt4
 Dim strGubun
 
 strBizAreaCD2 = Trim(Request("txtBizAreaCD2"))
 strIssueDt3   = Trim(Request("txtIssueDt3"))
 strIssueDt4   = Trim(Request("txtIssueDt4"))
 strGubun      = Trim(Request("rdoGubun"))
    
    Call SubMakeSQLStatements(strBizAreaCD2)                             '☆: Make sql statements

    If  FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then               'If data not exists
          Call DisplayMsgBox("124200", vbInformation, "", "", I_MKSCRIPT)            '☜: No data is found. 
  %>
<Script Language=vbscript>
 '------ Developer Coding part (Start ) ------------------------------------------------------------------
      Parent.Frm1.txtBizAreaNM2.Value = ""
</Script>       
<%             
          Call SetErrorStatus()
    Else
%>
<Script Language=vbscript>
 '------ Developer Coding part (Start ) ------------------------------------------------------------------
 ' Set condition area, contents area
 '--------------------------------------------------------------------------------------------------------

  With Parent.Frm1
      .txtBizAreaNM2.Value = "<%=lgObjRs("tax_biz_area_nm")%>"
   
  End With          
 '------ Developer Coding part (End   ) ------------------------------------------------------------------
</Script>       
<%     
       Call SubCloseRs(lgObjRs)                                                          '☜ : Release RecordSSet
       Call SubBizSaveSingleUpdate()
       
    End If
End Sub 
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
    
    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜: Create
              Call SubBizSaveSingleCreate()
        Case  OPMD_UMODE                                                             '☜: Update
              Call SubBizSaveSingleUpdate()
    End Select

End Sub 
     
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti(pKey1)
  
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    

End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
 Dim strBizAreaCD2
 Dim strIssueDt3
 Dim strIssueDt4
 Dim strGubun
 Dim strfileGubun
 
 strBizAreaCD2 = Trim(Request("txtBizAreaCD2"))
 strIssueDt3   = Trim(Request("txtIssueDt3"))
 strIssueDt4   = Trim(Request("txtIssueDt4"))
 strGubun      = Trim(Request("rdoGubun"))
 strfileGubun  = Trim(Request("rdofileGubun"))   
 
 lgStrSQL = "UPDATE  A_Vat" 
 lgStrSQL = lgStrSQL & " Set  Made_Vat_Fg = " & FilterVar("N", "''", "S") & " " 
 If strfileGubun = "B"  Then
     lgStrSQL = lgStrSQL & " ,  MISS_FG = '' " 
 End IF
 lgStrSQL = lgStrSQL & " FROM A_Vat A, B_Configuration C"
 lgStrSQL = lgStrSQL & " WHERE  A.Report_Biz_Area_Cd = " & FilterVar(strBizAreaCD2, "''", "S")
 lgStrSQL = lgStrSQL & " AND    A.Issued_Dt >=" & FilterVar(UNIConvDate(strIssueDt3), "''", "S")
 lgStrSQL = lgStrSQL & " AND    A.Issued_Dt <=" & FilterVar(UNIConvDate(strIssueDt4), "''", "S")
 lgStrSQL = lgStrSQL & " AND    (A.Io_Fg = " & FilterVar("I", "''", "S") & "  OR A.Io_Fg = " & FilterVar("O", "''", "S") & " )"
 lgStrSQL = lgStrSQL & " AND    A.Conf_Fg = " & FilterVar("C", "''", "S") & " " 
 If strfileGubun = "A"  Then
     lgStrSQL = lgStrSQL & " AND    A.Made_Vat_Fg = " & FilterVar("Y", "''", "S") & " " 
 Else
     lgStrSQL = lgStrSQL & " AND    A.MISS_FG = " & FilterVar("Y", "''", "S") & " "  
 End IF
 lgStrSQL = lgStrSQL & " AND    A.Vat_Type = C.Minor_Cd"
 lgStrSQL = lgStrSQL & " AND    C.Major_Cd = " & FilterVar("B9001", "''", "S") & " "
 If strGubun = "6" Then
	lgStrSQL = lgStrSQL & " AND    C.Seq_No in (6,7) "
 Else
	lgStrSQL = lgStrSQL & " AND    C.Seq_No = " & strGubun
 End If	
 lgStrSQL = lgStrSQL & " AND    C.Reference = " & FilterVar("Y", "''", "S") & " "
 
 '//Response.Write "UNIConvDateCompanyToDB(strIssueDt3)" & UNIConvDate(strIssueDt3)
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
 Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
 
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)

  

End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

 
End Sub


'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pCode)
    '------ Developer Coding part (Start ) ------------------------------------------------------------------
        lgStrSQL = "SELECT TAX_Biz_Area_Nm " 
        lgStrSQL = lgStrSQL & " FROM  B_TAX_Biz_Area "
        lgStrSQL = lgStrSQL & " WHERE TAX_Biz_Area_Cd = " & FilterVar(pCode, "''", "S")  
        
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
    lgErrorStatus     = "YES"                                                         '☜: Set error status
 '------ Developer Coding part (Start ) ------------------------------------------------------------------
 '------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
 If CheckSYSTEMError(pErr,True) = True Then
       Call DisplayMsgBox("800407", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
       ObjectContext.SetAbort
       Call SetErrorStatus
    Else
       If CheckSQLError(pConn,True) = True Then
          Call DisplayMsgBox("800407", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
          ObjectContext.SetAbort
          Call SetErrorStatus
       Else
   '//성공했을 경우 
   Call DisplayMsgBox("183114", vbInformation, "", "", I_MKSCRIPT)     '성공했을 경우 
       End If
    End If
    
End Sub

%>

