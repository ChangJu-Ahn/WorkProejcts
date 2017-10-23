<%@ LANGUAGE="VBScript" CODEPAGE=949 TRANSACTION=Required%>
<% Option Explicit%>
<% session.CodePage=949 %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../wcm/inc_SvrDebug.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%    
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    Dim lgFISC_YEAR, lgREP_TYPE
	lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtCo_Cd"),gColSep)
	lgFISC_YEAR 		= Request("txtFISC_YEAR")
	lgREP_TYPE 			= Request("cboREP_TYPE")


	Call InitSpreadPosVariables	' 그리드 위치 초기화 함수 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

	Call CheckVersion(Request("txtFISC_YEAR"), Request("cboREP_TYPE"))	' 2005-03-11 버전관리기능 추가 
	
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)
             Call SubBizQuery()
        Case CStr(UID_M0002)
             Call SubBizSave()
        Case CStr(UID_M0003)
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)
    
'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()	' 데이타 넘겨주는 컬럼 기준 
End Sub


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	Dim iCnt
    Dim iKey1, iKey2, iKey3
    Dim strPreCD

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0),"''", "S")
	iKey2 = FilterVar(lgFISC_YEAR,"''", "S")
	iKey3 = FilterVar(lgREP_TYPE,"''", "S")
	
    Call SubMakeSQLStatements("R",iKey1,iKey2, iKey3)                                       '☜ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
          'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
          Call SetErrorStatus()
    Else


        lgstrData = ""
        iCnt = 1
        strPreCD = ""

%>
<Script Language=vbscript>
        Parent.Frm1.txtCO_CD.Value  = "<%=ConvSPChars(lgObjRs("CO_CD"))%>"
        Parent.Frm1.txtFISC_YEAR.text  = "<%=ConvSPChars(lgObjRs("FISC_YEAR"))%>"
        Parent.Frm1.cboREP_TYPE.Value  = "<%=ConvSPChars(lgObjRs("REP_TYPE"))%>"
        Parent.Frm1.txtCO_CD_Body.Value  = "<%=ConvSPChars(lgObjRs("CO_CD"))%>"
        Parent.Frm1.txtFISC_YEAR_Body.Value  = "<%=ConvSPChars(lgObjRs("FISC_YEAR"))%>"
        Parent.Frm1.txtREP_TYPE_Body.Value  = "<%=ConvSPChars(lgObjRs("REP_TYPE"))%>"
        Parent.Frm1.txtBS_PL_FG.Value  = "<%=ConvSPChars(lgObjRs("W1"))%>"
</Script>       
<%     

        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W2"))		' GP_CD
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAR_GP_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W3"))		' GP_NM
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W4"))		' FISC_CD
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W5"))		' AMT

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUM_FG"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GP_LVL"))
            
            lgstrData = lgstrData & Chr(11) & Chr(12)

			strPreCD = ConvSPChars(lgObjRs("GP_CD"))
		    lgObjRs.MoveNext
			iCnt = iCnt + 1
        Loop 
    End If

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)
End Sub	

'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next
    Err.Clear
  
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
 
    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜ : Create
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

    lgStrSQL = "DELETE  TB_3_3_3 WITH (ROWLOCK) "
    lgStrSQL = lgStrSQL & " WHERE CO_CD = " &  FilterVar(Trim(UCase(wgCO_CD)),"''","S") & vbCrLf
    lgStrSQL = lgStrSQL & " AND FISC_YEAR = " &  FilterVar(Request("txtFISC_YEAR"),"''","S") & vbCrLf
    lgStrSQL = lgStrSQL & " AND REP_TYPE = " &  FilterVar(Request("cboREP_TYPE"),"''","S") & vbCrLf
    lgStrSQL = lgStrSQL & " AND W1 = " &  FilterVar(Request("txtBS_PL_FG"),"''","S") & vbCrLf
    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)
End Sub


'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow
    Dim iDx , i

    On Error Resume Next
    Err.Clear 

	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
	PrintLog "Spead count: " & Request("txtSpread")
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)    
        
	    lgStrSQL =            " INSERT INTO TB_3_3_3 WITH (ROWLOCK) "
	    lgStrSQL = lgStrSQL & " (CO_CD, FISC_YEAR, REP_TYPE "
	    lgStrSQL = lgStrSQL & "   , W1, W2, W3, W4, W5 "
	    lgStrSQL = lgStrSQL & "   , INSRT_USER_ID, UPDT_USER_ID ) "
	    
	    lgStrSQL = lgStrSQL & " VALUES ( " 
	    lgStrSQL = lgStrSQL & FilterVar(Request("txtCO_CD_Body"),"''","S") & ","
	    lgStrSQL = lgStrSQL & FilterVar(Request("txtFISC_YEAR_Body"),"''","S") & ","
	    lgStrSQL = lgStrSQL & FilterVar(Request("txtREP_TYPE_Body"),"''","S") & ","
	    lgStrSQL = lgStrSQL & FilterVar(Request("txtBS_PL_FG"),"''","S") & ","
	    lgStrSQL = lgStrSQL & FilterVar(arrColVal(1),"''","S") & ","
	    lgStrSQL = lgStrSQL & FilterVar(arrColVal(2),"''","S") & ","
	    lgStrSQL = lgStrSQL & FilterVar(arrColVal(3),"''","S") & ","
		lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(4), "0"),"0","D")     & ","
	    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ","
	    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ""
	       
	    lgStrSQL = lgStrSQL & "   ) " 

		PrintLog "SubBizSaveSingleCreate = " & lgStrSQL
	    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
        
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
		
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

    On Error Resume Next
    Err.Clear

    
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow
    Dim iDx , i

    'On Error Resume Next
    Err.Clear 

	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
	PrintLog "Spead count: " & Request("txtSpread")
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)    
        
	    lgStrSQL =            "UPDATE TB_3_3_3 WITH (ROWLOCK) " & vbCrLf
	    lgStrSQL = lgStrSQL & "   SET "  & vbCrLf
	    lgStrSQL = lgStrSQL & "       W3 = " & FilterVar(arrColVal(2),"''","S") & "," & vbCrLf
	    lgStrSQL = lgStrSQL & "       W4 = " & FilterVar(arrColVal(3),"''","S") & "," & vbCrLf
	    lgStrSQL = lgStrSQL & "       W5 = " & FilterVar(UNICDbl(arrColVal(4), "0"),"0","D") & "," & vbCrLf
	    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf
	    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S") & vbCrLf
	
	    lgStrSQL = lgStrSQL & " WHERE CO_CD = " &  FilterVar(Request("txtCO_CD_Body"),"''","S") & vbCrLf
	    lgStrSQL = lgStrSQL & " AND FISC_YEAR = " &  FilterVar(Request("txtFISC_YEAR_Body"),"''","S") & vbCrLf
	    lgStrSQL = lgStrSQL & " AND REP_TYPE = " &  FilterVar(Request("txtREP_TYPE_Body"),"''","S") & vbCrLf
	    lgStrSQL = lgStrSQL & " AND W1 = " &  FilterVar(Request("txtBS_PL_FG"),"''","S") & vbCrLf
	    lgStrSQL = lgStrSQL & " AND W2 = " & FilterVar(arrColVal(1),"''","S") & vbCrLf

		PrintLog "SubBizSaveSingleUpdate = " & lgStrSQL
	    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
        
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
		
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1,pCode2,pCode3)

    Select Case pMode 
      Case "R"


			lgStrSQL =			  " SELECT a.CO_CD, a.FISC_YEAR, a.REP_TYPE, a.W1, a.W2, C.PAR_GP_CD, SPACE((C.GP_LVL * 2)) + a.W3 AS W3, a.W4, a.W5, C.SUM_FG, C.GP_LVL " & vbCrLf
            lgStrSQL = lgStrSQL & "  FROM TB_3_3_3 a (NOLOCK)  " & vbCrLf
            lgStrSQL = lgStrSQL & "		INNER JOIN dbo.ufn_TB_ACCT_GP('" & C_REVISION_YM & "') c ON A.W1=C.BS_PL_FG AND A.W2=C.GP_CD" & vbCrLf
            lgStrSQL = lgStrSQL & "  WHERE a.CO_CD = " & pCode1 	 & vbCrLf
            lgStrSQL = lgStrSQL & "  AND a.FISC_YEAR = " & pCode2 	 & vbCrLf
            lgStrSQL = lgStrSQL & "  AND a.REP_TYPE = " & pCode3 	 & vbCrLf
            lgStrSQL = lgStrSQL & "  AND a.W1 = '3'" & vbCrLf
            'lgStrSQL = lgStrSQL & "  AND c.BS_PL_FG = a.W1 " & vbCrLf
            'lgStrSQL = lgStrSQL & "  AND c.GP_CD = a.W2 " & vbCrLf
            

            lgStrSQL = lgStrSQL & "  ORDER by C.GP_SEQ, a.W2 " & vbCrLf

    End Select
	PrintLog "SubMakeSQLStatements = " & lgStrSQL
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
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .DBQueryOk
	         End with
	      Else
	      	Parent.FncNew
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