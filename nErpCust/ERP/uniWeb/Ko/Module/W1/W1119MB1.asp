<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<%Option Explicit%> 
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
    
    'On Error Resume Next
    Err.Clear
   
	Dim sCO_CD, sFISC_YEAR, sREP_TYPE
	Dim lgStrPrevKey

	lgErrorStatus    = "NO"
    lgOpModeCRUD     = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    sCO_CD			= Request("txtCO_CD")
    sFISC_YEAR		= Request("txtFISC_YEAR")
    sREP_TYPE		= Request("cboREP_TYPE")

    'lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    'lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    'lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    Call SubOpenDB(lgObjConn) 
    
    Call CheckVersion(sFISC_YEAR, sREP_TYPE)	' 2005-03-11 버전관리기능 추가 
    	
    Select Case lgOpModeCRUD 
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)
    
'========================================================================================
Sub SubBizSave()

'	PrintLog "SubBizSave.."
	
    'On Error Resume Next
    Err.Clear 
    
    ' 헤더 저장 
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
 
    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜ : Create
              Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE           
              Call SubBizSaveSingleUpdate()
    End Select
    
End Sub
'========================================================================================
Sub SubBizDelete()
    'On Error Resume Next
    Err.Clear

    lgStrSQL =            "DELETE TB_3_3_4 WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(sCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

'PrintLog "SubBizDelete = " & lgStrSQL 
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
 	
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iKey1, iKey2, iKey3, iStrData, iStrData2, iIntMaxRows, iLngRow
    Dim iDx
    Dim iLoopMax
    
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(sCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

    Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
  
         lgStrPrevKey = ""
        'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
        Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.DbQueryFalse                                   " & vbCr
	    Response.Write " End With                                           " & vbCr
	    Response.Write " </Script>                                          " & vbCr
        
    Else
        lgstrData = ""
        
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr

        If lgObjRs("W_TYPE") = "1" Then
		    Response.Write "	.frm1.txtW_TYPE.value = ""1""" & vbCr
		    Response.Write "	.frm1.txtW1.value = """ & ConvSPChars(lgObjRs("W1"))          & """" & vbCr
		    Response.Write "	.frm1.txtW2.value = """ & ConvSPChars(lgObjRs("W2"))          & """" & vbCr
		    Response.Write "	.frm1.txtW3.value = """ & ConvSPChars(lgObjRs("W3"))          & """" & vbCr
		    Response.Write "	.frm1.txtW4.value = """ & ConvSPChars(lgObjRs("W4"))          & """" & vbCr
		    Response.Write "	.frm1.txtW5.value = """ & ConvSPChars(lgObjRs("W5"))          & """" & vbCr
		    Response.Write "	.frm1.txtW6.value = """ & ConvSPChars(lgObjRs("W6"))          & """" & vbCr
		    Response.Write "	.frm1.txtW8.value = """ & ConvSPChars(lgObjRs("W8"))          & """" & vbCr
		    Response.Write "	.frm1.txtW10.value = """ & ConvSPChars(lgObjRs("W10"))          & """" & vbCr
		    Response.Write "	.frm1.txtW11.value = """ & ConvSPChars(lgObjRs("W11"))          & """" & vbCr
		    Response.Write "	.frm1.txtW12.value = """ & ConvSPChars(lgObjRs("W12"))          & """" & vbCr
		    Response.Write "	.frm1.txtW13.value = """ & ConvSPChars(lgObjRs("W13"))          & """" & vbCr
		    Response.Write "	.frm1.txtW14.value = """ & ConvSPChars(lgObjRs("W14"))          & """" & vbCr
		    Response.Write "	.frm1.txtW15.value = """ & ConvSPChars(lgObjRs("W15"))          & """" & vbCr
		    Response.Write "	.frm1.txtW16.value = """ & ConvSPChars(lgObjRs("W16"))          & """" & vbCr
		    Response.Write "	.frm1.txtW17.value = """ & ConvSPChars(lgObjRs("W17"))          & """" & vbCr
		    Response.Write "	.frm1.txtW18.value = """ & ConvSPChars(lgObjRs("W18"))          & """" & vbCr
		    Response.Write "	.frm1.txtW19.value = """ & ConvSPChars(lgObjRs("W19"))          & """" & vbCr
		    Response.Write "	.frm1.txtW20.value = """ & ConvSPChars(lgObjRs("W20"))          & """" & vbCr
		    Response.Write "	.frm1.txtW25.value = """ & ConvSPChars(lgObjRs("W25"))          & """" & vbCr
		    
		    ' -- 2006.03 개정 
		    Response.Write "	.frm1.txtW26.value = """ & ConvSPChars(lgObjRs("W26"))          & """" & vbCr
		    Response.Write "	.frm1.txtW27.value = """ & ConvSPChars(lgObjRs("W27"))          & """" & vbCr
		    Response.Write "	.frm1.txtW28.value = """ & ConvSPChars(lgObjRs("W28"))          & """" & vbCr
		    
		    Response.Write "	.frm1.txtRW1.value = """ & ConvSPChars(lgObjRs("W2"))          & """" & vbCr
		    Response.Write "	.frm1.txtRW2.value = """ & ConvSPChars(lgObjRs("W6"))          & """" & vbCr
		Else
		    Response.Write "	.frm1.txtW_TYPE.value = ""2""" & vbCr
		    Response.Write "	.frm1.txtW30.value = """ & ConvSPChars(lgObjRs("W30"))          & """" & vbCr
		    Response.Write "	.frm1.txtW31.value = """ & ConvSPChars(lgObjRs("W31"))          & """" & vbCr
		    Response.Write "	.frm1.txtW32.value = """ & ConvSPChars(lgObjRs("W32"))          & """" & vbCr
		    Response.Write "	.frm1.txtW33.value = """ & ConvSPChars(lgObjRs("W33"))          & """" & vbCr
		    Response.Write "	.frm1.txtW34.value = """ & ConvSPChars(lgObjRs("W34"))          & """" & vbCr
		    Response.Write "	.frm1.txtW35.value = """ & ConvSPChars(lgObjRs("W35"))          & """" & vbCr
		    Response.Write "	.frm1.txtW40.value = """ & ConvSPChars(lgObjRs("W40"))          & """" & vbCr
		    Response.Write "	.frm1.txtW41.value = """ & ConvSPChars(lgObjRs("W41"))          & """" & vbCr
		    Response.Write "	.frm1.txtW42.value = """ & ConvSPChars(lgObjRs("W42"))          & """" & vbCr
		    Response.Write "	.frm1.txtW43.value = """ & ConvSPChars(lgObjRs("W43"))          & """" & vbCr
		    Response.Write "	.frm1.txtW44.value = """ & ConvSPChars(lgObjRs("W44"))          & """" & vbCr
		    Response.Write "	.frm1.txtW50.value = """ & ConvSPChars(lgObjRs("W50"))          & """" & vbCr
		    Response.Write "	.frm1.txtRW1.value = """ & ConvSPChars(lgObjRs("W31"))          & """" & vbCr
		    Response.Write "	.frm1.txtRW2.value = """ & ConvSPChars(unicdbl(lgObjRs("W35"), "0") * -1)          & """" & vbCr
		End If
		Response.Write "	.frm1.txtW_DT.text = """ & ConvSPChars(lgObjRs("W_DT"))          & """" & vbCr
		    
		Response.Write "	.DbQueryOk                                      " & vbCr
	    Response.Write " End With                                           " & vbCr
	    Response.Write " </Script>                                          " & vbCr
        
        lgObjRs.Close
        Set lgObjRs = Nothing
 
    End If
    
     Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
     
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "R"
			lgStrSQL =			  " SELECT  "	 & vbCrLf
            lgStrSQL = lgStrSQL & "  W_TYPE, W1, W2, W3, W4, W5, W6, W8, W10 "	 & vbCrLf
            lgStrSQL = lgStrSQL & " , W11, W12, W13, W14, W15, W16, W17, W18, W19, W20, W25, W26, W27, W28 "	 & vbCrLf
            lgStrSQL = lgStrSQL & " , W30, W31, W32, W33, W34, W35, W40, W41, W42, W43, W44, W50, W_DT "	 & vbCrLf

            lgStrSQL = lgStrSQL & " FROM TB_3_3_4 WITH (NOLOCK) "	 & vbCrLf
			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & pCode3 	 & vbCrLf
            
    End Select
'	PrintLog "SubMakeSQLStatements = " & lgStrSQL
End Sub


'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    'On Error Resume Next
    Err.Clear

    lgStrSQL =            " INSERT INTO TB_3_3_4 WITH (ROWLOCK) "
    lgStrSQL = lgStrSQL & " (CO_CD, FISC_YEAR, REP_TYPE, W_TYPE, W1, W2, W3, W4, W5, W6, W8, W10 "
    lgStrSQL = lgStrSQL & " , W11, W12, W13, W14, W15, W16, W17, W18, W19, W20, W25, W26, W27, W28 "
    lgStrSQL = lgStrSQL & " , W30, W31, W32, W33, W34, W35, W40, W41, W42, W43, W44, W50, W_DT "
    lgStrSQL = lgStrSQL & "  , INSRT_USER_ID, UPDT_USER_ID ) " 
    lgStrSQL = lgStrSQL & " VALUES ( " 
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sCO_CD)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","    
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW_TYPE"),"''","S") & ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW1"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW2"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW3"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW4"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW5"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW6"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW8"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW10"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW11"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW12"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW13"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW14"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW15"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW16"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW17"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW18"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW19"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW20"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW25"), "0"),"0","D")		& ","
	
	' -- 2006.03 개정 
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW26"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW27"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW28"), "0"),"0","D")		& ","
	
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW30"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW31"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW32"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW33"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW34"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW35"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW40"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW41"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW42"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW43"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW44"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW50"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(Request("txtW_DT"),"''","S") & ","
	
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ""
       
    lgStrSQL = lgStrSQL & "   ) " 

	PrintLog "SubBizSaveSingleCreate = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub   

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    'On Error Resume Next
    Err.Clear

    lgStrSQL =            "UPDATE TB_3_3_4 WITH (ROWLOCK) " & vbCrLf
    lgStrSQL = lgStrSQL & "   SET "  & vbCrLf
    lgStrSQL = lgStrSQL & "       W_TYPE = " & FilterVar(Request("txtW_TYPE"),"''","S") & ","
	lgStrSQL = lgStrSQL & "       W1 = " & FilterVar(UNICDbl(Request("txtW1"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W2 = " & FilterVar(UNICDbl(Request("txtW2"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W3 = " & FilterVar(UNICDbl(Request("txtW3"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W4 = " & FilterVar(UNICDbl(Request("txtW4"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W5 = " & FilterVar(UNICDbl(Request("txtW5"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W6 = " & FilterVar(UNICDbl(Request("txtW6"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W8 = " & FilterVar(UNICDbl(Request("txtW8"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W10 = " & FilterVar(UNICDbl(Request("txtW10"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W11 = " & FilterVar(UNICDbl(Request("txtW11"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W12 = " & FilterVar(UNICDbl(Request("txtW12"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W13 = " & FilterVar(UNICDbl(Request("txtW13"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W14 = " & FilterVar(UNICDbl(Request("txtW14"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W15 = " & FilterVar(UNICDbl(Request("txtW15"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W16 = " & FilterVar(UNICDbl(Request("txtW16"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W17 = " & FilterVar(UNICDbl(Request("txtW17"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W18 = " & FilterVar(UNICDbl(Request("txtW18"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W19 = " & FilterVar(UNICDbl(Request("txtW19"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W20 = " & FilterVar(UNICDbl(Request("txtW20"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W25 = " & FilterVar(UNICDbl(Request("txtW25"), "0"),"0","D")		& ","
	' -- 2006.03 개정 
	lgStrSQL = lgStrSQL & "       W26 = " & FilterVar(UNICDbl(Request("txtW26"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W27 = " & FilterVar(UNICDbl(Request("txtW27"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W28 = " & FilterVar(UNICDbl(Request("txtW28"), "0"),"0","D")		& ","
	
	lgStrSQL = lgStrSQL & "       W30 = " & FilterVar(UNICDbl(Request("txtW30"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W31 = " & FilterVar(UNICDbl(Request("txtW31"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W32 = " & FilterVar(UNICDbl(Request("txtW32"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W33 = " & FilterVar(UNICDbl(Request("txtW33"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W34 = " & FilterVar(UNICDbl(Request("txtW34"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W35 = " & FilterVar(UNICDbl(Request("txtW35"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W40 = " & FilterVar(UNICDbl(Request("txtW40"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W41 = " & FilterVar(UNICDbl(Request("txtW41"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W42 = " & FilterVar(UNICDbl(Request("txtW42"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W43 = " & FilterVar(UNICDbl(Request("txtW43"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W44 = " & FilterVar(UNICDbl(Request("txtW44"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & "       W50 = " & FilterVar(UNICDbl(Request("txtW50"), "0"),"0","D")		& ","
    lgStrSQL = lgStrSQL & "       W_DT = " & FilterVar(Request("txtW_DT"),"''","S") & ","
	lgStrSQL = lgStrSQL & "		  UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & "		  UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  
	
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(sCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
 
	PrintLog "SubBizSaveSingleUpdate = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
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

'========================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    'On Error Resume Next
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