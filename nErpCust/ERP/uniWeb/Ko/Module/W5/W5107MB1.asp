<%@ Transaction=required Language=VBScript%>
<%Option Explicit%> 
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
   
	Const C_SHEETMAXROWS_D = 100
	Dim sFISC_YEAR, sREP_TYPE
	Dim lgStrPrevKey

	' -- 그리드 컬럼 정의 
	Dim C_W_YEAR
	Dim C_W_TYPE
	Dim C_W_NAME
	Dim C_W26
	Dim C_W27
	Dim C_W28
	Dim C_W29

	lgErrorStatus   = "NO"
    lgOpModeCRUD    = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    sFISC_YEAR		= Request("txtFISC_YEAR")
    sREP_TYPE		= Request("cboREP_TYPE")


    lgLngMaxRow     = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey	= UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    	
	Call InitSpreadPosVariables	' 그리드 위치 초기화 함수 

    Call SubOpenDB(lgObjConn) 
    
    Call CheckVersion(sFISC_YEAR, sREP_TYPE)	' 2005-03-11 버전관리기능 추가 
    	
    Select Case lgOpModeCRUD 
        Case CStr(UID_M0001)                                                         '☜: Query
			Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
			Call SubBizSave()
			Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
			Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)
    
'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()

	C_W_YEAR	= 1
	C_W_TYPE	= 2
	C_W_NAME	= 3
	C_W26		= 4
	C_W27		= 5
	C_W28		= 6
	C_W29		= 7
	
End Sub

'========================================================================================
'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next
    Err.Clear
  
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
 PrintLog "lgIntFlgMode = " & lgIntFlgMode & ";" &  OPMD_CMODE & ";" & OPMD_UMODE
    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜ : Create
              Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE           
              Call SubBizSaveSingleUpdate()
    End Select
End Sub	

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    On Error Resume Next
    Err.Clear

	lgStrSQL = "INSERT INTO TB_21H WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, W1, W2, W3, W4, W5, W6, W7, W8, W9, W10 "  
	lgStrSQL = lgStrSQL & " , W11, W12, W13, W14, W15, W16, W17, W18, W19, W20, W21, W22, W23, W24, W25, W_R1 "  
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")				& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW1"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW2"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW3"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW4"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW5"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW6"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW7"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW8"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW9"), "0"),"0","D")		& ","
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
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW21"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW22"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW23"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW24"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW25"), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtW_R1"), "0"),"0","D")		& ","

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveSingleCreate = " & lgStrSQL

	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    On Error Resume Next
    Err.Clear

    lgStrSQL =            "UPDATE TB_21H WITH (ROWLOCK) " & vbCrLf
    lgStrSQL = lgStrSQL & "   SET "  & vbCrLf

	lgStrSQL = lgStrSQL & "       W1 = " & FilterVar(UNICDbl(Request("txtW1"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W2 = " & FilterVar(UNICDbl(Request("txtW2"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W3 = " & FilterVar(UNICDbl(Request("txtW3"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W4 = " & FilterVar(UNICDbl(Request("txtW4"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W5 = " & FilterVar(UNICDbl(Request("txtW5"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W6 = " & FilterVar(UNICDbl(Request("txtW6"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W7 = " & FilterVar(UNICDbl(Request("txtW7"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W8 = " & FilterVar(UNICDbl(Request("txtW8"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W9 = " & FilterVar(UNICDbl(Request("txtW9"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W10 = " & FilterVar(UNICDbl(Request("txtW10"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W11 = " & FilterVar(UNICDbl(Request("txtW11"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W12 = " & FilterVar(UNICDbl(Request("txtW12"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W13 = " & FilterVar(UNICDbl(Request("txtW13"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W14 = " & FilterVar(UNICDbl(Request("txtW14"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W15 = " & FilterVar(UNICDbl(Request("txtW15"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W16 = " & FilterVar(UNICDbl(Request("txtW16"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W17 = " & FilterVar(UNICDbl(Request("txtW17"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W18 = " & FilterVar(UNICDbl(Request("txtW18"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W19 = " & FilterVar(UNICDbl(Request("txtW19"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W20 = " & FilterVar(UNICDbl(Request("txtW20"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W21 = " & FilterVar(UNICDbl(Request("txtW21"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W22 = " & FilterVar(UNICDbl(Request("txtW22"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W23 = " & FilterVar(UNICDbl(Request("txtW23"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W24 = " & FilterVar(UNICDbl(Request("txtW24"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W25 = " & FilterVar(UNICDbl(Request("txtW25"), "0"),"0","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W_R1 = " & FilterVar(UNICDbl(Request("txtW_R1"), "0"),"0","D")		& "," & vbCrLf

    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(wgCO_CD,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(sFISC_YEAR,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(sREP_TYPE,"''","S") 	 & vbCrLf

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords

	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
End Sub



'========================================================================================
Sub SubBizDelete()
    'On Error Resume Next
    Err.Clear

	' 디테일부터 제거한다.
    lgStrSQL =            "DELETE TB_21D WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

    lgStrSQL =            "DELETE TB_21H WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

PrintLog "SubMakeSQLStatements = " & lgStrSQL 
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
 	
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iKey1, iKey2, iKey3, iIntMaxRows, iLngRow
    Dim iDx
    Dim iLoopMax
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")	' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 

    Call SubMakeSQLStatements("R1",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
         lgStrPrevKey = ""
        'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else
%>
<Script Language=vbscript>
       With Parent	
                .Frm1.txtW1.Text  = "<%=ConvSPChars(lgObjRs("W1"))%>"
                .Frm1.txtW2.Text  = "<%=ConvSPChars(lgObjRs("W2"))%>"
                .Frm1.txtW3.Text  = "<%=ConvSPChars(lgObjRs("W3"))%>"
                .Frm1.txtW4.Text  = "<%=ConvSPChars(lgObjRs("W4"))%>"
                .Frm1.txtW5.Text  = "<%=ConvSPChars(lgObjRs("W5"))%>"
                .Frm1.txtW6.Text  = "<%=ConvSPChars(lgObjRs("W6"))%>"
                .Frm1.txtW7.Text  = "<%=ConvSPChars(lgObjRs("W7"))%>"
                .Frm1.txtW8.Text  = "<%=ConvSPChars(lgObjRs("W8"))%>"
                .Frm1.txtW9.Text  = "<%=ConvSPChars(lgObjRs("W9"))%>"
                .Frm1.txtW10.Text  = "<%=ConvSPChars(lgObjRs("W10"))%>"
                .Frm1.txtW11.Text  = "<%=ConvSPChars(lgObjRs("W11"))%>"
                .Frm1.txtW12.Text  = "<%=ConvSPChars(lgObjRs("W12"))%>"
                .Frm1.txtW13.Text  = "<%=ConvSPChars(lgObjRs("W13"))%>"
                .Frm1.txtW14.Text  = "<%=ConvSPChars(lgObjRs("W14"))%>"
                .Frm1.txtW15.Text  = "<%=ConvSPChars(lgObjRs("W15"))%>"
                .Frm1.txtW16.Text  = "<%=ConvSPChars(lgObjRs("W16"))%>"
                .Frm1.txtW17.Text  = "<%=ConvSPChars(lgObjRs("W17"))%>"

                .Frm1.txtW18.Text  = "<%=ConvSPChars(lgObjRs("W18"))%>"
                .Frm1.txtW19.Text  = "<%=ConvSPChars(lgObjRs("W19"))%>"
                .Frm1.txtW20.Text  = "<%=ConvSPChars(lgObjRs("W20"))%>"
                .Frm1.txtW21.Text  = "<%=ConvSPChars(lgObjRs("W21"))%>"
                .Frm1.txtW22.Text  = "<%=ConvSPChars(lgObjRs("W22"))%>"
                .Frm1.txtW23.Text  = "<%=ConvSPChars(lgObjRs("W23"))%>"
                .Frm1.txtW24.Text  = "<%=ConvSPChars(lgObjRs("W24"))%>"
                .Frm1.txtW25.Text  = "<%=ConvSPChars(lgObjRs("W25"))%>"

                .Frm1.txtW_R1.Value  = "<%=ConvSPChars(lgObjRs("W_R1"))%>"

       End With          
</Script>       
<%     

	    Call SubMakeSQLStatements("R2",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements
	
	    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then

	        iStrData = ""
	        
	        iDx = 1
	        Do While Not lgObjRs.EOF

				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W_YEAR"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W_TYPE"))
				lgstrData = lgstrData & Chr(11)
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W26"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W27"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W28"))			
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("W29"))
				lgstrData = lgstrData & Chr(11) & iDx
				lgstrData = lgstrData & Chr(11) & Chr(12)
	
				iDx = iDx + 1
			    lgObjRs.MoveNext
	
	        Loop 
	    End If

    End If
    

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "R1"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.W1, A.W2, A.W3, A.W4, A.W5, A.W6, A.W7, A.W8, A.W9, A.W10 "
            lgStrSQL = lgStrSQL & " , A.W11, A.W12, A.W13, A.W14, A.W15, A.W16, A.W17, A.W18, A.W19, A.W20, A.W21, A.W22, A.W23, A.W24, A.W25, A.W_R1 "
            lgStrSQL = lgStrSQL & " FROM TB_21H A WITH (NOLOCK) " & vbCrLf

            If pCode1 = "''" Then
            	lgStrSQL = lgStrSQL & " WHERE A.CO_CD <> " & pCode1 	 & vbCrLf
            Else
            	lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
            End If

            If pCode2 <> "''" Then
            	lgStrSQL = lgStrSQL & " AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
            End If

            If pCode3 <> "''" Then
            	lgStrSQL = lgStrSQL & " AND A.REP_TYPE = " & pCode3 	 & vbCrLf
            End If

'            lgStrSQL = lgStrSQL & " ORDER BY  A.SEQ_NO ASC" & vbcrlf
                        
      Case "R2"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & " A.W_YEAR, A.W_TYPE, A.W26, A.W27, A.W28, A.W29 "
            lgStrSQL = lgStrSQL & " FROM TB_21D A WITH (NOLOCK) " & vbCrLf

            If pCode1 = "''" Then
            	lgStrSQL = lgStrSQL & " WHERE A.CO_CD <> " & pCode1 	 & vbCrLf
            Else
            	lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
            End If

            If pCode2 <> "''" Then
            	lgStrSQL = lgStrSQL & " AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
            End If

            If pCode3 <> "''" Then
            	lgStrSQL = lgStrSQL & " AND A.REP_TYPE = " & pCode3 	 & vbCrLf
            End If

            lgStrSQL = lgStrSQL & " ORDER BY  A.W_YEAR ASC, A.W_TYPE" & vbcrlf
                        
    End Select
	PrintLog "SubMakeSQLStatements = " & lgStrSQL
End Sub


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow
    Dim iDx , i

	PrintLog "SubBizSaveMulti.."
	If lgErrorStatus    = "YES" Then Exit Sub
	
    'On Error Resume Next
    Err.Clear 
    
	PrintLog "1번째 그리드. .: " & Request("txtSpread") 
	' --- 1번째 그리드 
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)    
        
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
' Desc : 1번째 그리드 저장 
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i
	
	lgStrSQL = "INSERT INTO TB_21D WITH (ROWLOCK) ("
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, W_YEAR, W_TYPE, W26, W27, W28, W29 "  
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES (" 
	    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")				& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W_YEAR))),"''","S")	& ","
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W_TYPE))),"''","S")	& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W26), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W27), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W28), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W29), "0"),"0","D")		& ","

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate = " & lgStrSQL

	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : 1번 그리드 업데이트 
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_21D WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    
	lgStrSQL = lgStrSQL & " W26		= " &  FilterVar(UNICDbl(arrColVal(C_W26), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & " W27		= " &  FilterVar(UNICDbl(arrColVal(C_W27), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & " W28		= " &  FilterVar(UNICDbl(arrColVal(C_W28), "0"),"0","D")		& ","
	lgStrSQL = lgStrSQL & " W29		= " &  FilterVar(UNICDbl(arrColVal(C_W29), "0"),"0","D")		& ","
                  
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(wgCO_CD,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(sFISC_YEAR,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(sREP_TYPE,"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND W_YEAR = " & FilterVar(Trim(UCase(arrColVal(C_W_YEAR))),"''","S")  
	lgStrSQL = lgStrSQL & "		AND W_TYPE = " & FilterVar(Trim(UCase(arrColVal(C_W_TYPE))),"''","S")  

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	Response.Write lgStrSQL & "<br>" & vbCrLf
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : 1번째 그리드 삭제 
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_21D WITH (ROWLOCK) "
 	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND W_YEAR = " & FilterVar(Trim(UCase(arrColVal(C_W_YEAR))),"''","S")  
	lgStrSQL = lgStrSQL & "		AND W_TYPE = " & FilterVar(Trim(UCase(arrColVal(C_W_TYPE))),"''","S")  
	
	PrintLog "SubBizSaveMultiDelete = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

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
        Case "SC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "SD"
        Case "SR"
        Case "SU"
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
<%				If lgstrData <> "" Then %>
                .ggoSpread.SSShowData "<%=lgstrData%>"   
<%				End If %>             
                .DBQueryOk        
	         End with
          Else
			Call Parent.FncNew()	' -- 데이타비존재시신규호출
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
       
</Script>t                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         