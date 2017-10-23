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

	Const TYPE_1	= 0		' 그리드를 구분짓기 위한 상수 
	Const TYPE_2	= 1		
	Const TYPE_3	= 2		

	' -- 그리드 컬럼 정의 
	Dim	C_W1_CD
	Dim C_W1
	Dim C_W2_CD
	Dim C_W2
	Dim C_W3
	Dim C_W4
	Dim C_W7
	

	lgErrorStatus    = "NO"
    lgOpModeCRUD     = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    sFISC_YEAR		= Request("txtFISC_YEAR")
    sREP_TYPE		= Request("cboREP_TYPE")

    'lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    'lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    'lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    	
	Call InitSpreadPosVariables	' 그리드 위치 초기화 함수 

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
'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()	' 데이타 넘겨주는 컬럼 기준 
	C_W1_CD		= 1
	C_W1		= 2
	C_W2_CD		= 3
	C_W2		= 4
	C_W3		= 5
	C_W4		= 6
	C_W7		= 7

End Sub

'========================================================================================
Sub SubBizQuery()
    On Error Resume Next
    Err.Clear
End Sub

'========================================================================================
Sub SubBizDelete()
    'On Error Resume Next
    Err.Clear

	' 디테일부터 제거한다.
    lgStrSQL =            "DELETE TB_13_A WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
	lgStrSQL = lgStrSQL & "DELETE TB_13_B WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 
	

	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'========================================================================================
Sub SubBizQuery()
    Dim iDx, arrRs(3), iIntMaxRows, iLngCol
    Dim iRow, iKey1, iKey2, iKey3, sW_TYPE
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(wgCO_CD,"''", "S")		' 글로벌변수 컴퍼니코드 
    iKey2 = FilterVar(sFISC_YEAR,"''", "S")		' 사업연도 
    iKey3 = FilterVar(sREP_TYPE,"''", "S")		' 신고구분 
    
    
    
     Call SubMakeSQLStatements("Q" ,iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
       Call SetErrorStatus()
    Else 
				%>
				<Script Language=vbscript>
				       With Parent	
				                .Frm1.txtW1.Value		     = "<%=ConvSPChars(lgObjRs("w1"))%>"
				                .Frm1.txtW2.Value			 = "<%=ConvSPChars(lgObjRs("w2"))%>"
				                .Frm1.txtW2_VAL.Value		 = "<%=ConvSPChars(lgObjRs("w2_val"))%>"              
				                .Frm1.txtW3.Value			 = "<%=ConvSPChars(lgObjRs("w3"))%>"
				                .Frm1.txtW3_SUM.Value			 = "<%=ConvSPChars(lgObjRs("w3"))%>"
				                .Frm1.txtW4_A.Value			 = "<%=ConvSPChars(lgObjRs("w4_A"))%>"
				                .Frm1.txtW4_B.Value			 = "<%=ConvSPChars(lgObjRs("w4_B"))%>"
				                .Frm1.txtW4_Sum.Value		 = "<%=ConvSPChars(lgObjRs("W4"))%>"
				                .Frm1.txtW5_A.Value			 = "<%=ConvSPChars(lgObjRs("w5_A"))%>"
				                .Frm1.txtW5_A_VAL.Value		 = "<%=ConvSPChars(lgObjRs("w5_A_VAL"))%>"
				                .Frm1.txtW5_B.Value			 = "<%=ConvSPChars(lgObjRs("w5_B"))%>"
				                .Frm1.txtW5_B_VAL.Value		 = "<%=ConvSPChars(lgObjRs("w5_B_VAL"))%>"
				                .Frm1.txtW6_A.Value			 = "<%=ConvSPChars(lgObjRs("w6_A"))%>"
				                .Frm1.txtW6_B.Value			 = "<%=ConvSPChars(lgObjRs("w6_B"))%>"
				                .Frm1.txtW6_Sum.Value			 = "<%=ConvSPChars(lgObjRs("w6"))%>"
				                 .Frm1.txtW7.Value			 = "<%=ConvSPChars(lgObjRs("w7"))%>"


				              
				               
				       End With          
				</Script>       
				<%     
    End If



    Call SubMakeSQLStatements("R",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL, "", "") = False Then
  
         lgStrPrevKey = ""
        'Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
        
    Else
		iLngCol = lgObjRs.Fields.Count
		sW_TYPE = "" : lgstrData = ""
		iDx = 1

				lgstrData = lgstrData & " With parent.lgvspdData(0) " & vbCr
				lgstrData = lgstrData & "	.Redraw = false " & vbCr
				Do While Not lgObjRs.EOF
					lgstrData = lgstrData & "	.Row = " &iDx & "" & vbCrLf
					lgstrData = lgstrData & "	.Col = 0 : .value = """" " & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_W1_CD & " : .value = """ & lgObjRs("W1_CD") & """" & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_W1 & " : .value = """ & lgObjRs("W1") & """" & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_W2_CD & " : .value = """ & lgObjRs("W2_CD") & """" & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_W2 & " : .value = """ & lgObjRs("W2") & """" & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_W3 & " : .value = """ & lgObjRs("W3") & """" & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_W4 & " : .value = """ & lgObjRs("W4") & """" & vbCrLf
					lgstrData = lgstrData & "	.Col = " & C_W7 & " : .value = """ & lgObjRs("W_DES") & """" & vbCrLf
                    lgstrData = lgstrData & "	.Col = " & C_W7 + 1 & " : .value = """  &  iDx  & """" & vbCrLf

				If Err.number <> 0 Then
					PrintLog "iDx=" & iDx
					Exit Sub
				End If
		        iDx = iDx +1    
				lgObjRs.MoveNext
		
			Loop


		
		lgObjRs.Close
		Set lgObjRs = Nothing
			
		lgstrData = lgstrData & "	parent.lgIntFlgMode = parent.parent.OPMD_UMODE" & vbCrLf
    End If

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write lgstrData  &  vbCrLf
		
	If lgstrData <> "" Then	
		Response.Write "	.Redraw = True " & vbCr
		Response.Write " End With " & vbCrLf	' With 문 종료 
	End If
	
	If lgErrorStatus = "NO" Then
		Response.Write " Call parent.DbQueryOk                                      " & vbCr
	Else
		Response.Write " Call parent.SetDefaultVal                                     " & vbCr
	End If
	
	Response.Write " </Script>                                          " & vbCr
	    
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
	
	

	


End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "R"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & "   A.W1_CD, A.W1, A.W2_CD, A.W2, A.W3, A.W4, A.W_DES  "
            lgStrSQL = lgStrSQL & " FROM TB_13_A A WITH (NOLOCK) "
            lgStrSQL = lgStrSQL & " LEFT OUTER JOIN ufn_TB_CONFIGURATION('W1075', '" & C_REVISION_YM & "') B ON A.W2_CD = B.MINOR_CD "
			lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND A.REP_TYPE = " & pCode3 	 & vbCrLf
			lgStrSQL = lgStrSQL & "	ORDER BY A.W1_CD, B.REFERENCE_4"
	 Case "Q"
			lgStrSQL =			  " SELECT  "
            lgStrSQL = lgStrSQL & "   W1, W2, W2_VAL, W3, W4_A, W4_B , W4, W5_A,W5_A_VAL , W5_B, W5_B_VAL, W6_A, w6_B, W6 , W7  "
            lgStrSQL = lgStrSQL & " FROM TB_13_B WITH (NOLOCK) "
			lgStrSQL = lgStrSQL & " WHERE CO_CD = " & pCode1 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & pCode2 	 & vbCrLf
			lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & pCode3 	 & vbCrLf	

	
    End Select

End Sub


'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow
    Dim iDx , i


    Err.Clear 
     lgIntFlgMode = CInt(Request("txtFlgMode")) 
	
	' 신규입력 

	'PrintLog "txtSpread = " & Request("txtSpread")
			
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)
	
	For iDx = 1 To lgLngMaxRow

	    arrColVal = Split(arrRowVal(iDx-1), gColSep)    
			    
	    Select Case arrColVal(0)
	        Case "C"
	                Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
	        Case "U"
	                Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
	    End Select
			    
	    If lgErrorStatus    = "YES" Then
	       lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
	       Exit For
	    End If
			    
	Next

	if lgErrorStatus <> "YES"  then
	   select case lgIntFlgMode
	          case OPMD_CMODE
            	   Call SubBizSaveSingleCreate
              case OPMD_UMODE
            	   Call SubBizSaveSingleUpdate	   
        end select    	   
	end if
Response.Write "OPMD_CMODE = " & OPMD_CMODE
Response.Write "OPMD_UMODE = " & OPMD_UMODE
End Sub  

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
       On Error Resume Next
	dim iType

	lgStrSQL = "INSERT INTO TB_13_A WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE "  & vbCrLf 
    lgStrSQL = lgStrSQL & " , W1_CD, W1, W2_CD, W2, W3, W4, W_DES "  & vbCrLf 
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
	lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W1_CD))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W1))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W2_CD))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W2))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W3))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W7))),"''","S")      & "," & vbCrLf

	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
	lgStrSQL = lgStrSQL & ")"
	
	



	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
    Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
 

End Sub




'============================================================================================================
' Name : SubBizSaveSingleCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()


	dim iType


	
	
	lgStrSQL =   "INSERT INTO TB_13_B WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE "  & vbCrLf 
    lgStrSQL = lgStrSQL & " , W1, W2, W2_Val, W3, W4_A, W4_B, W4,  W5_A,  W5_A_VAL,  W5_B,  W5_B_VAL, w6_a  , w6_b  ,w6  , w7   "  & vbCrLf 
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID )"  & vbCrLf
	lgStrSQL = lgStrSQL & " VALUES ("  & vbCrLf
    
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtw1"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(Request("txtw2"))),"''","S")      & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtw2_VAL"), "0"),"0","D")       & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtw3"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtw4_A"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtw4_B"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtw4_SUM"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(Request("txtw5_A"))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL &  FilterVar(UNICDbl(Request("txtw5_A_VAL"), "0"),"0","D")      & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(Request("txtw5_B"))),"''","S")     & "," & vbCrLf
	lgStrSQL = lgStrSQL &  FilterVar(UNICDbl(Request("txtw5_B_VAL"), "0"),"0","D")      & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtw6_A"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtw6_B"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtw6_SUM"), "0"),"0","D")     & "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(Request("txtw7"), "0"),"0","D")     & "," & vbCrLf
	
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," & vbCrLf 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","  & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")             & vbCrLf            
	lgStrSQL = lgStrSQL & ")"



	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
		

End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
	dim i

	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_13_B WITH (ROWLOCK) "
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W1        = "  & FilterVar(UNICDbl(Request("txtw1"), "0"),"0","D")  	& "," & vbCrLf
    lgStrSQL = lgStrSQL & " W2        = "  & FilterVar(Trim(UCase(Request("txtw2"))),"''","S")   	& "," & vbCrLf
    lgStrSQL = lgStrSQL & " W2_Val    = "  & FilterVar(UNICDbl(Request("txtw2_VAL"), "0"),"0","D") 	& "," & vbCrLf
    lgStrSQL = lgStrSQL & " W3        = "  & FilterVar(UNICDbl(Request("txtw3"), "0"),"0","D")   	& "," & vbCrLf
    lgStrSQL = lgStrSQL & " W4_A      = "  & FilterVar(UNICDbl(Request("txtw4_A"), "0"),"0","D")  	& "," & vbCrLf
    lgStrSQL = lgStrSQL & " W4_B      = "  & FilterVar(UNICDbl(Request("txtw4_B"), "0"),"0","D") 	& "," & vbCrLf
    lgStrSQL = lgStrSQL & " W4        = "  & FilterVar(UNICDbl(Request("txtw4_SUM"), "0"),"0","D") 	& "," & vbCrLf
    lgStrSQL = lgStrSQL & " W5_A      = "  & FilterVar(Trim(UCase(Request("txtw5_A"))),"''","S")  	& "," & vbCrLf
    lgStrSQL = lgStrSQL & " W5_A_VAL  = "  &  FilterVar(UNICDbl(Request("txtw5_A_VAL"), "0"),"0","D") 	& "," & vbCrLf
    lgStrSQL = lgStrSQL & " W5_B      = "  & FilterVar(Trim(UCase(Request("txtw5_B"))),"''","S")  	& "," & vbCrLf
    lgStrSQL = lgStrSQL & " W5_B_VAL   = " &  FilterVar(UNICDbl(Request("txtw5_B_VAL"), "0"),"0","D")   	& "," & vbCrLf
    lgStrSQL = lgStrSQL & " W6_A      = "  & FilterVar(UNICDbl(Request("txtw6_A"), "0"),"0","D")  	& "," & vbCrLf
    lgStrSQL = lgStrSQL & " W6_B      = "  & FilterVar(UNICDbl(Request("txtw6_B"), "0"),"0","D")	& "," & vbCrLf
    lgStrSQL = lgStrSQL & " W6        = "  & FilterVar(UNICDbl(Request("txtw6_SUM"), "0"),"0","D") 	& "," & vbCrLf
    lgStrSQL = lgStrSQL & " W7        = "  & FilterVar(UNICDbl(Request("txtw7"), "0"),"0","D")  	& "," & vbCrLf

    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 

	

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub



'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_13_A WITH (ROWLOCK) "
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W2        = " &  FilterVar(Trim(UCase(arrColVal(C_W2))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W3        = " &  FilterVar(Trim(UCase(arrColVal(C_W3))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W4        = " &  FilterVar(UNICDbl(arrColVal(C_W4), "0"),"0","D") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " W_DES        = " &  FilterVar(Trim(UCase(arrColVal(C_W7))),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf                  
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND W1_CD = " & FilterVar(Trim(UCase(arrColVal(C_W1_CD))),"''","S") 	 & vbCrLf 
	lgStrSQL = lgStrSQL & "		AND W2_CD = " & FilterVar(Trim(UCase(arrColVal(C_W2_CD))),"''","S") 	 & vbCrLf 

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


Function RemovePercent(Byval pVal)
	RemovePercent = Replace(pVal, "%", "")
End Function
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
<%
'   **************************************************************
'	1.4 Transaction 처러 이벤트 
'   **************************************************************

Sub	onTransactionCommit()
	' 트랜잭션 완료후 이벤트 처리 
End Sub

Sub onTransactionAbort()
	' 트랜잭선 실패(에러)후 이벤트 처리 
'PrintForm
'	' 에러 출력 
	'Call SaveErrorLog(Err)	' 에러로그를 남긴 
	
End Sub
%>
