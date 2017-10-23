	<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/AdoVbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDB.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd                                                               '☜: Hide Processing message
Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "P", "NOCOOKIE", "RB")

'---------------------------------------Common-----------------------------------------------------------
lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
lgStrPrevKeyIndex = CInt(Request("lgStrPrevKeyIndex"))   
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '☜: Set to space
lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
'------ Developer Coding part (Start ) ------------------------------------------------------------------
Dim IntRetCD
	
Dim strPlantCd
Dim strItemCd
Dim strBomNo
Dim strBaseDt
Dim strExpFlg

Dim strItemNm
Dim strItemAcct
Dim strProcType
Dim strItemAcctNm
Dim strProcTypeNm
Dim strSpec
Dim strBasicUnit
Dim BaseDt
Dim idx
Dim TmpBuffer
Dim iTotalStr
	
Dim QueryType
Dim strSpId
Dim strLevel

ReDim TmpBuffer(0)

strPlantCd = FilterVar(Trim(Request("txtPlantCd")) ,"''", "S")
strItemCd = FilterVar(Trim(Request("txtItemCd")) ,"''", "S")
strBomNo = FilterVar(Trim(Request("txtBomNo")) ,"''", "S")
BaseDt = FilterVar(UNIConvYYYYMMDDToDate(gAPDateFormat,"1900","01","01"),"''","S")
strBaseDt = FilterVar(Trim(Request("txtBaseDt")), BaseDt, "D")
strExpFlg = FilterVar(Trim(Request("rdoSrchType"))	,"''", "S")

Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

Call SubBizQuery("P_CK")										
Call SubBizQuery("CK")				
Call SubBizQuery("B_CK")
Call SubCreateCommandObject(lgObjComm)
Call SubBizBatch()
Call SubBizQueryMulti()
			
Call SubCloseCommandObject(lgObjComm)
			
Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection
		
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery(pOpCode)
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	Select Case pOpCode
		
		Case "P_CK"
			'--------------
			'공장 체크		
			'--------------	
			lgStrSQL = ""
			Call SubMakeSQLStatements("P_CK", strPlantCd, "", "", "", "")           '☜ : Make sql statements
			
			If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists
		    
				IntRetCD = -1
				Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.txtPlantNm.Value  = """"" & vbCrLf   'Set condition area
					Response.Write "parent.txtItemCd.focus" & vbCrLf   'Set condition area
				Response.Write "</Script>" & vbcRLf
				Response.End
		    Else
				IntRetCD = 1
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.txtPlantNm.Value = """ & ConvSPChars(lgObjRs(1)) & """" & vbCrLf 'Set condition area
				Response.Write "</Script>" & vbcRLf
			End If
		
			Call SubCloseRs(lgObjRs) 
			
		Case "B_CK"

			'------------------
			' bom type 체크 
			'------------------
			lgStrSQL = ""
			
			Call SubMakeSQLStatements("BT_CK", strBomNo, "", "", "", "")           '☜ : Make sql statements
			
		    If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists
		    
				IntRetCD = -1
				Call DisplayMsgBox("182622", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()
				
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.hBomType.value = """"" & vbCrLf
					Response.Write "parent.txtBomNo.focus" & vbCrLf
				Response.Write "</Script>" & vbCrLf
				Response.End
		    Else
				IntRetCD = 1
			End If
		
			Call SubCloseRs(lgObjRs) 

			'------------------
			'품목, bom no 체크 
			'------------------
			lgStrSQL = ""
			Call SubMakeSQLStatements("B_CK", strPlantCd, strItemCd, strBomNo, "", "")           '☜ : Make sql statements
			
		    If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists
		    
				IntRetCD = -1
				
				Call DisplayMsgBox("182600", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.hBomType.value = """"" & vbCrLf
					Response.Write "Call parent.DbQueryNotOk()" & vbCrLf
				Response.Write "</Script>" & vbCrLf
				Response.End
		    Else
				IntRetCD = 1
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "With Parent" & vbCrLf
						Response.Write ".txtItemNm.Value = """ & ConvSPChars(Trim(lgObjRs(14))) & """" & vbCrLf
						Response.Write ".txtItemAcctNm.value = """ & ConvSPChars(lgObjRs(16)) & """" & vbCrLf
						Response.Write ".txtItemAcct.value = """ & ConvSPChars(Trim(lgObjRs(12))) & """" & vbCrLf
						Response.Write ".txtSpec.value = """ & ConvSPChars(lgObjRs(15)) & """" & vbCrLf
						Response.Write ".txtPlantItemFromDt.text = """ & UNIDateClientFormat(lgObjRs(19)) & """" & vbCrLf
						Response.Write ".txtPlantItemToDt.text = """ & UNIDateClientFormat(lgObjRs(20)) & """" & vbCrLf
						Response.Write ".txtBasicUnit.value = """ & ConvSPChars(Trim(lgObjRs(18))) & """" & vbCrLf
						Response.Write ".txtBOMDesc.value = """ & ConvSPChars(Trim(lgObjRs("DESCRIPTION"))) & """" & vbCrLf
						Response.Write ".txtDrawingPath.value = """ & ConvSPChars(Trim(lgObjRs("DRAWING_PATH"))) & """" & vbCrLf
				    Response.Write "End With" & vbCrLf
				Response.Write "</Script>" & vbCrLf
			End If
		
			Call SubCloseRs(lgObjRs) 
		
		Case "CK"
			'------------------
			'품목체크 
			'------------------
			lgStrSQL = ""
			Call SubMakeSQLStatements("I_CK", strPlantCd, strItemCd, "", "", "")           '☜ : Make sql statements

		    If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists
		    
				IntRetCD = -1
				
				Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()
				Response.Write "<Script Language = VBScript>" & vbCrLf
				
				Response.Write "parent.txtItemNm.Value = """"" & vbCrLf
				Response.Write "parent.txtItemCd.focus" & vbCrLf
				
				Response.Write "</Script>" & vbCrLf
				Response.End 
		    Else
				IntRetCD = 1
				strItemNm = Trim(lgObjRs("ITEM_NM"))
				strItemAcct = Trim(lgObjRs("ITEM_ACCT"))
				strProcType = Trim(lgObjRs("PROCUR_TYPE"))
				strItemAcctNm = Trim(lgObjRs("ITEM_ACCT_NM"))
				strProcTypeNm = Trim(lgObjRs("PROCUR_TYPE_NM"))
				strSpec		= Trim(lgObjRs("SPEC"))
				strBasicUnit = Trim(lgObjRs("BASIC_UNIT"))

			End If
		
			Call SubCloseRs(lgObjRs) 
		    
	End Select
End Sub    

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Dim PrntKey
	Dim NodX
	Dim Node
	Dim iIntCnt, iLevelCnt

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    strBomNo = FilterVar(Trim(Request("txtBomNo")), "''", "S")
    
    '========================================================================
	' 0 Level BOM 전개를 실시한다.
	'========================================================================
	IF lgStrPrevKeyIndex = 0 Then						'row수가 maxrow수를 넘어서 다시 query 하더라도 최상위품목이 다시 조회되지 않도록.
		Call SubMakeSQLStatements("B_CK", strPlantCd, strItemCd, strBomNo, "", "")           '☜ : Make sql statements
	    	
	    If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
	        Call SetErrorStatus()
	    Else
			IntRetCD = 1

			Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "With Parent" & vbCrLf
			Response.Write "	.hBomType.value = """ & ConvSPChars(Trim(lgObjRs(2))) & """" & vbCrLf
			Response.Write "End With" & vbCrLf
			Response.Write "</Script>" & vbCrLf
	'		 Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

	        lgstrData = ""
	        iDx       = 1
			
	        lgstrData = lgstrData & Chr(11)	& Chr(11) & "0"			'Select, 레벨 
	        lgstrData = lgstrData & Chr(11)							'순서 
			lgstrData = lgstrData & Chr(11) & Trim(lgObjRs(1))		'자품목코드 
	        lgstrData = lgstrData & Chr(11) & lgObjRs(14)			'자품목명 
	        lgstrData = lgstrData & Chr(11) & lgObjRs(15)			'규격 
	        lgstrData = lgstrData & Chr(11) & Trim(lgObjRs(18))		'단위 
	        lgstrData = lgstrData & Chr(11) & lgObjRs(12)		'품목계정코드 
	        lgstrData = lgstrData & Chr(11) & lgObjRs(16)		'계정명 
	        lgstrData = lgstrData & Chr(11) & lgObjRs(13)		'조달구분코드 
	        lgstrData = lgstrData & Chr(11) & lgObjRs(17)		'조달구분명 
	        lgstrData = lgstrData & Chr(11) & lgObjRs(2)		'bom type
	        lgstrData = lgstrData & Chr(11)									'자품목기준수 
	        lgstrData = lgstrData & Chr(11)									'단위 
	        lgstrData = lgstrData & Chr(11)									'모품목기준수 
	        lgstrData = lgstrData & Chr(11)									'단위 
	        lgstrData = lgstrData & Chr(11)									'안전L/T
	        lgstrData = lgstrData & Chr(11)									'loss율 
	        lgstrData = lgstrData & Chr(11)									'유무상구분 
	        lgstrData = lgstrData & Chr(11)									'유무상구분명 
	        lgstrData = lgstrData & Chr(11)									'시작일 
	        lgstrData = lgstrData & Chr(11)									'종료일		
	        lgstrData = lgstrData & Chr(11)									'변경번호 
	        lgstrData = lgstrData & Chr(11) 								'변경내용 
	        lgstrData = lgstrData & Chr(11)									'변경근거 
	        lgstrData = lgstrData & Chr(11) 								'변경근거명 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("DRAWING_PATH")		'도면경로 
	        lgstrData = lgstrData & Chr(11)									'비고 
	        lgstrData = lgstrData & Chr(11)									'모품목 
	        lgstrData = lgstrData & Chr(11) 								'모품목bom no
	        lgstrData = lgstrData & Chr(11) 								'모품목조달구분 
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(19))		'품목유효기간시작일 
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(20))		'품목유효기간종료일 
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
	        lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
	        lgstrData = lgstrData & Chr(11) & Chr(12)

	        iDx =  iDx + 1

		End If   
	    
	    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
	    Call SubCloseRs(lgObjRs) 
	
	End IF
		     
	'========================================================================
	' 하위품목 BOM 전개를 실시한다.
	'========================================================================
	
	lgStrSQL = ""
	strPlantCd = FilterVar(Trim(Request("txtPlantCd")), "''", "S")			
	
	Call SubMakeSQLStatements("M", strPlantCd, strSpId, "", "", "")					'☜ : Make sql statements

	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists
		lgStrPrevKeyIndex = ""    
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
	Else
		IntRetCD = 1
		Call SubSkipRs(lgObjRs, lgMaxCount * lgStrPrevKeyIndex)
		iDx       = 2
					
        Do While Not lgObjRs.EOF
			
			'-----------------------
			' Level Setting
			'-----------------------
			strLevel = ""
			iLevelCnt = lgObjRs(6)
		
			For iIntCnt = 1 To iLevelCnt
				strLevel = strLevel & "."
			Next 
		
			strLevel = strLevel & lgObjRs(6)
			
			lgstrData = ""				
	        lgstrData = lgstrData & Chr(11) & Chr(11) & strLevel			'레벨 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("CHILD_ITEM_SEQ")		'순서 
			lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("CHILD_ITEM_CD")) '자품목코드 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_NM")			'자품목명 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("SPEC")				'규격 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("BASIC_UNIT")			'기준단위 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_ACCT")			'품목계정코드 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_ACCT_NM")		'계정명 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("PROCUR_TYPE")		'조달구분코드 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("PROCUR_TYPE_NM")		'조달구분명 
	        lgstrData = lgstrData & Chr(11) & Trim(lgObjRs(11))				'BOM type

	        lgstrData = lgstrData & Chr(11)	& UniConvNumberDBToCompany(lgObjRs(14), 6, 3, "", 0)	'자품목기준수   'hanc
	        lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs(15))										'자품목단위 

	        lgstrData = lgstrData & Chr(11)	& UniConvNumberDBToCompany(lgObjRs(12), 6, 3, "", 0)	'모품목기준수   'hanc
	        lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs(13))										'모품목단위 

	        lgstrData = lgstrData & Chr(11)	& lgObjRs(17)											'안전L/T
	        lgstrData = lgstrData & Chr(11)	& UniConvNumberDBToCompany(lgObjRs(16), 3, 3, "", 0)	'loss율 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs(18)											'유무상구분 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("SUPPLY_TYPE_NM")											'유무상구분명 
	        lgstrData = lgstrData & Chr(11)	& UNIDateClientFormat(lgObjRs(21))						'시작일 
	        lgstrData = lgstrData & Chr(11)	& UNIDateClientFormat(lgObjRs(22))						'종료일		
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("ECN_NO")										'변경번호 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("ECN_DESC") 									'변경내용 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("REASON_CD")									'변경근거 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("REASON_NM")									'변경근거명 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("DRAWING_PATH")								'도면경로 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs(20)		'비고 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs(7)		'모품목 
	        lgstrData = lgstrData & Chr(11) & lgObjRs(8)		'모품목bom no
	        lgstrData = lgstrData & Chr(11) & lgObjRs("PRNT_PROC_TYPE")		'모품목조달구분 
	        lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("ITEM_F_DT"))	'품목유효기간시작일 
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("ITEM_T_DT"))		'품목유효기간종료일 
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
	        lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
	        lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext

			ReDim Preserve TmpBuffer(iDx-2)
			TmpBuffer(iDx-2) = lgstrData	 
	        iDx =  iDx + 1
	        If iDx > lgMaxCount + 1  Then			'처음에 최상위품목row를 한줄 써주었으므로 
	           lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1    
	           Exit Do
	        End If   
        Loop 
        
		If iDx <= lgMaxCount + 1 Then
		   lgStrPrevKeyIndex = ""
		End If   
		
		iTotalStr = Join(TmpBuffer, "")

		Call SubHandleError("MR", lgObjConn, lgObjRs, Err)
		Call SubCloseRs(lgObjRs)       
    End If
 
    lgStrSQL = ""
	'-------------------------
	' 생성된 temp table 삭제 
	'-------------------------
    lgStrSQL = "DELETE FROM P_BOM_FOR_EXPLOSION "
	lgStrSQL = lgStrSQL & " WHERE PLANT_CD = " & FilterVar(Trim(Request("txtPlantCd"))	,"''", "S")
	lgStrSQL = lgStrSQL & " AND USER_ID = " & strSpId
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords 
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub    

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2,pCode3,pCode4)
    Dim iSelCount
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case pDataType

		Case "M"
			lgStrSQL = "SELECT a.*, b.ITEM_NM, b.PHANTOM_FLG, b.SPEC, b.BASIC_UNIT, c.ITEM_ACCT, dbo.ufn_GetCodeName('P1001', c.ITEM_ACCT) ITEM_ACCT_NM, c.PROCUR_TYPE, dbo.ufn_GetCodeName('P1003', c.PROCUR_TYPE) PROCUR_TYPE_NM, dbo.ufn_GetCodeName('M2201', a.SUPPLY_TYPE) SUPPLY_TYPE_NM, c.VALID_FROM_DT ITEM_F_DT, c.VALID_TO_DT ITEM_T_DT, g.PROCUR_TYPE PRNT_PROC_TYPE, "
			lgStrSQL = lgStrSQL & "  d.ECN_DESC,d.REASON_CD, dbo.ufn_GetCodeName('P1402', d.REASON_CD) REASON_NM, f.DRAWING_PATH "
			lgStrSQL = lgStrSQL & " FROM P_BOM_FOR_EXPLOSION a LEFT OUTER JOIN P_ECN_MASTER d ON a.ECN_NO = d.ECN_NO, P_BOM_FOR_EXPLOSION aa LEFT OUTER JOIN P_BOM_HEADER f ON (aa.PLANT_CD = f.PLANT_CD AND aa.CHILD_ITEM_CD = f.ITEM_CD AND aa.PRNT_BOM_NO = f.BOM_NO), B_ITEM b, B_ITEM_BY_PLANT c, B_ITEM_BY_PLANT g"
			lgStrSQL = lgStrSQL & " WHERE (a.PLANT_CD = aa.PLANT_CD AND a.USER_ID = aa.USER_ID AND a.SEQ = aa.SEQ)"
			lgStrSQL = lgStrSQL & " AND a.CHILD_ITEM_CD = b.ITEM_CD AND b.ITEM_CD = c.ITEM_CD AND a.PLANT_CD = c.PLANT_CD "
			lgStrSQL = lgStrSQL & " AND g.PLANT_CD = a.PLANT_CD AND g.ITEM_CD = a.PRNT_ITEM_CD " 
			lgStrSQL = lgStrSQL & " AND a.PLANT_CD = " & pCode
			lgStrSQL = lgStrSQL & " AND a.USER_ID = " & pCode1
			lgStrSQL = lgStrSQL & " ORDER BY a.SEQ "
		Case "B_CK"
			lgStrSQL = "SELECT a.*, b.ITEM_ACCT, b.PROCUR_TYPE, c.ITEM_NM, c.SPEC, d.MINOR_NM  ,e.MINOR_NM, c.BASIC_UNIT, b.VALID_FROM_DT, b.VALID_TO_DT "
			lgStrSQL = lgStrSQL & " FROM P_BOM_HEADER a, B_ITEM_BY_PLANT b, B_ITEM c, B_MINOR d, B_MINOR e  "
			lgStrSQL = lgStrSQL & " WHERE a.PLANT_CD = b.PLANT_CD AND a.ITEM_CD = b.ITEM_CD AND b.ITEM_CD = c.ITEM_CD AND b.ITEM_ACCT = d.MINOR_CD AND d.MAJOR_CD = 'P1001' AND b.PROCUR_TYPE = e.MINOR_CD AND e.MAJOR_CD='P1003' "
			lgStrSQL = lgStrSQL & " AND a.PLANT_CD = " & pCode 
			lgStrSQL = lgStrSQL & " AND a.ITEM_CD = " & pCode1
			lgStrSQL = lgStrSQL & " AND a.BOM_NO= " & pCode2
			
		Case "BT_CK"
			lgStrSQL = "SELECT * FROM B_MINOR WHERE MAJOR_CD = 'P1401'"
			lgStrSQL = lgStrSQL & " AND MINOR_CD = " & pCode 
			
		Case "I_CK"
			lgStrSQL = "SELECT a.*, b.ITEM_NM, b.SPEC, dbo.ufn_GetCodeName('P1001', a.ITEM_ACCT) ITEM_ACCT_NM, b.PHANTOM_FLG, b.BASIC_UNIT, dbo.ufn_GetCodeName('P1003', a.PROCUR_TYPE) PROCUR_TYPE_NM"
			lgStrSQL = lgStrSQL & " FROM B_ITEM_BY_PLANT a, B_ITEM b "
			lgStrSQL = lgStrSQL & " WHERE a.ITEM_CD = b.ITEM_CD "
			lgStrSQL = lgStrSQL & " AND a.PLANT_CD = " & pCode 
			lgStrSQL = lgStrSQL & " AND a.ITEM_CD = " & pCode1
			
		Case "P_CK"
			lgStrSQL = "SELECT * FROM B_PLANT WHERE PLANT_CD = " & pCode 
		
    End Select

	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()
	
	Dim strMsg_cd
    Dim strMsg_text
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	If Request("txtBomNo") = "" Then
		strBomNo = " "
	Else 
		strBomNo = Request("txtBomNo")
	End If
	
    With lgObjComm
        .CommandText = "usp_BOM_explode_main"
        .CommandType = adCmdStoredProc
        
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@srch_type",	advarXchar,adParamInput,2, Request("rdoSrchType"))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_cd",	advarXchar,adParamInput,4, Request("txtPlantCd"))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@par_item_cd",	advarXchar,adParamInput,18, Request("txtItemCd"))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@par_bom_no",advarXchar,adParamInput,4,strBomNo)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@base_dt_s",	advarXchar,adParamInput,10,UniConvDate(Request("txtBaseDt")))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@base_qty",	adInteger,adParamInput,2,1)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd",	advarXchar,adParamOutput,6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text",	advarXchar,adParamOutput,60)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@user_id",	advarXchar,adParamOutput,13)

        lgObjComm.Execute ,, adExecuteNoRecords
        
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
        If  IntRetCD <> 1 then
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            strMsg_text = lgObjComm.Parameters("@msg_text").Value
            strSpId = FilterVar(lgObjComm.Parameters("@user_id").Value, "''", "S")
            
            If strMsg_cd <> MSG_OK_STR Then
				Call DisplayMsgBox(strMsg_cd, vbInformation, strMsg_text, "", I_MKSCRIPT)
			End If
            IntRetCD = -1
            Exit Sub
        Else
			IntRetCD = 1
        End if
    Else           
        Call SvrMsgBox(Err.Description, vbinformation, i_mkscript)
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
        IntRetCD = -1
    End if
    
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
        Case "MB"
			ObjectContext.SetAbort
            Call SetErrorStatus        
    End Select
End Sub
		
Response.Write "<Script Language = VBScript>" & vbCrLf
	If Trim(lgErrorStatus) = "NO" And IntRetCd <> -1 Then
		 Response.Write "With Parent" & vbCrLf
	        Response.Write ".ggoSpread.Source = .vspdData" & vbCrLf
	        Response.Write ".lgStrPrevKeyIndex = """ & lgStrPrevKeyIndex & """" & vbCrLf
	                
	        Response.Write ".ggoSpread.SSShowDataByClip """ & ConvSPChars(iTotalStr) & """" & vbCrLf

	        Response.Write ".DBQueryOk(" & lgLngMaxRow & " + 1)" & vbCrLf
	     Response.Write "End with" & vbCrLf
	End If   
Response.Write "</Script>" & vbCrLf
Response.End
%>