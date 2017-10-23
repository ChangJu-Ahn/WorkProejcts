<%@ LANGUAGE=VBSCript%>
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
Call LoadinfTB19029B("I", "*", "NOCOOKIE", "MB")

'---------------------------------------Common-----------------------------------------------------------
lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    
lgMaxCount        = 500									'2004-03-18
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")), 0)   
    
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
Dim strItemAcctGrp
Dim BaseDt
Dim idx
	
Dim QueryType
Dim strSpId
Dim strLevel

Dim TmpBuffer
Dim iTotalStr

ReDim TmpBuffer(0)
	
strPlantCd = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
strItemCd = FilterVar(Trim(Request("txtItemCd"))	, "''", "S")
strBomNo = FilterVar(Trim(Request("txtBomNo"))	, "''", "S")
	
QueryType = Request("QueryType")
'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	
Select Case QueryType
    Case "A"				'☜: 전체Query
		BaseDt = FilterVar(UNIConvYYYYMMDDToDate(gAPDateFormat,"1900","01","01"), "''", "S")
		strBaseDt = FilterVar(Trim(Request("txtBaseDt")), BaseDt, "D")
		strExpFlg = FilterVar(Trim(Request("rdoSrchType"))	, "''", "S")

		Call SubBizQuery("P_CK")										
		Call SubBizQuery("CK")				
		Call SubBizQuery("B_CK")
		Call SubCreateCommandObject(lgObjComm)
		Call SubBizBatch()
		Call SubBizQueryMulti()
			
		Call SubCloseCommandObject(lgObjComm)
			
    Case "I"								        							'☜: Header Query
		Call SubBizQuery("CK")

End Select
      
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
			Call SubMakeSQLStatements("P_CK",strPlantCd,"","","","")           '☜ : Make sql statements
			
			If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		    
				IntRetCD = -1
				Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.Frm1.txtPlantNm.Value  = """"" & vbCrLf   'Set condition area
					Response.Write "parent.Frm1.txtPlantCd.focus" & vbCrLf   'Set condition area
				Response.Write "</Script>" & vbcRLf
				Response.End
		    Else
				IntRetCD = 1
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.Frm1.txtPlantNm.Value = """ & ConvSPChars(lgObjRs(1)) & """" & vbCrLf 'Set condition area
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
					Response.Write "parent.frm1.hBomType.value = """"" & vbCrLf
					Response.Write "parent.frm1.txtBomNo.focus" & vbCrLf
				Response.Write "</Script>" & vbCrLf
				Response.End
		    Else
				IntRetCD = 1
				Response.Write "<Script Language = VBScript>" & vbCrLf
					'Response.Write "parent.frm1.txtBomNm.value = """ & lgObjRs(0) & """" & vbCrLf
				Response.Write "</Script>" & vbCrLf
			End If
		
			Call SubCloseRs(lgObjRs) 

			'------------------
			'품목, bom no 체크 
			'------------------
			lgStrSQL = ""
			Call SubMakeSQLStatements("B_CK",strPlantCd,strItemCd,strBomNo,"","")           '☜ : Make sql statements
			
		    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		    
				IntRetCD = -1
				
				Call DisplayMsgBox("182600", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.frm1.hBomType.value = """"" & vbCrLf
					Response.Write "Call parent.DbQueryNotOk()" & vbCrLf
				Response.Write "</Script>" & vbCrLf
				Response.End
		    Else
				IntRetCD = 1
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "With Parent" & vbCrLf
						Response.Write ".Frm1.txtItemNm.Value = """ & ConvSPChars(Trim(lgObjRs(14))) & """" & vbCrLf
						Response.Write ".frm1.txtItemAcctNm.value = """ & ConvSPChars(lgObjRs(16)) & """" & vbCrLf
						Response.Write ".frm1.txtItemAcct.value = """ & ConvSPChars(Trim(lgObjRs(12))) & """" & vbCrLf
						Response.Write ".frm1.txtSpec.value = """ & ConvSPChars(lgObjRs(15)) & """" & vbCrLf
						Response.Write ".frm1.txtPlantItemFromDt.text = """ & UNIDateClientFormat(lgObjRs(19)) & """" & vbCrLf
						Response.Write ".frm1.txtPlantItemToDt.text = """ & UNIDateClientFormat(lgObjRs(20)) & """" & vbCrLf
						Response.Write ".frm1.txtBasicUnit.value = """ & ConvSPChars(Trim(lgObjRs(18))) & """" & vbCrLf
						Response.Write ".frm1.txtBOMDesc.value = """ & ConvSPChars(Trim(lgObjRs("DESCRIPTION"))) & """" & vbCrLf
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

		    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		    
				IntRetCD = -1
				
				Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()
				Response.Write "<Script Language = VBScript>" & vbCrLf
				
				If QueryType = "A" Then
					Response.Write "parent.frm1.txtItemNm.Value = """"" & vbCrLf
					Response.Write "parent.frm1.txtItemCd.focus" & vbCrLf
				Else
					Response.Write "Call parent.LookUpItemByPlantNotOk" & vbCrLf
				End If
				
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
				strItemAcctGrp = Trim(lgObjRs("ITEM_ACCT_GRP"))
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
    
    strBomNo = FilterVar(Trim(Request("txtBomNo"))	, "''", "S")
    
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
					Response.Write ".frm1.hBomType.value = """ & ConvSPChars(Trim(lgObjRs(2))) & """" & vbCrLf
			    Response.Write "End With" & vbCrLf
			Response.Write "</Script>" & vbCrLf

	        lgstrData = ""
	        iDx       = 1
			
	        lgstrData = lgstrData & Chr(11)	& "0"								'레벨 
	        lgstrData = lgstrData & Chr(11)									'순서 
			lgstrData = lgstrData & Chr(11) & Trim(lgObjRs(1))		'자품목코드 
			lgstrData = lgstrData & Chr(11)									'자품목팝업 
	        lgstrData = lgstrData & Chr(11) & lgObjRs(14)		'자품목명 
	        lgstrData = lgstrData & Chr(11) & lgObjRs(15)		'규격 
	        lgstrData = lgstrData & Chr(11) & Trim(lgObjRs(18))		'단위 
	   
	        lgstrData = lgstrData & Chr(11) & lgObjRs(12)		'품목계정코드 
	        lgstrData = lgstrData & Chr(11) & lgObjRs(16)		'계정명 
	        lgstrData = lgstrData & Chr(11) & lgObjRs(13)		'조달구분코드 
	        lgstrData = lgstrData & Chr(11) & lgObjRs(17)		'조달구분명 
	        lgstrData = lgstrData & Chr(11) & lgObjRs(2)		'bom type
	        lgstrData = lgstrData & Chr(11)									'bom type popup
	        lgstrData = lgstrData & Chr(11)									'자품목기준수 
	        lgstrData = lgstrData & Chr(11)									'단위 
	        lgstrData = lgstrData & Chr(11)									'단위팝업			
	        lgstrData = lgstrData & Chr(11)									'모품목기준수 
	        lgstrData = lgstrData & Chr(11)									'단위 
	        lgstrData = lgstrData & Chr(11)									'단위팝업	
	        lgstrData = lgstrData & Chr(11)									'안전L/T
	        lgstrData = lgstrData & Chr(11)									'loss율 
	        lgstrData = lgstrData & Chr(11)									'유무상구분 
	        lgstrData = lgstrData & Chr(11)									'유무상구분명 
	        lgstrData = lgstrData & Chr(11)									'시작일 
	        lgstrData = lgstrData & Chr(11)									'종료일		
	        lgstrData = lgstrData & Chr(11)									'변경번호 
	        lgstrData = lgstrData & Chr(11)									'변경번호 팝업 
	        lgstrData = lgstrData & Chr(11) 								'변경내용 
	        lgstrData = lgstrData & Chr(11)									'변경근거 
	        lgstrData = lgstrData & Chr(11)									'변경근거 팝업 
	        lgstrData = lgstrData & Chr(11) 								'변경근거명 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("DRAWING_PATH")		'도면경로 
	        lgstrData = lgstrData & Chr(11)									'비고 
	        lgstrData = lgstrData & Chr(11)									'모품목 
	        lgstrData = lgstrData & Chr(11) 								'모품목bom no
	        lgstrData = lgstrData & Chr(11) 								'모품목조달구분 
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(19))		'품목유효기간시작일 
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(20))		'품목유효기간종료일 
			lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_ACCT_GRP")		'품목계정그룹 
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
	        lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
	        lgstrData = lgstrData & Chr(11) & Chr(12)
			
			ReDim Preserve TmpBuffer(iDx-1)
			TmpBuffer(iDx-1) = lgstrData
			
	        iDx =  iDx + 1

		End If   
	    
	    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
	    Call SubCloseRs(lgObjRs) 
	
	End IF
		     
	'========================================================================
	' 하위품목 BOM 전개를 실시한다.
	'========================================================================
	
	lgStrSQL = ""
	strPlantCd = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")			
	
	Call SubMakeSQLStatements("M", strPlantCd, strSpId, "", "", "")					'☜ : Make sql statements

	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists
		lgStrPrevKeyIndex = ""
	
	Else
		IntRetCD = 1
		Call SubSkipRs(lgObjRs, lgMaxCount * lgStrPrevKeyIndex)
		iDx       = 2
		
        Do While Not lgObjRs.EOF
			
			'-----------------------
			' Level Setting
			'-----------------------
			strLevel = ""
			iLevelCnt = lgObjRs("LEVEL_CD")
		
			For iIntCnt = 1 To iLevelCnt
				strLevel = strLevel & "."
			Next 
		
			strLevel = strLevel & iLevelCnt
			
			lgstrData = ""
			
	        lgstrData = lgstrData & Chr(11) & strLevel			'레벨 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("CHILD_ITEM_SEQ")		'순서 
			lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("CHILD_ITEM_CD"))	& Chr(11) '자품목코드, 자품목팝업 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_NM")		'자품목명 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("SPEC")		'규격 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("BASIC_UNIT")		'기준단위 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_ACCT")		'품목계정코드 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_ACCT_NM")		'계정명 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("PROCUR_TYPE")		'조달구분코드 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("PROCUR_TYPE_NM")		'조달구분명 
	        lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("PRNT_BOM_NO"))	& Chr(11) 'BOM type, BOM type popup

	        lgstrData = lgstrData & Chr(11)	& UniConvNumberDBToCompany(lgObjRs("CHILD_ITEM_QTY"), 6, 3, "", 0)		'자품목기준수 
	        lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("CHILD_ITEM_UNIT"))	& Chr(11)	'자품목단위, 단위팝업			

	        lgstrData = lgstrData & Chr(11)	& UniConvNumberDBToCompany(lgObjRs("PRNT_ITEM_QTY"), 6, 3, "", 0)		'모품목기준수 
	        lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("PRNT_ITEM_UNIT"))	& Chr(11)	'모품목단위, 단위팝업	

	        lgstrData = lgstrData & Chr(11)	& lgObjRs("SAFETY_LT")		'안전L/T
	        lgstrData = lgstrData & Chr(11)	& UniConvNumberDBToCompany(lgObjRs("LOSS_RATE"), 3, 3, "", 0)		'loss율 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("SUPPLY_TYPE")		'유무상구분 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("SUPPLY_TYPE_NM")		'유무상구분명 
	        lgstrData = lgstrData & Chr(11)	& UNIDateClientFormat(lgObjRs("BOM_FROM_DT"))		'시작일 
	        lgstrData = lgstrData & Chr(11)	& UNIDateClientFormat(lgObjRs("BOM_TO_DT"))		'종료일		
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("ECN_NO")	& Chr(11)	'변경번호, 변경번호 popup
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("ECN_DESC") 		'변경내용	        
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("REASON_CD")	& Chr(11)	'변경근거, 변경근거 popup
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("REASON_NM")		'변경근거명 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("DRAWING_PATH")	'도면경로 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("REMARK")		'비고 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("PRNT_ITEM_CD")		'모품목 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("PRNT_BOM_NO")		'모품목bom no
	        lgstrData = lgstrData & Chr(11) & lgObjRs("PRNT_PROC_TYPE")		'모품목조달구분 
	        lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("BOM_FROM_DT"))	'품목유효기간시작일 
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("BOM_TO_DT"))		'품목유효기간종료일 
			lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_ACCT_GRP")		'품목계정그룹 
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
	        lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
	        lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext
			
			ReDim Preserve TmpBuffer(iDx-1)
			TmpBuffer(iDx-1) = lgstrData
			
	        iDx =  iDx + 1
	        If iDx > lgMaxCount + 1  Then			'처음에 최상위품목row를 한줄 써주었으므로 
	           lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
	               
	           Exit Do
	        End If   
        Loop 
        
		If iDx <= lgMaxCount + 1 Then
		   lgStrPrevKeyIndex = ""
		End If   
		
		Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
		Call SubCloseRs(lgObjRs)       
    End If
	
	iTotalStr = Join(TmpBuffer,"")
	
    lgStrSQL = ""
	'-------------------------
	' 생성된 temp table 삭제 
	'-------------------------
    lgStrSQL = "DELETE FROM P_BOM_FOR_EXPLOSION "
	lgStrSQL = lgStrSQL & " WHERE PLANT_CD = " & FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
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
			lgStrSQL = "SELECT a.LEVEL_CD, a.PRNT_ITEM_CD, a.PRNT_BOM_NO, a.CHILD_ITEM_SEQ, a.CHILD_ITEM_CD, a.PRNT_ITEM_QTY, a.PRNT_ITEM_UNIT, a.CHILD_ITEM_QTY, a.CHILD_ITEM_UNIT, a.LOSS_RATE, a.SAFETY_LT, a.SUPPLY_TYPE, a.REMARK, a.VALID_FROM_DT BOM_FROM_DT, a.VALID_TO_DT BOM_TO_DT, a.ECN_NO, "
			lgStrSQL = lgStrSQL & " b.ITEM_NM, b.PHANTOM_FLG, b.SPEC, b.BASIC_UNIT, c.ITEM_ACCT, dbo.ufn_GetCodeName(" & FilterVar("P1001", "''", "S") & ", c.ITEM_ACCT) ITEM_ACCT_NM, c.PROCUR_TYPE, dbo.ufn_GetCodeName(" & FilterVar("P1003", "''", "S") & ", c.PROCUR_TYPE) PROCUR_TYPE_NM, dbo.ufn_GetCodeName(" & FilterVar("M2201", "''", "S") & ", a.SUPPLY_TYPE) SUPPLY_TYPE_NM, g.PROCUR_TYPE PRNT_PROC_TYPE, "
			lgStrSQL = lgStrSQL & " d.ECN_DESC, d.REASON_CD, dbo.ufn_GetCodeName(" & FilterVar("P1402", "''", "S") & ", d.REASON_CD) REASON_NM,  f.DRAWING_PATH, dbo.ufn_GetItemAcctGrp(c.ITEM_ACCT) ITEM_ACCT_GRP "
			lgStrSQL = lgStrSQL & " FROM P_BOM_FOR_EXPLOSION a LEFT OUTER JOIN P_ECN_MASTER d ON a.ECN_NO = d.ECN_NO, P_BOM_FOR_EXPLOSION aa LEFT OUTER JOIN P_BOM_HEADER f ON (aa.PLANT_CD = f.PLANT_CD AND aa.CHILD_ITEM_CD = f.ITEM_CD AND aa.PRNT_BOM_NO = f.BOM_NO), B_ITEM b, B_ITEM_BY_PLANT c, B_ITEM_BY_PLANT g"
			lgStrSQL = lgStrSQL & " WHERE (a.PLANT_CD = aa.PLANT_CD AND a.USER_ID = aa.USER_ID AND a.SEQ = aa.SEQ)"
			lgStrSQL = lgStrSQL & " AND a.CHILD_ITEM_CD = b.ITEM_CD AND b.ITEM_CD = c.ITEM_CD AND a.PLANT_CD = c.PLANT_CD "
			lgStrSQL = lgStrSQL & " AND g.PLANT_CD = a.PLANT_CD AND g.ITEM_CD = a.PRNT_ITEM_CD " 
			lgStrSQL = lgStrSQL & " AND a.PLANT_CD = " & pCode
			lgStrSQL = lgStrSQL & " AND a.USER_ID = " & pCode1
			lgStrSQL = lgStrSQL & " ORDER BY a.SEQ "
		Case "B_CK"
			lgStrSQL = "SELECT a.*, b.ITEM_ACCT, b.PROCUR_TYPE, c.ITEM_NM, c.SPEC, d.MINOR_NM  ,e.MINOR_NM, c.BASIC_UNIT, b.VALID_FROM_DT, b.VALID_TO_DT, dbo.ufn_GetItemAcctGrp(b.ITEM_ACCT) ITEM_ACCT_GRP "
			lgStrSQL = lgStrSQL & " FROM P_BOM_HEADER a, B_ITEM_BY_PLANT b, B_ITEM c, B_MINOR d, B_MINOR e  "
			lgStrSQL = lgStrSQL & " WHERE a.PLANT_CD = b.PLANT_CD AND a.ITEM_CD = b.ITEM_CD AND b.ITEM_CD = c.ITEM_CD AND b.ITEM_ACCT = d.MINOR_CD AND d.MAJOR_CD = " & FilterVar("P1001", "''", "S") & " AND b.PROCUR_TYPE = e.MINOR_CD AND e.MAJOR_CD=" & FilterVar("P1003", "''", "S") & " "
			lgStrSQL = lgStrSQL & " AND a.PLANT_CD = " & pCode 
			lgStrSQL = lgStrSQL & " AND a.ITEM_CD = " & pCode1
			lgStrSQL = lgStrSQL & " AND a.BOM_NO= " & pCode2
			
		Case "BT_CK"
			lgStrSQL = "SELECT MINOR_NM FROM B_MINOR WHERE MAJOR_CD = " & FilterVar("P1401", "''", "S") & " AND MINOR_CD = " & pCode 
			
		Case "I_CK"
			lgStrSQL = "SELECT a.*, b.ITEM_NM, b.SPEC, dbo.ufn_GetCodeName(" & FilterVar("P1001", "''", "S") & ", a.ITEM_ACCT) ITEM_ACCT_NM, b.PHANTOM_FLG, b.BASIC_UNIT, " _
						& " dbo.ufn_GetCodeName(" & FilterVar("P1003", "''", "S") & ", a.PROCUR_TYPE) PROCUR_TYPE_NM, dbo.ufn_GetItemAcctGrp(a.ITEM_ACCT) ITEM_ACCT_GRP "
			lgStrSQL = lgStrSQL & " FROM B_ITEM_BY_PLANT a, B_ITEM b "
			lgStrSQL = lgStrSQL & " WHERE a.ITEM_CD = b.ITEM_CD "
			lgStrSQL = lgStrSQL & " AND a.PLANT_CD = " & pCode 
			lgStrSQL = lgStrSQL & " AND a.ITEM_CD = " & pCode1
			
		Case "P_CK"
			lgStrSQL = "SELECT * FROM B_PLANT A, P_PLANT_CONFIGURATION B"
			lgStrSQL = lgStrSQL & " WHERE A.PLANT_CD = B.PLANT_CD"
			lgStrSQL = lgStrSQL & " AND B.ENG_BOM_FLAG = 'Y'"
			lgStrSQL = lgStrSQL & " AND A.PLANT_CD = " & pCode			
		
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
	Select Case QueryType
		Case "A"                                                         '☜ : Query
			If Trim(lgErrorStatus) = "NO" And IntRetCd <> -1 Then
				 Response.Write "With Parent" & vbCrLf
	                Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
	                Response.Write ".lgStrPrevKeyIndex = """ & lgStrPrevKeyIndex & """" & vbCrLf
	                
	                Response.Write ".ggoSpread.SSShowDataByClip """ & ConvSPChars(iTotalStr) & """" & vbCrLf

	                Response.Write ".DBQueryOk(" & lgLngMaxRow & " + 1)" & vbCrLf
		         Response.Write "End with" & vbCrLf
	        End If   

		Case "I"
			If Trim(lgErrorStatus) = "NO" And IntRetCd <> -1 Then
				Response.Write "Call parent.LookUpItemByPlantOk(""" & ConvSPChars(strItemNm) & """, """ & strItemAcct & """, """ & strItemAcctNm & """, """ & strProcType & """, """ & strProcTypeNm & """, """ & ConvSPChars(strSpec) & """, """ & ConvSPChars(strBasicUnit) & """, """ & Request("IRow") & """, """ & ConvSPChars(strBasicUnit) & """)" & vbCrLf
			Else			
				Response.Write "Call parent.LookUpItemByPlantNotOk()" & vbCrLf
			End If   
	End Select    
Response.Write "</Script>" & vbCrLf
Response.End
%>
