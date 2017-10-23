<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : 설계BOM관리 
'*  2. Function Name        : 
'*  3. Program ID           : p1713qb1.asp
'*  4. Program Name         : BOM변경이력 조회 
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2005/01/25
'*  8. Modified date(Last)  : 2005/01/25
'*  9. Modifier (First)     : Cho Yong Chill
'* 10. Modifier (Last)      : Cho Yong Chill
'* 11. Comment              :
'**********************************************************************************************-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/AdoVbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDB.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%

'On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Const C_SHEETMAXROWS_D = 30

Call HideStatusWnd                                                               '☜: Hide Processing message

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "P", "NOCOOKIE", "MB")

'---------------------------------------Common-----------------------------------------------------------
lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
lgMaxCount        = C_SHEETMAXROWS_D			                                 '☜: Fetch count at a time for VspdData
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")), 0)
    
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '☜: Set to space
'lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

'------ Developer Coding part (Start ) ------------------------------------------------------------------
Dim IntRetCD
Dim strPlantCd
Dim strBomNo
Dim strChgFromDt
Dim strChgToDt
Dim strECNNo
Dim strItemCd
Dim strChildItemCd
Dim strExpFlg
Dim BaseDt
Dim i, iDx

Dim TmpBuffer
Dim iTotalStr
	
Dim strSpId
Dim strLevel

ReDim TmpBuffer(0)
	
strPlantCd		= FilterVar(Request("txtPlantCd"), "''", "S")
strBomNo		= FilterVar(Request("txtBomNo"), "''", "S")
strChgFromDt	= FilterVar(Trim(UniConvDate(Request("txtChgFromDt"))) & " 00:00:00", "''", "S")
strChgToDt		= FilterVar(Trim(UniConvDate(Request("txtChgToDt"))) & " 23:59:59", "''", "S")
strECNNo		= FilterVar(Request("txtECNNo"), "''", "S")
strItemCd		= FilterVar(Request("txtItemCd"), "''", "S")
strChildItemCd  = FilterVar(Request("txtChildItemCd"), "''", "S")
strExpFlg 		= FilterVar(Trim(Request("rdoSrchType"))	, "''", "S")

'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	
Call SubBizQuery("P_CK")	'공장 Check
Call SubBizQuery("CD_CK")	'변경기간 Check
Call SubBizQuery("ENC_CK")	'ENC No Check
Call SubBizQuery("IP_CK")	'모품목 Check
Call SubBizQuery("IC_CK")	'자품목 Check

Call SubCreateCommandObject(lgObjComm)
Call SubBizBatch()
Call SubBizQueryMulti()

Call SubCloseCommandObject(lgObjComm)
Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery(pDataType)

	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear   
                                                                         '☜: Clear Error status
 Select Case pDataType

		'--------------
		'공장 체크		
		'--------------	
		Case "P_CK"
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
		
		'------------------
		'변경일체크 
		'------------------
		Case "CD_CK"
			IF strChgToDt = FilterVar("1900-01-01 23:59:59", "''", "S") THEN strChgToDt = FilterVar("2999-12-31 23:59:59", "''", "S")
			
			IF strChgFromDt > strChgToDt THEN
				Call DisplayMsgBox("800111", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()
				
				Response.Write "<Script Language = VBScript>" & vbCrLf
				Response.Write "parent.Frm1.txtChgFromDt.Focus" & vbCrLf 
				Response.Write "</Script>" & vbcRLf
				Response.End
		
				Call SubCloseRs(lgObjRs) 
			END IF
		
		'------------------
		'설계변경번호체크 
		'------------------
		Case "ENC_CK"			
			If strECNNo = FilterVar("", "''", "S") THEN
				Response.Write "<Script Language = VBScript>" & vbCrLf
				Response.Write "parent.Frm1.txtECNNoDesc.Value  = """"" & vbCrLf		'Set condition area
				Response.Write "</Script>" & vbcRLf
			Else
				lgStrSQL = ""
				Call SubMakeSQLStatements("ECN_CK",strECNNo,"","","","")           '☜ : Make sql statements
						
				If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
					    
					IntRetCD = -1
							
					Call DisplayMsgBox("182801", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
					Call SetErrorStatus()
					
					Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.Frm1.txtECNNoDesc.Value  = """"" & vbCrLf   'Set condition area
					Response.Write "parent.frm1.txtECNNo.focus" & vbCrLf
					Response.Write "</Script>" & vbCrLf
					Response.End							
				Else
					IntRetCD = 1
					Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.Frm1.txtECNNoDesc.Value = """ & ConvSPChars(Trim(lgObjRs(1))) & """" & vbCrLf
					Response.Write "</Script>" & vbcRLf
				End If
					
				Call SubCloseRs(lgObjRs) 
			End If
			
		'------------------
		'모품목체크 
		'------------------
		Case "IP_CK"			
			IF strItemCd = FilterVar("", "''", "S") THEN
				Response.Write "<Script Language = VBScript>" & vbCrLf
				Response.Write "parent.Frm1.txtItemNm.Value  = """"" & vbCrLf		'Set condition area
				Response.Write "</Script>" & vbcRLf
			ELSE
				lgStrSQL = ""
				Call SubMakeSQLStatements("I_CK",strPlantCd,strItemCd,"","","")          '☜ : Make sql statements
		
				If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                   'If data not exists
					    
					IntRetCD = -1
		
					Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
					Call SetErrorStatus()
		
					Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.Frm1.txtItemNm.Value  = """"" & vbCrLf   'Set condition area
					Response.Write "parent.Frm1.txtItemCd.Focus" & vbCrLf 
					Response.Write "</Script>" & vbcRLf
					Response.End
				Else
					IntRetCD = 1
					Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.Frm1.txtItemNm.Value = """ & ConvSPChars(Trim(lgObjRs(0))) & """" & vbCrLf
					Response.Write "</Script>" & vbcRLf
				End If
					
				Call SubCloseRs(lgObjRs) 
		
			END IF
			
		'------------------
		'자품목체크 
		'------------------
		Case "IC_CK"						
			IF strChildItemCd = FilterVar("", "''", "S") THEN
				Response.Write "<Script Language = VBScript>" & vbCrLf
				Response.Write "parent.Frm1.txtChildItemNm.Value  = """"" & vbCrLf   'Set condition area
				Response.Write "</Script>" & vbcRLf
			ELSE
				lgStrSQL = ""
				Call SubMakeSQLStatements("I_CK",strPlantCd,strChildItemCd,"","","")          '☜ : Make sql statements

				If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
					    
					IntRetCD = -1
							
					Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
					Call SetErrorStatus()
		
					Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.Frm1.txtChildItemNm.Value  = """"" & vbCrLf   'Set condition area
					Response.Write "parent.Frm1.txtChildItemCd.Focus" & vbCrLf 
					Response.Write "</Script>" & vbcRLf
					Response.End
				Else
					IntRetCD = 1
					Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.Frm1.txtChildItemNm.Value = """ & ConvSPChars(Trim(lgObjRs(0))) & """" & vbCrLf
					Response.Write "</Script>" & vbcRLf
				End If
			END IF
				
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
	Dim sModifiedDate

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    strBomNo = FilterVar(Trim(Request("txtBomNo"))	,"''", "S")
    
    '========================================================================
	' 0 Level BOM 전개를 실시한다.
	'========================================================================
	
	If lgStrPrevKeyIndex = 0 Then					'row수가 maxrow수를 넘어서 다시 query 하더라도 최상위품목이 다시 조회되지 않도록.
		If strExpFlg = "1" OR strExpFlg = "2" Then
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
			
		        lgstrData = lgstrData & Chr(11)	& ""							'설계변경구분 
		        lgstrData = lgstrData & Chr(11)									'설계변경일		
		        lgstrData = lgstrData & Chr(11)	& "0"							'레벨 
		        lgstrData = lgstrData & Chr(11)									'순서 
				lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("ITEM_CD"))		'자품목코드 
		        lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_NM")			'자품목명 
		        lgstrData = lgstrData & Chr(11) & lgObjRs("SPEC")				'규격 
		        lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_ACCT_NM")		'품목계정명  
		        lgstrData = lgstrData & Chr(11) & lgObjRs("PROCUR_TYPE_NM")		'조달구분명 
		        lgstrData = lgstrData & Chr(11)									'자품목기준수 
		        lgstrData = lgstrData & Chr(11)									'단위 
		        lgstrData = lgstrData & Chr(11)									'모품목기준수 
		        lgstrData = lgstrData & Chr(11)									'단위 
		        lgstrData = lgstrData & Chr(11)									'안전L/T
		        lgstrData = lgstrData & Chr(11)									'loss율 
		        lgstrData = lgstrData & Chr(11)									'유무상구분명 
		        lgstrData = lgstrData & Chr(11)									'시작일 
		        lgstrData = lgstrData & Chr(11)									'종료일		
		        lgstrData = lgstrData & Chr(11)									'설계변경번호 
		        lgstrData = lgstrData & Chr(11)									'설계변경내용 
		        lgstrData = lgstrData & Chr(11)									'설계변경근거명 
		        lgstrData = lgstrData & Chr(11) & lgObjRs("DRAWING_PATH")		'도면경로 
		        lgstrData = lgstrData & Chr(11)									'비고 
		        lgstrData = lgstrData & Chr(11)									'모품목명 
		        lgstrData = lgstrData & Chr(11) 								'모품목bom no
		'------ Developer Coding part (End   ) ------------------------------------------------------------------
		        lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
		        lgstrData = lgstrData & Chr(11) & Chr(12)
		        
		        ReDim Preserve TmpBuffer(0)
				TmpBuffer(0) = lgstrData
		        
		        iDx =  iDx + 1

			End If   

		    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
		    Call SubCloseRs(lgObjRs) 
		    
		Else
			Call SubMakeSQLStatements("B_CK",strPlantCd,strItemCd,strBomNo,"","")           '☜ : Make sql statements
		 
		    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		        Call SetErrorStatus()
		    Else
		    	
				IntRetCD = 1

				 Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

		        lgstrData = ""
		        iDx       = 1

		        lgstrData = lgstrData & Chr(11)	& ""							'설계변경구분 
		        lgstrData = lgstrData & Chr(11)									'설계변경일		
		        lgstrData = lgstrData & Chr(11)	& "0"							'레벨 
		        lgstrData = lgstrData & Chr(11)									'순서 
				lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("ITEM_CD"))		'자품목코드 
		        lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_NM")			'자품목명 
		        lgstrData = lgstrData & Chr(11) & lgObjRs("SPEC")				'규격 
		        lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_ACCT_NM")		'품목계정명  
		        lgstrData = lgstrData & Chr(11) & lgObjRs("PROCUR_TYPE_NM")		'조달구분명 
		        lgstrData = lgstrData & Chr(11)									'자품목기준수 
		        lgstrData = lgstrData & Chr(11)									'단위 
		        lgstrData = lgstrData & Chr(11)									'모품목기준수 
		        lgstrData = lgstrData & Chr(11)									'단위 
		        lgstrData = lgstrData & Chr(11)									'안전L/T
		        lgstrData = lgstrData & Chr(11)									'loss율 
		        lgstrData = lgstrData & Chr(11)									'유무상구분명 
		        lgstrData = lgstrData & Chr(11)									'시작일 
		        lgstrData = lgstrData & Chr(11)									'종료일		
		        lgstrData = lgstrData & Chr(11)									'설계변경번호 
		        lgstrData = lgstrData & Chr(11)									'설계변경내용 
		        lgstrData = lgstrData & Chr(11)									'설계변경근거명 
		        lgstrData = lgstrData & Chr(11) & lgObjRs("DRAWING_PATH")		'도면경로 
		        lgstrData = lgstrData & Chr(11)									'비고 
		'------ Developer Coding part (End   ) ------------------------------------------------------------------
		        lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
		        lgstrData = lgstrData & Chr(11) & Chr(12)
				
				ReDim Preserve TmpBuffer(0)
				TmpBuffer(0) = lgstrData
				
		        iDx =  iDx + 1

			End If   
		    
		    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
		    Call SubCloseRs(lgObjRs) 
		End If	 
	
	End If
		     
	'========================================================================
	' BOM 전개를 실시한다.(하위품목)
	'========================================================================
	
	lgStrSQL = ""
	strPlantCd = FilterVar(Trim(Request("txtPlantCd"))	,"''", "S")			
	
	Call SubMakeSQLStatements("M",strPlantCd,strSpId,"","","")					'☜ : Make sql statements

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		
		lgStrPrevKeyIndex = ""    
		
		If strBomNo = FilterVar("", "''", "S") Then
			IntRetCD = -1
			Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
			Call SetErrorStatus()
			Response.End
		End If
	
	Else
		IntRetCD = 1
		Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)
		iDx       = 2							'먼저 뿌려지는 최상위품목 다음에 나오도록..

        Do While Not lgObjRs.EOF
			
			'-----------------------
			' Level Setting
			'-----------------------
			strLevel = ""
			iLevelCnt = lgObjRs("LEVEL_CD")
		
			For iIntCnt = 1 To iLevelCnt
				strLevel = strLevel & "."
			Next 
			
			If Trim(lgObjRs("ACTION_NM")) = "" Then
				sModifiedDate = ""
			Else
				sModifiedDate = UNIDateClientFormat(lgObjRs("MODIFIED_DATE"))
			End If

			strLevel = strLevel & lgObjRs("LEVEL_CD")
			
			lgstrData = ""

	        lgstrData = lgstrData & Chr(11)	& lgObjRs("ACTION_NM")					'설계변경구분 
	        lgstrData = lgstrData & Chr(11)	& sModifiedDate							'설계변경일						
	        lgstrData = lgstrData & Chr(11) & strLevel								'레벨 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("CHILD_ITEM_SEQ")				'순서 
			lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("CHILD_ITEM_CD"))		'자품목코드 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_NM")					'자품목명 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("SPEC")						'규격 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_ACCT_NM")				'계정명 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("PROCUR_TYPE_NM")				'조달구분명 
	        lgstrData = lgstrData & Chr(11)	& UniConvNumberDBToCompany(lgObjRs("CHILD_ITEM_QTY"), 6, 3, "", 0)	'자품목기준수 
	        lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("CHILD_ITEM_UNIT"))		'단위 
	        lgstrData = lgstrData & Chr(11)	& UniConvNumberDBToCompany(lgObjRs("PRNT_ITEM_QTY"), 6, 3, "", 0)	'모품목기준수 
	        lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("PRNT_ITEM_UNIT"))		'단위 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("SAFETY_LT")					'안전L/T
	        lgstrData = lgstrData & Chr(11)	& UniConvNumberDBToCompany(lgObjRs("LOSS_RATE"), 3, 3, "", 0)	'loss율 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("SUPPLY_TYPE_NM")				'유무상구분명 
	        lgstrData = lgstrData & Chr(11)	& UNIDateClientFormat(lgObjRs("ITEM_F_DT"))		'시작일 
	        lgstrData = lgstrData & Chr(11)	& UNIDateClientFormat(lgObjRs("ITEM_T_DT"))		'종료일		
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("ECN_NO")						'설계변경번호 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("ECN_DESC")					'설계변경내용 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("REASON_NM")					'설계변경근거 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("DRAWING_PATH")				'Drawing Path
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("REMARK")						'비고 
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
	        lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
	        lgstrData = lgstrData & Chr(11) & Chr(12)

	    	lgObjRs.MoveNext
	 
			ReDim Preserve TmpBuffer(iDx-1)
			TmpBuffer(iDx-1) = lgstrData
	        iDx =  iDx + 1
	        If iDx > lgMaxCount + 1 Then		'최상위 품목 row가 처음에 조회되고 나서 groupview가 조회되므로 
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
	
	iTotalStr = Join(TmpBuffer, "")
	
    lgStrSQL = ""
	'-------------------------
	' 생성된 temp table 삭제 
	'-------------------------
    lgStrSQL = "DELETE FROM p_bom_history_for_exp "
	lgStrSQL = lgStrSQL & " WHERE plant_cd = " & FilterVar(Trim(Request("txtPlantCd"))	,"''", "S")
	lgStrSQL = lgStrSQL & " AND user_id = " & strSpId
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords 
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub    

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2,pCode3,pCode4)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
	Dim iSelCount
	
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Select Case pDataType

		Case "M"
			lgStrSQL = "SELECT ( CASE A.ACTION_FLG WHEN 'A' THEN 'Add' WHEN 'D' THEN 'Delete' WHEN 'C' THEN 'Change' ELSE '' END ) ACTION_NM, a.*, b.ITEM_NM, b.PHANTOM_FLG, b.SPEC, b.BASIC_UNIT, c.ITEM_ACCT, dbo.ufn_GetCodeName('P1001', c.ITEM_ACCT) ITEM_ACCT_NM, dbo.ufn_GetCodeName('P1003', c.PROCUR_TYPE) PROCUR_TYPE_NM, dbo.ufn_GetCodeName('M2201', a.SUPPLY_TYPE) SUPPLY_TYPE_NM, a.VALID_FROM_DT ITEM_F_DT, a.VALID_TO_DT ITEM_T_DT, g.PROCUR_TYPE PRNT_PROC_TYPE, "
			lgStrSQL = lgStrSQL & " A.ECN_NO, d.ECN_DESC, d.REASON_CD, dbo.ufn_GetCodeName('P1402', d.REASON_CD) REASON_NM, f.DRAWING_PATH, A.REMARK "
			lgStrSQL = lgStrSQL & " FROM P_BOM_HISTORY_FOR_EXP a LEFT OUTER JOIN P_ECN_MASTER d ON a.ECN_NO = d.ECN_NO, P_BOM_HISTORY_FOR_EXP aa LEFT OUTER JOIN P_BOM_HEADER f ON (aa.PLANT_CD = f.PLANT_CD AND aa.CHILD_ITEM_CD = f.ITEM_CD AND aa.PRNT_BOM_NO = f.BOM_NO), B_ITEM b, B_ITEM_BY_PLANT c, B_ITEM_BY_PLANT g"
			lgStrSQL = lgStrSQL & " WHERE (a.PLANT_CD = aa.PLANT_CD AND a.USER_ID = aa.USER_ID AND a.SEQ = aa.SEQ)"
			lgStrSQL = lgStrSQL & " AND a.CHILD_ITEM_CD = b.ITEM_CD AND b.ITEM_CD = c.ITEM_CD AND a.PLANT_CD = c.PLANT_CD "
			lgStrSQL = lgStrSQL & " AND g.PLANT_CD = a.PLANT_CD AND g.ITEM_CD = a.PRNT_ITEM_CD " 
			lgStrSQL = lgStrSQL & " AND a.MODIFIED_DATE > " & FilterVar("1900-01-01", "''", "S") 
			lgStrSQL = lgStrSQL & " AND a.PLANT_CD = " & pCode
			lgStrSQL = lgStrSQL & " AND a.USER_ID = " & pCode1
			If strChildItemCd <> FilterVar("", "''", "S") Then
				lgStrSQL = lgStrSQL & " AND a.CHILD_ITEM_CD = " & strChildItemCd
			End If
			If Trim(strECNNo) <> FilterVar("", "''", "S") Then
				lgStrSQL = lgStrSQL & " AND a.ECN_NO = " & strECNNo
			End If
			If strChgFromDt <> "''" Then
				lgStrSQL = lgStrSQL & " AND a.MODIFIED_DATE >= " & strChgFromDt
			End If
			If strChgToDt <> "''" Then
				lgStrSQL = lgStrSQL & " AND a.MODIFIED_DATE <= " & strChgToDt
			End If
			
			
			lgStrSQL = lgStrSQL & " ORDER BY a.own_node "' a.SEQ "

		Case "B_CK"
			lgStrSQL = "SELECT a.*, b.item_acct, b.procur_type, c.item_nm, c.spec, d.minor_nm  ,e.minor_nm, c.basic_unit, b.valid_from_Dt, b.valid_to_dt, dbo.ufn_GetCodeName('P1001', B.ITEM_ACCT) ITEM_ACCT_NM, B.PROCUR_TYPE, dbo.ufn_GetCodeName('P1003', B.PROCUR_TYPE) PROCUR_TYPE_NM "
			lgStrSQL = lgStrSQL & " FROM p_bom_header a, b_item_by_plant b, b_item c, b_minor d, b_minor e  "
			lgStrSQL = lgStrSQL & " WHERE a.plant_cd = b.plant_cd and a.item_cd = b.item_cd and b.item_cd = c.item_cd and b.item_acct = d.minor_cd and d.major_cd='p1001' and b.procur_type = e.minor_cd and e.major_cd='p1003' "
			lgStrSQL = lgStrSQL & " AND a.plant_cd = " & pCode 
			lgStrSQL = lgStrSQL & " AND a.item_cd = " & pCode1
			lgStrSQL = lgStrSQL & " AND a.bom_no= " & pCode2
			
		Case "BT_CK"
			lgStrSQL = "SELECT * FROM b_minor WHERE major_cd = " & FilterVar("P1401", "''", "S") & ""
			lgStrSQL = lgStrSQL & " AND minor_cd = " & pCode 

		Case "I_CK"
			lgStrSQL = "SELECT b.item_nm, b.spec, c.minor_nm, b.phantom_flg, b.basic_unit, a.valid_from_dt, a.valid_to_dt, a.item_cd, d.minor_nm"
			lgStrSQL = lgStrSQL & " FROM b_item_by_plant a, b_item b, b_minor c, b_minor d"
			lgStrSQL = lgStrSQL & " WHERE a.item_cd = b.item_cd and c.minor_cd = a.item_acct and c.major_cd ='p1001' and d.minor_cd = a.procur_type and d.major_cd='p1003'"
			lgStrSQL = lgStrSQL & " AND a.plant_cd = " & pCode 
			lgStrSQL = lgStrSQL & " AND a.item_cd = " & pCode1

		Case "P_CK"
			lgStrSQL = "SELECT * FROM b_plant where plant_cd = " & pCode 

		Case "ECN_CK"
			lgStrSQL = "SELECT * FROM p_ecn_master where ecn_no = " & pCode 

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
    Dim strFromDt
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
   
    With lgObjComm
        .CommandText = "usp_BOM_history_exp_main"
        .CommandType = adCmdStoredProc
        
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@srch_type",	advarXchar,adParamInput,2, Request("rdoSrchType"))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_cd",	advarXchar,adParamInput,4, Request("txtPlantCd"))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@par_item_cd",	advarXchar,adParamInput,18, Request("txtItemCd"))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@par_bom_no",advarXchar,adParamInput,4, Request("txtBomNo"))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@chg_from_dt_s",	advarXchar,adParamInput,10,UniConvDate(Request("txtChgFromDt")))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@base_dt_s",	advarXchar,adParamInput,10,UniConvDate(Request("txtChgToDt")))
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
	        Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
	        Response.Write ".lgStrPrevKeyIndex = """ & lgStrPrevKeyIndex & """" & vbCrLf
	        Response.Write ".ggoSpread.SSShowDataByClip """ & ConvSPChars(iTotalStr) & """" & vbCrLf
		
	        Response.Write ".DBQueryOk(" & lgLngMaxRow & " + 1)" & vbCrLf
	     Response.Write "End With" & vbCrLf
	End If   
Response.Write "</Script>" & vbCrLf
%>	