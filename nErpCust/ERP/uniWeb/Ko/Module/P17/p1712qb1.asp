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
lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
'------ Developer Coding part (Start ) ------------------------------------------------------------------
Dim IntRetCD
Dim strPlantCd
Dim strItemCd
Dim strBomNo
Dim strBaseDt
Dim strExpFlg
Dim BaseDt
Dim i, iDx

Dim TmpBuffer
Dim iTotalStr
	
Dim strSpId
Dim strLevel

ReDim TmpBuffer(0)
	
strPlantCd = FilterVar(Trim(Request("txtPlantCd"))	,"''", "S")
strItemCd = FilterVar(Trim(Request("txtItemCd"))	,"''", "S")
strBomNo = FilterVar(Trim(Request("txtBomNo"))	,"''", "S")
	
'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	
BaseDt = FilterVar(UNIConvYYYYMMDDToDate(gAPDateFormat,"1900","01","01"),"''","S")
strBaseDt = FilterVar(Trim(Request("txtBaseDt"))	, BaseDt, "D")
strExpFlg = Trim(Request("rdoSrchType"))

Call SubCreateCommandObject(lgObjComm)

If strExpFlg = "1" Or strExpFlg = "2" Then				'정전개 혹은 bom no를 갖는 역전개 
	Call SubBizQuery("B_CK")
	Call SubBizBatch()
	Call SubBizQueryMulti()
			
ElseIf strExpFlg = "3" Or strExpFlg = "4" Then		'bom no를 갖지 않는 역전개 
	Call SubBizQuery("CK")
	Call SubBizBatch()
	Call SubBizQueryMulti()

Else
	Call DisplayMsgBox("182600", vbOKOnly, "", "", I_MKSCRIPT)
	Response.End 
End If

Call SubCloseCommandObject(lgObjComm)
      
Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery(pOpCode)

	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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

	'------------------
	'품목체크 
	'------------------
	lgStrSQL = ""
	Call SubMakeSQLStatements("I_CK",strPlantCd,strItemCd,"","","")           '☜ : Make sql statements

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		    
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
			
	'------------------
	' bom type 체크 
	'------------------
	lgStrSQL = ""
			
	Call SubMakeSQLStatements("BT_CK",strBomNo,"","","","")           '☜ : Make sql statements
			
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		    
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
	End If
		
	Call SubCloseRs(lgObjRs) 
		
	If pOpCode = "B_CK" Then	
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
			Response.Write "</Script>" & vbCrLf
			Response.End
		Else
			IntRetCD = 1
			Response.Write "<Script Language = VBScript>" & vbCrLf
				Response.Write "With Parent" & vbCrLf
					Response.Write ".Frm1.txtItemNm.Value = """ & ConvSPChars(Trim(lgObjRs(14))) & """" & vbCrLf
					Response.Write ".frm1.cboItemAcct.value = """ & ConvSPChars(lgObjRs("item_acct")) & """" & vbCrLf
					Response.Write ".frm1.txtItemAcct.value = """ & ConvSPChars(Trim(lgObjRs(12))) & """" & vbCrLf
					Response.Write ".frm1.txtSpec.value = """ & ConvSPChars(lgObjRs(15)) & """" & vbCrLf
					Response.Write ".frm1.txtPlantItemFromDt.text = """ & UNIDateClientFormat(lgObjRs(19)) & """" & vbCrLf
					Response.Write ".frm1.txtPlantItemToDt.text = """ & UNIDateClientFormat(lgObjRs(20)) & """" & vbCrLf
					Response.Write ".frm1.txtBasicUnit.value = """ & ConvSPChars(Trim(lgObjRs(18))) & """" & vbCrLf
			    Response.Write "End With" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
				
		Call SubCloseRs(lgObjRs) 
		
	ElseIf pOpCode = "CK" Then
	
		lgStrSQL = ""
		Call SubMakeSQLStatements("CK",strPlantCd,strItemCd,strBomNo,"","")           '☜ : Make sql statements
					
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
				    
			IntRetCD = -1
						
			Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
			Call SetErrorStatus()
			Response.Write "<Script Language = VBScript>" & vbCrLf
				Response.Write "parent.frm1.hBomType.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			Response.End
		Else
			IntRetCD = 1
			Response.Write "<Script Language = VBScript>" & vbCrLf
				Response.Write "With Parent" & vbCrLf
					Response.Write ".Frm1.txtItemNm.Value = """ & ConvSPChars(Trim(lgObjRs("item_nm"))) & """" & vbCrLf
					Response.Write ".frm1.cboItemAcct.value = """ & ConvSPChars(lgObjRs("item_acct")) & """" & vbCrLf
					Response.Write ".frm1.txtItemAcct.value = """ & ConvSPChars(Trim(lgObjRs("minor_nm"))) & """" & vbCrLf
					Response.Write ".frm1.txtSpec.value = """ & ConvSPChars(lgObjRs("spec")) & """" & vbCrLf
					Response.Write ".frm1.txtPlantItemFromDt.text = """ & UNIDateClientFormat(lgObjRs("valid_from_Dt")) & """" & vbCrLf
					Response.Write ".frm1.txtPlantItemToDt.text = """ & UNIDateClientFormat(lgObjRs("valid_to_dt")) & """" & vbCrLf
					Response.Write ".frm1.txtBasicUnit.value = """ & ConvSPChars(Trim(lgObjRs("basic_unit"))) & """" & vbCrLf
			    Response.Write "End With" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
				
		Call SubCloseRs(lgObjRs) 
		
	End If
    
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
			
		        lgstrData = lgstrData & Chr(11)	& "0"						'레벨 
		        lgstrData = lgstrData & Chr(11)								'순서 
				lgstrData = lgstrData & Chr(11) & Trim(lgObjRs(1))			'자품목코드 
		        lgstrData = lgstrData & Chr(11) & lgObjRs(14)				'자품목명 
		        lgstrData = lgstrData & Chr(11) & lgObjRs(15)				'규격 
		        lgstrData = lgstrData & Chr(11) & Trim(lgObjRs(18))			'단위 
		        lgstrData = lgstrData & Chr(11) & lgObjRs(16)				'품목계정명 
		        lgstrData = lgstrData & Chr(11) & lgObjRs(17)				'조달구분명 
		        lgstrData = lgstrData & Chr(11) & lgObjRs(2)				'bom type
		        lgstrData = lgstrData & Chr(11)								'자품목기준수 
		        lgstrData = lgstrData & Chr(11)								'단위 
		        lgstrData = lgstrData & Chr(11)								'모품목기준수 
		        lgstrData = lgstrData & Chr(11)								'단위 
		        lgstrData = lgstrData & Chr(11)								'안전L/T
		        lgstrData = lgstrData & Chr(11)								'loss율 
		        lgstrData = lgstrData & Chr(11)								'유무상구분명 
		        lgstrData = lgstrData & Chr(11)								'시작일 
		        lgstrData = lgstrData & Chr(11)								'종료일		
		        lgstrData = lgstrData & Chr(11)								'설계변경번호 
		        lgstrData = lgstrData & Chr(11)								'설계변경내용 
		        lgstrData = lgstrData & Chr(11)								'설계변경근거 
		        lgstrData = lgstrData & Chr(11) & lgObjRs("DRAWING_PATH")	'Drawing Path
		        lgstrData = lgstrData & Chr(11)								'비고 
		        lgstrData = lgstrData & Chr(11)								'모품목명 
		        lgstrData = lgstrData & Chr(11) 							'모품목bom no
		        lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
		        lgstrData = lgstrData & Chr(11) & Chr(12)
		        
		        ReDim Preserve TmpBuffer(0)
				TmpBuffer(0) = lgstrData
		        
		        iDx =  iDx + 1

			End If   

		    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
		    Call SubCloseRs(lgObjRs) 
		    
		Else
			Call SubMakeSQLStatements("I_CK",strPlantCd,strItemCd,strBomNo,"","")           '☜ : Make sql statements
		 
		    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		        Call SetErrorStatus()
		    Else
		    	
				IntRetCD = 1

				 Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

		        lgstrData = ""
		        iDx       = 1
		
		        lgstrData = lgstrData & Chr(11)	& "0"							'레벨 
		        lgstrData = lgstrData & Chr(11)									'순서 
				lgstrData = lgstrData & Chr(11) & Trim(lgObjRs(7))				'자품목코드 

		        lgstrData = lgstrData & Chr(11) & lgObjRs(0)					'자품목명 
		        lgstrData = lgstrData & Chr(11) & lgObjRs(1)					'규격 
		        lgstrData = lgstrData & Chr(11) & Trim(lgObjRs(4))				'단위 
		   
		        lgstrData = lgstrData & Chr(11) & lgObjRs(2)					'계정명 
		        lgstrData = lgstrData & Chr(11) & lgObjRs(8)					'조달구분명 
		        lgstrData = lgstrData & Chr(11)	& UCase(Trim(Request("txtBomNo")))	'bom type	'2003-09-08
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
		        lgstrData = lgstrData & Chr(11)									'설계변경근거 
		        lgstrData = lgstrData & Chr(11) & lgObjRs("DRAWING_PATH")		'Drawing Path
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
		End If	 
	
	End IF
		     
	'========================================================================
	' BOM 전개를 실시한다.(하위품목)
	'========================================================================
	
	lgStrSQL = ""
	strPlantCd = FilterVar(Trim(Request("txtPlantCd"))	,"''", "S")			
	
	Call SubMakeSQLStatements("M",strPlantCd,strSpId,"","","")					'☜ : Make sql statements

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		
		lgStrPrevKeyIndex = ""    
		
		If strBomNo = "''" Then
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
			iLevelCnt = lgObjRs(6)
		
			For iIntCnt = 1 To iLevelCnt
				strLevel = strLevel & "."
			Next 
		
			strLevel = strLevel & lgObjRs(6)
			
			lgstrData = ""
				
	        lgstrData = lgstrData & Chr(11) & strLevel								'레벨 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("CHILD_ITEM_SEQ")				'순서 
			lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("CHILD_ITEM_CD"))		'자품목코드 

	        lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_NM")					'자품목명 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("SPEC")						'규격 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("BASIC_UNIT")					'단위 
	   
	        lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_ACCT_NM")				'계정명 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("PROCUR_TYPE_NM")				'조달구분명 
	        lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("PRNT_BOM_NO"))			'bom type

	        lgstrData = lgstrData & Chr(11)	& UniNumClientFormat(lgObjRs("CHILD_ITEM_QTY"), "", 0)	'자품목기준수 
	        lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("CHILD_ITEM_UNIT"))		'단위 

	        lgstrData = lgstrData & Chr(11)	& UniNumClientFormat(lgObjRs("PRNT_ITEM_QTY"), "", 0)	'모품목기준수 
	        lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("PRNT_ITEM_UNIT"))		'단위 

	        lgstrData = lgstrData & Chr(11)	& lgObjRs("SAFETY_LT")					'안전L/T
	        lgstrData = lgstrData & Chr(11)	& UniNumClientFormat(lgObjRs("LOSS_RATE"), ggQty.DecPoint, 0)	'loss율 

	        lgstrData = lgstrData & Chr(11)	& lgObjRs("SUPPLY_TYPE_NM")				'유무상구분명 
	        lgstrData = lgstrData & Chr(11)	& UNIDateClientFormat(lgObjRs("ITEM_F_DT"))		'시작일 
	        lgstrData = lgstrData & Chr(11)	& UNIDateClientFormat(lgObjRs("ITEM_T_DT"))		'종료일		
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("ECN_NO")						'설계변경번호 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("ECN_DESC")					'설계변경내용 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("REASON_NM")					'설계변경근거 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("DRAWING_PATH")				'Drawing Path
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("REMARK")						'비고 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("PRNT_ITEM_CD")				'모품목 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("PRNT_BOM_NO")				'모품목bom no
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
    lgStrSQL = "DELETE FROM p_bom_for_explosion "
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
    Dim iSelCount
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case pDataType

		Case "M"
			lgStrSQL = "SELECT a.*, b.ITEM_NM, b.PHANTOM_FLG, b.SPEC, b.BASIC_UNIT, c.ITEM_ACCT, dbo.ufn_GetCodeName('P1001', c.ITEM_ACCT) ITEM_ACCT_NM, dbo.ufn_GetCodeName('P1003', c.PROCUR_TYPE) PROCUR_TYPE_NM, dbo.ufn_GetCodeName('M2201', a.SUPPLY_TYPE) SUPPLY_TYPE_NM, a.VALID_FROM_DT ITEM_F_DT, a.VALID_TO_DT ITEM_T_DT, g.PROCUR_TYPE PRNT_PROC_TYPE, "
			lgStrSQL = lgStrSQL & " A.ECN_NO, d.ECN_DESC, d.REASON_CD, dbo.ufn_GetCodeName('P1402', d.REASON_CD) REASON_NM, f.DRAWING_PATH, A.REMARK "
			lgStrSQL = lgStrSQL & " FROM P_BOM_FOR_EXPLOSION a LEFT OUTER JOIN P_ECN_MASTER d ON a.ECN_NO = d.ECN_NO, P_BOM_FOR_EXPLOSION aa LEFT OUTER JOIN P_BOM_HEADER f ON (aa.PLANT_CD = f.PLANT_CD AND aa.CHILD_ITEM_CD = f.ITEM_CD AND aa.PRNT_BOM_NO = f.BOM_NO), B_ITEM b, B_ITEM_BY_PLANT c, B_ITEM_BY_PLANT g"
			lgStrSQL = lgStrSQL & " WHERE (a.PLANT_CD = aa.PLANT_CD AND a.USER_ID = aa.USER_ID AND a.SEQ = aa.SEQ)"
			lgStrSQL = lgStrSQL & " AND a.CHILD_ITEM_CD = b.ITEM_CD AND b.ITEM_CD = c.ITEM_CD AND a.PLANT_CD = c.PLANT_CD "
			lgStrSQL = lgStrSQL & " AND g.PLANT_CD = a.PLANT_CD AND g.ITEM_CD = a.PRNT_ITEM_CD " 
			lgStrSQL = lgStrSQL & " AND a.PLANT_CD = " & pCode
			lgStrSQL = lgStrSQL & " AND a.USER_ID = " & pCode1
			lgStrSQL = lgStrSQL & " ORDER BY a.SEQ "

		Case "B_CK"
			lgStrSQL = "SELECT a.*, b.item_acct, b.procur_type, c.item_nm, c.spec, d.minor_nm  ,e.minor_nm, c.basic_unit, b.valid_from_Dt, b.valid_to_dt "
			lgStrSQL = lgStrSQL & " FROM p_bom_header a, b_item_by_plant b, b_item c, b_minor d, b_minor e  "
			lgStrSQL = lgStrSQL & " WHERE a.plant_cd = b.plant_cd and a.item_cd = b.item_cd and b.item_cd = c.item_cd and b.item_acct = d.minor_cd and d.major_cd='p1001' and b.procur_type = e.minor_cd and e.major_cd='p1003' "
			lgStrSQL = lgStrSQL & " AND a.plant_cd = " & pCode 
			lgStrSQL = lgStrSQL & " AND a.item_cd = " & pCode1
			lgStrSQL = lgStrSQL & " AND a.bom_no= " & pCode2
			
		Case "CK"
			lgStrSQL = "SELECT b.item_acct, b.procur_type, c.item_nm, c.spec, d.minor_nm  ,e.minor_nm, c.basic_unit, b.valid_from_Dt, b.valid_to_dt "
			lgStrSQL = lgStrSQL & " FROM b_item_by_plant b, b_item c, b_minor d, b_minor e  "
			lgStrSQL = lgStrSQL & " WHERE b.item_cd = c.item_cd and b.item_acct = d.minor_cd and d.major_cd='p1001' and b.procur_type = e.minor_cd and e.major_cd='p1003' "
			lgStrSQL = lgStrSQL & " AND b.plant_cd = " & pCode 
			lgStrSQL = lgStrSQL & " AND b.item_cd = " & pCode1
			
		Case "BT_CK"
			lgStrSQL = "SELECT * FROM b_minor WHERE major_cd = 'P1401'"
			lgStrSQL = lgStrSQL & " AND minor_cd = " & pCode 

		Case "I_CK"
			lgStrSQL = "SELECT b.item_nm, b.spec, c.minor_nm, b.phantom_flg, b.basic_unit, a.valid_from_dt, a.valid_to_dt, a.item_cd, d.minor_nm"
			lgStrSQL = lgStrSQL & " FROM b_item_by_plant a, b_item b, b_minor c, b_minor d"
			lgStrSQL = lgStrSQL & " WHERE a.item_cd = b.item_cd and c.minor_cd = a.item_acct and c.major_cd ='p1001' and d.minor_cd = a.procur_type and d.major_cd='p1003'"
			lgStrSQL = lgStrSQL & " AND a.plant_cd = " & pCode 
			lgStrSQL = lgStrSQL & " AND a.item_cd = " & pCode1

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
	
    With lgObjComm
        .CommandText = "usp_BOM_explode_main"
        .CommandType = adCmdStoredProc
        
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@srch_type",	advarXchar,adParamInput,2, Request("rdoSrchType"))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_cd",	advarXchar,adParamInput,4, Request("txtPlantCd"))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@par_item_cd",	advarXchar,adParamInput,18, Request("txtItemCd"))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@par_bom_no",advarXchar,adParamInput,4, Request("txtBomNo"))
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
	        Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
	        Response.Write ".lgStrPrevKeyIndex = """ & lgStrPrevKeyIndex & """" & vbCrLf
	        Response.Write ".ggoSpread.SSShowDataByClip """ & ConvSPChars(iTotalStr) & """" & vbCrLf
		
	        Response.Write ".DBQueryOk(" & lgLngMaxRow & " + 1)" & vbCrLf
	     Response.Write "End With" & vbCrLf
	End If   
Response.Write "</Script>" & vbCrLf
%>	