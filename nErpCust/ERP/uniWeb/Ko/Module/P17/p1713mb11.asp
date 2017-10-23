<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/AdoVbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDB.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next
Err.Clear 

Call HideStatusWnd   

Call LoadBasisGlobalInf
Call loadInfTB19029B("Q", "P", "NOCOOKIE", "MB")

'---------------------------------------Common-----------------------------------------------------------
lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    
lgMaxCount        = 500									'2004-03-18
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")), 0)   
    
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '☜: Set to space
lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
'------ Developer Coding part (Start ) ------------------------------------------------------------------

Dim strDestPlantCd
Dim strItemCd
Dim strBomType
Dim i

Const C_Level				= 1
Const C_Seq					= 2
Const C_ChildItemCd			= 3
Const C_ChildItemPopUp		= 4
Const C_ChildItemNm			= 5
Const C_Spec				= 6
Const C_ChildItemUnit		= 7
Const C_ItemAcct			= 8
Const C_ItemAcctNm			= 9
Const C_ProcType			= 10
Const C_ProcTypeNm			= 11
Const C_BomType				= 12
Const C_BomTypePopup		= 13
Const C_ChildItemBaseQty	= 14
Const C_ChildBasicUnit		= 15
Const C_ChildBasicUnitPopup	= 16
Const C_PrntItemBaseQty		= 17
Const C_PrntBasicUnit		= 18
Const C_PrntBasicUnitPopup	= 19
Const C_SafetyLT			= 20
Const C_LossRate			= 21
Const C_SupplyFlg			= 22
Const C_SupplyFlgNm			= 23
Const C_ValidFromDt			= 24
Const C_ValidToDt			= 25
Const C_ECNNo				= 26
Const C_ECNNoPopup			= 27
Const C_ECNDesc				= 28
Const C_ReasonCd			= 29
Const C_ReasonCdPopup		= 30
Const C_ReasonNm			= 31
Const C_DrawingPath			= 32
Const C_Remark				= 33
Const C_HdrItemCd			= 34
Const C_HdrBomNo			= 35
Const C_HdrProcType			= 36
Const C_ItemValidFromDt		= 37
Const C_ItemValidToDt		= 38
Const C_ItemAcctGrp			= 39
Const C_ReqTransNo			= 40
Const C_ReqTransDt			= 41
Const C_TransStatus			= 42
Const C_TransDt				= 43
Const C_Row					= 44		


strDestPlantCd 	= Trim(Request("txtDestPlantCd"))
strItemCd 		= Trim(Request("txtItemCd"))
strBomType 		= Trim(Request("txtBomType")) 



Call SubOpenDB(lgObjConn)
	Call SubBizQuery("P")	'공장체크 
	Call SubBizQuery("I")	'품목체크 
	Call SubBizQuery("S")	'생산BOM QUERY (MaxSeq)
	Call SubBizQuery("B")	'생산BOM QUERY
Call SubCloseDB(lgObjConn) 

Response.Write "<Script Language = VBScript>" & vbCrLf                                                      '☜ : Query
	If Trim(lgErrorStatus) = "NO" And IntRetCd <> -1 Then
		 Response.Write "With Parent" & vbCrLf
            Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
            Response.Write ".ggoSpread.SSShowDataByClip """ & ConvSPChars(lgstrData) & """" & vbCrLf
			Response.Write ".lgIntFlgMode = .parent.OPMD_CMODE"  & vbCrLf
			Response.Write ".frm1.txtReqTransNo.value = """""  & vbCrLf
			Response.Write ".frm1.txtReqTransNo2.value = """""  & vbCrLf
			Response.Write ".frm1.hReqTransNo.value = """""  & vbCrLf
			Response.Write ".frm1.hReqTransDt.value = """""  & vbCrLf
			Response.Write ".frm1.btnCopy.disabled = False" & vbCrLf
			Response.Write ".frm1.hStatus.value = ""N"""  & vbCrLf
			Response.Write ".frm1.hReqTransNo.value = """""  & vbCrLf     
			       
			Response.Write " If .frm1.vspdData.MaxRows > 0 Then"  & vbCrLf

			Response.Write " 	For i = 1 To .frm1.vspdData.MaxRows"  & vbCrLf
			Response.Write " 		.frm1.vspdData.Row = i "  & vbCrLf
	    	Response.Write " 		.frm1.vspdData.Col = 0"  & vbCrLf
	    	Response.Write " 		Call .frm1.vspdData.SetText(0, i, ""입력"")"  & vbCrLf
			Response.Write " 	Next"  & vbCrLf
			Response.Write " End If"  & vbCrLf			
            
            Response.Write "Call .SetSpreadColor (2, .frm1.vspddata.MaxRows, 1, 1)" & vbCrLf
            Response.Write "Call .SetSpreadColor (1, 1, 0, 1)" & vbCrLf
			Response.Write "Call .SetToolbar(""11101101000011"")" & vbCrLf	
         Response.Write "End with" & vbCrLf
    End If 
Response.Write "</Script>" & vbCrLf
Response.End

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery(pOpCode)
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	Select Case pOpCode
		
		Case "P"		'공장체크 
			lgStrSQL = "SELECT 1 FROM B_PLANT WHERE PLANT_CD = " & FilterVar(strDestPlantCd, "''", "S")
			
			If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
				Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Response.End
			End If
				
			Call SubCloseRs(lgObjRs) 			
		
		Case "I"		'품목체크 
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & " SELECT 1 FROM B_ITEM_BY_PLANT "
			lgStrSQL = lgStrSQL & " WHERE PLANT_CD = " & FilterVar(strDestPlantCd, "''", "S") & "AND "
			lgStrSQL = lgStrSQL & " ITEM_CD = " & FilterVar(strItemCd, "''", "S")
			
			If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    	'If data not exists
				Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)      				'☜ : No data is found. 
				Response.End
			End If
				
			Call SubCloseRs(lgObjRs) 			

		Case "S"		'Seq의 Max Value
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & " SELECT Max(CHILD_ITEM_SEQ) MAXSEQ FROM UV_Y_M_BOM_HDR_DTL "
			lgStrSQL = lgStrSQL & " WHERE PLANT_CD = " & FilterVar(strDestPlantCd, "''", "S") & "AND "
			lgStrSQL = lgStrSQL & " PRNT_ITEM_CD = " & FilterVar(strItemCd, "''", "S") & "AND "
			lgStrSQL = lgStrSQL & " PRNT_BOM_NO = " & FilterVar(strBomType, "''", "S")
			
		    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		    	Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "With Parent" & vbCrLf
						Response.Write ".frm1.hMaxSeq.value = """"" & vbCrLf
					Response.Write "End With" & vbCrLf
				Response.Write "</Script>" & vbCrLf
			Else
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "With Parent" & vbCrLf
						Response.Write ".frm1.hMaxSeq.value = """ & lgObjRs("MAXSEQ") & """" & vbCrLf
				    Response.Write "End With" & vbCrLf
				Response.Write "</Script>" & vbCrLf
			End If			
				
			Call SubCloseRs(lgObjRs) 
				
		Case "B"		'생산BOM QUERY
			lgStrSQL = ""
			lgStrSQL = lgStrSQL & " SELECT * FROM UV_Y_M_BOM_HDR_DTL "
			lgStrSQL = lgStrSQL & " WHERE PLANT_CD = " & FilterVar(strDestPlantCd, "''", "S") & "AND "
			lgStrSQL = lgStrSQL & " PRNT_ITEM_CD = " & FilterVar(strItemCd, "''", "S") & "AND "
			lgStrSQL = lgStrSQL & " PRNT_BOM_NO = " & FilterVar(strBomType, "''", "S")
			lgStrSQL = lgStrSQL & " ORDER BY 1, 2, 3"

			If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                 'If data not exists
				Call DisplayMsgBox("182600", vbInformation, "", "", I_MKSCRIPT)      				'☜ : No data is found. 
				Call SubCloseRs(lgObjRs)  
				Response.End
			Else
		
				IntRetCD = 1
				lgstrData = ""
				i = 0
				
				Do While Not lgObjRs.EOF
		        	lgstrData = lgstrData & Chr(11) & lgObjRs("LEVEL_CD")
			        lgstrData = lgstrData & Chr(11)	& lgObjRs("CHILD_ITEM_SEQ")						'순서 
					lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("CHILD_ITEM_CD"))	& Chr(11) 	'자품목코드, 자품목팝업 
			        lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_NM")							'자품목명 
			        lgstrData = lgstrData & Chr(11) & lgObjRs("SPEC")								'규격 
			        lgstrData = lgstrData & Chr(11) & lgObjRs("BASIC_UNIT")							'기준단위 
			        lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_ACCT")							'품목계정코드 
			        lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_ACCT_NM")						'계정명 
			        lgstrData = lgstrData & Chr(11) & lgObjRs("PROCUR_TYPE")						'조달구분코드 
			        lgstrData = lgstrData & Chr(11) & lgObjRs("PROCUR_TYPE_NM")						'조달구분명 
			        lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("PRNT_BOM_NO"))	& Chr(11) 		'BOM type, BOM type popup
		
			        lgstrData = lgstrData & Chr(11)	& UniConvNumberDBToCompany(lgObjRs("CHILD_ITEM_QTY"), 6, 3, "", 0)		'자품목기준수 
			        lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("CHILD_ITEM_UNIT"))	& Chr(11)	'자품목단위, 단위팝업			
		
			        lgstrData = lgstrData & Chr(11)	& UniConvNumberDBToCompany(lgObjRs("PRNT_ITEM_QTY"), 6, 3, "", 0)		'모품목기준수 
			        lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("PRNT_ITEM_UNIT"))	& Chr(11)	'모품목단위, 단위팝업	
		
			        lgstrData = lgstrData & Chr(11)	& lgObjRs("SAFETY_LT")							'안전L/T
			        lgstrData = lgstrData & Chr(11)	& UniConvNumberDBToCompany(lgObjRs("LOSS_RATE"), 3, 3, "", 0)		'loss율 
			        lgstrData = lgstrData & Chr(11)	& lgObjRs("SUPPLY_TYPE")						'유무상구분 
			        lgstrData = lgstrData & Chr(11)	& lgObjRs("SUPPLY_TYPE_NM")						'유무상구분명 
			        lgstrData = lgstrData & Chr(11)	& UNIDateClientFormat(lgObjRs("BOM_FROM_DT"))	'시작일 
			        lgstrData = lgstrData & Chr(11)	& UNIDateClientFormat(lgObjRs("BOM_TO_DT"))		'종료일		
			        lgstrData = lgstrData & Chr(11)	& lgObjRs("ECN_NO")	& Chr(11)					'변경번호, 변경번호 popup
			        lgstrData = lgstrData & Chr(11)	& lgObjRs("ECN_DESC") 							'변경내용	        
			        lgstrData = lgstrData & Chr(11)	& lgObjRs("REASON_CD")	& Chr(11)				'변경근거, 변경근거 popup
			        lgstrData = lgstrData & Chr(11)	& lgObjRs("REASON_NM")							'변경근거명 
			        lgstrData = lgstrData & Chr(11)	& lgObjRs("DRAWING_PATH")						'도면경로 
			        lgstrData = lgstrData & Chr(11)	& lgObjRs("REMARK")								'비고 
			        lgstrData = lgstrData & Chr(11)	& lgObjRs("PRNT_ITEM_CD")						'모품목 
			        lgstrData = lgstrData & Chr(11) & lgObjRs("PRNT_BOM_NO")						'모품목bom no
			        lgstrData = lgstrData & Chr(11) & lgObjRs("PRNT_PROC_TYPE")						'모품목조달구분 
			        lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("ITEM_VALID_FROM_DT"))	'품목유효기간시작일 
					lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("ITEM_VALID_TO_DT"))		'품목유효기간종료일 
					lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_ACCT_GRP")						'품목계정그룹 
					lgstrData = lgstrData & Chr(11)													'이관의뢰번호 
					lgstrData = lgstrData & Chr(11) & parent.frm1.StartDate							'이관요청일 
					lgstrData = lgstrData & Chr(11) & "N"											'이관상태 
					lgstrData = lgstrData & Chr(11) 												'이관일 
			'------ Developer Coding part (End   ) ------------------------------------------------------------------
			        lgstrData = lgstrData & Chr(11) & i
			        lgstrData = lgstrData & Chr(11) & Chr(12)	
			        
					lgObjRs.MoveNext
					
					i = i + 1
				Loop
			End If
			
			Call SubCloseRs(lgObjRs)  
			
			iViewStr = Join(TmpBuffer,"")
			
	End Select
    
End Sub 		

'
''============================================================================================================
'' Name : CommonOnTransactionCommit
'' Desc : This Sub is called by OnTransactionCommit Error handler
''============================================================================================================
'Sub CommonOnTransactionCommit()
'	'------ Developer Coding part (Start ) ------------------------------------------------------------------
'	'------ Developer Coding part (End   ) ------------------------------------------------------------------
'End Sub
'
''============================================================================================================
'' Name : CommonOnTransactionAbort
'' Desc : This Sub is called by OnTransactionAbort Error handler
''============================================================================================================
'Sub CommonOnTransactionAbort()
'    lgErrorStatus    = "YES"
'	'------ Developer Coding part (Start ) ------------------------------------------------------------------
'	'------ Developer Coding part (End   ) ------------------------------------------------------------------
'End Sub
'
''============================================================================================================
'' Name : SetErrorStatus
'' Desc : This Sub set error status
''============================================================================================================
'Sub SetErrorStatus()
'    lgErrorStatus     = "YES"                                                         '☜: Set error status
'	'------ Developer Coding part (Start ) ------------------------------------------------------------------
'	'------ Developer Coding part (End   ) ------------------------------------------------------------------
'End Sub
''============================================================================================================
'' Name : SubHandleError
'' Desc : This Sub handle error
''============================================================================================================
'Sub SubHandleError(pOpCode,pConn,pRs,pErr)
'    On Error Resume Next                                                             '☜: Protect system from crashing
'    Err.Clear                                                                        '☜: Clear Error status
'
'    Select Case pOpCode
'        Case "MC"
'            If CheckSYSTEMError(pErr,True) = True Then
'               ObjectContext.SetAbort
'               Call SetErrorStatus
'            Else
'               If CheckSQLError(pConn,True) = True Then
'                  ObjectContext.SetAbort
'                  Call SetErrorStatus
'               End If
'            End If
'        Case "MD"
'        Case "MR"
'        Case "MU"
'            If CheckSYSTEMError(pErr,True) = True Then
'               ObjectContext.SetAbort
'               Call SetErrorStatus
'            Else
'               If CheckSQLError(pConn,True) = True Then
'                  ObjectContext.SetAbort
'                  Call SetErrorStatus
'               End If
'            End If
'        Case "MB"
'			ObjectContext.SetAbort
'            Call SetErrorStatus        
'    End Select
'End Sub

%>