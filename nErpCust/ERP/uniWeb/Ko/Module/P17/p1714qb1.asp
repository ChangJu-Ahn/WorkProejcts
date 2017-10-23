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
lgLngMaxRow       = Request("txtMaxRows")                                       
    
lgMaxCount        = 500									
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")), 0)   
    
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '☜: Set to space
'lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
'------ Developer Coding part (Start ) ------------------------------------------------------------------

Dim IntRetCD
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

Dim strBasePlantCd
Dim strDestPlantCd	
Dim strItemCd
Dim strReqTransNo	

Dim TmpBuffer
Dim iTotalStr

ReDim TmpBuffer(0)
	
strBasePlantCd 	= Trim(Request("txtBasePlantCd"))
strDestPlantCd 	= Trim(Request("txtDestPlantCd"))
strItemCd 		= Trim(Request("txtItemCd"))
strReqTransNo 	= Trim(Request("txtReqTransNo"))

'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

	Call SubBizQuery("PB_CK") 				'기준공장 체크 
	Call SubBizQuery("PD_CK") 				'대상공장 체크 
	Call SubBizQuery("I_CK")				'모품목체크	
	
	Call SubCreateCommandObject(lgObjComm)
	Call SubBizQueryMulti()						
	Call SubCloseCommandObject(lgObjComm)	
	

Call SubCloseDB(lgObjConn)              '☜: Close DB Connection
	
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery(pOpCode)
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	Select Case pOpCode
		
		Case "PB_CK"
			'--------------
			'설계 공장 체크		
			'--------------	
			lgStrSQL = ""
			Call SubMakeSQLStatements("PB_CK",strBasePlantCd,"","","","")           '☜ : Make sql statements
			
			If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
				IntRetCD = -1
					
				If Trim(strBasePlantCd) <> "" Then
					Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
					Call SetErrorStatus()
					Response.Write "<Script Language = VBScript>" & vbCrLf
						Response.Write "parent.Frm1.txtBasePlantNm.Value  = """"" & vbCrLf   'Set condition area
						Response.Write "parent.Frm1.txtBasePlantCd.focus" & vbCrLf   'Set condition area
					Response.Write "</Script>" & vbcRLf
					Response.End
				Else
					Response.Write "<Script Language = VBScript>" & vbCrLf
						Response.Write "parent.Frm1.txtBasePlantNm.Value  = """"" & vbCrLf   'Set condition area
					Response.Write "</Script>" & vbcRLf
				End If
		    Else
				IntRetCD = 1
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.Frm1.txtBasePlantNm.Value = """ & ConvSPChars(lgObjRs(1)) & """" & vbCrLf 'Set condition area
				Response.Write "</Script>" & vbcRLf
			End If
			
			Call SubCloseRs(lgObjRs) 

		Case "PD_CK"
			'--------------
			'대상 공장 체크		
			'--------------	
			lgStrSQL = ""
			Call SubMakeSQLStatements("PD_CK",strDestPlantCd,"","","","")           '☜ : Make sql statements
			
			If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		    
				IntRetCD = -1
				Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.Frm1.txtDestPlantNm.Value  = """"" & vbCrLf   'Set condition area
					Response.Write "parent.Frm1.txtDestPlantCd.focus" & vbCrLf   'Set condition area
				Response.Write "</Script>" & vbcRLf
				Response.End
		    Else
				IntRetCD = 1
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.Frm1.txtDestPlantNm.Value = """ & ConvSPChars(lgObjRs(1)) & """" & vbCrLf 'Set condition area
				Response.Write "</Script>" & vbcRLf
			End If
				
			Call SubCloseRs(lgObjRs) 			
		
		Case "I_CK"
			'------------------
			'품목체크 
			'------------------
			lgStrSQL = ""
			Call SubMakeSQLStatements("I_CK", strDestPlantCd, strItemCd, "", "", "")           '☜ : Make sql statements

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
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write  "parent.frm1.txtItemNm.Value = """ & ConvSPChars(Trim(lgObjRs("ITEM_NM"))) & """" & vbCrLf
				Response.Write "</Script>" & vbCrLf
			
			End If
		
			Call SubCloseRs(lgObjRs) 
		    
	End Select
    
End Sub 
  
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : 설계BOM전개결과조회 
'============================================================================================================
Sub SubBizQueryMulti()
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Dim PrntKey
	Dim NodX
	Dim Node
	Dim iIntCnt, iLevelCnt

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
'	lgStrSQL = ""

	Call SubMakeSQLStatements("MB", strDestPlantCd, strItemCd, strBasePlantCd, strReqTransNo, "")					'☜ : Make sql statements

	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists
		lgStrPrevKeyIndex = ""
	Else
		IntRetCD = 1
		Call SubSkipRs(lgObjRs, lgMaxCount * lgStrPrevKeyIndex)
		iDx       = 1
		
        Do While Not lgObjRs.EOF
			
			lgstrData = ""
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("CHILD_ITEM_SEQ")											'순서 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("REQ_TRANS_NO")											'이관의뢰번호 
			lgstrData = lgstrData & Chr(11) & lgObjRs("STATUS_NM")												'이관상태 
			lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("CHILD_ITEM_CD"))	 								'자품목 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_NM")												'자품목명 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("SPEC")													'규격 
	        lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("REQ_TRANS_DT"))						'이관요청일 
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("TRANS_DT"))							'이관일 
			lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_ACCT_NM")											'계정명 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("PROCUR_TYPE_NM")											'조달구분명 
	        lgstrData = lgstrData & Chr(11)	& UniConvNumberDBToCompany(lgObjRs("CHILD_ITEM_QTY"), 4, 3, "", 0)	'자품목기준수 
	        lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("CHILD_ITEM_UNIT"))									'자품목단위 
	        lgstrData = lgstrData & Chr(11)	& UniConvNumberDBToCompany(lgObjRs("PRNT_ITEM_QTY"), 4, 3, "", 0)	'모품목기준수 
	        lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("PRNT_ITEM_UNIT"))									'모품목단위 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("SAFETY_LT")												'안전L/T
	        lgstrData = lgstrData & Chr(11)	& UniConvNumberDBToCompany(lgObjRs("LOSS_RATE"), 3, 3, "", 0)		'loss율 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("SUPPLY_TYPE_NM")											'유무상구분명 
	        lgstrData = lgstrData & Chr(11)	& UNIDateClientFormat(lgObjRs("BOM_FROM_DT"))						'시작일 
 	        lgstrData = lgstrData & Chr(11)	& UNIDateClientFormat(lgObjRs("BOM_TO_DT"))							'종료일 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("REASON_NM")												'변경근거명 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("ECN_NO")													'변경번호 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("ECN_DESC") 												'변경내용	        
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("DRAWING_PATH")											'도면경로 
	        
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
	        lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
	        lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext
			
			ReDim Preserve TmpBuffer(iDx-1)
			TmpBuffer(iDx-1) = lgstrData
			
	        iDx =  iDx + 1
	        If iDx > lgMaxCount  Then			'처음에 최상위품목row를 한줄 써주었으므로 
	           lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
	               
	           Exit Do
	        End If   
        Loop 

		If iDx <= lgMaxCount Then
		   lgStrPrevKeyIndex = ""
		End If   
		
		Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
		Call SubCloseRs(lgObjRs)       
    End If
	
	iTotalStr = Join(TmpBuffer,"")


End Sub    


'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2,pCode3,pCode4)
 '   Dim iSelCount
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case pDataType

		Case "MB"
						lgStrSQL = " SELECT A.REQ_TRANS_NO, B.CHILD_ITEM_SEQ, B.CHILD_ITEM_CD, E.ITEM_NM, E.SPEC, A.REQ_TRANS_DT, "
			lgStrSQL = 	lgStrSQL &        " A.TRANS_DT, C.ITEM_ACCT, dbo.ufn_GetCodeName('P1001', C.ITEM_ACCT) ITEM_ACCT_NM, "
			lgStrSQL = 	lgStrSQL &        " C.PROCUR_TYPE, dbo.ufn_GetCodeName('P1003', C.PROCUR_TYPE) PROCUR_TYPE_NM, "
			lgStrSQL = 	lgStrSQL &        " B.CHILD_ITEM_QTY, B.CHILD_ITEM_UNIT, B.PRNT_ITEM_QTY, B.PRNT_ITEM_UNIT, "
			lgStrSQL = 	lgStrSQL &        " B.SAFETY_LT, B.LOSS_RATE,  "
			lgStrSQL = 	lgStrSQL &        " B.SUPPLY_TYPE, dbo.ufn_GetCodeName('M2201', B.SUPPLY_TYPE) SUPPLY_TYPE_NM, "
			lgStrSQL = 	lgStrSQL &        " B.VALID_FROM_DT AS BOM_FROM_DT, B.VALID_TO_DT AS BOM_TO_DT, B.ECN_NO, D.ECN_DESC, "
			lgStrSQL = 	lgStrSQL &        " D.REASON_CD,  dbo.ufn_GetCodeName('P1402', D.REASON_CD) REASON_NM, A.DRAWING_PATH, A.STATUS, dbo.ufn_GetCodeName('Y4001', A.STATUS) STATUS_NM "
			lgStrSQL = 	lgStrSQL &   " FROM P_EBOM_TO_PBOM_DETAIL B "
			lgStrSQL = 	lgStrSQL &              " LEFT OUTER JOIN B_ITEM_BY_PLANT C ON B.PRNT_PLANT_CD = C.PLANT_CD AND B.CHILD_ITEM_CD = C.ITEM_CD "
			lgStrSQL = 	lgStrSQL &              " LEFT OUTER JOIN P_ECN_MASTER D    ON B.ECN_NO = D.ECN_NO "
			lgStrSQL = 	lgStrSQL &              " LEFT OUTER JOIN B_ITEM E          ON B.CHILD_ITEM_CD = E.ITEM_CD "
			lgStrSQL = 	lgStrSQL &              " LEFT JOIN P_EBOM_TO_PBOM_MASTER A ON  A.PLANT_CD = B.PRNT_PLANT_CD "
			lgStrSQL = 	lgStrSQL &                                                " AND A.ITEM_CD = B.PRNT_ITEM_CD "
			lgStrSQL = 	lgStrSQL &                                                " AND A.BOM_NO = B.PRNT_BOM_NO "
			lgStrSQL = 	lgStrSQL &                                                " AND A.REQ_TRANS_NO = B.REQ_TRANS_NO "
			lgStrSQL = 	lgStrSQL & " WHERE A.PLANT_CD = " & FilterVar(pCode, "''", "S")
			lgStrSQL = 	lgStrSQL & " AND A.ITEM_CD = " & FilterVar(pCode1, "''", "S")
			
			If Trim(pcode2) <> "" Then
				lgStrSQL = 	lgStrSQL & " AND A.DESIGN_PLANT_CD = " & FilterVar(pCode2, "''", "S")	
			End If
			
			If Trim(pcode3) <> "" Then
				lgStrSQL = 	lgStrSQL & " AND A.REQ_TRANS_NO = " & FilterVar(pCode3, "''", "S")	
			End If
			
			lgStrSQL = 	lgStrSQL & " ORDER BY B.CHILD_ITEM_SEQ, A.REQ_TRANS_NO, B.CHILD_ITEM_CD "

		Case "I_CK"
			lgStrSQL = "SELECT a.*, b.ITEM_NM, b.SPEC, dbo.ufn_GetCodeName(" & FilterVar("P1001", "''", "S") & ", a.ITEM_ACCT) ITEM_ACCT_NM, b.PHANTOM_FLG, b.BASIC_UNIT, " _
						& " dbo.ufn_GetCodeName(" & FilterVar("P1003", "''", "S") & ", a.PROCUR_TYPE) PROCUR_TYPE_NM, dbo.ufn_GetItemAcctGrp(a.ITEM_ACCT) ITEM_ACCT_GRP "
			lgStrSQL = lgStrSQL & " FROM B_ITEM_BY_PLANT a, B_ITEM b "
			lgStrSQL = lgStrSQL & " WHERE a.ITEM_CD = b.ITEM_CD "
			lgStrSQL = lgStrSQL & " AND a.PLANT_CD = " & FilterVar(pCode, "''", "S")
			lgStrSQL = lgStrSQL & " AND a.ITEM_CD = " & FilterVar(pCode1, "''", "S")
			
		Case "PB_CK"
			lgStrSQL = "SELECT * FROM B_PLANT A, P_PLANT_CONFIGURATION B"
			lgStrSQL = lgStrSQL & " WHERE A.PLANT_CD = B.PLANT_CD"
			lgStrSQL = lgStrSQL & " AND B.ENG_BOM_FLAG = 'Y'"
			lgStrSQL = lgStrSQL & " AND A.PLANT_CD = " & FilterVar(pCode, "''", "S")		
		Case "PD_CK"
			lgStrSQL = "SELECT * FROM B_PLANT WHERE PLANT_CD = " & FilterVar(pCode, "''", "S")
		
    End Select
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

Response.Write "<Script Language = VBScript>" & vbCrLf                                                      '☜ : Query
	If Trim(lgErrorStatus) = "NO" And IntRetCd <> -1 Then
		 Response.Write "With Parent" & vbCrLf
            Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
            Response.Write ".lgStrPrevKeyIndex = """ & lgStrPrevKeyIndex & """" & vbCrLf
            Response.Write ".ggoSpread.SSShowDataByClip """ & ConvSPChars(iTotalStr) & """" & vbCrLf				
            Response.Write ".DBQueryOk(" & lgLngMaxRow & " + 1)" & vbCrLf	
         Response.Write "End with" & vbCrLf
    End If 
Response.Write "</Script>" & vbCrLf

Response.End
%>
