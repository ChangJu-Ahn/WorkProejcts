<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../inc/IncSvrMain.asp" -->
<!-- #Include file="../inc/IncSvrDate.inc" -->
<!-- #Include file="../inc/AdoVbs.inc" -->
<!-- #Include file="../inc/lgSvrVariables.inc" -->
<!-- #Include file="../inc/incServerAdoDB.asp" -->
<%
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd                                                               '☜: Hide Processing message
Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "*", "NOCOOKIE", "PB")

'---------------------------------------Common-----------------------------------------------------------
lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
lgStrPrevKeyIndex = CInt(Request("lgStrPrevKeyIndex"))   
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '☜: Set to space

'------ Developer Coding part (Start ) ------------------------------------------------------------------
Dim IntRetCD
	
Dim strItemCd
Dim strItemNm
Dim strItemAcct
Dim strFromItemAcctGrp
Dim strToItemAcctGrp
Dim strFromItemClass
Dim strToItemClass
Dim strItemGroupCd
Dim strBaseDt
Dim strSpec
Dim strValidFlg
Dim strWhere
Dim strAssignDtFlag
Dim strNextKey
Const C_SHEETMAXROWS_D = 100

strItemCd = FilterVar(Trim(Request("txtItemCd")) ,"''", "S")
strItemNm = Trim(Request("txtItemNm"))
strNextKey = Trim(Request("strNextKey"))

If Request("cboItemAccount") <> "" Then
	strItemAcct = FilterVar(Trim(Request("cboItemAccount"))	,"''", "S")
Else
	strItemAcct = ""
End If

If Request("ToItemAcctGrp") <> "zz" Then
	strFromItemAcctGrp = FilterVar(Trim(Request("FromItemAcctGrp"))	,"''", "S")
	strToItemAcctGrp = FilterVar(Cint(Request("ToItemAcctGrp") + 1)	,"''", "S")
Else
	strFromItemAcctGrp = FilterVar(Trim(Request("FromItemAcctGrp"))	,"''", "S")
	strToItemAcctGrp = FilterVar(Trim(Request("ToItemAcctGrp"))	,"''", "S")	
End If
	
IF Trim(Request("txtItemGroup")) <> "" Then		'품목그룹 
	strItemGroupCd = FilterVar(Trim(Request("txtItemGroup"))	,"''", "S")
Else
	strItemGroupCd = "''"
End If
	
IF Trim(Request("cboItemClass")) <> "" Then		'품목클래스 
	strFromItemClass = FilterVar(Trim(Request("cboItemClass"))	,"''", "S")
	strToItemClass = FilterVar(Trim(Request("cboItemClass"))	,"''", "S")
Else
	strFromItemClass = "''"
	strToItemClass = "'zzzzzzzzzzzz' OR A.ITEM_CLASS IS NULL"
End If

strItemNm = Replace(strItemNm, "[", "[[]")
strItemNm = "%" & Replace(strItemNm, "%", "[%]") & "%"

strSpec = Replace(Trim(Request("txtItemSpec")), "[", "[[]")
strSpec = "%" & Replace(strSpec, "%", "[%]") & "%"

strItemNm = FilterVar(strItemNm, "''", "S")
strNextKey = FilterVar(strNextKey, "''", "S")
strSpec = FilterVar(strSpec, "''", "S")
strValidFlg = FilterVar(Trim(Request("rdoValidFlg")), "''", "S")
strBaseDt = FilterVar(UniConvDate(Request("lgCurDate")), "''", "S")
strWhere = Trim(Request("txtWhere"))
strAssignDtFlag = Trim(Request("txtAssignDtFlag"))
	
'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

Call CheckItemGroupCd(strItemGroupCd)

If Trim(Request("pType")) = "" Then
	IF strItemNm = "'%%'" Or (strItemCd <> "''" And strItemNm <> "'%%'" ) Then
		Call SubBizQueryMulti("ITEM_CD")
	Else		
	    Call SubBizQueryMulti("ITEM_NM")
	End If
Else
	Call SubBizQueryMulti(Trim(Request("pType")))
End If
    
Call SubCloseDB(lgObjConn)

Response.End

'============================================================================================================
' Name : CheckItemGroupCd
' Desc : Check Item Group Cd
'============================================================================================================
Sub CheckItemGroupCd(ByVal pITemGroupCd)

	If pITemGroupCd = "''" Then Exit Sub
	
	lgStrSQL =			  " SELECT ITEM_GROUP_CD "
	lgStrSQL = lgStrSQL & " FROM B_ITEM_GROUP "
	lgStrSQL = lgStrSQL & " WHERE ITEM_GROUP_CD = " & pITemGroupCd

	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
		lgStrPrevKeyIndex = ""    
		IntRetCD = -1
		Call DisplayMsgBox("127400", vbInformation, "", "", I_MKSCRIPT)
		Call SetErrorStatus()

		Response.End
	End If

End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti(pType)
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Dim PrntKey
	Dim NodX
	Dim Node
	Dim iIntCnt
	Dim TmpBuffer
	Dim iTotalStr
	
	If pType = "ITEM_CD" Then
		Call SubMakeSQLStatements("CD",strItemCd,strItemNm,strFromItemClass,strToItemClass,strFromItemAcctGrp,strToItemAcctGrp,strItemGroupCd,strBaseDt,strSpec,strValidFlg,strNextKey, strItemAcct)           '☜ : Make sql statements
	Else	
		Call SubMakeSQLStatements("NM",strItemCd,strItemNm,strFromItemClass,strToItemClass,strFromItemAcctGrp,strToItemAcctGrp,strItemGroupCd,strBaseDt,strSpec,strValidFlg,strNextKey, strItemAcct)           '☜ : Make sql statements
	End If
	
	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists
		lgStrPrevKeyIndex = ""    
		IntRetCD = -1
		Call DisplayMsgBox("122600", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()
		Call SubCloseRs(lgObjRs)  
		Call SubCloseDB(lgObjConn)
		Response.End
	End If

	IntRetCD = 1
	iIntCnt = 1
	ReDim TmpBuffer(0)
	
	Do While Not lgObjRs.EOF
		lgstrData = ""
		lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("ITEM_CD"))	'품목코드 
		lgstrData = lgstrData & Chr(11)	& lgObjRs("ITEM_NM")		'품목명 
		lgstrData = lgstrData & Chr(11) & lgObjRs("SPEC")		'규격 
		lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("BASIC_UNIT"))		'단위 
		lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("ITEM_ACCT"))		'계정 
		lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("NM_ITEM_ACCT"))		'계정명 
		lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("NM_ITEM_CLASS"))		'품목그룹			
	        
		lgstrData = lgstrData & Chr(11) & lgObjRs("PHANTOM_FLG")		'팬텀 
		lgstrData = lgstrData & Chr(11)	& lgObjRs("BASE_ITEM_CD")		'기준품목 
		lgstrData = lgstrData & Chr(11)	& lgObjRs("BASE_ITEM_NM")		'기준품목명 
		lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("ITEM_GROUP_CD"))		'품목그룹 
		lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("ITEM_GROUP_NM"))		'품목그룹명 
	        
		lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("ITEM_IMAGE_FLG"))		'품목사진플래그 
		lgstrData = lgstrData & Chr(11)	& lgObjRs("FORMAL_NM")		'품목정식명칭 
		lgstrData = lgstrData & Chr(11)	& UNIDateClientFormat(lgObjRs("VALID_FROM_DT"))		'시작일 
		lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("VALID_TO_DT"))				'종료일			
		lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("HS_CD"))		'HS코드 
		lgstrData = lgstrData & Chr(11)	& lgObjRs("HS_UNIT")		'HS단위 
		lgstrData = lgstrData & Chr(11)	& ""		
		lgstrData = lgstrData & Chr(11)	& ""
		lgstrData = lgstrData & Chr(11)	& lgObjRs("UNIT_WEIGHT")		'단위중량 
		lgstrData = lgstrData & Chr(11)	& lgObjRs("UNIT_OF_WEIGHT")		'중량단위 
		lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("GROSS_WEIGHT"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
		lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GROSS_UNIT"))
		lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("CBM"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
		lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CBM_DESCRIPTION"))
		lgstrData = lgstrData & Chr(11)	& lgObjRs("DRAW_NO")		'도면번호 
		lgstrData = lgstrData & Chr(11)	& lgObjRs("BLANKET_PUR_FLG")		'BLANKET_PUR_FLG
		lgstrData = lgstrData & Chr(11)	& lgObjRs("PROPORTION_RATE")		'PROPORTION_RATE		
		lgstrData = lgstrData & Chr(11) & lgObjRs("VALID_FLG")		'VALID_FLG
		lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_ACCT_GRP")		'ITEM_ACCT_GRP
			
		lgstrData = lgstrData & Chr(11) & (lgLngMaxRow + iIntCnt)
		lgstrData = lgstrData & Chr(11) & Chr(12)

		lgObjRs.MoveNext
		
		ReDim Preserve TmpBuffer(iIntCnt-1)
		
		TmpBuffer(iIntCnt-1) = lgstrData
		
		iIntCnt =  iIntCnt + 1

	    If iIntCnt > C_SHEETMAXROWS_D Then
			Exit Do
	    End If
	Loop
		
	If lgObjRs.EOF Then
		lgStrPrevKeyIndex = ""
	Else
		lgStrPrevKeyIndex = iIntCnt
	End If
	
	iTotalStr = Join(TmpBuffer, "")
			
	Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "With parent" & vbCrLf
	
	If Trim(lgErrorStatus) = "NO" And IntRetCd <> -1 Then
	    Response.Write ".ggoSpread.Source = .vspdData" & vbCrLf
		Response.Write ".lgStrPrevKeyIndex = """ & lgStrPrevKeyIndex & """" & vbCrLf
	    Response.Write ".ggoSpread.SSShowDataByClip " & """" & ConvSPChars(iTotalStr) & """" & vbCrLf
	End If
			
	Response.Write ".hItemCd.value = """ & ConvSPChars(Trim(lgObjRs("ITEM_CD"))) & """" & vbCrLf
	Response.Write ".hpType.value = """ & pType & """" & vbCrLf
	If pType = "ITEM_NM" Then
		Response.Write "	.hItemNm.value = """ & ConvSPChars(Trim(Request("txtItemNm"))) & """" & vbCrLf	'from TextBox
		Response.Write "	.strNextKey = """ & ConvSPChars(Trim(lgObjRs("ITEM_NM"))) & """" & vbCrLf		'from Queried Values									
	End If
			
	
	Response.Write "	.DbQueryOk" & vbCrLf

	Response.Write ".vspdData.Focus" & vbCrLf
	
	Response.Write "End With" & vbCrLf
	Response.Write "</Script>" & vbCrLf
        
	Call SubCloseRs(lgObjRs)       
    
End Sub    

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2,pCode3,pCode4,pCode5,pCode6,pCode7,pCode8,pCode9,pCode10, pCode11)
    Dim iSelCount
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case pDataType

		Case "CD"
			lgStrSQL = "SELECT TOP " & CStr(C_SHEETMAXROWS_D + 1) & " A.*, B.ITEM_NM BASE_ITEM_NM, C.ITEM_GROUP_NM, A.ITEM_ACCT, dbo.ufn_GetCodeName('P1001', A.ITEM_ACCT) NM_ITEM_ACCT, dbo.ufn_GetCodeName('P1002', A.ITEM_CLASS) NM_ITEM_CLASS, "
			lgStrSQL = lgStrSQL & " dbo.ufn_GetItemAcctGrp(A.ITEM_ACCT) ITEM_ACCT_GRP "
			lgStrSQL = lgStrSQL & " FROM B_ITEM A, B_ITEM B, B_ITEM_GROUP C, B_ITEM_ACCT_INF D "
			lgStrSQL = lgStrSQL & " WHERE A.BASE_ITEM_CD *= B.ITEM_CD AND A.ITEM_GROUP_CD *= C.ITEM_GROUP_CD "
			lgStrSQL = lgStrSQL & " AND A.ITEM_ACCT = D.ITEM_ACCT "
			lgStrSQL = lgStrSQL & " AND A.ITEM_CD >= " & pCode
			lgStrSQL = lgStrSQL & " AND A.ITEM_NM LIKE  " & pCode1
			lgStrSQL = lgStrSQL & " AND A.ITEM_NM >= " & pCode10	'2003-09-02
			lgStrSQL = lgStrSQL & " AND (A.ITEM_CLASS >= " & pCode2
			lgStrSQL = lgStrSQL & " AND A.ITEM_CLASS <= " & pCode3
			lgStrSQL = lgStrSQL & " ) AND D.ITEM_ACCT_GROUP >= " & pCode4
			lgStrSQL = lgStrSQL & " AND D.ITEM_ACCT_GROUP <= " & pCode5
			If pCode6 <> "''" Then			
				lgStrSQL = lgStrSQL & " AND A.ITEM_GROUP_CD in (select item_group_cd from ufn_P_ListItemGrp(" & pCode6 & " )) "
			End If
			If strAssignDtFlag = "Y" Then
				lgStrSQL = lgStrSQL & " AND A.VALID_FROM_DT <= " & pCode7
			End If
			lgStrSQL = lgStrSQL & " AND A.VALID_TO_DT >= " & pCode7
			lgStrSQL = lgStrSQL & " AND A.SPEC LIKE " & pCode8
			lgStrSQL = lgStrSQL & " AND A.VALID_FLG = " & pCode9
			
			If pCode11 <> "" Then
				lgStrSQL = lgStrSQL & " AND A.ITEM_ACCT = " & pCode11
			End If

			lgStrSQL = lgStrSQL & " " & strWhere
			
			lgStrSQL = lgStrSQL & " ORDER BY A.ITEM_CD, A.ITEM_NM " 
			
			        
		Case "NM"
			lgStrSQL = "SELECT TOP " & CStr(C_SHEETMAXROWS_D + 1) & " A.*, B.ITEM_NM BASE_ITEM_NM, C.ITEM_GROUP_NM, A.ITEM_ACCT, dbo.ufn_GetCodeName('P1001', A.ITEM_ACCT) NM_ITEM_ACCT, dbo.ufn_GetCodeName('P1002', A.ITEM_CLASS) NM_ITEM_CLASS, "
			lgStrSQL = lgStrSQL & " dbo.ufn_GetItemAcctGrp(A.ITEM_ACCT) ITEM_ACCT_GRP "
			lgStrSQL = lgStrSQL & " FROM B_ITEM A, B_ITEM B, B_ITEM_GROUP C, B_ITEM_ACCT_INF D "
			lgStrSQL = lgStrSQL & " WHERE A.BASE_ITEM_CD *= B.ITEM_CD AND A.ITEM_GROUP_CD *= C.ITEM_GROUP_CD "
			lgStrSQL = lgStrSQL & " AND A.ITEM_ACCT = D.ITEM_ACCT "
			lgStrSQL = lgStrSQL & " AND A.ITEM_NM LIKE  " & pCode1
			lgStrSQL = lgStrSQL & " AND ((A.ITEM_CD >= " & pCode & " AND A.ITEM_NM = " & pCode10 & ") OR (A.ITEM_NM > " & pCode10
			lgStrSQL = lgStrSQL & ")) AND (A.ITEM_CLASS >= " & pCode2
			lgStrSQL = lgStrSQL & " AND A.ITEM_CLASS <= " & pCode3
			lgStrSQL = lgStrSQL & " ) AND D.ITEM_ACCT_GROUP >= " & pCode4
			lgStrSQL = lgStrSQL & " AND D.ITEM_ACCT_GROUP <= " & pCode5
			If pCode6 <> "''" Then			
				lgStrSQL = lgStrSQL & " AND A.ITEM_GROUP_CD in (select item_group_cd from ufn_P_ListItemGrp(" & pCode6 & " )) "
			End If
			If strAssignDtFlag = "Y" Then
				lgStrSQL = lgStrSQL & " AND A.VALID_FROM_DT <= " & pCode7
			End If
			lgStrSQL = lgStrSQL & " AND A.VALID_TO_DT >= " & pCode7
			lgStrSQL = lgStrSQL & " AND A.SPEC LIKE " & pCode8
			lgStrSQL = lgStrSQL & " AND A.VALID_FLG = " & pCode9
			If pCode11 <> "" Then
				lgStrSQL = lgStrSQL & " AND A.ITEM_ACCT = " & pCode11
			End If
			
			lgStrSQL = lgStrSQL & " " & strWhere
			
			lgStrSQL = lgStrSQL & " ORDER BY A.ITEM_NM, A.ITEM_CD " 
			
			

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

%>
