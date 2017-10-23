<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../inc/IncSvrMain.asp" -->
<!-- #Include file="../inc/IncSvrDate.inc" -->
<!-- #Include file="../inc/AdoVbs.inc" -->
<!-- #Include file="../inc/lgSvrVariables.inc" -->
<!-- #Include file="../inc/incServerAdoDB.asp" -->
<%
On Error Resume Next                                                             '��: Protect system from crashing
Err.Clear                                                                        '��: Clear Error status

Call HideStatusWnd                                                               '��: Hide Processing message
Call LoadBasisGlobalInf
'---------------------------------------Common-----------------------------------------------------------
lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
lgStrPrevKeyIndex = CInt(Request("lgStrPrevKeyIndex"))   
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '��: Set to space
'------ Developer Coding part (Start ) ------------------------------------------------------------------
Dim IntRetCD
	
Dim strPlantCd
Dim strItemCd
Dim strItemNm
Dim strFromItemAcctGrp
Dim strToItemAcctGrp
Dim strItemAcct
Dim strFromItemClass
Dim strToItemClass
Dim strFromProcType
Dim strToProcType
Dim strFromProdEnv
Dim strToProdEnv
Dim strBaseDt
Dim strSpec
Dim strInspClass
Dim strTrackingFlg
Dim strWhere
Dim strAssignDtFlag
Dim strNextKey
Dim strItemGroupCd

Const C_SHEETMAXROWS_D = 100

strPlantCd = FilterVar(Trim(Request("PlantCd"))	,"''", "S")
strItemCd = FilterVar(Trim(Request("txtItemCd"))	,"''", "S")
strItemNm = Trim(Request("txtItemNm"))
strNextKey = Trim(Request("strNextKey"))

If Request("cboProcurType") <> "" Then		'���ޱ��� 
	strFromProcType = FilterVar(Trim(Request("cboProcurType"))	,"''", "S")
	strToProcType = FilterVar(Trim(Request("cboProcurType"))	,"''", "S")
Else
	If Request("ToProcType") <> "" Then
		strFromProcType = FilterVar(Trim(Request("FromProcType"))	,"''", "S")
		strToProcType = FilterVar(Trim(Request("ToProcType"))	,"''", "S")
	End If	
End If
	
IF Request("cboProdtEnv") <> "" Then		'�������� 
	strFromProdEnv = FilterVar(Trim(Request("cboProdtEnv"))	,"''", "S")
	strToProdEnv = FilterVar(Trim(Request("cboProdtEnv"))	,"''", "S")
Else
	strFromProdEnv = FilterVar(Trim(Request("cboProdtEnv"))	,"''", "S")
	strToProdEnv = FilterVar(Trim(Request("cboProdtEnv"))	,"'zz'", "S")
End If

If Trim(Request("cboItemAccount")) <> "" Then
	strItemAcct = FilterVar(Trim(Request("cboItemAccount"))	,"''", "S")
Else
	strItemAcct = ""
End If	

If Request("ToItemAcctGrp") <> "zz" Then	'ǰ������׷� 
	strFromItemAcctGrp = FilterVar(Trim(Request("FromItemAcctGrp"))	,"''", "S")
	strToItemAcctGrp = FilterVar(Cint(Request("ToItemAcctGrp") + 1)	,"''", "S")
Else
	strFromItemAcctGrp = FilterVar(Trim(Request("FromItemAcctGrp"))	,"''", "S")
	strToItemAcctGrp = FilterVar(Trim(Request("ToItemAcctGrp"))	,"''", "S")	
End If

	
IF Request("cboItemClass") <> "" Then		'�������� 
	strFromItemClass = FilterVar(Trim(Request("cboItemClass"))	,"''", "S")
	strToItemClass = FilterVar(Trim(Request("cboItemClass"))	,"''", "S")
Else
	strFromItemClass = FilterVar(Trim(Request("cboItemClass"))	,"''", "S")
	strToItemClass = FilterVar(Trim(Request("cboItemClass"))	,"'zzzzzzzzzzzz' OR B.ITEM_CLASS IS NULL", "S")
End If

IF Request("txtItemGroupCd") <> "" Then
	strItemGroupCd = Filtervar(Request("txtItemGroupCd"), "''", "S")
Else
	strItemGroupCd = ""
End If

strItemNm = Replace(strItemNm, "[", "[[]")
strItemNm = "%" & Replace(strItemNm, "%", "[%]") & "%"

strSpec = Replace(Trim(Request("txtItemSpec")), "[", "[[]")
strSpec = "%" & Replace(strSpec, "%", "[%]") & "%"

strItemNm = FilterVar(strItemNm, "''", "S")
strNextKey = FilterVar(strNextKey, "''", "S")
strSpec = FilterVar(strSpec, "''", "S")
strTrackingFlg = FilterVar(Trim(Request("rdoTrackingItem"))	,"''", "S")
strInspClass = Trim(Request("cboInspType"))
strBaseDt = FilterVar(UniConvDate(Request("lgCurDate"))	,"''", "S")
strWhere = Trim(Request("txtWhere"))
strAssignDtFlag = Trim(Request("txtAssignDtFlag"))

	
'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection


' Add 2005-06-12
If strItemGroupCd <> "" Then
	Call FncGetData("ITEM_GP")
Else
	Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write " Parent.txtItemGroupNm.value = """" "& vbCrLf 	
	Response.Write "</Script>" & vbCrLf
End If	
	

If Trim(Request("pType")) = "" Then
	IF strItemNm = "'%%'" Or (strItemCd <> "''" And strItemNm <> "'%%'" ) Then
		Call SubBizQueryMulti("ITEM_CD")
	Else		
	    Call SubBizQueryMulti("ITEM_NM")
	End If
Else
	Call SubBizQueryMulti(Trim(Request("pType")))
End If
    
Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection

Response.End

'============================================================================================================
' Name : FncGetData
' Desc : Get Data from Db
'============================================================================================================
Sub FncGetData(pType)
	
	On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    
	Dim iIntCnt
	Dim TmpBuffer
	Dim iTotalStr
	
	Select Case pType
		Case "ITEM_GP"
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
			Call SubGetSQLStatements("GP",strItemGroupCd)           '�� : Make sql statements
			
			If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then                    'If data not exists
				Response.Write "<Script Language = VBScript>" & vbCrLf
				Response.Write "With parent" & vbCrLf
					Response.Write ".txtItemGroupNm.value = """ & ConvSpChars(lgObjRs("ITEM_GROUP_NM")) & """" & vbCrLf
				Response.Write "End With" & vbCrLf
				Response.Write "</Script>" & vbCrLf
			Else	
				Response.Write "<Script Language = VBScript>" & vbCrLf
				Response.Write "With parent" & vbCrLf
					Response.Write ".txtItemGroupNm.value = """" " & vbCrLf
				Response.Write "End With" & vbCrLf
				Response.Write "</Script>" & vbCrLf
			End If
			
	End Select

End Sub    


'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti(pType)
	
	On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    
    Dim PrntKey
	Dim NodX
	Dim Node
	Dim iIntCnt
	Dim TmpBuffer
	Dim iTotalStr
	
	If pType = "ITEM_CD" Then
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
		Call SubMakeSQLStatements("CD",strPlantCd,strItemCd,strItemNm,strFromItemClass,strToItemClass,strFromItemAcctGrp,strToItemAcctGrp,strFromProcType,strToProcType,strFromProdEnv,strToProdEnv,strBaseDt,strSpec,strTrackingFlg,strInspClass,strNextKey, strItemAcct, strItemGroupCd)           '�� : Make sql statements
	ElseIf pType = "ITEM_NM" Then	
		Call SubMakeSQLStatements("NM",strPlantCd,strItemCd,strItemNm,strFromItemClass,strToItemClass,strFromItemAcctGrp,strToItemAcctGrp,strFromProcType,strToProcType,strFromProdEnv,strToProdEnv,strBaseDt,strSpec,strTrackingFlg,strInspClass,strNextKey, strItemAcct, strItemGroupCd)           '�� : Make sql statements	
	End If

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		lgStrPrevKeyIndex = ""    
		IntRetCD = -1
		Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
		Call SetErrorStatus()
		Call SubCloseDB(lgObjRs)
		Call SubCloseDB(lgObjConn)		 
		Response.End
     
	End If
	
	IntRetCD = 1
	iIntCnt = 1
	ReDim TmpBuffer(0)
		
    Do While Not lgObjRs.EOF
		
		lgstrData = ""	
		lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("ITEM_CD"))		'ǰ���ڵ� 
        lgstrData = lgstrData & Chr(11)	& lgObjRs("ITEM_NM")			'ǰ��� 
        lgstrData = lgstrData & Chr(11) & lgObjRs("SPEC")				'�԰� 
        lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("BASIC_UNIT"))	'���� 
		lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("ITEM_ACCT"))		'���� 
		lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("ITEM_ACCT_NM"))		'������ 
		lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("ITEM_GROUP_CD"))	'ǰ��׷� 
        lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_CLASS_NM")			'�����ǰ��Ŭ���� 
        lgstrData = lgstrData & Chr(11) & lgObjRs("PROCUR_TYPE")		'���� 
        lgstrData = lgstrData & Chr(11) & lgObjRs("PROCUR_TYPE_NM")		'���ޱ��и� 
        lgstrData = lgstrData & Chr(11) & lgObjRs("PROD_ENV")		'�������� 
        lgstrData = lgstrData & Chr(11) & lgObjRs("PROD_ENV_NM")	'���������� 
        lgstrData = lgstrData & Chr(11) & lgObjRs("PHANTOM_FLG")		'���� 
        lgstrData = lgstrData & Chr(11) & lgObjRs("LOT_FLG")			'LOT���� 
	        
        lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("MAJOR_SL_CD"))	'�԰�â�� 
        lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("ISSUED_SL_CD"))	'���â��						'bom type popup
        lgstrData = lgstrData & Chr(11)	& lgObjRs(57)					'��ȿ���� 
        lgstrData = lgstrData & Chr(11)	& UNIDateClientFormat(lgObjRs(58))	'������ 
        lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(59))	'������			
	        
        lgstrData = lgstrData & Chr(11)	& lgObjRs("FORMAL_NM")				'ǰ�����ĸ�Ī 
        lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs(82))					'ǰ����� 
        lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("HS_CD"))			'HS�ڵ�	
        lgstrData = lgstrData & Chr(11)	& lgObjRs("HS_UNIT")			'HS���� 
        lgstrData = lgstrData & Chr(11)	& lgObjRs("BASE_ITEM_CD")		'����ǰ�� 
        lgstrData = lgstrData & Chr(11)	& lgObjRs("TRACKING_FLG")		'TRACKING ���� 
        lgstrData = lgstrData & Chr(11)	& lgObjRs("ORDER_UNIT_MFG")		'������������ 
        lgstrData = lgstrData & Chr(11)	& lgObjRs("ORDER_UNIT_PUR")		'���ſ������� 
        lgstrData = lgstrData & Chr(11)	& lgObjRs("ORDER_LT_MFG")		'��������L/T		
        lgstrData = lgstrData & Chr(11)	& lgObjRs("ORDER_LT_PUR")		'���ſ���L/T
        lgstrData = lgstrData & Chr(11) & lgObjRs("ORDER_TYPE")			'����Ÿ�� 
        lgstrData = lgstrData & Chr(11)	& lgObjRs("ORDER_RULE")			'���ֹ�ħ 
        lgstrData = lgstrData & Chr(11) & lgObjRs("FIXED_MRP_QTY")		'��������� 
        lgstrData = lgstrData & Chr(11) & lgObjRs("MIN_MRP_QTY")		'�ּҼ���� 
		lgstrData = lgstrData & Chr(11) & lgObjRs("MAX_MRP_QTY")		'�ִ����� 
		lgstrData = lgstrData & Chr(11) & lgObjRs("ROUND_QTY")			'�ø��� 
		lgstrData = lgstrData & Chr(11) & lgObjRs("ROUND_PERD")			'�ø��Ⱓ 
		lgstrData = lgstrData & Chr(11) & lgObjRs("MPS_FLG")			'MPS���� 
		lgstrData = lgstrData & Chr(11) & lgObjRs("ISSUE_MTHD")		'����� 
		lgstrData = lgstrData & Chr(11) & lgObjRs("PUR_ORG")			'�������� 
		lgstrData = lgstrData & Chr(11) & lgObjRs("OPTION_FLG")			'OPTION ���� 
		lgstrData = lgstrData & Chr(11) & lgObjRs("CYCLE_CNT_PERD")		'�ǻ��ֱ� 
		lgstrData = lgstrData & Chr(11) & lgObjRs("ISSUED_UNIT")		'������ 
		lgstrData = lgstrData & Chr(11) & lgObjRs("RECV_INSPEC_FLG")	'���԰˻籸�� 
		lgstrData = lgstrData & Chr(11) & lgObjRs("PROD_INSPEC_FLG")	'�����˻籸�� 
		lgstrData = lgstrData & Chr(11) & lgObjRs("FINAL_INSPEC_FLG")	'�����˻籸�� 
		lgstrData = lgstrData & Chr(11) & lgObjRs("SHIP_INSPEC_FLG")	'���ϰ˻籸�� 
		lgstrData = lgstrData & Chr(11) & lgObjRs("INSPEC_LT_MFG")		'�����˻�L/T
		lgstrData = lgstrData & Chr(11) & lgObjRs("INSPEC_LT_PUR")		'���Ű˻�L/T
		lgstrData = lgstrData & Chr(11) & lgObjRs("INSPEC_MGR")			'�˻����� 
		lgstrData = lgstrData & Chr(11) & lgObjRs(96)					'ǰ����ȿ���� 
		lgstrData = lgstrData & Chr(11) & lgObjRs("SINGLE_ROUT_FLG")	'�ܰ������� 
		lgstrData = lgstrData & Chr(11) & lgObjRs("WORK_CENTER")	    '���۾��� 
		lgstrData = lgstrData & Chr(11) & lgObjRs("ABC_FLG")			'ABC���� 
		lgstrData = lgstrData & Chr(11) & lgObjRs("ITEM_ACCT_GRP")		'ǰ������׷� 
			
'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
 		
	Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "With parent" & vbCrLf
			If Trim(lgErrorStatus) = "NO" And IntRetCd <> -1 Then
		        Response.Write ".ggoSpread.Source = .vspdData" & vbCrLf
				Response.Write ".lgStrPrevKeyIndex = """ & lgStrPrevKeyIndex & """" & vbCrLf
		        Response.Write ".ggoSpread.SSShowDataByClip """ & ConvSPChars(iTotalStr) & """" & vbCrLf
	        End If
			
			If lgObjRs.EOF Then
				Response.Write ".hItemCd.value = """"" & vbCrLf
				Response.Write ".hpType.value = """ & pType & """" & vbCrLf

				If pType = "ITEM_NM" Then
					Response.Write "	.hItemNm.value = """ & ConvSPChars(Trim(Request("txtItemNm"))) & """" & vbCrLf	'from TextBox
					Response.Write "	.strNextKey = """"" & vbCrLf		'from Queried Values
				End If
			Else
				Response.Write ".hItemCd.value = """ & ConvSPChars(Trim(lgObjRs("ITEM_CD"))) & """" & vbCrLf
				Response.Write ".hpType.value = """ & pType & """" & vbCrLf

				If pType = "ITEM_NM" Then
					Response.Write "	.hItemNm.value = """ & ConvSPChars(Trim(Request("txtItemNm"))) & """" & vbCrLf	'from TextBox
					Response.Write "	.strNextKey = """ & ConvSPChars(Trim(lgObjRs("ITEM_NM"))) & """" & vbCrLf		'from Queried Values
				End If
			End If	
			
			Response.Write "	.hitemGroupCd.value	= """ & ConvSPChars(Request("txtItemGroupCd")) & """" & vbCrLf

			Response.Write ".DbQueryOk" & vbCrLf

			Response.Write ".vspdData.Focus" & vbCrLf
		Response.Write "End With" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	
	Call SubCloseDB(lgObjRs)                                                       '��: Close DB Connection
	
End Sub    

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2,pCode3,pCode4,pCode5,pCode6,pCode7,pCode8,pCode9,pCode10,pCode11,pCode12,pCode13,pCode14,pCode15,pCode16, pCode17)
    Dim iSelCount
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case pDataType

		Case "CD"
			lgStrSQL = "SELECT TOP " & CStr(C_SHEETMAXROWS_D + 1) & " A.*, B.*, dbo.ufn_GetCodeName('P1001', A.ITEM_ACCT) ITEM_ACCT_NM, "
			lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName('P1002', B.ITEM_CLASS) ITEM_CLASS_NM, "
			lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName('P1003', A.PROCUR_TYPE) PROCUR_TYPE_NM, "
			lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName('P1004', A.PROD_ENV) PROD_ENV_NM, "
			lgStrSQL = lgStrSQL & " dbo.ufn_GetItemAcctGrp(A.ITEM_ACCT) ITEM_ACCT_GRP "
			lgStrSQL = lgStrSQL & " FROM B_ITEM_BY_PLANT A, B_ITEM B, B_ITEM_ACCT_INF C "
			lgStrSQL = lgStrSQL & " WHERE A.ITEM_CD = B.ITEM_CD "
			lgStrSQL = lgStrSQL & " AND A.ITEM_ACCT = C.ITEM_ACCT "
			lgStrSQL = lgStrSQL & " AND B.PHANTOM_FLG = 'N'"
			lgStrSQL = lgStrSQL & " AND  A.PLANT_CD = " & pCode
			lgStrSQL = lgStrSQL & " AND A.ITEM_CD >= " & pCode1
			lgStrSQL = lgStrSQL & " AND B.ITEM_NM LIKE  " & pCode2
			lgStrSQL = lgStrSQL & " AND B.ITEM_NM >= " & pCode15	'2003-09-02
			lgStrSQL = lgStrSQL & " AND (B.ITEM_CLASS >= " & pCode3
			lgStrSQL = lgStrSQL & " AND B.ITEM_CLASS <= " & pCode4
			lgStrSQL = lgStrSQL & " ) AND C.ITEM_ACCT_GROUP >= " & pCode5
			lgStrSQL = lgStrSQL & " AND C.ITEM_ACCT_GROUP <= " & pCode6
			lgStrSQL = lgStrSQL & " AND A.PROCUR_TYPE >= " & pCode7
			lgStrSQL = lgStrSQL & " AND A.PROCUR_TYPE <= " & pCode8
			lgStrSQL = lgStrSQL & " AND A.PROD_ENV >= " & pCode9
			lgStrSQL = lgStrSQL & " AND A.PROD_ENV <= " & pCode10
			If strAssignDtFlag = "Y" Then
				lgStrSQL = lgStrSQL & " AND A.VALID_FROM_DT <= " & pCode11
			End If
			lgStrSQL = lgStrSQL & " AND A.VALID_TO_DT >= " & pCode11
			lgStrSQL = lgStrSQL & " AND B.SPEC LIKE " & pCode12
			lgStrSQL = lgStrSQL & " AND A.TRACKING_FLG LIKE " & pCode13
	
			If pCode14 = "R" Then
				lgStrSQL = lgStrSQL & " AND A.RECV_INSPEC_FLG = 'Y'"
			ElseIf pCode14 = "F" Then
				lgStrSQL = lgStrSQL & " AND A.FINAL_INSPEC_FLG = 'Y'"
			ElseIf pCode14 = "P" Then
				lgStrSQL = lgStrSQL & " AND A.PROD_INSPEC_FLG = 'Y'"
			ElseIf pCode14 = "S" Then
				lgStrSQL = lgStrSQL & " AND A.SHIP_INSPEC_FLG = 'Y'"
			End If
			
			If pCode16 <> "" Then
				lgStrSQL = lgStrSQL & " AND A.ITEM_ACCT = " & pCode16
			End If
			
			If PCode17 <> "" Then
				lgStrSQL = lgStrSQL & " AND b.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & PCode17 & " )) "
			End If
			
		    lgStrSQL = lgStrSQL & " " & strWhere

			lgStrSQL = lgStrSQL & " ORDER BY A.ITEM_CD, B.ITEM_NM " 
			
		Case "NM"
			lgStrSQL = "SELECT TOP " & CStr(C_SHEETMAXROWS_D + 1) & " A.*, B.*, dbo.ufn_GetCodeName('P1001', A.ITEM_ACCT) ITEM_ACCT_NM, "
			lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName('P1002', B.ITEM_CLASS) ITEM_CLASS_NM, "
			lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName('P1003', A.PROCUR_TYPE) PROCUR_TYPE_NM, "
			lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName('P1004', A.PROD_ENV) PROD_ENV_NM, "
			lgStrSQL = lgStrSQL & " dbo.ufn_GetItemAcctGrp(A.ITEM_ACCT) ITEM_ACCT_GRP "
			lgStrSQL = lgStrSQL & " FROM B_ITEM_BY_PLANT A, B_ITEM B, B_ITEM_ACCT_INF C "
			lgStrSQL = lgStrSQL & " WHERE A.ITEM_CD = B.ITEM_CD "
			lgStrSQL = lgStrSQL & " AND A.ITEM_ACCT = C.ITEM_ACCT "
			lgStrSQL = lgStrSQL & " AND B.PHANTOM_FLG = 'N'"
			lgStrSQL = lgStrSQL & " AND  A.PLANT_CD = " & pCode
			lgStrSQL = lgStrSQL & " AND B.ITEM_NM LIKE  " & pCode2
			lgStrSQL = lgStrSQL & " AND ((A.ITEM_CD >= " & pCode1 & " AND B.ITEM_NM = " & pCode15 & ") OR (B.ITEM_NM > " & pCode15
			lgStrSQL = lgStrSQL & ")) AND (B.ITEM_CLASS >= " & pCode3
			lgStrSQL = lgStrSQL & " AND B.ITEM_CLASS <= " & pCode4
			lgStrSQL = lgStrSQL & " ) AND C.ITEM_ACCT_GROUP >= " & pCode5
			lgStrSQL = lgStrSQL & " AND C.ITEM_ACCT_GROUP <= " & pCode6
			lgStrSQL = lgStrSQL & " AND A.PROCUR_TYPE >= " & pCode7
			lgStrSQL = lgStrSQL & " AND A.PROCUR_TYPE <= " & pCode8
			lgStrSQL = lgStrSQL & " AND A.PROD_ENV >= " & pCode9
			lgStrSQL = lgStrSQL & " AND A.PROD_ENV <= " & pCode10
			If strAssignDtFlag = "Y" Then
				lgStrSQL = lgStrSQL & " AND A.VALID_FROM_DT <= " & pCode11
			End If
			lgStrSQL = lgStrSQL & " AND A.VALID_TO_DT >= " & pCode11
			lgStrSQL = lgStrSQL & " AND B.SPEC LIKE " & pCode12
			lgStrSQL = lgStrSQL & " AND A.TRACKING_FLG LIKE " & pCode13
			
			If pCode14 = "R" Then
				lgStrSQL = lgStrSQL & " AND A.RECV_INSPEC_FLG = 'Y'"
			ElseIf pCode14 = "F" Then
				lgStrSQL = lgStrSQL & " AND A.FINAL_INSPEC_FLG = 'Y'"
			ElseIf pCode14 = "P" Then
				lgStrSQL = lgStrSQL & " AND A.PROD_INSPEC_FLG = 'Y'"
			ElseIf pCode14 = "S" Then
				lgStrSQL = lgStrSQL & " AND A.SHIP_INSPEC_FLG = 'Y'"
			End If
			
			If pCode16 <> "" Then
				lgStrSQL = lgStrSQL & " AND A.ITEM_ACCT = " & pCode16
			End If
			
			If PCode17 <> "" Then
				lgStrSQL = lgStrSQL & " AND b.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & PCode17 & " ))"
			End If
			
			lgStrSQL = lgStrSQL & " " & strWhere
			
			lgStrSQL = lgStrSQL & " ORDER BY B.ITEM_NM, A.ITEM_CD " 
			
   End Select

	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubGetSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubGetSQLStatements(pDataType, pData1)
	 Select Case pDataType

		Case "GP"
			lgStrSQL = "SELECT TOP 1 ITEM_GROUP_NM FROM B_ITEM_GROUP "
			lgStrSQL = lgStrSQL & " WHERE ITEM_GROUP_CD = " & pData1
	End Select	
	

End Sub



'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '��: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>