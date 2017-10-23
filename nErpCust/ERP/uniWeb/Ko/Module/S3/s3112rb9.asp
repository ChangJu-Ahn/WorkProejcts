<%@ LANGUAGE=VBSCript%>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업																		*
'*  2. Function Name        : 																			*
'*  3. Program ID           : S3112RB9																	*
'*  4. Program Name         : 클래스참조(수주내역등록)													*
'*  5. Program Desc         :																			*
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/02/06																*
'*  8. Modified date(Last)  : 																			*
'*  9. Modifier (First)     : Hwang Seong Bae															*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 																			*
'********************************************************************************************************
%>
<% Option Explicit %>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

On Error Resume Next                                                             
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd                                                               '☜: Hide Processing message
'---------------------------------------Common-----------------------------------------------------------
CONST C_QryPlant	= 1
CONST C_QryMain		= 2
	
lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '☜: Query
        Call SubBizQuery() 
    Case CStr(UID_M0002)                                                         '☜: Save,Update
        'Call SubBizSave()
    Case CStr(UID_M0003)                                                         '☜: Delete
        'Call SubBizDelete()
End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	On Error Resume Next
	
	Dim iStrSvrData, iStrCurrency
	Dim iLngRow
	
    Call SubOpenDB(lgObjConn)

    ' 공장코드의 존재유무 Check
	If Request("txtPlantCd") <> "" Then
	    Call SubMakeSQLStatements(C_QryPlant)
	    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then                 'If data not exists
		    Response.Write "<Script Language=vbscript>" & vbCr
			Response.Write "With parent.frm1"           & vbCr
			Response.Write ".txtConPlantCd.value = """ & ConvSPChars(lgObjRs("PLANT_CD")) & """" & vbCr		
			Response.Write ".txtConPlantNm.value = """ & ConvSPChars(lgObjRs("PLANT_NM")) & """" & vbCr
			Response.Write "End With "           & vbCr
		    Response.Write "</Script>" & vbCr
		Else
			If ChkNotFound Then
				Call DisplayMsgBox("970000", vbInformation, "공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
				Response.Write "<Script Language=vbscript>" & vbCr
				Response.Write "parent.frm1.txtConPlantNm.value = """"" & vbCr
				Response.Write "</Script>" & vbCr
			End If
			
		    Exit Sub
		End If
	End If
    
    ' 관련 품목 조회 
    Call SubMakeSQLStatements(C_QryMain)
 
	If Not FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then                 'If data not exists
		If ChkNotFound Then
			'해당되는 자료가 없습니다.
			Call DisplayMsgBox("800076", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found. 
		End If
	    Exit Sub
	End If

	' 화폐단위 
	iStrCurrency = Request("txtCurrency")
	iStrSvrData = ""
	iLngRow = 0

	While (Not lgObjRs.EOF)
		iLngRow = iLngRow + 1
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs("PLANT_CD"))			' 공장 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs("PLANT_NM"))			' 공장명 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs("CLASS_CD"))			' 클래스 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs("CLASS_NM"))			' 클래스명 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs("CHAR_VALUE_CD1"))	' 사양1
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs("CHAR_VALUE_NM1"))	' 사양설명1
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs("CHAR_VALUE_CD2"))	' 사양2
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs("CHAR_VALUE_NM2"))	' 사양설명2
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs("BASIC_UNIT"))		' 단위 
   		iStrSvrData = iStrSvrData & gColSep
   		iStrSvrData = iStrSvrData & gColSep & "0"										' 수량 
   		iStrSvrData = iStrSvrData & gColSep & UNIConvNumDBToCompanyByCurrency(lgObjRs("PRICE"),iStrCurrency,ggUnitCostNo, "X" , "X") ' 단가 
   		iStrSvrData = iStrSvrData & gColSep & "0"										' 금액 
   		iStrSvrData = iStrSvrData & gColSep & UNINumClientFormat(lgObjRs("INV_QTY"), ggQty.DecPoint, 0)		' 재고수량 
   		iStrSvrData = iStrSvrData & gColSep & UNINumClientFormat(lgObjRs("RCPT_QTY"), ggQty.DecPoint, 0)	' 입고예정량 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs("ITEM_CD"))			' 품목코드 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs("ITEM_NM"))			' 품목명 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs("SPEC"))				' 규격 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs("HS_CD"))				' H.S. 부호 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs("VAT_TYPE"))			' VAT 유형 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs("VAT_NM"))			' VAT 유형명 
   		iStrSvrData = iStrSvrData & gColSep & lgObjRs("VAT_RATE")						' VAT 율 
   		iStrSvrData = iStrSvrData & gColSep & "0"										' 이전수량 
   		iStrSvrData = iStrSvrData & gColSep & "0"										' 이전금액 
   		iStrSvrData = iStrSvrData & gColSep												' SPREAD2의 관련 ROW NUMBER
   		iStrSvrData = iStrSvrData & gColSep & iLngRow 
   		iStrSvrData = iStrSvrData & gColSep & gRowSep
		lgObjRs.MoveNext
	Wend
	
	lgObjRs.Close
	lgObjConn.Close
	Set lgObjRs = Nothing
	Set lgObjConn = Nothing
	
	Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> " & vbCr   
	Response.Write " Parent.SetColHiddenByClass " & vbCr   
    Response.Write " Parent.ggoSpread.Source = Parent.frm1.vspdData " & vbCr
    Response.Write " Parent.ggoSpread.SSShowDataByClip """ & iStrSvrData & """" & vbCr
    Response.Write " Parent.DbQueryOk" & vbCr   
	Response.Write "</SCRIPT> "
End Sub    

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
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
    Call SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

Function ChkNotFound()
	ChkNotFound = False
	
	' Error가 발행한 경우는 recordset이 닫혔있다.
	If lgObjRs.State = adStateOpen  Then
		lgObjRs.Close
		lgObjConn.Close
		Set lgObjRs = Nothing
		Set lgObjConn = Nothing
		ChkNotFound = True
	Else
		Set lgObjRs = Nothing
		Set lgObjConn = Nothing
	End If
End Function
'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(ByVal pvIntWhere)
	Dim iStrSelectList, iStrFromList, iStrWhere, iStrOrderBy
	
	Select Case pvIntWhere
		Case C_QryPlant
			iStrSelectList = " SELECT PLANT_CD, PLANT_NM "
			iStrFromList = " FROM dbo.B_PLANT "
			iStrWhere = " WHERE PLANT_CD =  " & FilterVar(Request("txtPlantCd"), "''", "S") & " "
			iStrOrderBy = ""

		Case C_QryMain
			iStrSelectList = "SELECT ITP.PLANT_CD, PT.PLANT_NM, IT.CLASS_CD, CL.CLASS_NM, "
			iStrSelectList = iStrSelectList & " IT.CHAR_VALUE_CD1, CV1.CHAR_VALUE_NM AS CHAR_VALUE_NM1, IT.CHAR_VALUE_CD2, CV2.CHAR_VALUE_NM AS CHAR_VALUE_NM2, "
			iStrSelectList = iStrSelectList & " IT.BASIC_UNIT, dbo.ufn_s_GetItemSalesPrice("
			iStrSelectList = iStrSelectList & FilterVar(Request("txtSoldToParty"), "''", "S") & ", IT.ITEM_CD, "
			iStrSelectList = iStrSelectList & FilterVar(Request("txtDealType"), "''", "S") & ", "
			iStrSelectList = iStrSelectList & FilterVar(Request("txtPayMeth"), "''", "S") & ", IT.BASIC_UNIT, "
			iStrSelectList = iStrSelectList & FilterVar(Request("txtCurrency"), "''", "S") & ", '"
			iStrSelectList = iStrSelectList & UniConvDateToYYYYMMDD(Request("txtSoDt"),gDateFormat, "") & "') AS PRICE, "
			'iStrSelectList = iStrSelectList & UNIConvDate(Request("txtSoDt")) & "') AS PRICE, "
			iStrSelectList = iStrSelectList & " ISNULL(ST.GOOD_ON_HAND_QTY - ST.SCHD_ISSUE_QTY, 0) AS INV_QTY, "
			iStrSelectList = iStrSelectList & " ISNULL(ST.SCHD_RCPT_QTY, 0) AS RCPT_QTY, "
			iStrSelectList = iStrSelectList & " IT.ITEM_CD,	IT.ITEM_NM,	IT.SPEC, "
			iStrSelectList = iStrSelectList & " ISNULL(IT.HS_CD, '') AS HS_CD, ISNULL(IT.VAT_TYPE, '') AS VAT_TYPE, ISNULL(VT.MINOR_NM, '') AS VAT_NM, ISNULL(CF.REFERENCE, 0) AS VAT_RATE "

			iStrFromList = " FROM dbo.B_ITEM IT "
			iStrFromList = iStrFromList & " INNER JOIN dbo.B_ITEM_BY_PLANT ITP ON (ITP.ITEM_CD = IT.ITEM_CD) "
			iStrFromList = iStrFromList & " INNER JOIN dbo.B_PLANT PT ON (PT.PLANT_CD = ITP.PLANT_CD) "
			iStrFromList = iStrFromList & " LEFT OUTER JOIN dbo.I_ONHAND_STOCK ST ON (ST.PLANT_CD = ITP.PLANT_CD AND ST.ITEM_CD = ITP.ITEM_CD) "
			iStrFromList = iStrFromList & " INNER JOIN dbo.B_CLASS CL ON (CL.CLASS_CD = IT.CLASS_CD) "
			iStrFromList = iStrFromList & " INNER JOIN dbo.B_CHAR_VALUE CV1 ON (CV1.CHAR_CD = CL.CHAR_CD1 AND CV1.CHAR_VALUE_CD = IT.CHAR_VALUE_CD1) "
			iStrFromList = iStrFromList & " LEFT OUTER JOIN dbo.B_CHAR_VALUE CV2 ON (CV2.CHAR_CD = CL.CHAR_CD2 AND CV2.CHAR_VALUE_CD = IT.CHAR_VALUE_CD2) "
			iStrFromList = iStrFromList & " LEFT OUTER JOIN dbo.B_MINOR VT ON (VT.MAJOR_CD = " & FilterVar("B9001", "''", "S") & " AND VT.MINOR_CD = IT.VAT_TYPE) "
			iStrFromList = iStrFromList & " LEFT OUTER JOIN dbo.B_CONFIGURATION CF ON (CF.MAJOR_CD = " & FilterVar("B9001", "''", "S") & " AND CF.MINOR_CD = IT.VAT_TYPE AND SEQ_NO = 1) "

			iStrWhere = " WHERE IT.CLASS_CD =  " & FilterVar(Request("txtClassCd"), "''", "S") & " "
			If Request("txtCharValueCd1") <> "" Then
				iStrWhere = iStrWhere & " AND IT.CHAR_VALUE_CD1 >=  " & FilterVar(Request("txtCharValueCd1"), "''", "S") & " "
			End If
			
			If Request("txtCharValueCd2") <> "" Then
				iStrWhere = iStrWhere & " AND IT.CHAR_VALUE_CD2 >=  " & FilterVar(Request("txtCharValueCd2"), "''", "S") & " "
			End If

			If Request("txtPlantCd") <> "" Then
				iStrWhere = iStrWhere & " AND ITP.PLANT_CD =  " & FilterVar(Request("txtPlantCd"), "''", "S") & " "
			End If
			
			iStrOrderBy = " ORDER BY ITP.PLANT_CD, IT.CLASS_CD, IT.CHAR_VALUE_CD1, IT.CHAR_VALUE_CD2"
	End Select
	
	lgStrSql = iStrSelectList & iStrFromList & iStrWhere & iStrOrderBy
End Sub
%>
