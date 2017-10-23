<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : L2111MB1
'*  4. Program Name         : 수주정보등록(IF)
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/04/16
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwang seong bae
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf()
Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB")
%>
<%													

On Error Resume Next														

Call HideStatusWnd

' 조회조건 
Const	C_SoldToParty = 3

' 조회결과 Output Index
Const	C_INF_NO			= 0
Const	C_INF_SEQ			= 1
Const	C_RCPT_DT			= 2
Const	C_DOC_ISSUE_DT		= 3
Const	C_SOLD_TO_PARTY		= 4
Const	C_SOLD_TO_PARTY_NM	= 5
Const	C_DEAL_TYPE			= 6
Const	C_PAY_METH			= 7
Const	C_DOC_NO			= 8
Const	C_DOC_SEQ			= 9
Const	C_SHIP_TO_PARTY		= 10
Const	C_SHIP_TO_PARTY_NM	= 11
Const	C_ITEM_CD			= 12
Const	C_ITEM_NM			= 13
Const	C_QTY				= 14
Const	C_UNIT				= 15
Const	C_CUR				= 16
Const	C_PRICE				= 17
Const	C_PRICE_FLAG		= 18
Const	C_VAT_INC_FLAG 		= 19
Const	C_VAT_TYPE	 		= 20
Const	C_VAT_TYPE_NM 		= 21
Const	C_VAT_RATE			= 22
Const	C_DLVY_DT	 		= 23
Const	C_REMARK	 		= 24

Dim iStrMode
Dim iStrSvrData, iStrSvrData2, iStrNextKey
Dim iObjPL2G111
Dim iArrListOut			' Result of recordset.getrow(), it means iArrListOut is two dimension array (column, row)
Dim iArrListGroupOut	' Result of recordset.getrow(), it means iArrListGroupOut is two dimension array (column, row)
Dim iArrWhereIn, iArrWhereOut
Dim iLngRow
Dim iLngLastRow			' The last row number in the spread
Dim iLngSheetMaxRows	' Row numbers to be displayed in the spread.
Dim iLngErrorPosition
Const C_SHEETMAXROWS = 30  

iStrMode = Request("txtMode")												'☜ : 현재 상태를 받음 

Select Case iStrMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
    Err.Clear                                                                

	iLngSheetMaxRows = CLng(C_SHEETMAXROWS)
	
    Set iObjPL2G111 = Server.CreateObject("PL2G111.cListSInfSoHdrForSSoHdr")    
  
    If CheckSYSTEMError(Err,True) = True Then
		Response.End		
    End If

    Call iObjPL2G111.ListRows (gStrGlobalCollection, iLngSheetMaxRows, Request("txtWhere"), Request("lgStrPrevKey"), _
						  iArrListOut, iArrWhereOut)
	
    Set iObjPL2G111 = Nothing

	If CheckSYSTEMError(Err,True) = True Then
	   Response.End 
	End If
    
    ' Check Query Condition - 처음 조회시 
    If Request("lgStrPrevKey") = "" Then
		Call BeginScriptTag()

		iArrWhereIn = Split(Request("txtWhere"), gColSep)
		' 거래처 
		If iArrWhereIn(C_SoldToParty) = iArrWhereOut(0, C_SoldToParty) Then
			Call WriteConDesc("txtConSoldToPartyNm", iArrWhereOut(1, C_SoldToParty))
		Else
			Call WriteConDesc("txtConSoldToPartyNm", "")
			Call ConNotFound("txtConSoldToParty")
'			Response.Write "Call parent.SetToolbar(""11000000000011"") " & vbCr   
			Response.End		
		End If

		' 해당되는 자료가 존재하지 않습니다.
		If IsArray(iArrListOut) = 0 Then
			Call DataNotFound("txtConRcptFromDt")
			Response.End		
		End If
		
		Call EndScriptTag()
	End If
    
	'------------------------
	'Result data display area
	'------------------------ 
	iLngLastRow = CLng(Request("txtLastRow"))

	' Set Next key
	If Ubound(iArrListOut,2) = iLngSheetMaxRows Then
		iStrNextKey = iArrListOut(C_INF_NO, iLngSheetMaxRows) & gColSep & iArrListOut(C_INF_SEQ, iLngSheetMaxRows)
		iLngSheetMaxRows  = iLngSheetMaxRows - 1
	Else
		iStrNextKey = ""
		iLngSheetMaxRows = Ubound(iArrListOut,2)
	End If

	' Spread1
	Response.Write "'iLngSheetMaxRows = " & iLngSheetMaxRows & vbcr
	iStrSvrData = ""
   	For iLngRow = 0 To iLngSheetMaxRows
   		iStrSvrData = iStrSvrData & gColSep & "" & _
   									gColSep & UNIDateClientFormat(iArrListOut(C_RCPT_DT,iLngRow)) & _
   									gColSep & UNIDateClientFormat(iArrListOut(C_DOC_ISSUE_DT,iLngRow)) & _
   									gColSep & iArrListOut(C_SOLD_TO_PARTY,iLngRow) & _
   									gColSep & iArrListOut(C_SOLD_TO_PARTY_NM,iLngRow) & _
   									gColSep & iArrListOut(C_SHIP_TO_PARTY,iLngRow) & _
   									gColSep & iArrListOut(C_SHIP_TO_PARTY_NM,iLngRow) & _
   									gColSep & iArrListOut(C_ITEM_CD,iLngRow) & _
   									gColSep & iArrListOut(C_ITEM_NM,iLngRow) & _
   									gColSep & UNINumClientFormat(iArrListOut(C_QTY,iLngRow), ggQty.DecPoint, 0) & _
   									gColSep & iArrListOut(C_UNIT,iLngRow) & _
   									gColSep & "0" & _
   									gColSep & "0" & _
   									gColSep & iArrListOut(C_CUR,iLngRow) & _
   									gColSep & UNIConvNumDBToCompanyByCurrency(iArrListOut(C_PRICE,iLngRow),iArrListOut(C_CUR,iLngRow),ggUnitCostNo, "X" , "X") & _
   									gColSep & "0" & _
   									gColSep & iArrListOut(C_PRICE_FLAG,iLngRow) & _
   									gColSep & "" & _
   									gColSep & iArrListOut(C_VAT_INC_FLAG,iLngRow) & _
   									gColSep & "" & _
   									gColSep & UNIDateClientFormat(iArrListOut(C_DLVY_DT,iLngRow)) & _
   									gColSep & "" & _
   									gColSep & iArrListOut(C_VAT_TYPE,iLngRow) & _
   									gColSep & "" & _
   									gColSep & iArrListOut(C_VAT_TYPE_NM,iLngRow) & _
   									gColSep & UNINumClientFormat(iArrListOut(C_VAT_RATE,iLngRow), ggExchRate.DecPoint, 0) & _
   									gColSep & "" & _
   									gColSep & "" & _
   									gColSep & "" & _
   									gColSep & "" & _
   									gColSep & "" & _
   									gColSep & "" & _
   									gColSep & "" & _
   									gColSep & iArrListOut(C_DOC_NO,iLngRow) & _
   									gColSep & iArrListOut(C_DOC_SEQ,iLngRow) & _
   									gColSep & iArrListOut(C_REMARK,iLngRow) & _
   									gColSep & "" & _
   									gColSep & iArrListOut(C_DEAL_TYPE,iLngRow) & _
   									gColSep & iArrListOut(C_PAY_METH,iLngRow) & _
   									gColSep & iArrListOut(C_INF_NO,iLngRow) & _
   									gColSep & iArrListOut(C_INF_SEQ,iLngRow) & _
   									gColSep & iLngLastRow + iLngRow + 1 & gColSep & gRowSep
	
   	Next
   	
   	Response.Write "'" & iStrSvrData
	Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> " & vbCr   
    Response.Write " Parent.ggoSpread.Source = Parent.frm1.vspdData " & vbCr
    Response.Write  "Parent.frm1.vspdData.Redraw = False  "      & vbCr      
    Response.Write  "Parent.ggoSpread.SSShowData   """ & ConvSPChars(iStrSvrData) & """ ,""F""" & vbCr
    Response.Write " Parent.lgStrPrevKey = """ & ConvSPChars(iStrNextKey) & """" & vbCr
    
    Response.Write " Parent.DbQueryOk" & vbCr   
	Response.Write  "Parent.frm1.vspdData.Redraw = True  "       & vbCr      
	Response.Write "</SCRIPT> "		

Case CStr(UID_M0002)																'☜: 저장 요청을 받음 

    Err.Clear																		

	Dim iVarSoldToParty, iVarOverLimitAmt
	
    Set iObjPL2G111 = Server.CreateObject("PL2G111.cMaintSSoHdrByInf")  
    
    If CheckSYSTEMError(Err,True) = True Then
		Response.End		
    End If

	Call iObjPL2G111.Maintain (gStrGlobalCollection, Trim(Request("txtHeader")), Trim(Request("txtSpread")), iLngErrorPosition, iVarSoldToParty, iVarOverLimitAmt)

    Set iObjPL2G111 = Nothing

	If Err.number <> 0 Then
		' 여신한도가 초과된 경우 
		If InStr(1, err.Description, "B_MESSAGE" & Chr(11) & "201929") > 0 Or InStr(1, err.Description, "B_MESSAGE" & Chr(11) & "201722") > 0 Then
			Call BeginScriptTag()
			' 여신한도를 초과하였습니다.(주문처: - )
			Response.Write("Call parent.DisplayMsgBox(""201931"", ""X"", """", ""주문처 : " & iVarSoldToParty & " - " & UNINumClientFormat(iVarOverLimitAmt, ggAmtOfMoney.DecPoint, 0) & gCurrency & """)" & vbCr )
			Call EndScriptTag()
		ElseIf CheckSYSTEMError2(Err, True, iLngErrorPosition & "행","","","","") = True Then
			Call BeginScriptTag()
			Response.Write " Call Parent.SubSetErrPos(" & iLngErrorPosition & ")" & vbCr
			Call EndScriptTag()
		End If
		Response.End 
	End If
	
	Call BeginScriptTag()
    Response.Write " Parent.DBSaveOk "      & vbCr   
	Call EndScriptTag()
    
' 삭제 
Case CStr(UID_M0003)																'☜: 삭제 
									
End Select

'----------------------------------------------------------------------------------------------------------
' Write the Result
' 결과Html을 작성한다.
'----------------------------------------------------------------------------------------------------------
Sub BeginScriptTag()
	Response.Write "<Script language=VBScript> " & VbCr
End Sub

Sub EndScriptTag()
	Response.Write "</Script> " & VbCr
End Sub

' 데이터가 존재하지 않는 경우 처리 Script 작성(조회조건 포함)
Sub ConNotFound(ByVal pvStrField)
	Response.Write " Call Parent.DisplayMsgBox(""970000"", ""X"", parent.frm1." & pvStrField & ".alt, ""X"") " & VbCr
	Response.Write "Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' 조회조건에 해당하는 명을 Display하는 Script 작성 
Sub WriteConDesc(ByVal pvStrField, Byval pvStrFieldDesc)
	Response.Write "Parent.frm1." & pvStrField & ".value = """ & ConvSPChars(pvStrFieldDesc) & """" &VbCr
End Sub

' 데이터가 존재하지 않는 경우 처리 Script 작성 
Sub DataNotFound(ByVal pvStrFocusField)
	Response.Write " Call Parent.DisplayMsgBox(""900014"", ""X"", ""X"", ""X"") " & VbCr
	Response.Write "Parent.frm1." & pvStrFocusField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

%>
