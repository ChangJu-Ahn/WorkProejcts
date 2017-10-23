<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029B("I", "*", "NOCOOKIE","MB") %>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2115MB2
'*  4. Program Name         : 부적합처리 등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

On Error Resume Next												
Call HideStatusWnd 

Dim strinsp_class_cd
strinsp_class_cd = "F"	'@@@주의	
	
Dim PQIG110																	'☆ : 조회용 ComProxy Dll 사용 변수 
	
Dim lgIntFlgMode
Dim LngMaxRow
	
Dim arrRowVal								'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim arrColVal								'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim strStatus								'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
	
Dim strInspReqNo
Dim strInspResultNo
Dim LngRow
	
Dim IG1_import_group
Const IG1_select_char = 0
Const IG1_disposition_cd = 1
Const IG1_qty = 2
Const IG1_remark = 3
	
Set PQIG110 = Server.CreateObject("PQIG110.cQMtInspDispSvr")

If CheckSystemError(Err,True) Then											'☜: ComProxy Unload
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

LngMaxRow = CInt(Request("txtMaxRows"))					'☜: 최대 업데이트된 갯수 
lgIntFlgMode = CInt(Request("txtFlgMode"))					'☜: 저장시 Create/Update 판별 
strInspReqNo = UCase(Request("txtInspReqNo"))
strInspResultNo = 1
	
If Request("txtSpread") <> "" Then
	arrRowVal = Split(Request("txtSpread"), gRowSep)
	LngMaxRow = UBound(arrRowVal) - 1
	Redim IG1_import_group(LngMaxRow,3)
	For LngRow = 0 To LngMaxRow
		arrColVal = Split(arrRowVal(LngRow), gColSep)
		strStatus = UCase(arrColVal(0))
		IG1_import_group(LngRow,IG1_disposition_cd)	= arrColVal(1)
		IG1_import_group(LngRow,IG1_Select_Char)	= strStatus
		If strStatus = "C" or strStatus = "U" Then
			IG1_import_group(LngRow,IG1_Qty) = UniConvNum(arrColVal(2), 0)
			IG1_import_group(LngRow,IG1_Remark)	 = arrColVal(3)
		End If
	Next
		
	Call PQIG110.Q_MAINT_INSP_DISPOSIT_SVR(gStrGlobalCollection,strInspReqNo,strInspResultNo,IG1_import_group)
		
	If CheckSYSTEMError(Err,True) Then
		Set PQIG110 = Nothing 
		Response.End
	End If
End If
Set PQIG110 = Nothing
%>
<Script Language=vbscript>
With parent																		'☜: 화면 처리 ASP 를 지칭함 
	.DbSaveOk
End With
</Script>