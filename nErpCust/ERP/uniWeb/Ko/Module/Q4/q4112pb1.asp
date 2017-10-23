<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029B("I", "*", "NOCOOKIE","PB") %>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q4112PB1
'*  4. Program Name         : 
'*  5. Program Desc         : 부적합처리 팝업 
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

Const C_SHEETMAXROWS_D = 100

'EXPORT VIEW
Const E1_disposition_cd = 0
Const E1_disposition_nm = 1
Const E1_insp_class_cd = 2
Const E1_insp_class_nm = 3
Const E1_stock_type_cd = 4
Const E1_stock_type_nm = 5
    
Dim PQIG360

Dim LngMaxRow
Dim StrNextKey1
Dim StrNextKey2
Dim i
Dim StrData

Dim iExportDisposition

If Request("lgQueryFlag") = "0" Then		'추가 조회 
	LngMaxRow = Request("txtMaxRows")
	StrNextKey1 = Request("txtDispositionCd")
	StrNextKey2 = Request("txtDispositionNm")
Else										'신규 조회 
	LngMaxRow = 0
	StrNextKey1 = ""
	StrNextKey2 = ""
End If

' 해당 Business Object 생성 
Set PQIG360 = Server.CreateObject("PQIG360.cQLiDispositionSimple")
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PQIG360.Q_LIST_DISPOSITION_SIMPLE_SVR(gStrGlobalCollection, _
									C_SHEETMAXROWS_D, _
									Request("QueryType"), _
									Request("txtDispositionCd"), _
									Request("txtDispositionNm"), _
									Request("txtInspClassCd"), _
									iExportDisposition)

If CheckSYSTEMError(Err,True) = True Then
	Set PQIG360 = Nothing
	Response.End
End If

For i = 0 To UBound(iExportDisposition, 1)
    If i < C_SHEETMAXROWS_D Then
    	strData = strData & Chr(11) & ConvSPChars(Trim(iExportDisposition(i, E1_disposition_cd))) _
						  & Chr(11) & ConvSPChars(Trim(iExportDisposition(i, E1_disposition_nm))) _
						  & Chr(11) & ConvSPChars(Trim(iExportDisposition(i, E1_insp_class_nm))) _
						  & Chr(11) & ConvSPChars(Trim(iExportDisposition(i, E1_stock_type_nm))) _
						  & Chr(11) & ConvSPChars(Trim(iExportDisposition(i, E1_insp_class_cd))) _
						  & Chr(11) & ConvSPChars(Trim(iExportDisposition(i, E1_stock_type_nm))) _
						  & Chr(11) & Chr(12)
    Else
		StrNextKey1 = ConvSPChars(Trim(iExportDisposition(i, E1_disposition_cd)))
		StrNextKey2 = ConvSPChars(Trim(iExportDisposition(i, E1_disposition_nm)))
    End If
Next  

Set PQIG360 = Nothing
%>
<Script Language="vbscript">   
    With parent
		.ggoSpread.Source = .vspdData 
		.ggoSpread.SSShowDataByClip "<%=strData%>"
		.vspdData.focus
		
		.lgStrPrevKey1 = "<%=StrNextKey1%>"
		.lgStrPrevKey2 = "<%=StrNextKey2%>"
		
		' Request값을 hidden Varialbe로 넘겨줌 
		.hInspClassCd = "<%=ConvSPChars(Request("txtInspClassCd"))%>"
		.DbQueryOk
	    
	End with
</Script>
