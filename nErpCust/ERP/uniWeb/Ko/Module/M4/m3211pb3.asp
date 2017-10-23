<%
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3211pb3.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Local L/C No POPUP Query Transaction 처리용 ASP							*
'*  7. Modified date(First) : 2000/04/10																*
'*  8. Modified date(Last)  : 2000/04/10																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/30 : Coding Start												*
'********************************************************************************************************

Response.Expires = -1															'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True															'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncServer.asp" -->
<%
																				'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Dim strMode																		'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

strMode = Request("txtMode")													'☜ : 현재 상태를 받음 

Select Case strMode
	Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
		Dim S32118																' S/O Referenc 조회용 Object
		Dim B1H019																' Business Partner Lookup용 Object
		Dim strApplicantNm
		Dim LngRow
		Dim intGroupCount

		Err.Clear																'☜: Protect system from crashing

		If Request("txtApplicant") = "" Then									'⊙: 조회를 위한 값이 들어왔는지 체크 
			Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)
			Response.End
		End If

		lgStrPrevKey = Request("lgStrPrevKey")

		'---------------------------------- L/C Header Data Query ----------------------------------

		Set S32118 = Server.CreateObject("S32118.S32118ListLcHdrSvr")

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set B1H019 = Nothing												'☜: ComProxy UnLoad
			Set S32118 = Nothing												'☜: ComProxy UnLoad
			Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)
			Response.End														'☜: Process End
		End If

		'-----------------------
		'Data manipulate  area(import view match)
		'-----------------------
		S32118.ImportBBizPartnerBpCd = Request("txtApplicant")
		S32118.ImportNextSLcHdrLcKindAsString = "M"
		S32118.ImportNextSLcHdrLcNo = Request("lgStrPrevKey")
		S32118.CommandSent = "LIST"
		S32118.ServerLocation = ggServerIP
		S32118.ComCfg = gConnectionString
		'-----------------------
		'Com action area
		'-----------------------
		S32118.Execute

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set B1H019 = Nothing												'☜: ComProxy UnLoad
			Set S32118 = Nothing												'☜: ComProxy UnLoad
			Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)
			Response.End														'☜: Process End
		End If

		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------
		If Not (S32118.OperationStatusMessage = MSG_OK_STR) Then
			Call DisplayMsgBox(S32118.OperationStatusMessage, vbInformation, "", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code

			Set B1H019 = Nothing												'☜: ComProxy UnLoad
			Set S32118 = Nothing												'☜: ComProxy UnLoad
			Response.End														'☜: Process End
		End If

		intGroupCount = S32118.ExportGroupCount

		If S32118.ExportItemSLcHdrLcNo(intGroupCount) = S32118.ExportNextSLcHdrLcNo Then
			StrNextKey = ""
		Else
			StrNextKey = S32118.ExportNextSLcHdrLcNo
		End If

		'-----------------------
		'Result data display area
		'-----------------------
%>
<Script Language=VBScript>
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp
    Dim strData

	With parent
		.txtApplicantNm.value = "<%=ConvSPChars(strApplicantNm)%>"
		LngMaxRow = .vspdData.MaxRows											'Save previous Maxrow

<%

	For LngRow = 1 To intGroupCount

%>
    
        strData = strData & Chr(11) & "<%=ConvSPChars(S32118.ExportItemSLcHdrLcNo(LngRow))%>"							'1
        strData = strData & Chr(11) & "<%=ConvSPChars(S32118.ExportItemSLcHdrLcDocNo(LngRow))%>"							'2
        strData = strData & Chr(11) & "<%=ConvSPChars(S32118.ExportItemSLcHdrLcAmendSeq(LngRow))%>"						'3
        strData = strData & Chr(11) & "<%=UNIDateClientFormat(S32118.ExportItemSLcHdrOpenDt(LngRow))%>"		'4
        strData = strData & Chr(11) & "<%=UNIDateClientFormat(S32118.ExportItemSLcHdrExpiryDt(LngRow))%>"	'5
        strData = strData & Chr(11) & "<%=ConvSPChars(S32118.ExportItemBBankBankCd(LngRow))%>"							'6
        strData = strData & Chr(11) & "<%=ConvSPChars(S32118.ExportItemSLcHdrLcType(LngRow))%>"							'7
        strData = strData & Chr(11) & Chr(12)

<%
    Next
%>

		.ggoSpread.Source = .vspdData 
		.ggoSpread.SSShowData strData
		
		.lgStrPrevKey = "<%=StrNextKey%>"

		If .vspdData.MaxRows < .C_SHEETMAXROWS And .lgStrPrevKey <> "" Then	<% ' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 %>
			.DbQuery
		Else
			.DbQueryOk
		End If

		.vspdData.focus
	End With
</Script>
<%
	Set B1H019 = Nothing														'☜: Unload Comproxy
	Set S32118 = Nothing														'☜: ComProxy UnLoad
	Response.End																'☜: Process End

	Case Else
		Response.End
End Select
%>