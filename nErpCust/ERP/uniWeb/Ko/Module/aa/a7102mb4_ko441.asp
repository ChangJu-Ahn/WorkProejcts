<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

fod<% 
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7102mb4
'*  4. Program Name         : 고정자산취득내역등록
'*  5. Program Desc         : 취득번호별 자산Master 조회(6th)
'*  6. Comproxy List        : +As0029LookupSvr
'                             +B1a028ListMinorCode
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2001/05/25
'*  9. Modifier (First)     : 김희정
'* 10. Modifier (Last)      : 김희정
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.

'#########################################################################################################
'												1. Include
'##########################################################################################################
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'#########################################################################################################
'												2. 조건부
'##########################################################################################################
Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
''On Error Resume Next														'☜: 

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
strMode = Request("txtMode")		
										'☜ : 현재 상태를 받음
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
	
'#########################################################################################################
'												2.1 조건 체크
'##########################################################################################################
If strMode = "" Then
	Response.End 
ElseIf strMode <> CStr(UID_M0001) Then											'☜: 조회 전용 Biz 이므로 다른값은 그냥 종료함
	Call ServerMesgBox("700118", vbInformation, I_MKSCRIPT)		'⊙: 조회 전용인데 다른 상태로 요청이 왔을 경우, 필요없으면 빼도 됨, 메세지는 ID값으로 사용해야 함
	Response.End 
ElseIf Trim(Request("txtAcqNo")) = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크
	Call ServerMesgBox("700112", vbInformation, I_MKSCRIPT)						'⊙:
	Response.End 
End If

'#########################################################################################################
'												2. 업무 처리 수행부
'##########################################################################################################
'#########################################################################################################
'												2.1. 변수, 상수 선언 
'##########################################################################################################
Dim strAs0029																	'☆ : 조회용 ComProxy Dll 사용 변수
Dim IntRows
Dim IntCols
Dim sList
Dim strData_i, strData_m
Dim vbIntRet
Dim intCount,IntCount1
Dim IntCurSeq
Dim LngMaxRow
Dim StrNextKey_m,plgStrPrevKey_m

'#########################################################################################################
'												2.2. 요청 변수 처리
'##########################################################################################################
	plgStrPrevKey_m = Request("lgStrPrevKey_m")

	Set strAs0029 = Server.CreateObject("As0029.As0029AcqLookupSvr")

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Set strAs0029 = Nothing												'☜: ComProxy Unload
		Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)	'⊙:
		Response.End														'☜: 비지니스 로직 처리를 종료함
	End If
	
	'-----------------------
	'Data manipulate  area(import view match)
	'-----------------------
	If plgStrPrevKey_m = "" Then
		strAs0029.ImportNextAAssetMasterAsstNo = ""
	Else
		strAs0029.ImportNextAAssetMasterAsstNo = plgStrPrevKey_m
	End If
	
	strAs0029.ImportAAssetAcqAcqNo	= Trim(Request("txtAcqNo"))
	strAs0029.ImportFgIefSuppliedSelectChar = "B"      ' 두번째 Tab에 대한 조회 표시
	strAs0029.ServerLocation		= ggServerIP

	strAs0029.ComCfg = gConnectionString
	strAs0029.Execute
   
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.number <> 0 Then
		Set strAs0029 = Nothing
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'⊙:
		Response.End 								'☜: 비지니스 로직 처리를 종료함
	End If

	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If Not (strAs0029.OperationStatusMessage = MSG_OK_STR) Then
		Call DisplayMsgBox(strAs0029.OperationStatusMessage, vbInformation, "", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
		Set strAs0029 = Nothing												'☜: ComProxy Unload
		Response.End
	End If

	'#########################################################################################################
	'												2.4. HTML 결과 생성부
	'##########################################################################################################

%>
	<Script Language=vbscript>
	Dim LngMaxRow
	Dim varMakeFg
	Dim varApAmt

		'**************************************************************************************
		'        The Query Part Of Multi for Asset Master
		'**************************************************************************************	
<%
		LngMaxRow = Request("txtMaxRows_m")											'Save previous Maxrow                                                
		intCount = strAs0029.ExpGroupCount

		If strAs0029.ExportItemAAssetMasterAsstNo(intCount) = strAs0029.ExportNextAAssetMasterAsstNo Then
			StrNextKey_m = ""
		Else
			StrNextKey_m = strAs0029.ExportNextAAssetMasterAsstNo
		End If
    
%>


	With parent
	
		LngMaxRow = .frm1.vspdData.MaxRows
		.frm1.vspdData.MaxRows = LngMaxRow

<%    
		For IntRows = 1 To intCount
%>


<% 
	Dim lgCurrency
	lgCurrency = ConvSPChars(strAs0029.ExportItemAAssetMasterdoccur(IntRows))
%>	
			strData_m = strData_m & Chr(11) & "<%=ConvSPChars(strAs0029.ExportItemBAcctDeptDeptCd(IntRows))%>"
			strData_m = strData_m & Chr(11) & ""	    
			strData_m = strData_m & Chr(11) & "<%=ConvSPChars(strAs0029.ExportItemBAcctDeptDeptNm(IntRows))%>"

			strData_m = strData_m & Chr(11) & "<%=ConvSPChars(strAs0029.ExportItemAAcctAcctCd(IntRows))%>"
			strData_m = strData_m & Chr(11) & ""
			strData_m = strData_m & Chr(11) & "<%=ConvSPChars(strAs0029.ExportItemAAcctAcctNm(IntRows))%>"

			strData_m = strData_m & Chr(11) & "<%=ConvSPChars(strAs0029.ExportItemAAssetMasterAsstNo(IntRows))%>"                		
        	strData_m = strData_m & Chr(11) & "<%=ConvSPChars(strAs0029.ExportItemAAssetMasterAsstNm(IntRows))%>"        
    
			'strData_m = strData_m & Chr(11) & "<%=UNINumClientFormat(strAs0029.ExportItemAAssetMasterAcqAmt(IntRows),    ggAmtOfMoney.DecPoint, 0)%>"
			'strData_m = strData_m & Chr(11) & "<%=UNINumClientFormat(strAs0029.ExportItemAAssetMasterAcqLocAmt(IntRows), ggAmtOfMoney.DecPoint, 0)%>"

			strData_m = strData_m & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(strAs0029.ExportItemAAssetMasterAcqAmt(IntRows),lgCurrency,ggAmtOfMoneyNo, "X" , "X")%>"
			strData_m = strData_m & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(strAs0029.ExportItemAAssetMasterAcqLocAmt(IntRows),gCurrency,ggAmtOfMoneyNo, gLocRndPolocyNo, "X")%>"
			
			strData_m = strData_m & Chr(11) & "<%=UNINumClientFormat(strAs0029.ExportItemAAssetMasterAcqQty(IntRows),    ggQty.DecPoint, 0)%>"

			strData_m = strData_m & Chr(11) & "<%=ConvSPChars(strAs0029.ExportItemAAssetMasterRefNo(IntRows))%>"   	    	    	    	    
			strData_m = strData_m & Chr(11) & "<%=ConvSPChars(strAs0029.ExportItemAAssetMasterAssetDesc(IntRows))%>"

			strData_m = strData_m & Chr(11) & LngMaxRow + <%=IntRows%>
			strData_m = strData_m & Chr(11) & Chr(12)

<%
		Next
%>    
	    .ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData strData_m
		
		.lgStrPrevKey_m = "<%=StrNextKey_m%>"		
		
		If .frm1.vspdData.MaxRows < .C_SHEETMAXROWS_m And .lgStrPrevKey_m <> "" Then	
		 	.DbQuery2
		Else    
			.DbQueryOk2
		End If
	
	End With	

<%
	Set strAs0029 = Nothing
%>

</Script>
