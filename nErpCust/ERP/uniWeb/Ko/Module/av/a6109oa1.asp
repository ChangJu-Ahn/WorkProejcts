<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Account Management
'*  3. Program ID           : A6109MA1
'*  4. Program Name         : 공문및표지출력 
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/12/13
'*  8. Modified date(Last)  : 2002/09/11
'*  9. Modifier (First)     : Hersheys
'* 10. Modifier (Last)      : Lee hye young
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->								<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '☆: 해당 위치에 따라 달라짐, 상대 경로  -->

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript"	SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                                                              '☜: indicates that All variables must be declared in advance 

'========================================================================================================= 

<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim lgMpsFirmDate, lgLlcGivenDt											 '☜: 비지니스 로직 ASP에서 참조하므로 Dim 

Dim lgCurName()															'☆ : 개별 화면당 필요한 로칼 전역 변수 
    
 Dim IsOpenPop          

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2
Const TAB3 = 3

Dim  gSelframeFlg

	Dim strYear
	dim strMonth
	Dim strDay
	Dim StartDate
   	Dim strSvrDate

	strSvrDate = "<%=GetSvrDate%>"
	Call ExtractDateFrom(strSvrDate, parent.gServerDateFormat, parent.gServerDateType, strYear,strMonth,strDay)

	StartDate= UNIConvYYYYMMDDToDate(gDateFormat, strYear, strMonth, "01")		'☆: 초기화면에 뿌려지는 시작 날짜 
'	EndDate= UNIConvYYYYMMDDToDate(gDateFormat, strYear, strMonth, strDay)		'☆: 초기화면에 뿌려지는 마지막 날짜 


'========================================================================================================= 
Sub SetDefaultVal()
    Dim svrDate

	svrDate               = UNIDateClientFormat("<%=GetSvrDate%>")

	'If gSelframeFlg = TAB1 Then
		frm1.txtFromIssueDt.text = svrDate
		frm1.txtToIssueDt.text   = svrDate
		'frm1.txtBizAreaCD.value	= gBizArea
		frm1.txtFiscCnt.value	= parent.gFiscCnt
	'Else
		frm1.txtFromIssueDt2.text = svrDate
		frm1.txtToIssueDt2.text   = svrDate
		'frm1.txtBizAreaCD2.value	= gBizArea
		frm1.txtFiscYear.text	= strYear
		
		frm1.txtFromIssueDt3.text = svrDate
		frm1.txtToIssueDt3.text   = svrDate
		'frm1.txtBizAreaCD.value	= gBizArea
		frm1.txtFisc.value	= parent.gFiscCnt
		
		If gSelframeFlg = TAB1 Then
		frm1.txtBizAreaCD.focus
		elseIf gSelframeFlg = TAB2 Then
		frm1.txtBizAreaCD2.focus  
		elseIf gSelframeFlg = TAB3 Then
		frm1.txtBizAreaCD3.focus  
		end if 
		
	'End If	
End Sub

'=======================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","QA") %>
End Sub


'*****************************************  2.1 Pop-Up 함수   ********************************************
'	기능: Pop-Up 
'********************************************************************************************************* 

'------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0, 1,2
			arrParam(0) = "세금신고사업장 팝업"				' 팝업 명칭 
			arrParam(1) = "B_TAX_BIZ_AREA"	 				' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "세금신고사업장코드"				' 조건필드의 라벨 명칭 

			arrField(0) = "TAX_BIZ_AREA_CD"					' Field명(0)
			arrField(1) = "TAX_BIZ_AREA_NM"					' Field명(0)
    
			arrHeader(0) = "세금신고사업장코드"				' Header명(0)
			arrHeader(1) = "세금신고사업장명"				' Header명(0)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0		' 사업장 
				frm1.txtBizAreaCd.focus
			Case 1		' 사업장(두번째탭)
				frm1.txtBizAreaCd2.focus
			Case 2		' 사업장(세번째탭)
				frm1.txtBizAreaCd3.focus	
		End Select
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function



'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------

Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		' 사업장 
				.txtBizAreaCd.focus
				.txtBizAreaCd.value = UCase(Trim(arrRet(0)))
				.txtBizAreaNm.value = arrRet(1)
			Case 1		' 사업장(두번째탭)
				.txtBizAreaCd2.focus
				.txtBizAreaCd2.value = UCase(Trim(arrRet(0)))
				.txtBizAreaNm2.value = arrRet(1)	
			Case 2		' 사업장(두번째탭)
				.txtBizAreaCd3.focus
				.txtBizAreaCd3.value = UCase(Trim(arrRet(0)))
				.txtBizAreaNm3.value = arrRet(1)		
		End Select
	End With	
End Function


'======================================================================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)	 '~~~ 첫번째 Tab 
	gSelframeFlg = TAB1
	Call SetDefaultVal()
	
						 
End Function

Function ClickTab2()

	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)	 '~~~ 두번째 Tab 
	gSelframeFlg = TAB2
	Call SetDefaultVal()
	

End Function

Function ClickTab3()

	If gSelframeFlg = TAB3 Then Exit Function
	Call changeTabs(TAB3)	 '~~~ 두번째 Tab 
	gSelframeFlg = TAB3
	if frm1.Rd_gbn2.checked = true then
		frm1.Rdo_gbn1.checked = true
		Call ggoOper.SetReqAttr(frm1.Rdo_gbn1		, "Q")
		Call ggoOper.SetReqAttr(frm1.Rdo_gbn2		, "Q")
	End if
	Call SetDefaultVal()
    
End Function



Function FncBtnPrint() 
	On Error Resume Next
	
	Dim Var1
	Dim Var2
	Dim Var3
	Dim Var4
	Dim Var5
	Dim Var6
	Dim strUrl
	
	Dim arrParam, arrField, arrHeader
	Dim StrEbrFile
	Dim ObjName
	
    lngPos = 0	

  '//	tab 에따라 체크할 항목이 다르기 때문에 주석으로함 
   '// If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
   '//    Exit Function
   '// End If
	
	If gSelframeFlg = TAB1 Then 
	
		If Trim(frm1.txtBizAreaCD.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtBizAreaCD.Alt, "X") 	
			Exit Function
		End If	
	
		If Trim(frm1.txtFromIssueDt.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtFromIssueDt.Alt, "X") 	
			Exit Function
		End If	
		If Trim(frm1.txtToIssueDt.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtToIssueDt.Alt, "X") 	
			Exit Function
		End If	
		If Trim(frm1.txtFiscCnt.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtFiscCnt.Alt, "X") 	
			Exit Function
		End If	
		If Trim(frm1.cboVatDiv.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.cboVatDiv.Alt, "X") 	
			Exit Function
		End If	
		If Trim(frm1.txtFileName.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtFileName.Alt, "X") 	
			Exit Function
		End If	
		
		If UniConvDateToYYYYMMDD(frm1.txtFromIssueDt.text, parent.gDateFormat, "") > UniConvDateToYYYYMMDD(frm1.txtToIssueDt.text, parent.gDateFormat, "") Then
			IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'⊙: "Will you destory previous data"
			Exit Function
		End If
	
		If IsNumeric(frm1.txtFiscCnt.value) = False Then
			IntRetCD = DisplayMsgBox("229924", "X", "X", "X")							'필수입력 check!!
			frm1.txtFiscCnt.focus
			' 숫자를 입력하십시오 
			Exit Function
		End If
	
		var1 = UCase(Trim(frm1.txtBizAreaCD.value))
		var2 = UniConvDateToYYYYMMDD(frm1.fpDateTime1.text, parent.gDateFormat, "") 
		var3 = UniConvDateToYYYYMMDD(frm1.fpDateTime2.text, parent.gDateFormat, "") 
		var4 = frm1.txtFiscCnt.value 
		var5 = frm1.cboVatDiv.value 
		var6 = frm1.txtFileName.value 
	
		If var1 = "" Then
			var1 = "*"
		Else
		    var1 = UCase(Trim(frm1.txtBizAreaCD.value))
		End If

		If frm1.Rb_WK1.checked = True Then
			' 디스켓출력 
			StrEbrFile = "a6109ma2"
		ElseIf frm1.Rb_WK2.checked = True Then
			' 공문및 표지 출력 
			StrEbrFile = "a6109ma1"
		End If
		StrUrl = StrUrl & "ReportBizAreaCd|" & FilterVar(var1,"","SNM")
		StrUrl = StrUrl & "|FromIssueDt|"	  & var2
		StrUrl = StrUrl & "|ToIssueDt|"	      & var3
		StrUrl = StrUrl & "|FiscCnt|"	      & FilterVar(var4,"","SNM")
		StrUrl = StrUrl & "|VatDiv|"	      & FilterVar(var5,"","SNM")
		StrUrl = StrUrl & "|FileName|"	      & FilterVar(var6,"","SNM")
		StrUrl = StrUrl & "|TmpBizAreaCd|"    & FilterVar(var1,"","SNM")

	ElseIf gSelframeFlg = TAB2 Then 
		
		If Trim(frm1.txtBizAreaCD2.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtBizAreaCD2.Alt, "X") 	
			Exit Function
		End If	
	
		If Trim(frm1.txtFromIssueDt2.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtFromIssueDt2.Alt, "X") 	
			Exit Function
		End If	
		If Trim(frm1.txtToIssueDt2.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtToIssueDt2.Alt, "X") 	
			Exit Function
		End If	
		
		If UniConvDateToYYYYMMDD(frm1.txtFromIssueDt2.text, parent.gDateFormat, "") > UniConvDateToYYYYMMDD(frm1.txtToIssueDt2.text, parent.gDateFormat, "") Then
			IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'⊙: "Will you destory previous data"
			Exit Function
		End If
	
		If Trim(frm1.txtFiscYear.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtFiscYear.Alt, "X") 	
			Exit Function
		End If	
		If Trim(frm1.cboVatDiv2.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.cboVatDiv2.Alt, "X") 	
			Exit Function
		End If	
		var1 = UCase(Trim(frm1.txtBizAreaCD2.value))
		var2 = UniConvDateToYYYYMMDD(frm1.fpDateTime3.text, parent.gDateFormat, "") 
		var3 = UniConvDateToYYYYMMDD(frm1.fpDateTime4.text, parent.gDateFormat, "") 
	
		
		If var1 = "" Then
			var1 = "%"
		Else
		    var1 = UCase(Trim(frm1.txtBizAreaCD2.value))
		End If
		
		If frm1.chkDari.checked = true then
			var4 =  "Y"										'☆: 조회 조건 데이타 
		Else
			var4 =  "N"										'☆: 조회 조건 데이타 
		End If
		
		var5 = 	Trim(frm1.txtFiscYear.text)
		var6 =  Trim(frm1.cboVatDiv2.value)
		
		
		If frm1.Rb_WK3.checked = True Then
			' 전산매체제출집계표 
			StrEbrFile = "a6109ma4"
			var4 =  "N"
		ElseIf frm1.Rb_WK4.checked = True Then
			' 공문및 표지 출력 
			StrEbrFile = "a6109ma3"
		ElseIf frm1.Rb_WK5.checked = True Then
			' 공문및 표지 출력 
			StrEbrFile = "a6109ma5"
			var4 =  "N"
		ElseIf frm1.Rb_WK6.checked = True Then
			' 디스켓출력 
			StrEbrFile = "a6109ma6"
			var4 =  "N"
		End If
		

		StrUrl = StrUrl & "ReportBizAreaCd|" & FilterVar(var1,"","SNM")
		StrUrl = StrUrl & "|FromIssueDt|"	  & var2
		StrUrl = StrUrl & "|ToIssueDt|"	      & var3
		StrUrl = StrUrl & "|ChkDari|"	      & FilterVar(var4,"","SNM")
		StrUrl = StrUrl & "|FiscYear|"	      & var5
		StrUrl = StrUrl & "|VatDiv|"	      & FilterVar(var6,"","SNM")
		
	ElseIf gSelframeFlg = TAB3 Then 
	
		If Trim(frm1.txtBizAreaCD3.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtBizAreaCD3.Alt, "X") 	
			Exit Function
		End If	
	
		If Trim(frm1.txtFromIssueDt3.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtFromIssueDt3.Alt, "X") 	
			Exit Function
		End If	
		If Trim(frm1.txtToIssueDt3.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtToIssueDt3.Alt, "X") 	
			Exit Function
		End If	
		If Trim(frm1.txtFisc.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtFisc.Alt, "X") 	
			Exit Function
		End If	
		If Trim(frm1.cboVatDiv3.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.cboVatDiv3.Alt, "X") 	
			Exit Function
		End If	
		If Trim(frm1.txtFileName3.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtFileName3.Alt, "X") 	
			Exit Function
		End If	

		If UniConvDateToYYYYMMDD(frm1.txtFromIssueDt3.text, parent.gDateFormat, "") > UniConvDateToYYYYMMDD(frm1.txtToIssueDt3.text, parent.gDateFormat, "") Then
			IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'⊙: "Will you destory previous data"
			Exit Function
		End If
	
		If IsNumeric(frm1.txtFisc.value) = False Then
			IntRetCD = DisplayMsgBox("229924", "X", "X", "X")							'필수입력 check!!
			frm1.txtFisc.focus
			' 숫자를 입력하십시오 
			Exit Function
		End If
	
		var1 = UCase(Trim(frm1.txtBizAreaCD3.value))
		var2 = UniConvDateToYYYYMMDD(frm1.fpDateTime6.text, parent.gDateFormat, "") 
		var3 = UniConvDateToYYYYMMDD(frm1.fpDateTime7.text, parent.gDateFormat, "") 
		var4 = frm1.txtFisc.value 
		var5 = frm1.cboVatDiv3.value 
		var6 = frm1.txtFileName2.value 
	
		If var1 = "" Then
			var1 = "*"
		Else
		    var1 = UCase(Trim(frm1.txtBizAreaCD.value))
		End If

		If frm1.Rb_gb1.checked = True Then
			' 디스켓출력 
			If frm1.Rd_gbn1.checked = True Then
					if frm1.Rdo_gbn1.checked = True Then
					StrEbrFile = "a6109ma8_1"
					Else
					StrEbrFile = "a6109ma8"
					End if
			ElseIf frm1.Rd_gbn2.checked = True Then
					StrEbrFile = "a6109ma7"
			End if		
		ElseIf frm1.Rb_gb2.checked = True Then
			' 시디 
			If frm1.Rd_gbn1.checked = True Then
					if frm1.Rdo_gbn1.checked = True Then
					StrEbrFile = "a6109ma10_1"
					Else
					StrEbrFile = "a6109ma10"
					End if
			ElseIf frm1.Rd_gbn2.checked = True Then
					StrEbrFile = "a6109ma9"
			End if		
		End If
		StrUrl = StrUrl & "ReportBizAreaCd|" & FilterVar(var1,"","SNM")
		StrUrl = StrUrl & "|FromIssueDt|"	  & var2
		StrUrl = StrUrl & "|ToIssueDt|"	      & var3
		StrUrl = StrUrl & "|FiscCnt|"	      & FilterVar(var4,"","SNM")
		StrUrl = StrUrl & "|VatDiv|"	      & FilterVar(var5,"","SNM")
		StrUrl = StrUrl & "|FileName|"	      & FilterVar(var6,"","SNM")

	
	End If	
    ObjName = AskEBDocumentName(StrEbrFile,"ebr")	
	Call FncEBRPrint(EBAction,ObjName,StrUrl)	
		
End Function

Function FncBtnPreview()
	On Error Resume Next
	
	Dim Var1
	Dim Var2
	Dim Var3
	Dim Var4
	Dim Var5
	Dim Var6
	Dim strUrl
	
	Dim arrParam, arrField, arrHeader
	Dim StrEbrFile
	Dim ObjName
	
   
   '//	tab 에따라 체크할 항목이 다르기 때문에 주석으로함 
   '// If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
   '//    Exit Function
   '// End If
	If gSelframeFlg = TAB1 Then 
		If Trim(frm1.txtBizAreaCD.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtBizAreaCD.Alt, "X") 	
			Exit Function
		End If	
	
		If Trim(frm1.txtFromIssueDt.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtFromIssueDt.Alt, "X") 	
			Exit Function
		End If	
		If Trim(frm1.txtToIssueDt.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtToIssueDt.Alt, "X") 	
			Exit Function
		End If	
		If Trim(frm1.txtFiscCnt.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtFiscCnt.Alt, "X") 	
			Exit Function
		End If	
		If Trim(frm1.cboVatDiv.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.cboVatDiv.Alt, "X") 	
			Exit Function
		End If	
		If Trim(frm1.txtFileName.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtFileName.Alt, "X") 	
			Exit Function
		End If	
		
		If UniConvDateToYYYYMMDD(frm1.txtFromIssueDt.text, parent.gDateFormat, "") > UniConvDateToYYYYMMDD(frm1.txtToIssueDt.text, parent.gDateFormat, "") Then
			IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'⊙: "Will you destory previous data"
			Exit Function
		End If
	
		If IsNumeric(frm1.txtFiscCnt.value) = False Then
			IntRetCD = DisplayMsgBox("229924", "X", "X", "X")							'필수입력 check!!
			frm1.txtFiscCnt.focus
			' 숫자를 입력하십시오 
			Exit Function
		End If
	
		var1 = UCase(Trim(frm1.txtBizAreaCD.value))
		var2 = UniConvDateToYYYYMMDD(frm1.fpDateTime1.text, parent.gDateFormat, "") 
		var3 = UniConvDateToYYYYMMDD(frm1.fpDateTime2.text, parent.gDateFormat, "") 
		var4 = frm1.txtFiscCnt.value 
		var5 = frm1.cboVatDiv.value 
		var6 = frm1.txtFileName.value 
	
		If var1 = "" Then
			var1 = "*"
		Else
		    var1 = UCase(Trim(frm1.txtBizAreaCD.value))
		End If

		If frm1.Rb_WK1.checked = True Then
			' 디스켓출력 
			StrEbrFile = "a6109ma2"
		ElseIf frm1.Rb_WK2.checked = True Then
			' 공문및 표지 출력 
			StrEbrFile = "a6109ma1"
		End If
		StrUrl = StrUrl & "ReportBizAreaCd|" & FilterVar(var1,"","SNM")
		StrUrl = StrUrl & "|FromIssueDt|"	  & var2
		StrUrl = StrUrl & "|ToIssueDt|"	      & var3
		StrUrl = StrUrl & "|FiscCnt|"	      & var4
		StrUrl = StrUrl & "|VatDiv|"	      & var5
		StrUrl = StrUrl & "|FileName|"	      & FilterVar(var6,"","SNM")
		StrUrl = StrUrl & "|TmpBizAreaCd|"    & var1

	ElseIf gSelframeFlg = TAB2 Then 
	
		If Trim(frm1.txtBizAreaCD2.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtBizAreaCD2.Alt, "X") 	
			Exit Function
		End If	
	
		If Trim(frm1.txtFromIssueDt2.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtFromIssueDt2.Alt, "X") 	
			Exit Function
		End If	
		If Trim(frm1.txtToIssueDt2.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtToIssueDt2.Alt, "X") 	
			Exit Function
		End If	
		If UniConvDateToYYYYMMDD(frm1.txtFromIssueDt2.text, parent.gDateFormat, "") > UniConvDateToYYYYMMDD(frm1.txtToIssueDt2.text, parent.gDateFormat, "") Then
			IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'⊙: "Will you destory previous data"
			Exit Function
		End If
		
		If Trim(frm1.txtFiscYear.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtFiscYear.Alt, "X") 	
			Exit Function
		End If	
		
		If Trim(frm1.cboVatDiv2.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.cboVatDiv2.Alt, "X") 	
			Exit Function
		End If	
		
		var1 = UCase(Trim(frm1.txtBizAreaCD2.value))
		var2 = UniConvDateToYYYYMMDD(frm1.fpDateTime3.text, parent.gDateFormat, "") 
		var3 = UniConvDateToYYYYMMDD(frm1.fpDateTime4.text, parent.gDateFormat, "") 
		
		
		If var1 = "" Then
			var1 = "%"
		Else
		    var1 = UCase(Trim(frm1.txtBizAreaCD2.value))
		End If



		If frm1.chkDari.checked = true then
			var4 =  "Y"										'☆: 조회 조건 데이타 
		Else
			var4 =  "N"										'☆: 조회 조건 데이타 
		End If
			
		var5 = 	Trim(frm1.txtFiscYear.text)
		var6 =  Trim(frm1.cboVatDiv2.value)
		
		
		If frm1.Rb_WK3.checked = True Then
			' 전산매체제출집계표 
			StrEbrFile = "a6109ma4"
			var4 =  "N"
		ElseIf frm1.Rb_WK4.checked = True Then
			' 공문및 표지 출력 
			StrEbrFile = "a6109ma3"
		ElseIf frm1.Rb_WK5.checked = True Then
			' 일괄대리제출 
			StrEbrFile = "a6109ma5"
			var4 =  "N"
		ElseIf frm1.Rb_WK6.checked = True Then
			' 디스켓출력 
			StrEbrFile = "a6109ma6"
			var4 =  "N"
		End If

		StrUrl = StrUrl & "ReportBizAreaCd|" & var1
		StrUrl = StrUrl & "|FromIssueDt|"	  & var2
		StrUrl = StrUrl & "|ToIssueDt|"	      & var3
		StrUrl = StrUrl & "|ChkDari|"	      & var4
		StrUrl = StrUrl & "|FiscYear|"	      & var5
		StrUrl = StrUrl & "|VatDiv|"	      & var6
	
	ElseIf gSelframeFlg = TAB3 Then 
	

	
		If Trim(frm1.txtBizAreaCD3.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtBizAreaCD3.Alt, "X") 	
			Exit Function
		End If	

		If Trim(frm1.txtFromIssueDt3.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtFromIssueDt3.Alt, "X") 	
			Exit Function
		End If	
		If Trim(frm1.txtToIssueDt3.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtToIssueDt3.Alt, "X") 	
			Exit Function
		End If	
		If Trim(frm1.txtFisc.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtFisc.Alt, "X") 	
			Exit Function
		End If	
		If Trim(frm1.cboVatDiv3.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.cboVatDiv3.Alt, "X") 	
			Exit Function
		End If	
		If Trim(frm1.txtFileName3.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtFileName3.Alt, "X") 	
			Exit Function
		End If	

		If UniConvDateToYYYYMMDD(frm1.txtFromIssueDt3.text, parent.gDateFormat, "") > UniConvDateToYYYYMMDD(frm1.txtToIssueDt3.text, parent.gDateFormat, "") Then
			IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'⊙: "Will you destory previous data"
			Exit Function
		End If
	
		If IsNumeric(frm1.txtFisc.value) = False Then
			IntRetCD = DisplayMsgBox("229924", "X", "X", "X")							'필수입력 check!!
			frm1.txtFisc.focus
			' 숫자를 입력하십시오 
			Exit Function
		End If

		var1 = UCase(Trim(frm1.txtBizAreaCD3.value))
		var2 = UniConvDateToYYYYMMDD(frm1.fpDateTime6.text, parent.gDateFormat, "") 
		var3 = UniConvDateToYYYYMMDD(frm1.fpDateTime7.text, parent.gDateFormat, "") 
		var4 = frm1.txtFisc.value 
		var5 = frm1.cboVatDiv3.value 
		var6 = frm1.txtFileName3.value 
	
		If var1 = "" Then
			var1 = "*"
		Else
		    var1 = UCase(Trim(frm1.txtBizAreaCD3.value))
		End If

		If frm1.Rb_gb1.checked = True Then
			' 디스켓출력 
			If frm1.Rd_gbn1.checked = True Then
					if frm1.Rdo_gbn1.checked = True Then
					StrEbrFile = "a6109ma8_1"
					Else
					StrEbrFile = "a6109ma8"
					End if
			ElseIf frm1.Rd_gbn2.checked = True Then
					StrEbrFile = "a6109ma7"
			End if		
		ElseIf frm1.Rb_gb2.checked = True Then
			' 시디 
			If frm1.Rd_gbn1.checked = True Then
					if frm1.Rdo_gbn1.checked = True Then
					StrEbrFile = "a6109ma10_1"
					Else
					StrEbrFile = "a6109ma10"
					End if
			ElseIf frm1.Rd_gbn2.checked = True Then
					StrEbrFile = "a6109ma9"
			End if		
		End If
		StrUrl = StrUrl & "ReportBizAreaCd|" & FilterVar(var1,"","SNM")
		StrUrl = StrUrl & "|FromIssueDt|"	  & var2
		StrUrl = StrUrl & "|ToIssueDt|"	      & var3
		StrUrl = StrUrl & "|FiscCnt|"	      & FilterVar(var4,"","SNM")
		StrUrl = StrUrl & "|VatDiv|"	      & FilterVar(var5,"","SNM")
		StrUrl = StrUrl & "|FileName|"	      & FilterVar(var6,"","SNM")

	
	End If		
	
		
    ObjName = AskEBDocumentName(StrEbrFile,"ebr")	
	Call FncEBRPreview(ObjName,StrUrl)
		
End Function


'===========================================  3.1.1 Form_Load()  =========================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029															'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    Call ggoOper.FormatDate(frm1.txtFiscYear, parent.gDateFormat, 3)
    Call ggoOper.FormatNumber(frm1.txtFiscCnt, "99", "0", False)	
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("10000000000011")
	Call ClickTab1()
    '//Call SetDefaultVal  : ClickTab1안에서 호출함 
    Call InitComboBox
	
	gIsTab     = "Y" 
	gTabMaxCnt = 3     
	frm1.txtBizAreaCD.focus 
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================


Sub InitComboBox()

		Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", _
                         " MAJOR_CD = " & FilterVar("A1025", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    Call SetCombo2(frm1.cboVatDiv,lgF0  ,lgF1  ,Chr(11))
    Call SetCombo2(frm1.cboVatDiv2,lgF0  ,lgF1  ,Chr(11))
    Call SetCombo2(frm1.cboVatDiv3,lgF0  ,lgF1  ,Chr(11))
    Call SetCombo(frm1.txtFisc, "1", "1기")
	Call SetCombo(frm1.txtFisc, "2", "2기")
    
    
    
End Sub




'=======================================================================================================
'   Event Name : txtFromIssueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromIssueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromIssueDt.Action = 7
  		Call SetFocusToDocument("M")
		frm1.txtFromIssueDt.Focus
    End If
End Sub


'=======================================================================================================
'   Event Name : txtToIssueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtToIssueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToIssueDt.Action = 7
  		Call SetFocusToDocument("M")
		frm1.txtToIssueDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFromIssueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================

Sub txtFromIssueDt2_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromIssueDt2.Action = 7
  		Call SetFocusToDocument("M")
		frm1.txtFromIssueDt2.Focus
    End If
End Sub


'=======================================================================================================
'   Event Name : txtToIssueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtToIssueDt2_DblClick(Button)
    If Button = 1 Then
        frm1.txtToIssueDt2.Action = 7
  		Call SetFocusToDocument("M")
		frm1.txtToIssueDt2.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFromIssueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================

Sub txtFromIssueDt2_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromIssueDt2.Action = 7
  		Call SetFocusToDocument("M")
		frm1.txtFromIssueDt2.Focus
    End If
End Sub


'=======================================================================================================
'   Event Name : txtToIssueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtToIssueDt3_DblClick(Button)
    If Button = 1 Then
        frm1.txtToIssueDt3.Action = 7
  		Call SetFocusToDocument("M")
		frm1.txtToIssueDt3.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtDrawnDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDrawnUpDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDrawnUpDt.Action = 7
  		Call SetFocusToDocument("M")
		frm1.txtDrawnUpDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFiscYear_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFiscYear_DblClick(Button)
    If Button = 1 Then
        frm1.txtFiscYear.Action = 7
  		Call SetFocusToDocument("M")
		frm1.txtFiscYear.Focus
    End If
End Sub
'===========================================================================================================
'	Event Name :Radio3_Click
'	Event Desc : 세금계산서, 계산서구분 라디오버튼 선택시 
	
'===========================================================================================================
Sub Radio34_Click()
	If gSelFrameFlg = Tab2 and (frm1.Rb_WK4.checked = true) Then
		Call ElementVisible(frm1.chkDari,"1")
		spnDari.innerHTML = "일괄대리제출"		
	Else	
		Call ElementVisible(frm1.chkDari,"0")
		spnDari.innerHTML = ""
	
	End If	
	
	
End Sub

Sub Rd_gbn2_onClick()
    frm1.Rdo_gbn1.checked = true
	Call ggoOper.SetReqAttr(frm1.Rdo_gbn1		, "Q")
	Call ggoOper.SetReqAttr(frm1.Rdo_gbn2		, "Q")
End Sub

Sub Rd_gbn1_onClick()
	Call ggoOper.SetReqAttr(frm1.Rdo_gbn1		, "D")
	Call ggoOper.SetReqAttr(frm1.Rdo_gbn2		, "D")	
End Sub


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
	Call Parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	


<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
    //alert(value)
    //alert(escape(value))
	document.cookie = name + "=" + escape(value)
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/bin"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/lib"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
}

setCookie("client", "-1", null)
setCookie("owner", "admin", null)
setCookie("identity", "admin", null)
 -->
</SCRIPT>

</HEAD>

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><!-- ' 상위 여백 --></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>공문및표지출력(세금)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">	
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>공문및표지출력(계산서)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">	
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab3()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>공문및표지출력(신용카드)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<!--첫번째 TAB  -->
						<DIV ID="TabDiv"  SCROLL="no">
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>출력구분</TD>
									<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio1 ID=Rb_WK2 Checked> <LABEL FOR=Rb_WK2>공문</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
													<INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio1 ID=Rb_WK1 ><LABEL FOR=Rb_WK1>디스켓</LABEL></TD>
									                
								</TR>
								<TR>
									<TD CLASS="TD5">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">세금신고사업장</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizAreaCD" NAME="txtBizAreaCD" SIZE=12 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" ALT="세금신고사업장" tag="12X1XU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD.Value, 0)">&nbsp;
												    <INPUT TYPE=TEXT ID="txtBizAreaNM" NAME="txtBizAreaNM" SIZE=20 MAXLENGTH=50 STYLE="TEXT-ALIGN: Left" ALT="세금신고사업장" tag="14X" ></TD>
								</TR>
								<TR>
								 	<TD CLASS="TD5">발행일</TD>
									<TD CLASS="TD6"><script language =javascript src='./js/a6109oa1_fpDateTime1_txtFromIssueDt.js'></script>
													 &nbsp;~&nbsp;
												    <script language =javascript src='./js/a6109oa1_fpDateTime2_txtToIssueDt.js'></script></TD>
								</TR>
								<TR>
								 	<TD CLASS="TD5">회기</TD>
									<TD CLASS="TD6"><script language =javascript src='./js/a6109oa1_fpDoubleSingle1_txtFiscCnt.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">부가세구분</TD>
									<TD CLASS="TD6" COLSPAN=3><SELECT ID="cboVatDiv" NAME="cboVatDiv" ALT="부가세구분" STYLE="WIDTH: 100px" tag="12X1"></SELECT></TD>
								</TR>
								<TR>
								 	<TD CLASS="TD5">파일명</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtFileName" NAME="txtFileName" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="파일명" tag="12X1" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
								</TR>
							</TABLE>
						</div>
						<!--두번째 TAB  -->
						<DIV ID="TabDiv"  SCROLL="no">
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>출력구분</TD>
									<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio2 ID=Rb_WK4 Checked onclick="Radio34_ClicK">        <LABEL FOR=Rb_WK4>공문</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
													<INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio2 ID=Rb_WK6  onclick="Radio34_ClicK"><LABEL FOR=Rb_WK6>디스켓</LABEL></TD>
									                
								</TR>

								<TR>
									<TD CLASS="TD5">&nbsp;</TD>
									<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio2 ID=Rb_WK5 onclick="Radio34_ClicK">        <LABEL FOR=Rb_WK5>일괄대리제출</LABEL>&nbsp;&nbsp;
													<INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio2 ID=Rb_WK3  onclick="Radio34_ClicK"><LABEL FOR=Rb_WK3>전산매체제출집계표</LABEL></TD>
									
								</TR>
								<TR>
									<TD CLASS="TD5">세금신고사업장</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizAreaCD2" NAME="txtBizAreaCD2" SIZE=12 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" ALT="세금신고사업장" tag="12X1XU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD2.Value, 1)">&nbsp;
												    <INPUT TYPE=TEXT ID="txtBizAreaNM2" NAME="txtBizAreaNM2" SIZE=20 MAXLENGTH=50 STYLE="TEXT-ALIGN: Left" ALT="세금신고사업장" tag="14X" ></TD>
								</TR>
								<TR>
								 	<TD CLASS="TD5">발행일</TD>
									<TD CLASS="TD6"><script language =javascript src='./js/a6109oa1_fpDateTime3_txtFromIssueDt2.js'></script>
													 &nbsp;~&nbsp;
												    <script language =javascript src='./js/a6109oa1_fpDateTime4_txtToIssueDt2.js'></script></TD>
								</TR>
								<TR>
								 	<TD CLASS="TD5">귀속년도</TD>
									<TD CLASS="TD6"><script language =javascript src='./js/a6109oa1_fpDateTime5_txtFiscYear.js'></script></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">부가세구분</TD>
									<TD CLASS="TD6" COLSPAN=3><SELECT ID="cboVatDiv2" NAME="cboVatDiv2" ALT="부가세구분" STYLE="WIDTH: 100px" tag="12X1"></SELECT></TD>
								</TR>
							
								<TR>	
									<TD CLASS=TD5 NOWRAP><span id="spnDari">일괄대리제출</span></TD>
									<TD CLASS="TD6"><input type="checkbox" class = "check" name="chkDari" value="Y"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
								</TR>
							</TABLE>
						</DIV>
						<!--세번째 TAB  -->
						<DIV ID="TabDiv"  SCROLL="no">
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>출력구분</TD>
									<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio3 ID=Rb_gb2 Checked> <LABEL FOR=Rb_gb2>CD(Compact Disk)</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
													<INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio3 ID=Rb_gb1 ><LABEL FOR=Rb_gb1>디스켓</LABEL></TD>
									                
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>구분</TD>
									<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio4 ID=Rd_gbn2 Checked> <LABEL FOR=Rd_gbn2>공문</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
													<INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio4 ID=Rd_gbn1 ><LABEL FOR=Rd_gbn1>label양식</LABEL></TD>
									                
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>일괄제출여부</TD>
									<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio5 ID=Rdo_gbn1 Checked> <LABEL FOR=Rdo_gbn1>예</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
													<INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio5 ID=Rdo_gbn2 ><LABEL FOR=Rdo_gbn2>아니요</LABEL></TD>
									                
								</TR>
								<TR>
									<TD CLASS="TD5">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">세금신고사업장</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizAreaCD3" NAME="txtBizAreaCD3" SIZE=12 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" ALT="세금신고사업장" tag="12X1XU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD3.Value, 2)">&nbsp;
												    <INPUT TYPE=TEXT ID="txtBizAreaNM3" NAME="txtBizAreaNM3" SIZE=20 MAXLENGTH=50 STYLE="TEXT-ALIGN: Left" ALT="세금신고사업장" tag="14X" ></TD>
								</TR>
								<TR>
								 	<TD CLASS="TD5">발행일</TD>
									<TD CLASS="TD6"><script language =javascript src='./js/a6109oa1_fpDateTime6_txtFromIssueDt3.js'></script>
													 &nbsp;~&nbsp;
												    <script language =javascript src='./js/a6109oa1_fpDateTime7_txttoIssueDt3.js'></script></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">회기</TD>
									<TD CLASS="TD6" COLSPAN=3><SELECT ID="txtFisc" NAME="txtFisc" ALT="회기" STYLE="WIDTH: 100px" tag="12X1"></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">부가세구분</TD>
									<TD CLASS="TD6" COLSPAN=3><SELECT ID="cboVatDiv3" NAME="cboVatDiv3" ALT="부가세구분" STYLE="WIDTH: 100px" tag="12X1"></SELECT></TD>
								</TR>
								<TR>
								 	<TD CLASS="TD5">파일명</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtFileName3" NAME="txtFileName2" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" ALT="파일명" tag="12X1" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">&nbsp;</TD>
									<TD CLASS="TD6">&nbsp;</TD>
								</TR>
							</TABLE>
						</div>	
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" OnClick="VBScript:FncBtnPreview()" Flag = 1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN"   OnClick="VBScript:FncBtnPrint()" Flag = 1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX="-1">	
</FORM>
</BODY>
</HTML>

