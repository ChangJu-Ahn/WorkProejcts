  <%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID		    : A6102BA1
'*  4. Program Name         : 부가세신고디스켓생성
'*  5. Program Desc         : 부가세신고디스켓생성 배치
'*  6. Component List       : +
'*  7. Modified date(First) : 2000/04/20
'*  8. Modified date(Last)  : 2001/03/05
'*  9. Modifier (First)     : Jong Hwan, Kim
'* 10. Modifier (Last)      : Hee Jung, Kim
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">


<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																		'☜: indicates that All variables must be declared in advance 

'==========================================================================================================

Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
EndDate     =   "<%=GetSvrDate%>"

Call ExtractDateFrom(EndDate, parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)
StartDate   = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth-3, "01")
EndDate     = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth-1, strDay) 

Const BIZ_PGM_ID = "a6102bb1.asp"													'☆: 비지니스 로직 ASP명
Const BIZ_PGM_ID2 = "a6102bb2.asp"													'☆: 비지니스 로직 ASP명
Const BIZ_PGM_ID3 = "a6102bb3.asp"	
Const BIZ_PGM_ID4 = "a6102bb4.asp"	

Const TAB1 = 1																		'☜: Tab의 위치
Const TAB2 = 2
										 '☆: 비지니스 로직 ASP명
 '==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨
'========================================================================================================= 
Dim lgBlnFlgConChg																	'☜: Condition 변경 Flag
Dim lgBlnFlgChgValue																'☜: Variable is for Dirty flag
Dim lgIntFlgMode																	'☜: Variable is for Operation Status

'==========================================  1.2.3 Global Variable값 정의  ===============================

Dim lgMpsFirmDate, lgLlcGivenDt														'☜: 비지니스 로직 ASP에서 참조하므로 

Dim  lgCurName()																	'☆ : 개별 화면당 필요한 로칼 전역 변수 
Dim  cboOldVal          
Dim  IsOpenPop          
Dim  gSelframeFlg

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE												'⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False														'⊙: Indicates that no value changed

    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False																'☆: 사용자 변수 초기화
    lgMpsFirmDate=""
    lgLlcGivenDt=""
    
End Sub

'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtIssueDT1.Text = StartDate
	frm1.txtIssueDT2.Text = EndDate
	frm1.txtReportDt.Text = EndDate
	frm1.txtYear.text	= strYear
	'frm1.txtBizAreaCD.value	= gBizArea
	frm1.txtBizAreaNM.value	= ""

	frm1.txtIssueDT3.Text = StartDate
	frm1.txtIssueDT4.Text = EndDate
		
	frm1.txtIssueDT5.Text = StartDate
	frm1.txtIssueDT6.Text = EndDate
		
	'frm1.txtBizAreaCD2.value	= gBizArea
	frm1.txtBizAreaNM2.value	= ""
End Sub

 '==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	'//파일구분
	Call SetCombo(frm1.cbofileGubun, "A", "발행일분")
    Call SetCombo(frm1.cbofileGubun, "B", "누락분")
	Call SetCombo(frm1.cbofileGubun, "C", "발행일분+누락분")
    
	'//기구분
	Call SetCombo(frm1.cboGiGubun, "1", "1기")
	Call SetCombo(frm1.cboGiGubun, "2", "2기")
	'//Call SetCombo(frm1.cboGiGubun, "3", "전체")
	
	'//신고구분
	Call SetCombo(frm1.cboSingoGubun, "1", "예정")
	Call SetCombo(frm1.cboSingoGubun, "2", "확정")
End Sub

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
		Case 0,1
			arrParam(0) = "세금신고사업장 팝업"					' 팝업 명칭
			arrParam(1) = "B_TAX_BIZ_AREA"	 						' TABLE 명칭
			arrParam(2) = strCode									' Code Condition
			arrParam(3) = ""										' Name Cindition
			arrParam(4) = ""										' Where Condition
			arrParam(5) = "세금신고사업장코드"					' 조건필드의 라벨 명칭

			arrField(0) = "TAX_BIZ_AREA_CD"							' Field명(0)
			arrField(1) = "TAX_BIZ_AREA_NM"							' Field명(0)
    
			arrHeader(0) = "세금신고사업장코드"					' Header명(0)
			arrHeader(1) = "세금신고사업장명"					' Header명(0)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0		' 사업장
				frm1.txtBizAreaCD.focus
			Case 1		' 사업장
				frm1.txtBizAreaCD2.focus
		End Select
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

'------------------------------------------  SetPopUp()  -------------------------------------------------
'	Name : SetPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		' 사업장
				.txtBizAreaCD.focus
				.txtBizAreaCD.value = UCase(Trim(arrRet(0)))
				.txtBizAreaNM.value = arrRet(1)
			Case 1		' 사업장
				.txtBizAreaCD2.focus
				.txtBizAreaCD2.value = UCase(Trim(arrRet(0)))
				.txtBizAreaNM2.value = arrRet(1)	
		End Select
	End With
End Function

'======================================================================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function ClickTab1()
	If lgIntFlgMode <> parent.OPMD_UMODE Then
	    Call SetToolbar("1000000000001111")										'⊙: 버튼 툴바 제어
	Else                 
	    Call SetToolbar("1000000000001111")										'⊙: 버튼 툴바 제어
	End If
	
	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)	 '~~~ 첫번째 Tab 
	gSelframeFlg = TAB1
	Call SetDefaultVal()

	frm1.txtBizAreaCD.focus
End Function

Function ClickTab2()

	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)	 '~~~ 두번째 Tab 
	gSelframeFlg = TAB2
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetToolbar("1000000000001111")
	ELSE                 
		Call SetToolbar("1000000000001111")
	END IF	
	Call SetDefaultVal()
	frm1.txtBizAreaCD2.focus

End Function

'========================================================================================================= 
Sub Form_Load()

    Call InitVariables							'⊙: Initializes local global variables
    Call LoadInfTB19029							'⊙: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")		'⊙: Lock  Suitable  Field
    Call ggoOper.FormatDate(frm1.txtYear, parent.gDateFormat, 3)
    '----------  Coding part  -------------------------------------------------------------
    Call ClickTab1()
   '// Call SetDefaultVal : ClickTab1안에서 호출함
	Call InitComboBox()
	Call Radio3_Click
    Call SetToolbar("1000000000001111")										'⊙: 버튼 툴바 제어
   
    gIsTab     = "Y" 
	gTabMaxCnt = 2     
	'//msgbox "본화면은 현재 테스트중입니다." & vbcrlf & "-- 이남요"
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'=======================================================================================================
'   Event Name : txtIssueDt1_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt1_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt1.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIssueDt1.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt1_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt1_Change()
    'lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt2_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt2_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt2.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIssueDt2.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt2_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt2_Change()
	frm1.cbofileGubun.value = ""
	call cbofileGubun_onChange()
    'lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtReportDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtReportDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReportDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtReportDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtReportDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYear_DblClick(Button)
    If Button = 1 Then
        frm1.txtYear.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtYear.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtReportDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtReportDt_Change()
    'lgBlnFlgChgValue = True
End Sub
'=======================================================================================================
'   Event Name : txtIssueDt1_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt3_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt3.Action = 7
 		Call SetFocusToDocument("M")
		frm1.txtIssueDt3.Focus
   End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt1_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt3_Change()
    'lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt2_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt4_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt4.Action = 7
 		Call SetFocusToDocument("M")
		frm1.txtIssueDt4.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt2_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt4_Change()
    'lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt5_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt5_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt5.Action = 7
 		Call SetFocusToDocument("M")
		frm1.txtIssueDt5.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt6_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt6_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt6.Action = 7
 		Call SetFocusToDocument("M")
		frm1.txtIssueDt6.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : cbofileGubun_OnChange()
'   Event Desc : 파일구분선택시 누락분계산서발행일을 선택한다.
'=======================================================================================================
Sub cbofileGubun_onChange()
	Select Case Trim(frm1.cbofileGubun.value)
		Case "C"	
   		    call  ExtractDateFrom(frm1.txtIssueDT2.Text,  parent.gDateFormatYYYYMM, parent.gComDateType,strYear, strMonth, strDay)   	                

		    IF  strMonth = "06" or strMonth = "12" Then
				Call ElementVisible(frm1.txtIssueDT5, "1")
		    	call ElementVisible(frm1.txtissueDT6, "1")	            
				spndate.innerHTML = "누락분발행일"
		        spnSign.innerHTML = "~"		
		    Else
				DisplayMsgBox "115116","X" , frm1.txtFileName.Alt, frm1.txtIssueDt2.Alt
				frm1.cbofileGubun.value=""
			End If
		Case Else		
			Call ElementVisible(frm1.txtIssueDT5, "0")
			call ElementVisible(frm1.txtissueDT6, "0")	            
			spndate.innerHTML = ""
		    spnSign.innerHTML = ""	
	End Select
End Sub

'===========================================================================================================
'	Event Name :Radio3_Click
'	Event Desc : 세금계산서, 계산서구분 라디오버튼 선택시
'===========================================================================================================
Sub Radio3_Click()
	frm1.txtFileName.value = ""
	If gSelFrameFlg = Tab1 Then
		If frm1.Rb_TA1.checked = True Then
			Call ElementVisible(frm1.txtIssueDT5, "0")
			call ElementVisible(frm1.txtissueDT6,"0")
			Call ElementVisible(frm1.txtYear,"0")
			Call ElementVisible(frm1.cboGiGubun,"0")
			Call ElementVisible(frm1.cboSingoGubun,"0")
			Call ElementVisible(frm1.chkDari,"0")
			
			spndate.innerHTML = ""
			spnSign.innerHTML = ""
			spnYear.innerHTML = ""
			spnGiGubun.innerHTML = ""
			spnSingoGubun.innerHTML = ""
			spnDari.innerHTML = ""
			'frm1.txtFileName.className = "Required"
			frm1.txtFileName.className = "protected"
			frm1.txtFileName.readonly = false
			frm1.cbofileGubun.className="Required"
			Call ggoOper.SetReqAttr(frm1.cbofileGubun, "N")
		ElseIf frm1.Rb_TA2.checked = True Then
			Call ElementVisible(frm1.txtIssueDT5, "0")
			call ElementVisible(frm1.txtissueDT6,"0")
			Call ElementVisible(frm1.txtYear,"1")
			Call ElementVisible(frm1.cboGiGubun,"1")
			Call ElementVisible(frm1.cboSingoGubun,"1")
			Call ElementVisible(frm1.chkDari,"1")
			frm1.txtFileName.className = "protected"
			frm1.txtFileName.readonly = True
			frm1.cbofileGubun.value = "A"
			frm1.cbofileGubun.className = "protected"
			Call ggoOper.SetReqAttr(frm1.cbofileGubun, "Q")

			spndate.innerHTML = ""
			spnSign.innerHTML = ""
			spnYear.innerHTML = "귀속년도"
			spnGiGubun.innerHTML = "기구분"
			spnSingoGubun.innerHTML = "신고구분"
			spnDari.innerHTML = "일괄대리제출"
		Else
			Call ElementVisible(frm1.txtIssueDT5, "0")
			call ElementVisible(frm1.txtissueDT6,"0")
			Call ElementVisible(frm1.txtYear,"1")
			Call ElementVisible(frm1.cboGiGubun,"1")
			Call ElementVisible(frm1.cboSingoGubun,"1")
			Call ElementVisible(frm1.chkDari,"1")
			frm1.txtFileName.className = "protected"
			frm1.txtFileName.readonly = True			
			
			spndate.innerHTML = ""
			spnSign.innerHTML = ""
			spnYear.innerHTML = "귀속년도"
			spnGiGubun.innerHTML = "기구분"
			spnSingoGubun.innerHTML = "신고구분"
			spnDari.innerHTML = "일괄대리제출"

			frm1.cbofileGubun.className="Required"
			Call ggoOper.SetReqAttr(frm1.cbofileGubun, "N")
		End If
	End If	
	frm1.txtBizAreaCD.focus
	
End Sub

 '#########################################################################################################
'												4. Common Function부
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수
'######################################################################################################### 
Function subVatDisk() 
	Dim RetFlag
	Dim strVal
	Dim intRetCD
	Dim intI, strFileName, intChrChk	'특수문자 Check
	Dim strYear1,strMonth1, strDay1, strDate1
	Dim strYear2,strMonth2, strDay2	, strDate2
	Dim strMsg
	
    '-----------------------
    'Check content area
    '-----------------------
	If gSelFrameFlg = Tab1 Then	
	'화일명으로 사용할 수 없는 특수문자 \/:*?"<>|&. 포함여부 확인
		'2006.11.01
		'자동으로발생토록함.lee wol san
		
		'If frm1.Rb_TA1.checked = True Then		
		'	strFileName = frm1.txtFileName.value
		'	For intI = 1 To Len(strFileName)
		''		intChrChk = ASC(Mid(strFileName, intI, 1))
		'		If intChrChk = ASC("\") Or intChrChk = ASC("/") Or intChrChk = ASC(":") Or intChrChk = ASC("*") Or _
		'			intChrChk = ASC("?") Or intChrChk = 34 Or intChrChk = ASC("<") Or intChrChk = ASC(">") Or _
		'			intChrChk = ASC("|") OR intChrChk = ASC("&") OR intChrChk = ASC(".") Then
		'				intRetCD =  DisplayMsgBox("970029","X" , frm1.txtFileName.Alt, frm1.txtIssueDt2.Alt)
		'				Exit Function
		'		End If
		'	Next
		' End IF	
		 
		 
		'//아래의 코드를 주석으로 막아놓은 이유는 탭에따라 체크해야할 항목이 다르기때문에 막음
		' Required로 표시된 Element들의 입력 [유/무]를 Check 한다.
		  ' ChkField(pDoc, pStrGrp) As Boolean
		'If Not chkField(Document, "1") Then        '⊙: Check contents area
		'  Exit Function
		'End If
		
		'*************************************************************************
		'//필수항목 체크 : 탭에따라 체크해야할 항목이 다르기때문에 막음
		'*************************************************************************
		If Trim(frm1.txtIssueDt1.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtIssueDt1.Alt, "X") 	
			Exit Function
		End If
		If Trim(frm1.txtIssueDt2.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtIssueDt2.Alt, "X") 	
			Exit Function
		End If
		
		If Trim(frm1.txtBizAreaCD.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtBizAreaCD.Alt, "X") 	
			Exit Function
		End If

		If Trim(frm1.txtReportDt.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtReportDt.Alt, "X") 	
			Exit Function
		End If
		
		If frm1.Rb_TA1.checked = True Then		
			If Trim(frm1.txtFileName.value) = "" Then
				'RetFlag = DisplayMsgBox("970029","X" , frm1.txtFileName.Alt, "X") 	
				'Exit Function
			End If
		ElseIf frm1.Rb_TA2.checked = True Then		
		    frm1.txtFileName.value = ""
			If Trim(frm1.txtYear.text) = "" Then
				RetFlag = DisplayMsgBox("970029","X" , frm1.txtYear.Alt, "X") 		'귀속년도을 확인하세요
				Exit Function
			End If
		
			If Trim(frm1.cboGiGubun.value) = "" Then			
				RetFlag = DisplayMsgBox("970029","X" , frm1.cboGiGubun.Alt, "X") 	'기구분을 확인하세요
				Exit Function
			End If
			If Trim(frm1.cboSingoGubun.value) = "" Then								'신고 구분을 선택하지 않았을경우
				RetFlag = DisplayMsgBox("970029","X" , frm1.cboSingoGubun.Alt, "X")
				Exit Function
			End If
		End If	

		If frm1.cbofileGubun.value = "C" Then
		    If Trim(frm1.txtIssueDt5.text) = "" Then
			    RetFlag = DisplayMsgBox("970029","X" , frm1.txtIssueDt5.Alt, "X") 	
			    Exit Function
		    End If
		    If Trim(frm1.txtIssueDt6.text) = "" Then
			    RetFlag = DisplayMsgBox("970029","X" , frm1.txtIssueDt6.Alt, "X") 	
			    Exit Function
		    End If
		Else
		    frm1.txtIssueDt5.text = ""
		    frm1.txtIssueDt6.text = ""
		End If
		
		If CompareDateByFormat(frm1.txtIssueDt1.text,frm1.txtIssueDt2.text,frm1.txtIssueDt1.Alt,frm1.txtIssueDt2.Alt, _
	     	               "970025",frm1.txtIssueDt1.UserDefinedFormat,parent.gComDateType, True) = False Then
		   frm1.txtIssueDt1.focus
		   Exit Function
		End If

		If CompareDateByFormat(frm1.txtIssueDt5.text,frm1.txtIssueDt6.text,frm1.txtIssueDt5.Alt,frm1.txtIssueDt6.Alt, _
	     	               "970025",frm1.txtIssueDt5.UserDefinedFormat,parent.gComDateType, True) = False Then
		   frm1.txtIssueDt1.focus
		   Exit Function
		End If

		'발행일 시작일자는 누락일 종료일자보다 반드시 이전일자 이어야 함. (2005-04-14 JYK)		
		If CompareDateByFormat(frm1.txtIssueDt6.text,frm1.txtIssueDt1.text,frm1.txtIssueDt6.Alt,frm1.txtIssueDt1.Alt, _
	     	               "970024",frm1.txtIssueDt5.UserDefinedFormat,parent.gComDateType, True) = False Then
		   frm1.txtIssueDt1.focus
		   Exit Function
		End If		
	ElseIf gSelFrameFlg = Tab2 Then	 '//취소탭
		If Trim(frm1.txtIssueDt3.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtIssueDt3.Alt, "X") 	
			Exit Function
		End If
		If Trim(frm1.txtIssueDt4.text) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtIssueDt4.Alt, "X") 	
			Exit Function
		End If
		If Trim(frm1.txtBizAreaCD2.value) = "" Then
			RetFlag = DisplayMsgBox("970029","X" , frm1.txtBizAreaCD2.Alt, "X") 	
			Exit Function
		End If
		
		If CompareDateByFormat(frm1.txtIssueDt3.text,frm1.txtIssueDt4.text,frm1.txtIssueDt3.Alt,frm1.txtIssueDt4.Alt, _
	     	               "970025",frm1.txtIssueDt3.UserDefinedFormat,parent.gComDateType, True) = False Then
		   frm1.txtIssueDt3.focus
		   Exit Function
		End If
	Else
		Exit Function
	End If
		
	RetFlag = DisplayMsgBox("900018", parent.VB_YES_NO,"x","x")   '☜ 바뀐부분
	'RetFlag = Msgbox("작업을 수행 하시겠습니까?", vbOKOnly + vbInformation, "정보")
	If RetFlag = VBNO Then
		Exit Function
	End IF

    Err.Clear                                                               '☜: Protect system from crashing

	Call LayerShowHide(1)
    dim chkYn 
    

    
    With frm1
    
	if frm1.chkYN(0).checked then 
		chkYn="N"
    else
		chkYn="Y"
    end if
    
    
		If gSelFrameFlg = Tab1 Then
			If .Rb_TA1.checked = True Then
				strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001					'☜: 비지니스 처리 ASP의 상태 
				strVal = strVal & "&txtIssueDt1=" & Trim(.txtIssueDt1.text)				'☆: 처리 조건 데이타
				strVal = strVal & "&txtIssueDt2=" & Trim(.txtIssueDt2.text)				'☆: 처리 조건 데이타
				strVal = strVal & "&txtBizAreaCD=" & UCase(Trim(.txtBizAreaCD.value))	'☆: 처리 조건 데이타
				strVal = strVal & "&txtReportDt=" & Trim(.txtReportDt.text)				'☆: 처리 조건 데이타
				strVal = strVal & "&txtFileName=" & Trim(.txtFileName.value)			'☆: 처리 조건 데이타
				strVal = strVal & "&cbofileGubun=" & Trim(.cbofileGubun.value)          '☆: 정상,누락분포함,누락분만
				strVal = strVal & "&txtIssueDT5=" & Trim(.txtIssueDT5.text)             '☆: 누락분의계산서발행일From
				strVal = strVal & "&txtIssueDT6=" & Trim(.txtIssueDT6.text)		    	'☆: 누락분의계산서발행일To	
			ElseIf .Rb_TA2.checked = True Then
				strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0001					'☜: 비지니스 처리 ASP의 상태 
				strVal = strVal & "&txtIssueDt1=" & Trim(.txtIssueDt1.text)				'☆: 처리 조건 데이타
				strVal = strVal & "&txtIssueDt2=" & Trim(.txtIssueDt2.text)				'☆: 처리 조건 데이타
				strVal = strVal & "&txtBizAreaCD=" & UCase(Trim(.txtBizAreaCD.value))	'☆: 처리 조건 데이타
				strVal = strVal & "&txtReportDt=" & Trim(.txtReportDt.text)				'☆: 처리 조건 데이타
				strVal = strVal & "&txtFileName=" & Trim(.txtFileName.value)			'☆: 처리 조건 데이타
				strVal = strVal & "&txtYear=" & Trim(.txtYear.text)						'☆: 처리 조건 데이타
				strVal = strVal & "&cboGiGubun=" & Trim(.cboGiGubun.value)				'☆: 처리 조건 데이타
				strVal = strVal & "&cboSingoGubun=" & Trim(.cboSingoGubun.value)		'☆: 처리 조건 데이타
				If .chkDari.checked = True Then
					strVal = strVal & "&chkDaeri=" & "Y"								'☆: 처리 조건 데이타
				Else
					strVal = strVal & "&chkDaeri=" & "N"								'☆: 처리 조건 데이타
				End If
				strVal = strVal & "&rdoGubun=" & "1"									'☆: 처리 조건 데이타
			ElseIf 	.Rb_TA7.checked = True Then
				strVal = BIZ_PGM_ID4 & "?txtMode=" & parent.UID_M0001					'☜: 비지니스 처리 ASP의 상태 
				strVal = strVal & "&txtIssueDt1=" & Trim(.txtIssueDt1.text)				'☆: 처리 조건 데이타
				strVal = strVal & "&txtIssueDt2=" & Trim(.txtIssueDt2.text)				'☆: 처리 조건 데이타
				strVal = strVal & "&txtBizAreaCD=" & UCase(Trim(.txtBizAreaCD.value))	'☆: 처리 조건 데이타
				strVal = strVal & "&txtReportDt=" & Trim(.txtReportDt.text)				'☆: 처리 조건 데이타
				strVal = strVal & "&txtFileName=" & Trim(.txtFileName.value)			'☆: 처리 조건 데이타
				strVal = strVal & "&txtYear=" & Trim(.txtYear.text)						'☆: 처리 조건 데이타
				strVal = strVal & "&cboGiGubun=" & Trim(.cboGiGubun.value)				'☆: 처리 조건 데이타
				strVal = strVal & "&cboSingoGubun=" & Trim(.cboSingoGubun.value)		'☆: 처리 조건 데이타								
				strVal = strVal & "&cbofileGubun=" & Trim(.cbofileGubun.value)          '☆: 정상,누락분포함,누락분만
				strVal = strVal & "&txtIssueDT5=" & Trim(.txtIssueDT5.text)             '☆: 누락분의계산서발행일From
				strVal = strVal & "&txtIssueDT6=" & Trim(.txtIssueDT6.text)		    	'☆: 누락분의계산서발행일To
				If .chkDari.checked = True Then
					strVal = strVal & "&chkDaeri=" & "Y"								'☆: 처리 조건 데이타
				Else
					strVal = strVal & "&chkDaeri=" & "N"								'☆: 처리 조건 데이타
				End If										
			End If
		Else
				strVal = BIZ_PGM_ID3 & "?txtMode=" & parent.UID_M0001					'☜: 비지니스 처리 ASP의 상태 
				strVal = strVal & "&txtIssueDt3=" & Trim(.txtIssueDt3.text)				'☆: 처리 조건 데이타
				strVal = strVal & "&txtIssueDt4=" & Trim(.txtIssueDt4.text)				'☆: 처리 조건 데이타
				strVal = strVal & "&txtBizAreaCD2=" & UCase(Trim(.txtBizAreaCD2.value))	'☆: 처리 조건 데이타
				If frm1.Rb_TA3.checked = True Then
					strVal = strVal & "&rdoGubun=" & "3"								'☆: 처리 조건 데이타
				ElseIf frm1.Rb_TA4.checked = True Then
					strVal = strVal & "&rdoGubun=" & "4"								'☆: 처리 조건 데이타
				ElseIf 	frm1.Rb_TA8.checked = True Then
					strVal = strVal & "&rdoGubun=" & "6"								'☆: 처리 조건 데이타				
				End If	
				If frm1.Rb_TA5.checked = True Then
					strVal = strVal & "&rdofileGubun=" & "A"							'☆: 정상
				ElseIf frm1.Rb_TA6.checked = True Then
					strVal = strVal & "&rdofileGubun=" & "B"							'☆: 누락분
				End If	
		End If	

		strVal = strVal & "&chkYn=" & chkYn

		Call RunMyBizASP(MyBizASP, strVal)												'☜: 비지니스 ASP 를 가동
	End With
    
End Function

Function subVatDiskOK(ByVal pFileName) 
	Dim strVal
    Err.Clear																			'☜: Protect system from crashing

	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0002								'☜: 비지니스 처리 ASP의 상태 
	strVal = strVal & "&txtFileName=" & pFileName										'☆: 조회 조건 데이타
	Call RunMyBizASP(MyBizASP, strVal)													'☜: 비지니스 ASP 를 가동
End Function

Function subVatDiskOK2(ByVal strVal) 
    Err.Clear
	On Error Resume Next
	Dim IntRetCD

	If strVal = "OK" Then
		IntRetCD = DisplayMsgBox("183114", "X", "X", "X")
	End If
End Function

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 

End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow() 
     On Error Resume Next                                                   '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
End Function


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
    Call parent.FncFind(parent.C_SINGLE, True)                                         '☜:화면 유형, Tab 유무 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

 '*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 
End Function


'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행
'========================================================================================

Function DbDeleteOk()														'☆: 삭제 성공후 실행 로직
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김
'========================================================================================

Function DbQueryOk()							'☆: 조회 성공후 실행로직
End Function


'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨
'========================================================================================

Function DbSave() 
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김
'========================================================================================

Function DbSaveOk()			'☆: 저장 성공후 실행 로직
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="no">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>부가세디스켓생성</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">	
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>부가세디스켓취소</font></td>
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
									<TD CLASS=TD5>&nbsp;</TD>
									<TD CLASS=TD6>&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>세금계산서구분</TD>
									<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio3 ID=Rb_TA1 Checked onclick="Radio3_Click()" value="0"><LABEL FOR=Rb_TA1>세금계산서</LABEL>&nbsp;&nbsp;
									                <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio3 ID=Rb_TA2 onclick="Radio3_Click()" value="1"><LABEL FOR=Rb_TA2>계산서</LABEL>&nbsp;&nbsp;
									                <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio3 ID=Rb_TA7 onclick="Radio3_Click()" value="2"><LABEL FOR=Rb_TA7>신용카드</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>계산서발행일</TD>
									<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssueDt1 CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="계산서발행일(From)" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>
												  &nbsp; ~ &nbsp;
												  <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssueDt2 CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="계산서발행일(To)" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								</TR>
								
						
								<TR>
									<TD CLASS=TD5 NOWRAP>통합과세구분</TD>
									<TD CLASS=TD6>
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="chkYN" TAG="1X" VALUE="N" CHECKED ID="chkYN0"><LABEL FOR="chkYN0">사업장별</LABEL>&nbsp;
				        	        <INPUT TYPE="RADIO" CLASS="RADIO" NAME="chkYN" TAG="1X" VALUE="Y"  ID="chkYN1"><LABEL FOR="chkYN1">통합</LABEL>
				
									 </TD>
								</TR>
								
								<TR>
									<TD CLASS=TD5 NOWRAP>세금신고사업장</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT ID="txtBizAreaCD" NAME="txtBizAreaCD" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="12XXXU" ALT="세금신고사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD.Value, 0)">&nbsp;<INPUT TYPE=TEXT ID="txtBizAreaNM" NAME="txtBizAreaNM" SIZE=30 MAXLENGTH=50 STYLE="TEXT-ALIGN: left" tag="14X" ALT="세금신고사업장"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>신고일자</TD>
									<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtReportDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="신고일자" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>화일명</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT ID="txtFileName" NAME="txtFileName" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="11X" ALT="화일명"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>파일구분</TD>
									<TD CLASS="TD6"><SELECT ID="cbofileGubun" NAME="cbofileGubun" ALT="파일구분" STYLE="WIDTH: 130px" tag="12X"><OPTION VALUE="" selected></OPTION></SELECT></TD>									
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP><span id="spnDate">누락분발행일</span></TD>
									<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssueDt5 CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="누락분발행일(From)" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>
												  &nbsp; <span id="spnSign">~</span> &nbsp;
												  <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssueDt6 CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="누락분발행일(To)" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP><span id="spnYear">귀속년도</span></TD>
									<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtYear CLASS=FPDTYYYY title=FPDATETIME ALT="귀속년도" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP><span id="spnGiGubun">기구분</span></TD>
									<TD CLASS="TD6"><SELECT ID="cboGiGubun" NAME="cboGiGubun" ALT="기구분" STYLE="WIDTH: 98px" tag="12X"><OPTION VALUE="" selected></OPTION></SELECT></TD>
								</TR>	
								<TR>
									<TD CLASS=TD5 NOWRAP><span id="spnSingoGubun">신고구분</span></TD>
									<TD CLASS="TD6"><SELECT ID="cboSingoGubun" NAME="cboSingoGubun" ALT="신고구분" STYLE="WIDTH: 98px" tag="12X"><OPTION VALUE="" selected></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP><span id="spnDari">일괄대리제출</span></TD>
									<TD CLASS="TD6"><input type="checkbox" class = "check" name="chkDari" value="Y"></TD>
								</TR>																		
								<TR>
									<TD CLASS=TD5></TD>
									<TD CLASS=TD6>&nbsp;</TD>
								</TR>
							</TABLE>
						</div>
						<!--두번째 TAB  -->
						<DIV ID="TabDiv"  SCROLL="no">
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5>&nbsp;</TD>
									<TD CLASS=TD6>&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>세금계산서구분</TD>
									<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio4 ID=Rb_TA3 Checked  value="0"><LABEL FOR=Rb_TA3>세금계산서</LABEL>&nbsp;&nbsp;
									                <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio4 ID=Rb_TA4  value="1"><LABEL FOR=Rb_TA4>계산서</LABEL>&nbsp;&nbsp;
									                <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio4 ID=Rb_TA8  value="2"><LABEL FOR=Rb_TA8>신용카드</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>파일구분</TD>
									<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio5 ID=Rb_TA5 Checked  value="A"><LABEL FOR=Rb_TA5>정상</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									                <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio5 ID=Rb_TA6  value="B"><LABEL FOR=Rb_TA6>누락분</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>계산서발행일</TD>
									<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssueDt3 CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="계산서발행일(From)" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>
												  &nbsp; ~ &nbsp;
												  <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssueDt4 CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="계산서발행일(To)" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>세금신고사업장</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT ID="txtBizAreaCD2" NAME="txtBizAreaCD2" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="121XXU" ALT="세금신고사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD2.Value, 1)">&nbsp;<INPUT TYPE=TEXT ID="txtBizAreaNM2" NAME="txtBizAreaNM2" SIZE=30 MAXLENGTH=50 STYLE="TEXT-ALIGN: left" tag="14X" ALT="세금신고사업장"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5>&nbsp;</TD>
									<TD CLASS=TD6>&nbsp;</TD>
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
					<TD><BUTTON NAME="btnExecute" CLASS="CLSMBTN" OnClick="VBScript:Call subVatDisk()" Flag=1>실 행</BUTTON>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</FORM>
</BODY>
</HTML>

