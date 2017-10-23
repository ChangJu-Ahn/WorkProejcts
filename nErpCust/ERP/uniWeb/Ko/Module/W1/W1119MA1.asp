<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 각 과목별 조정 
'*  3. Program ID           : W1119MA1
'*  4. Program Name         : W1119MA1.asp
'*  5. Program Desc         : 제3호 
'*  6. Modified date(First) : 2005/01/07
'*  7. Modified date(Last)  : 2005/
'*  8. Modifier (First)     : LSHSAT
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'*  미적용 내용 --  일반법인 : 표준손익계산서(갑)  당기순이익(순손실) <>  06 or 35	오류	WC0025	손익계산서의 당기순이익(순손실) 금액과 이익잉여금(결손금)처리 계산서의 당기순이익(순손실) 금액이 일치하지 않습니다.
'					금융법인 : 표준손익계산서(을)  당기순이익(순손실) <>  06 or 35	오류	WC0025	손익계산서의 당기순이익(순손실) 금액과 이익잉여금(결손금)처리 계산서의 당기순이익(순손실) 금액이 일치하지 않습니 
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%
'=========================  로긴중인 유저의 법인코드를 출력하기 위해  ======================
    Call LoadBasisGlobalInf()
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID		= "W1119MA1"
Const BIZ_PGM_ID		= "W1119MB1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_REF_PGM_ID	= "W1119MB2.asp"
Const EBR_RPT_ID		= "W1119OA1"


Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

'============================================  초기화 함수  ====================================

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgRefMode = False
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub



'============================================  신고구분 콤보 박스 채우기  ====================================

Sub InitComboBox()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
    
End Sub



'============================== 레퍼런스 함수  ========================================
'사용자가 1,이익잉여금처분계산서와 2결손금처리계산서를 			
'선택하지 아니하고 금액불러오기를 실행하는 경우	오류(W10001 : 계산서 종류을 선택하지 아니하였습니다. 먼저 계산서 종류을 선택하여 주십시요)
'
'02		대차대조표의 이월이익잉여금(또는 이월결손금)의 기초잔액을 입력함.		
'06		표준손익계산서의 당기순이익(순손실)을 입력함.		
'
'31		대차대조표의 이월이익잉여금(또는 이월결손금)의 기초잔액을 입력함.		
'35		표준손익계산서의 당기순이익(순손실) × (-1)을 입력함.		

Function GetRef()	' 금액불러오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	Dim sMesg
	DIm BackColor_w,BackColor_g

	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
    If NOT Frm1.chkW_TYPE1.checked And NOT Frm1.chkW_TYPE2.checked Then
        Call DisplayMsgBox("W10003", "X", "X", "X")
	    Frm1.chkW_TYPE1.focus
	    Exit Function
    End If

	sMesg = wgRefDoc & vbCrLf & vbCrLf
	
	' 화면상에 가져올 데이타의 색깔을 표시한다.
	If Frm1.chkW_TYPE1.checked Then
		BackColor_g = frm1.txtW1.BackColor
		BackColor_w = frm1.txtW2.BackColor
		Frm1.txtW2.BackColor = &H009BF0A2&
		Frm1.txtW6.BackColor = &H009BF0A2&
	Else
		BackColor_g = frm1.txtW35.BackColor
		BackColor_w = frm1.txtW31.BackColor
		Frm1.txtW31.BackColor = &H009BF0A2&
		Frm1.txtW35.BackColor = &H009BF0A2&
	End If
	
    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"

	If Frm1.chkW_TYPE1.checked Then
		Frm1.txtW2.BackColor = BackColor_w
		Frm1.txtW6.BackColor = BackColor_g
	Else
		Frm1.txtW31.BackColor = BackColor_w
		Frm1.txtW35.BackColor = BackColor_g
	End If

	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_REF_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtCO_CD="      	 & Frm1.txtCO_CD.Value	      '☜: Query Key        
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key           
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  
    	
End Function

Sub GetRefOK()
	If Frm1.chkW_TYPE1.checked Then
		Frm1.txtW2.Value = Frm1.txtRW1.Value
		Frm1.txtW6.Value = Frm1.txtRW2.Value
	Else
		Frm1.txtW31.Value = Frm1.txtRW1.Value
		Frm1.txtW35.Value = unicdbl(Frm1.txtRW2.Value) * -1
	End If
	Call SetAllTxtChkCalc
End Sub


'============================================  조회조건 함수  ====================================
Sub CheckFISC_DATE()	' 요청법인의 조회조건에 만족하는 당기시작,종료일을 가져온다.
	Dim sFiscYear, sRepType, sCoCd, sFISC_START_DT, sFISC_END_DT, datMonCnt, i, datNow
	
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If lgF0 <> "" Then 
		sFISC_START_DT = CDate(lgF0)
	Else
		sFISC_START_DT = ""
	End if

    If lgF1 <> "" Then 
		sFISC_END_DT = CDate(lgF1)
	Else
		sFISC_END_DT = ""
	End if
	
End Sub

'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100111100101011")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
    Call AppendNumberRange("0","","")
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call InitData

    call fncquery()
    
End Sub


'============================================  이벤트 함수  ====================================

<%
'==========================================================================================
'   Event Name : chkW_TYPE ...
'   Event Desc : 체크박스 Value Change
'==========================================================================================
%>
Sub chkW_TYPE1_OnClick()
	If Frm1.chkW_TYPE1.checked Then
		Call SetW_TYPE(1)
	Else
		Frm1.chkW_TYPE1.checked = True
	End If
End Sub

Sub chkW_TYPE2_OnClick()	
	If Frm1.chkW_TYPE2.checked Then
		Call SetW_TYPE(2)
	Else
		Frm1.chkW_TYPE2.checked = True
	End If
End Sub

Sub SetW_TYPE(ByVal chkNum)
	If chkNum = 1 Then
		With Frm1
			.chkW_TYPE1.checked = True
			.chkW_TYPE2.checked = False
			.txtW30.Value = 0 :	call ggoOper.SetReqAttr(.txtW30, "Q")
			.txtW31.Value = 0 :	call ggoOper.SetReqAttr(.txtW31, "Q")
			.txtW32.Value = 0 :	call ggoOper.SetReqAttr(.txtW32, "Q")
			.txtW33.Value = 0 :	call ggoOper.SetReqAttr(.txtW33, "Q")
			.txtW34.Value = 0 :	call ggoOper.SetReqAttr(.txtW34, "Q")
			.txtW35.Value = 0 :	call ggoOper.SetReqAttr(.txtW35, "Q")
			.txtW40.Value = 0 :	call ggoOper.SetReqAttr(.txtW40, "Q")
			.txtW41.Value = 0 :	call ggoOper.SetReqAttr(.txtW41, "Q")
			.txtW42.Value = 0 :	call ggoOper.SetReqAttr(.txtW42, "Q")
			.txtW43.Value = 0 :	call ggoOper.SetReqAttr(.txtW43, "Q")
			.txtW44.Value = 0 :	call ggoOper.SetReqAttr(.txtW44, "Q")
			.txtW50.Value = 0 :	call ggoOper.SetReqAttr(.txtW50, "Q")

'			call ggoOper.SetReqAttr(.txtW1, "D")
			call ggoOper.SetReqAttr(.txtW2, "D")
			call ggoOper.SetReqAttr(.txtW3, "D")
			call ggoOper.SetReqAttr(.txtW4, "D")
			call ggoOper.SetReqAttr(.txtW5, "D")
'			call ggoOper.SetReqAttr(.txtW6, "D")
			call ggoOper.SetReqAttr(.txtW8, "D")
'			call ggoOper.SetReqAttr(.txtW10, "D")
'			call ggoOper.SetReqAttr(.txtW11, "D")
			call ggoOper.SetReqAttr(.txtW12, "D")
			call ggoOper.SetReqAttr(.txtW13, "D")
			call ggoOper.SetReqAttr(.txtW14, "D")
'			call ggoOper.SetReqAttr(.txtW15, "D")
			call ggoOper.SetReqAttr(.txtW16, "D")
			call ggoOper.SetReqAttr(.txtW17, "D")
			call ggoOper.SetReqAttr(.txtW18, "D")
			call ggoOper.SetReqAttr(.txtW19, "D")
			call ggoOper.SetReqAttr(.txtW20, "D")
'			call ggoOper.SetReqAttr(.txtW25, "D")

			call ggoOper.SetReqAttr(.txtW26, "D")
			call ggoOper.SetReqAttr(.txtW27, "D")
			call ggoOper.SetReqAttr(.txtW28, "D")

			.txtW2.Value = .txtRW1.Value
			.txtW6.Value = .txtRW2.Value
		End With
	Else
		With Frm1
			.chkW_TYPE1.checked = False
			.chkW_TYPE2.checked = True
			.txtW1.Value = 0 :	call ggoOper.SetReqAttr(.txtW1, "Q")
			.txtW2.Value = 0 :	call ggoOper.SetReqAttr(.txtW2, "Q")
			.txtW3.Value = 0 :	call ggoOper.SetReqAttr(.txtW3, "Q")
			.txtW4.Value = 0 :	call ggoOper.SetReqAttr(.txtW4, "Q")
			.txtW5.Value = 0 :	call ggoOper.SetReqAttr(.txtW5, "Q")
			.txtW6.Value = 0 :	call ggoOper.SetReqAttr(.txtW6, "Q")
			.txtW8.Value = 0 :	call ggoOper.SetReqAttr(.txtW8, "Q")
			.txtW10.Value = 0 :	call ggoOper.SetReqAttr(.txtW10, "Q")
			.txtW11.Value = 0 :	call ggoOper.SetReqAttr(.txtW11, "Q")
			.txtW12.Value = 0 :	call ggoOper.SetReqAttr(.txtW12, "Q")
			.txtW13.Value = 0 :	call ggoOper.SetReqAttr(.txtW13, "Q")
			.txtW14.Value = 0 :	call ggoOper.SetReqAttr(.txtW14, "Q")
			.txtW15.Value = 0 :	call ggoOper.SetReqAttr(.txtW15, "Q")
			.txtW16.Value = 0 :	call ggoOper.SetReqAttr(.txtW16, "Q")
			.txtW17.Value = 0 :	call ggoOper.SetReqAttr(.txtW17, "Q")
			.txtW18.Value = 0 :	call ggoOper.SetReqAttr(.txtW18, "Q")
			.txtW19.Value = 0 :	call ggoOper.SetReqAttr(.txtW19, "Q")
			.txtW20.Value = 0 :	call ggoOper.SetReqAttr(.txtW20, "Q")
			.txtW25.Value = 0 :	call ggoOper.SetReqAttr(.txtW25, "Q")
			
			.txtW26.Value = 0 :	call ggoOper.SetReqAttr(.txtW26, "Q")
			.txtW27.Value = 0 :	call ggoOper.SetReqAttr(.txtW27, "Q")
			.txtW28.Value = 0 :	call ggoOper.SetReqAttr(.txtW28, "Q")

'			call ggoOper.SetReqAttr(.txtW30, "D")
			call ggoOper.SetReqAttr(.txtW31, "D")
			call ggoOper.SetReqAttr(.txtW32, "D")
			call ggoOper.SetReqAttr(.txtW33, "D")
			call ggoOper.SetReqAttr(.txtW34, "D")
'			call ggoOper.SetReqAttr(.txtW35, "D")
'			call ggoOper.SetReqAttr(.txtW40, "D")
			call ggoOper.SetReqAttr(.txtW41, "D")
			call ggoOper.SetReqAttr(.txtW42, "D")
			call ggoOper.SetReqAttr(.txtW43, "D")
			call ggoOper.SetReqAttr(.txtW44, "D")
'			call ggoOper.SetReqAttr(.txtW50, "D")
			.txtW31.Value = .txtRW1.Value
			.txtW35.Value = unicdbl(.txtRW2.Value) * -1
		End With
	End If
    Call SetAllTxtChkCalc()
End Sub

Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub


Sub txtW2_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw3_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw4_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw5_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw8_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw12_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw13_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw14_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw16_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw17_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw18_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw19_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw20_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw31_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw32_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw33_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw34_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw41_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw42_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw43_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub

Sub txtw44_Change( )
    lgBlnFlgChgValue = True
    Call SetAllTxtChkCalc()
End Sub


Sub SetAllTxtChkCalc()
	'10 < 0  &  15 = 0	오류	W10003	처분전 이익잉여금과 임의적립금 이입액의 합계금액이 0보다 적습니다. 결손금처리계산서를 작성하여 주십시요 
	'30 < 0	오류	W10002	처분전결손금의 금액이 0보다 적습니다. 이익잉여금처분계산서를 작성하여 주십시요 
	'50 < 0	오류	W10002	차기이월결손금의 금액이 0보다 적습니다. 이익잉여금처분계산서를 작성하여 주십시요 

	
	'01		02+03+04-05+06 를 계산하여 입력함.		=> 02+03+04+05+06 를 계산하여 입력함 
	'10		01 + 08 를 계산하여 입력함.	
	'11		12+13+14+15+18+19+20 를 계산하여 입력함.	(+26+27+28 추가 2006.03개정)	
	'15		16+17 를 계산하여 입력함.	
	'25		10 - 11 를 계산하여 입력함.	
	'30		(31+32+33-34-35) × (-1) 를 계산하여 입력함.		
	'35		표준손익계산서의 당기순이익(순손실) × (-1)을 입력함.	불러오기 
	'40		41+42+43+44 를 계산하여 입력함.		
	'50		30 - 40  를 계산하여 입력함.	

    Frm1.txtW1.value = unicdbl(Frm1.txtw2.value) + unicdbl(Frm1.txtw3.value) + unicdbl(Frm1.txtw4.value) - unicdbl(Frm1.txtw5.value) + unicdbl(Frm1.txtw6.value)

    Frm1.txtW10.value = unicdbl(Frm1.txtw1.value) + unicdbl(Frm1.txtw8.value)
	
    Frm1.txtW15.value = unicdbl(Frm1.txtw16.value) + unicdbl(Frm1.txtw17.value)
	
    Frm1.txtW11.value = unicdbl(Frm1.txtw12.value) + unicdbl(Frm1.txtw13.value) + unicdbl(Frm1.txtw14.value) + unicdbl(Frm1.txtw15.value) + unicdbl(Frm1.txtw18.value) + unicdbl(Frm1.txtw19.value) + unicdbl(Frm1.txtw20.value) + unicdbl(Frm1.txtw26.value) + unicdbl(Frm1.txtw27.value) + unicdbl(Frm1.txtw28.value)

    Frm1.txtW25.value = unicdbl(Frm1.txtw10.value) - unicdbl(Frm1.txtw11.value)
	
    Frm1.txtW30.value = (unicdbl(Frm1.txtw31.value) + unicdbl(Frm1.txtw32.value) + unicdbl(Frm1.txtw33.value) - unicdbl(Frm1.txtw34.value) - unicdbl(Frm1.txtw35.value)) * -1

    Frm1.txtW40.value = unicdbl(Frm1.txtw41.value) + unicdbl(Frm1.txtw42.value) + unicdbl(Frm1.txtw43.value) + unicdbl(Frm1.txtw44.value)

    Frm1.txtW50.value = unicdbl(Frm1.txtW30.value) - unicdbl(Frm1.txtW40.value)

End Sub

'==========================================================================================
Sub InitData()
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	Call CheckFISC_DATE
End Sub


'============================================  툴바지원 함수  ====================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

  '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
	
     
    If DbQuery = False Then
		Call RestoreToolBar()
        Exit Function
    End If

'    Call DbQuery
    FncQuery = True
End Function

Function FncSave() 
	Dim IntRetCD
        
    FncSave = False                                                         
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
	If lgBlnFlgChgValue = False Then
	    IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                          '⊙: No data changed!!
	    Exit Function
	End If

	'-----------------------
	'Check content area
	'-----------------------
	If Not chkField(Document, "A") Then                             '⊙: Check contents area
	   Exit Function
	End If

    If NOT Frm1.chkW_TYPE1.checked And NOT Frm1.chkW_TYPE2.checked Then
        Call DisplayMsgBox("W10001", "X", "X", "X")
	    Frm1.chkW_TYPE1.focus
	    Exit Function
    End If
    
	If Verification = False Then Exit Function

<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function


' ---------------------- 서식내 검증 -------------------------
Function  Verification()
	
	Verification = False

    If unicdbl(Frm1.txtw10.value) < 0 And unicdbl(Frm1.txtw15.value) = 0 Then
        Call DisplayMsgBox("W10005", "X", "X", "X")
	    Frm1.chkW_TYPE2.focus
	    Exit Function
    End If
    
    If unicdbl(Frm1.txtw30.value) < 0 Then
        Call DisplayMsgBox("W10004", "X", "처분전결손금", "X")
	    Frm1.chkW_TYPE1.focus
	    Exit Function
    End If

    If unicdbl(Frm1.txtw50.value) < 0 Then
        Call DisplayMsgBox("W10004", "X", "차기이월결손금", "X")
	    Frm1.chkW_TYPE1.focus
	    Exit Function
    End If

	Verification = True	
End Function


'========================================================================================
Function FncNew() 
    Dim IntRetCD 

    FncNew = False

  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    Call ggoOper.LockField(Document, "N")
    Call InitVariables

    Call SetToolbar("1100100000001011")
    lgIntFlgMode = parent.OPMD_CMODE

	frm1.txtCO_CD.focus

    FncNew = True

End Function


Function FncCopy() 
	exit Function
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

	
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
     On Error Resume Next
End Function

Function FncInsertRow(ByVal pvRowCnt) 
     On Error Resume Next
End Function

Function FncDeleteRow() 
     On Error Resume Next
End Function

Function FncPrint()
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)											 <%'☜: 화면 유형 %>
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                      <%'☜:화면 유형, Tab 유무 %>
End Function

Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")

		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
Function FncDelete() 
    Dim IntRetCD

    FncDelete = False
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF

    Call DbDelete

    FncDelete = True
End Function

'============================================  DB 억세스 함수  ====================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
	    strVal = strVal 	& "&txtCo_Cd=" 			 & Trim(frm1.txtCo_Cd.value)				'☆: 조회 조건 데이타 
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key   
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Call SetToolbar("1101100000010111")
    lgBlnFlgChgValue = False
'    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field

    lgIntFlgMode = parent.OPMD_UMODE
	Call SetW_TYPE(Frm1.txtW_TYPE.Value)

    		
End Function

Function DbQueryFalse()													<%'조회 성공후 실행로직 %>
	
    Call SetToolbar("1101100000010111")
    Frm1.chkW_TYPE1.checked = true
    Call chkW_TYPE1_OnClick()
    Call InitVariables   
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
     Err.Clear
	DbSave = False

    Dim strVal

    Call LayerShowHide(1) 

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value     = lgIntFlgMode
		
		If .chkW_TYPE1.checked Then
			.txtW_TYPE.Value = "1"
		Else
			.txtW_TYPE.Value = "2"
		End If

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	End With

    DbSave = True                                                         
End Function


Function DbSaveOk()													        <%' 저장 성공후 실행 로직 %>
    lgBlnFlgChgValue = False
    'FncQuery
End Function

'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal 	& "&txtCo_Cd=" 			 & Trim(frm1.txtCo_Cd.value)				'☆: 조회 조건 데이타 
    strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
    strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key            
	
	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function

Sub txtW_DT_DblClick(Button)
    If Button = 1 Then
       frm1.txtW_DT.Action = 7                                    ' 7 : Popup Calendar ocx
       Call SetFocusToDocument("M")	
       frm1.txtW_DT.Focus
    End If
End Sub


'========================================================================================
Function DbDeleteOk()
	Call FncNew()
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;
						<a href="vbscript:GetRef">금액불러오기</A>
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">사업연도</TD>
									<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFISC_YEAR CLASS=FPDTYYYY title=FPDATETIME ALT="사업연도" tag="14X1" id=txtFISC_YEAR></OBJECT>');</SCRIPT>
									<TD CLASS="TD5">법인명</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
									</TD>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">신고구분</TD>
									<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE" ALT="신고구분" STYLE="WIDTH: 50%" tag="14X"></SELECT>
									</TD>
									<TD CLASS="TD5"></TD>
									<TD CLASS="TD6"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<DIV ID="ViewDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto"><% ' -- overflow=auto : 컨텐츠 구역을 브라우저 크기에 따라 스크롤바가 생성되게 한다 %>
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=1>
                            <TR>
                                <TD WIDTH="100%" VALIGN=TOP>
                                   <FIELDSET CLASS="CLSFLD">
                                   <BR>
                                   <TABLE width = 100% bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
									   <TR>
									       <TD CLASS="TD51" width="50%" ALIGN=CENTER Colspan="3"><LABEL FOR="chkBpTypeC">1. 이익잉여금처분계산서</LABEL>
									       		<INPUT TYPE=CHECKBOX NAME="chkW_TYPE1" ID="chkW_TYPE1" tag="21" Class="Check">
									       </TD>
									       <TD CLASS="TD51" width="50%" ALIGN=CENTER Colspan="3"><LABEL FOR="chkBpTypeC">2. 결손금처리계산서</LABEL>
									       		<INPUT TYPE=CHECKBOX NAME="chkW_TYPE2" ID="chkW_TYPE2" tag="21" Class="Check">
								           </TD>
									  </TR>
									   <TR>
									       <TD CLASS="TD51" width="25%" ALIGN=CENTER>과목</TD>
									       <TD CLASS="TD51" width="7%" ALIGN=CENTER>코드</TD>
									       <TD CLASS="TD51" width="18%" ALIGN=CENTER>금액</TD>
									       <TD CLASS="TD51" width="25%" ALIGN=CENTER>과목</TD>
									       <TD CLASS="TD51" width="7%" ALIGN=CENTER>코드</TD>
									       <TD CLASS="TD51" width="18%" ALIGN=CENTER>금액</TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT><B>I. 처분전이익잉여금</B></TD>
									       <TD CLASS="TD51" ALIGN=CENTER>01</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW1" name=txtW1 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X20" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT><B>I. 처리전결손금</B></TD>
									       <TD CLASS="TD51" ALIGN=CENTER>30</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW30" name=txtW30 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X20" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;1. 전기이월이익잉여금(또는<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;전기이월결손금)</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>02</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW2" name=txtW2 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X20" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;1. 전기이월이익잉여금(또는<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;전기이월결손금)</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>31</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW31" name=txtW31 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X20" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;2. 회계변경의 누적효과</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>03</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW3" name=txtW3 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X20" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;2. 회계변경의 누적효과</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>32</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW32" name=txtW32 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X20" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;3. 전기오류수정이익(또는<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;전기오류수정손실)</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>04</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW4" name=txtW4 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X20" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;3. 전기오류수정이익(또는<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;전기오류수정손실)</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>33</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW33" name=txtW33 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X20" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;4. 중간배당액</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>05</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW5" name=txtW5 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X20" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;4. 중간배당액</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>34</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW34" name=txtW34 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X20" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;5. 당기순이익(또는<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;당기순손실)</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>06</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW6" name=txtW6 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X20" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;5. 당기순이익(또는<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;당기순손실)</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>35</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW35" name=txtW35 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X20" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT><B>II. 임의적립금 등의 이입액</B></TD>
									       <TD CLASS="TD51" ALIGN=CENTER>08</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8" name=txtW8 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT><B>II. 결손금 처리액</B></TD>
									       <TD CLASS="TD51" ALIGN=CENTER>40</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW40" name=txtW40 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X20" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=CENTER>합     계</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>10</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW10" name=txtW10 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X20" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;1. 임의적립금이입액</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>41</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW41" name=txtW41 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT><B>III. 이익잉여금 처분액</B></TD>
									       <TD CLASS="TD51" ALIGN=CENTER>11</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW11" name=txtW11 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X20" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;2. 기타법정적립금이입액</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>42</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW42" name=txtW42 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;1. 이익준비금</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>12</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW12" name=txtW12 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;3. 이익준비금이입액</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>43</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW43" name=txtW43 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;2. 기타법정적립금</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>13</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW13" name=txtW13 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;4. 자본잉여금이입액</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>44</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW44" name=txtW44 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;3. 주식할인발행차금상각액</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>14</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW14" name=txtW14 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT><B>III. 차기이월결손금</B></TD>
									       <TD CLASS="TD51" ALIGN=CENTER>50</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW50" name=txtW50 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X20" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;4. 배당금</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>15</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW15" name=txtW15 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD61" colspan="3"></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;&nbsp;&nbsp;가. 현금배당</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>16</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW16" name=txtW16 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD51" ALIGN=LEFT><B>처분(처리) 확정일</B></TD>
									       <TD CLASS="TD51" ALIGN=CENTER></TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtW_DT CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="처분(처리) 확정일" tag="22X1" id=txtW_DT style="width: 100%"></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;&nbsp;&nbsp;나. 주식배당</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>17</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW17" name=txtW17 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									       <TD CLASS="TD61" colspan="3" rowspan="8"></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;5. 이익처분에 의한 상여금</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>26</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW26" name=txtW26 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;6. 사업확장적립금</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>18</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW18" name=txtW18 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;7. 감채적립금</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>19</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW19" name=txtW19 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;8. 기타적립금</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>20</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW20" name=txtW20 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;9. 조세특례제한법 상 준비금 등 적립액</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>27</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW27" name=txtW27 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT>&nbsp;&nbsp;10. 기타 잉여금처분액</TD>
									       <TD CLASS="TD51" ALIGN=CENTER>28</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW28" name=txtW28 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
									  <TR>
									       <TD CLASS="TD51" ALIGN=LEFT><B>IV. 차기이월이익잉여금</B></TD>
									       <TD CLASS="TD51" ALIGN=CENTER>25</TD>
									       <TD CLASS="TD61" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW25" name=txtW25 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
									  </TR>
								  </table>
								  </FIELDSET>
								</TD>
							</TR>
						</TABLE>
						</DIV>
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
			<TABLE CLASS="TB3" CELLSPACING=0>
			
			
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('VIEW')" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('PRINT')"   Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtW_TYPE" tag="24">
<INPUT TYPE=hidden NAME="txtRW1" tag="24" tabindex="-1" value="0">
<INPUT TYPE=hidden NAME="txtRW2" tag="24" tabindex="-1" value="0">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
