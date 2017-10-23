<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 법인세 조정 
'*  3. Program ID           : W8101MA1
'*  4. Program Name         : W8101MA1.asp
'*  5. Program Desc         : 제3호 법인세과세표준 및 세액조정계산서 
'*  6. Modified date(First) : 2005/01/27
'*  7. Modified date(Last)  : 2005/01/27
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : 최영태 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
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
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  상수/변수 선언  ====================================
<!-- #Include file="../../inc/lgvariables.inc" -->



Const BIZ_MNU_ID		= "W8101MA1"
Const BIZ_PGM_ID		= "W8101mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_REF_PGM_ID	= "W8101mb2.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_POP_ID		= "W8101MA2.asp"
Const EBR_RPT_ID		 = "W8107OA1"

Const TYPE_1	= 0		' 그리드 배열번호 및 디비의 W_TYPE 컬럼의 값. 
Const TYPE_2	= 1		' 즉 멀티 그리드 PG이지만 단일 테이블의 코드로 관리된다.

' -- 그리드 컬럼 정의 
Dim C_W1	' 그리드 열 
Dim C_W2
Dim C_W2_1
Dim C_W2_2
Dim C_W3
Dim C_W4

Dim C_W01	' 그리드 행(디비의 열)
Dim C_W02	
Dim C_W03	
Dim C_W04	
Dim C_W05	
Dim C_W54	
Dim C_W06	
Dim C_W06_1
Dim C_W07	
Dim C_W08	
Dim C_W09	
Dim C_W10	
Dim C_W10_1
Dim C_W11	
Dim C_W12	
Dim C_W13	
Dim C_W14	
Dim C_W15	
Dim C_W16	
Dim C_W16_1
Dim C_W17	
Dim C_W18	
Dim C_W19	
Dim C_W20	
Dim C_W21	
Dim C_W22	
Dim C_W23	
Dim C_W24	
Dim C_W25	
Dim C_W26	
Dim C_W27	
Dim C_W28	
Dim C_W29	
Dim C_W30	
Dim C_W31	
Dim C_W32	
Dim C_W33	
Dim C_W34	
Dim C_W35	
Dim C_W36	
Dim C_W37	
Dim C_W38	
Dim C_W39	
Dim C_W40	
Dim C_W41	
Dim C_W42	
Dim C_W43	
Dim C_W44	
Dim C_W45	
Dim C_W46	
Dim C_W55	
Dim C_W47	
Dim C_W48	
Dim C_W49	
Dim C_W50	
Dim C_W51	
Dim C_W52	
Dim C_W53	

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 
Dim lgCurrGrid, lgvspdData(2)
Dim	lgFISC_START_DT, lgFISC_END_DT, lgW2018

Dim IsRunEvents	' ㅠㅠ 무한이벤트반복을 막음 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	' -- 그리드 1
	C_W1	= 1
	C_W2	= 2
	C_W2_1	= 3
	C_W2_2	= 4
	C_W3	= 5
	C_W4	= 6
	
	C_W01	= 0
	C_W02	= 1
	C_W03	= 2
	C_W04	= 3
	C_W05	= 4
	C_W54	= 5
	C_W06	= 6
	
	C_W06_1 = 7
	C_W07	= 8
	C_W08	= 9
	C_W09	= 10
	C_W10	= 11
	
	C_W10_1	= 12
	C_W11	= 13
	C_W12	= 14
	C_W13	= 15
	C_W14	= 16
	C_W15	= 17
	C_W16	= 18
	
	C_W16_1	= 19
	C_W17	= 20
	C_W18	= 21
	C_W19	= 22
	C_W20	= 23
	C_W21	= 24
	C_W22	= 25
	C_W23	= 26
	C_W24	= 27
	C_W25	= 28
	C_W26	= 29
	C_W27	= 30
	C_W28	= 31
	
	' -- 그리드2
	C_W29	= 32
	C_W30	= 33
	C_W31	= 34
	C_W32	= 35
	C_W33	= 36
	C_W34	= 37
	C_W35	= 38
	C_W36	= 39
	C_W37	= 40
	C_W38	= 41
	C_W39	= 42
	C_W40	= 43
	C_W41	= 44
	C_W42	= 45
	C_W43	= 46
	C_W44	= 47
	C_W45	= 48
	C_W46	= 49
	C_W55	= 50 ' <-- 2003.03.07 개정추가 
	C_W47	= 51
	C_W48	= 52
	C_W49	= 53
	C_W50	= 54
	C_W51	= 55
	C_W52	= 56
	C_W53	= 57
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""                          'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
    lgRefMode = False

    lgCurrGrid = TYPE_1
    IsRunEvents = False
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub



'============================================  신고구분 콤보 박스 채우기  ====================================

Sub InitComboBox()
	' 조회조건(구분)
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))

	Call InitSpreadPosVariables

End Sub

Sub InitSpreadSheet()

	Call AppendNumberPlace("6","3","2")
	Call AppendNumberPlace("8","15","0")	' -- 금액 15자리 고정 : 출하검사패치 
	
End Sub

Sub InitComboBox2()
	
	call CommonQueryRs("MINOR_CD, MINOR_NM + ' ('+ REFERENCE_2 + ')', REFERENCE_1, REFERENCE_2"," ufn_TB_Configuration('W2014','" & C_REVISION_YM & "') ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetComboX(frm1.txtData(C_W14) , lgF0, lgF1, lgF2, lgF3, Chr(11))
 
 	call CommonQueryRs("MINOR_CD, MINOR_NM+ ' ('+ REFERENCE_2 + ')', REFERENCE_1, REFERENCE_2"," ufn_TB_Configuration('W2013','" & C_REVISION_YM & "') ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetComboX(frm1.txtData(C_W35) , lgF0, lgF1, lgF2, lgF3, Chr(11))
    
 	call CommonQueryRs("REFERENCE_1"," ufn_TB_Configuration('W2018','" & C_REVISION_YM & "') ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
           
    lgW2018 = Split(lgF0 , chr(11))
    
End Sub

'============================================  그리드 함수  ====================================

Sub InitData()

	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
       
	Call GetFISC_DATE

	Call InitComboBox2()
End Sub

Sub InitSpreadComboBox()

End Sub

'============================== 레퍼런스 함수  ========================================

Function GetRef()	' 금액가져오기 링크 클릭시 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg, arrW1, arrW2, iRow, iMaxRows, sTmp

	' 온로드시 레퍼런스메시지 가져온다.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
			
	IntRetCD = CommonQueryRs("W1, W2"," dbo.ufn_TB_3_GetRef_" & C_REVISION_YM & "('" & sCoCd & "','" & sFiscYear & "','" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD = True Then
		arrW1		= Split(lgF0, chr(11))
		arrW2		= Split(lgF1, chr(11))
		iMaxRows	= UBound(arrW1)

		For iRow = 0 To iMaxRows -1
			sTmp = "frm1.txtData(C_W" & arrW1(iRow) & ").Text = """ & CStr(arrW2(iRow)) & """"	
			Execute sTmp	' -- 변수에 들어 있는 명령을 실행한다.  ** ufn_TB_3_GetRef의 W1필드의 값을 코드로 하지 않고, 코드의 배열인덱스값으로 하면 이렇게 안해도 됨.
		Next
	Else
		Call DisplayMsgBox("900014", parent.VB_INFORMATION, "", "")             '☜ : No data is found.
	End If
	
	Call SetHeadReCalc()
End Function

' 헤더 재계산 
Sub SetHeadReCalc()	
	Dim dblSum, dblW4(100), dblMonGap
	
	If IsRunEvents Then Exit Sub	' 아래 .vlaue = 에서 이벤트가 발생해 재귀함수로 가는걸 막는다.
	
	IsRunEvents = True
	
	With frm1
		dblW4(C_W01) = UNICDbl(.txtData(C_W01).value)
		dblW4(C_W02) = UNICDbl(.txtData(C_W02).value)
		dblW4(C_W03) = UNICDbl(.txtData(C_W03).value)
		
		dblW4(C_W04) = dblW4(C_W01) + dblW4(C_W02) - dblW4(C_W03)
		.txtData(C_W04).value = dblW4(C_W04)	' (104) 차가감소득금액 
		
		dblW4(C_W05) = UNICDbl(.txtData(C_W05).value)
		dblW4(C_W54) = UNICDbl(.txtData(C_W54).value)
		
		dblW4(C_W06) = dblW4(C_W04) + dblW4(C_W05) - dblW4(C_W54)
		.txtData(C_W06).value = dblW4(C_W06)	' (107) 각사업연도소득금액 
		.txtData(C_W06_1).value = dblW4(C_W06)
			
		dblW4(C_W07) = UNICDbl(.txtData(C_W07).value)
		dblW4(C_W08) = UNICDbl(.txtData(C_W08).value)
		dblW4(C_W09) = UNICDbl(.txtData(C_W09).value)		
		
		dblW4(C_W10) = dblW4(C_W06) - dblW4(C_W07) - dblW4(C_W08) - dblW4(C_W09)
		.txtData(C_W10).value = dblW4(C_W10)	' (112) 과세표준 
		
		' -- 2006-01-04: 200603 개정판 
		dblW4(C_W10_1) = dblW4(C_W10) + UNICDbl(.txtW55_1.value)
		.txtData(C_W10_1).value = dblW4(C_W10_1)

		' -- 114-115는 고려중: 설명서이해못함 
		If frm1.cboREP_TYPE.value = "2" Then
			dblMonGap = 6
		Else
			dblMonGap = DateDiff("m", lgFISC_START_DT, lgFISC_END_DT)+1
		End If
		
		' -- 2006-01-04: 200603 개정판 
		'dblSum = dblW4(C_W10) * 12 / dblMonGap
		dblSum = dblW4(C_W10_1) * 12 / dblMonGap

		If dblSum <= 100000000 Then
			.txtData(C_W11).value = lgW2018(0) ' 1억이하 
		Else
			.txtData(C_W11).value = lgW2018(1)	'1억초과 
		End If

		If dblSum <= 0 Then
			.txtData(C_W12).value = 0 ' 1억이하 
		Else
			If dblSum <= 100000000 Then
				.txtData(C_W12).value = (dblSum * lgW2018(0) * dblMonGap) / 12
			Else
				.txtData(C_W12).value = ((dblSum * lgW2018(1) * dblMonGap) / 12) - ( ( 100000000 * (lgW2018(1) - lgW2018(0)) * dblMonGap) / 12)
			End If
		End If
		'zzzzzzzzzzzzzz
			.txtData(C_W12).value = "782839149"	
		dblW4(C_W11) = UNICDbl(.txtData(C_W11).value)
		dblW4(C_W12) = UNICDbl(.txtData(C_W12).value)
		
		dblW4(C_W13) = UNICDbl(.txtData(C_W13).value)	
		If 	.txtData(C_W14).value <> "" Then
			dblW4(C_W14) = UNICDbl(.txtData(C_W14).options(.txtData(C_W14).selectedIndex).VAL)		' -- 다이나믹 콤보값: B_Configuration.SeqNo=1
			.txtW14.value = .txtData(C_W14).options(.txtData(C_W14).selectedIndex).VIEW				' -- 다이나믹 콤보값: B_Configuration.SeqNo=2
		Else
			dblW4(C_W14) = 0
		End If
		
		dblW4(C_W15) = dblW4(C_W13) * dblW4(C_W14)
		.txtData(C_W15).value = dblW4(C_W15)	' (118) 산출세액 

		dblW4(C_W16) = dblW4(C_W12) + dblW4(C_W15)	
		.txtData(C_W16).value = dblW4(C_W16)	' (119) 합계 
		.txtData(C_W16_1).value = dblW4(C_W16)
		
		dblW4(C_W17) = UNICDbl(.txtData(C_W17).value)
		dblW4(C_W18) = dblW4(C_W16) - dblW4(C_W17)	
		.txtData(C_W18).value = dblW4(C_W18)	' (122) 차감세액 
		
		dblW4(C_W19) = UNICDbl(.txtData(C_W19).value)
		dblW4(C_W20) = UNICDbl(.txtData(C_W20).value)

		dblW4(C_W21) = dblW4(C_W18) - dblW4(C_W19) + dblW4(C_W20)
		.txtData(C_W21).value = dblW4(C_W21)	' (125) 가감계 
		
		dblW4(C_W22) = UNICDbl(.txtData(C_W22).value)
		dblW4(C_W23) = UNICDbl(.txtData(C_W23).value)
		dblW4(C_W24) = UNICDbl(.txtData(C_W24).value)		
		dblW4(C_W25) = UNICDbl(.txtData(C_W25).value)		
		
		dblW4(C_W26) = dblW4(C_W22) + dblW4(C_W23) + dblW4(C_W24) + dblW4(C_W25)
		.txtData(C_W26).value = dblW4(C_W26)	' (130) 소계 
		
		dblW4(C_W27) = UNICDbl(.txtData(C_W27).value)	
		dblW4(C_W28) = dblW4(C_W26) + dblW4(C_W27)
		.txtData(C_W28).value = dblW4(C_W28)	' (132) 합계 
		
		dblW4(C_W29) = UNICDbl(.txtData(C_W29).value)	
		dblW4(C_W30) = dblW4(C_W21) - dblW4(C_W28) + dblW4(C_W29)
		.txtData(C_W30).value = dblW4(C_W30)	' (134) 차감납부할세액	
		
		dblW4(C_W31) = UNICDbl(.txtData(C_W31).value)
		dblW4(C_W32) = UNICDbl(.txtData(C_W32).value)
		dblW4(C_W33) = UNICDbl(.txtData(C_W33).value)	
		dblW4(C_W34) = dblW4(C_W31) + dblW4(C_W32) - dblW4(C_W33)
		.txtData(C_W34).value = dblW4(C_W34)	' (138) 과세표준 
		
		If .txtData(C_W35).value <> "" Then	
			dblW4(C_W35) = UNICDbl(.txtData(C_W35).options(.txtData(C_W35).selectedIndex).VAL)		' -- 다이나믹 콤보값: B_Configuration.SeqNo=1	
			'dblW4(C_W36) = dblW4(C_W34) * dblW4(C_W35)
			.txtW35.value = .txtData(C_W35).options(.txtData(C_W35).selectedIndex).VIEW				' -- 다이나믹 콤보값: B_Configuration.SeqNo=2
		Else
			'dblW4(C_W36) = 0
		End If
		
		'If pMode <> 1 Then	' -- 자동서식계산에서 제외 
		'	.txtData(C_W36).value = dblW4(C_W36)	' (140) 산출세액 
		'End If
		dblW4(C_W36) = UNICDbl(.txtData(C_W36).value)
		
		dblW4(C_W37) = UNICDbl(.txtData(C_W37).value)	
		dblW4(C_W38) = dblW4(C_W36) - dblW4(C_W37)
		.txtData(C_W38).value = dblW4(C_W38)	' (142) 차감세액 
		
		dblW4(C_W39) = UNICDbl(.txtData(C_W39).value)
		dblW4(C_W40) = UNICDbl(.txtData(C_W40).value)	
		dblW4(C_W41) = dblW4(C_W38) - dblW4(C_W39) + dblW4(C_W40)
		.txtData(C_W41).value = dblW4(C_W41)	' (145) 가감계 
			
		dblW4(C_W42) = UNICDbl(.txtData(C_W42).value)
		dblW4(C_W43) = UNICDbl(.txtData(C_W43).value)	
		dblW4(C_W44) = dblW4(C_W42) + dblW4(C_W43)
		.txtData(C_W44).value = dblW4(C_W44)	' (148) 계 
		
		dblW4(C_W45) = dblW4(C_W41) - dblW4(C_W44)
		.txtData(C_W45).value = dblW4(C_W45)	' (149) 차감납부할세액 
		
		dblW4(C_W46) = dblW4(C_W30) + dblW4(C_W45)
		.txtData(C_W46).value = dblW4(C_W46)	' (150) 차감납부할세액계 
		
		dblW4(C_W55) = UNICDbl(.txtData(C_W55).value)
		
		dblW4(C_W47) = dblW4(C_W46) - dblW4(C_W20) - dblW4(C_W29) - dblW4(C_W40) + dblW4(C_W27) - dblW4(C_W55)	' -- 200603 버그수정 
		If dblW4(C_W47) < 0 Then
			.txtData(C_W47).value = 0
		Else
			.txtData(C_W47).value = dblW4(C_W47)	' (151) 분납세액계산범위액 
		End If
		
		If dblW4(C_W47) <= 10000000 Then
			dblW4(C_W50) = 0
		ElseIf dblW4(C_W47) > 10000000 AND dblW4(C_W47) <= 20000000 Then
			dblW4(C_W50) = dblW4(C_W47) - 10000000
		ElseIf dblW4(C_W47) > 20000000 Then
			dblW4(C_W50) = Fix(dblW4(C_W47) * 0.5)
		End If
		.txtData(C_W50).value = dblW4(C_W50)	' (154) 계 

		dblW4(C_W49) = UNICDbl(.txtData(C_W49).value)
		'dblW4(C_W50) = UNICDbl(.txtData(C_W50).value)	
		dblW4(C_W48) = dblW4(C_W50) - dblW4(C_W49) 
		.txtData(C_W48).value = dblW4(C_W48)	' (152) 현금납부 

		dblW4(C_W55) = UNICDbl(.txtData(C_W55).value)
		dblW4(C_W53) = dblW4(C_W46) - dblW4(C_W55) - dblW4(C_W50)
		.txtData(C_W53).value = dblW4(C_W53)	' (157) 계 
						
		dblW4(C_W52) = UNICDbl(.txtData(C_W52).value)
		'dblW4(C_W53) = UNICDbl(.txtData(C_W53).value)
		dblW4(C_W51) = dblW4(C_W53) - dblW4(C_W52)
		.txtData(C_W51).value = dblW4(C_W51)	' (155) 현금납부 

	End With

	lgBlnFlgChgValue= True ' 변경여부 
	IsRunEvents = False	' 이벤트 발생금지를 해제함 
End Sub

Sub GetFISC_DATE()	' 법인의 조회조건에 만족하는 당기시작,종료일을 가져온다.
	Dim sFiscYear, sRepType, sCoCd, iGap
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	' 법인 기간은 필수입력 
	lgFISC_START_DT = CDate(lgF0)
	lgFISC_END_DT = CDate(lgF1)

End Sub

Function OpenW07()	'이월결손금 팝업 

    Dim arrRet
    Dim arrParam(4)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
	Dim arrRowVal
    Dim arrColVal, lLngMaxRow
    Dim iDx
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True
    
	arrParam(0) = frm1.txtCO_CD.Value
	arrParam(1) = frm1.txtFISC_YEAR.Text		
	arrParam(2) = frm1.cboREP_TYPE.Value		
	arrParam(3) = UNICDbl(frm1.txtData(C_W06_1).value)		

    arrRet = window.showModalDialog(BIZ_POP_ID, Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0) <> "" Then
		frm1.txtData(C_W07).value = arrRet(0)
		
		Call SetHeadReCalc
	End IF
    
    IsOpenPop = False
    
    
End Function


Function txtData_onchange()
	Call SetHeadReCalc
End Function

'====================================== 탭 함수 =========================================

'============================================  조회조건 함수  ====================================


'============================================  폼 함수  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         

    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100100000000111")										<%'버튼 툴바 제어 %>
	 
	Call InitComboBox
	
	' 변경한곳 
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
	
	Call InitData 
	
    'Call ggoOper.FormatDate(frm1.txtW2 , parent.gDateFormat,3)
	Call FncQuery
	'
End Sub


'============================================  이벤트 함수  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub

Sub cboREP_TYPE_onChange()	' 신고기준을 바꾸면..
	Call GetFISC_DATE
End Sub

'============================================  그리드 이벤트   ====================================

'============================================  툴바지원 함수  ====================================

Function FncQuery() 
    Dim IntRetCD , i, blnChange
    
    FncQuery = False                                                        
    blnChange = False
    
    Err.Clear                                                               <%'Protect system from crashing%>
	
	
    
	
<%  '-----------------------
    'Check previous data area
    '----------------------- %>  
    If lgBlnFlgChgValue Or blnChange Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call InitVariables													<%'Initializes local global variables%>
    'Call InitData                              
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
	 
    CALL DBQuery()
    
End Function

Function FncSave() 
    Dim blnChange, i, sMsg
    
    blnChange = False
    
    FncSave = False                                                         
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
        
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    If lgBlnFlgChgValue = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If
	    

    If Verification = False Then Exit Function
    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
        
    FncSave = True                                                          
    
End Function

' ----------------------  검증 -------------------------
Function  Verification()

	Verification = False

	With frm1
		If UNICDbl(.txtData(C_W06_1).value) > 0 Then
			If UNICDbl(.txtData(C_W06_1).value) < 0 And  UNICDbl(.txtData(C_W07).value) > 0 Then
				Call DisplayMsgBox("W80001", "X", "(108)각 사업연도 소득금액", "(109)이월결손금")                          <%'No data changed!!%>
				Exit Function
			End If

			If UNICDbl(.txtData(C_W06_1).value) < UNICDbl(.txtData(C_W07).value) Then
				Call DisplayMsgBox("WC0010", "X", "(109)이월결손", "(108)각 사업연도 소득금액")                          <%'No data changed!!%>
				Exit Function
			End If
		End If
		
		If UNICDbl(.txtData(C_W18).value) < UNICDbl(.txtData(C_W19).value) Then
			Call DisplayMsgBox("WC0010", "X", "(123)공제감면세액(ㄴ)", "(122)차감세액")                          <%'No data changed!!%>
			Exit Function
		End If
						
	End With
	
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
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")
    Call InitVariables
    Call InitData

    Call SetToolbar("1100100000000111")

	frm1.txtCO_CD.focus

    FncNew = True

End Function


Function FncCopy() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

 	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 

End Function

Function FncInsertRow(ByVal pvRowCnt) 

End Function

Function FncDeleteRow() 

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
	
    If lgBlnFlgChgValue Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
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
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key   
        'strVal = strVal     & "&txtMaxRows="         & lgvspdData(lgCurrGrid).MaxRows            
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '☜:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	lgBlnFlgChgValue = False
	
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = parent.OPMD_UMODE
		    
	' 세무정보 조사 : 컨펌되면 락된다.
	Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	'1 컨펌체크 : 그리드 락 
	If wgConfirmFlg = "N" Then
		' 세율 코드 환경값을 디비의 값과 비교함 
		With frm1
			If .txtData(C_W14).value <>"" Then
				If .txtData(C_W14).options(.txtData(C_W14).selectedIndex).VIEW <> .txtW14.value & "%" Then
					Call DisplayMsgBox("WC0029", "X", "(117) 세 율", "W2014")                          <%'No data changed!!%>
					Exit Function
				End If
			ElseIf .txtData(C_W35).value <>"" Then
				If .txtData(C_W35).options(.txtData(C_W35).selectedIndex).VIEW <> .txtW35.value & "%" Then
					Call DisplayMsgBox("WC0029", "X", "(139) 세 율", "W2013")                          <%'No data changed!!%>
					Exit Function
				End If
			End If
		End With
		Call SetToolbar("11011000000000111")
	Else
		
		'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
		Call SetToolbar("11001000000000111")										<%'버튼 툴바 제어 %>
	End If
	
	'lgvspdData(lgCurrGrid).focus			
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow, lCol, lMaxRows, lMaxCols , i    
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
	With frm1
	
		For i = C_W01 To C_W53	
			strVal = strVal & .txtData(i).Value & Parent.gColSep
		Next 

	End With

	Frm1.txtSpread.value      =  strVal
	Frm1.txtMode.value        =  Parent.UID_M0002
	frm1.txtHeadMode.value	  =  lgIntFlgMode
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		
	Dim iRow											        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	
    Call MainQuery()
End Function

'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '☜: Query Key        
    strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '☜: Query Key            
	
	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function


'========================================================================================
Function DbDeleteOk()
	Call FncNew()
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
<SCRIPT LANGUAGE=javascript FOR=txtData EVENT=Change>
<!--
	try {
		if (this.noevent == null)
			SetHeadReCalc();
    } catch(e) {
    }
//-->
</SCRIPT>
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>

		<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:GetRef()">금액불러오기</A></TD>
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
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto"><% ' -- overflow=auto : 컨텐츠 구역을 브라우저 크기에 따라 스크롤바가 생성되게 한다 %>
						<TABLE <%=LR_SPACE_TYPE_60%> BORDER=0>
                            <TR>
                                <TD WIDTH="50%" VALIGN=TOP HEIGHT=100%>
								<TABLE <%=LR_SPACE_TYPE_20%> border="0" height=100% width="100%">
								   <TR>
										<TD>
											<TABLE <%=LR_SPACE_TYPE_20%> border="1" height=100% width="100%">
											 <TR>
												   <TD CLASS="TD51" width="5%" ROWSPAN=7 ALIGN=CENTER>(1)<br>각<br>사<br>업<br>연<br>도<br>소<br>득<br>계<br>산</TD>
												   <TD CLASS="TD51" width="60%"COLSPAN=2>(101) 결 산 서 상 당 기 순 손 익</TD>
												   <TD CLASS="TD61" width="5%" ALIGN=CENTER>01</TD>
												   <TD CLASS="TD61" width="30%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
											       <TD CLASS="TD51"  width="10%" ROWSPAN=2 ALIGN=CENTER>소 득 조 정 금 액</TD>
												   <TD CLASS="TD51">(102) 익 금 산 입</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>02</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>										   
												   <TD CLASS="TD51">(103) 손 금 산 입</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>03</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
											     <TD CLASS="TD51" COLSPAN=2>(104) 차 가 감 소 득 금 액<br>[(101) + (102) - (103)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>04</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
											     <TD CLASS="TD51" COLSPAN=2>(105) 기 부 금 한 도 초 과 액</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>05</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>									  
											<TR>
											     <TD CLASS="TD51" COLSPAN=2>(106) 기 부 금 한 도 초 과 액<br>이 월 액 손 금 산 입</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>54</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>	
											<TR>
											     <TD CLASS="TD51" COLSPAN=2>(107)각 사 업 연 도 소 득 금 액<br>[(104) + (105) - (106)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>06</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											</TABLE>
										</TD>
									</TR>
									<TR>
									     <TD HEIGHT=2></TD>
									</TR>
									<TR>
										<TD>
											<TABLE <%=LR_SPACE_TYPE_20%> border="1" height=100% width="100%">
											 <TR>
												   <TD CLASS="TD51" width="5%" ROWSPAN=5 ALIGN=CENTER>(2)<br>과<br>세<br>표<br>준<br>계<br>산</TD>
												   <TD CLASS="TD51" width="60%">(108)각 사업연도 소득금액<br>[(108)=(107)]</TD>
												   <TD CLASS="TD61" width="5%" ALIGN=CENTER>&nbsp;</TD>
												   <TD CLASS="TD61" width="30%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											 <TR>
												   <TD CLASS="TD51">(109)이 월 결 손 금</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>07</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>											
											 <TR>
												   <TD CLASS="TD51">(110)비 과 세 소 득</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>08</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>	
											 <TR>
												   <TD CLASS="TD51">(111)소 득 공 제</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>09</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>	
											 <TR>
												   <TD CLASS="TD51">(112)과 세 표 준<br>[(108) - (109) - (110) - (111)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>10</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>	
											</TABLE>
										</TD>
									</TR>
									<TR>
									     <TD HEIGHT=2></TD>
									</TR>
									<TR>
										<TD>
											<TABLE <%=LR_SPACE_TYPE_20%> border="1" height=100% width="100%">
											 <TR>
												   <TD CLASS="TD51" width="5%" ALIGN=CENTER>&nbsp;&nbsp;</TD>
												   <TD CLASS="TD51" width="60%">(159) 선 박 표 준 이 익</TD>
												   <TD CLASS="TD61" width="5%" ALIGN=CENTER>55</TD>
												   <TD CLASS="TD61" width="30%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW55_1" name=txtW55_1 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											</TABLE>
										</TD>
									</TR>
									<TR>
									     <TD HEIGHT=2></TD>
									</TR>
									<TR>
										<TD>
											<TABLE <%=LR_SPACE_TYPE_20%> border="1" height=100% width="100%">
											 <TR>
												   <TD CLASS="TD51" width="5%" ROWSPAN=7 ALIGN=CENTER>(3)<br>산<br>출<br>세<br>액<br>계<br>산</TD>
												   <TD CLASS="TD51" width="60%">(113) 과세 표준 금액2 [(112) + (159)]</TD>
												   <TD CLASS="TD61" width="5%" ALIGN=CENTER>56</TD>
												   <TD CLASS="TD61" width="30%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											 <TR>
												   <TD CLASS="TD51">(114) 세 율</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>11</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X6Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>											
											 <TR>
												   <TD CLASS="TD51">(115) 산 출 세 액</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>12</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>	
											 <TR>
												   <TD CLASS="TD51">(116) 지 점 유 보 소 득<br>(법 제 96조)</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>13</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>	
											 <TR>
												   <TD CLASS="TD51">(117) 세 율</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>14</TD>
												   <TD CLASS="TD61"><SELECT NAME=txtData STYLE="Width: 100%" tag="25X8Z" onChange="vbscript:SetHeadReCalc()"><OPTION VALUE="" VAL="0" VIEW=""></OPTION></SELECT></TD>
											</TR>	
											 <TR>
												   <TD CLASS="TD51">(118) 산 출 세 액</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>15</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											 <TR>
												   <TD CLASS="TD51">(119) 합 계 {(115)+(118)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>16</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											</TABLE>
										</TD>
									</TR>
									<TR>
									     <TD HEIGHT=2></TD>
									</TR>
									<TR>
										<TD>
											<TABLE <%=LR_SPACE_TYPE_20%> border="1" height=100% width="100%">
											 <TR>
												   <TD CLASS="TD51" width="5%" ROWSPAN=13 ALIGN=CENTER>(4)<br>납<br>부<br>할<br><br>세<br>액<br>계<br>산</TD>
												   <TD CLASS="TD51" width="60%" COLSPAN=3>(120)산 출 세 액 [(120) = (119)]</TD>
												   <TD CLASS="TD61" width="5%" ALIGN=CENTER>&nbsp;</TD>
												   <TD CLASS="TD61" width="30%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											 <TR>
												   <TD CLASS="TD51" width="60%" COLSPAN=3>(121)공 제 감 면 세 액(ㄱ)</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>17</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>											
											 <TR>
												   <TD CLASS="TD51" COLSPAN=3>(122)차 감 세 액</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>18</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>	
											 <TR>
												   <TD CLASS="TD51" COLSPAN=3>(123)공 제 감 면 세 액 (ㄴ)</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>19</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>	
											 <TR>
												   <TD CLASS="TD51" COLSPAN=3>(124)가 산 세 액</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>20</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>	
											 <TR>
												   <TD CLASS="TD51" COLSPAN=3>(125)가감계 [(122)-(123)+(124)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>21</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											 <TR>
												   <TD CLASS="TD51" width="5%" ROWSPAN=7 ALIGN=CENTER>기부납부세액</TD>
												   <TD CLASS="TD51" width="5%" ROWSPAN=5 ALIGN=CENTER>기한내 납부세액</TD>
												   <TD CLASS="TD51" width="50%">(126)중간 예납 세액</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>22</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											 <TR>
												   <TD CLASS="TD51">(127)수시 부과 세액</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>23</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											 <TR>
												   <TD CLASS="TD51">(128)원천 납부 세액</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>24</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											 <TR>
												   <TD CLASS="TD51">(129)간접투자회사등의 외국납부세액<INPUT type=hidden name=txtW25_NM STYLE="WIDTH: 50%" tag="25" maxlength=20></TD>
												   <TD CLASS="TD61" ALIGN=CENTER>25</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											 <TR>
												   <TD CLASS="TD51">(130)소 계</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>26</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											 <TR>
												   <TD CLASS="TD51" COLSPAN=2>(131)신고 납부전 가산세액</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>27</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>											
											 <TR>
												   <TD CLASS="TD51" COLSPAN=2>(132)합 계 [(130)+(131)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>28</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>												
											
											</TABLE>
										</TD>
									</TR>										
								  </TABLE>
								</TD>
                                <TD WIDTH="50%" VALIGN=TOP HEIGHT=100%>
								<TABLE <%=LR_SPACE_TYPE_20%> border="1" width="100%">
								   <TR>
										<TD>
											<TABLE <%=LR_SPACE_TYPE_20%> border="1" height=100% width="100%">
											 <TR>
												   <TD CLASS="TD51" width="5%" ROWSPAN=2></TD>
												   <TD CLASS="TD51" width="60%">(133) 감 면 분 추 가 납 부 세 액</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>29</TD>
												   <TD CLASS="TD61" width="30%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51">(134) 차 감 납 부 부 할 세 액<br>[(125) - (132) + (133)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>30</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											</TABLE>
										</TD>
									</TR>
									<TR>
									     <TD HEIGHT=2></TD>
									</TR>
								   <TR>
										<TD>
											<TABLE <%=LR_SPACE_TYPE_20%> border="1" width="100%">
											 <TR>
												   <TD CLASS="TD51" width="5%" ROWSPAN=15 ALIGN=CENTER>(5)<br>토<br>지<br>등<br><br>양<br>도<br>소<br>득<br>에<br>대<br>한<br><br>법<br>인<br>세<br><br>계<br>산</TD>
												   <TD CLASS="TD51" width="10%" ROWSPAN=2 ALIGN=CENTER VALIGN=CENTER>양 도<br>차 액</TD>
												   <TD CLASS="TD51" width="50%">(135) 등 기 자 산</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>31</TD>
												   <TD CLASS="TD61" width="30%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51">(136) 미 등 기 자 산</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>32</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(137) 비 과 세 소 득</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>33</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(138) 과세 표준 [(135) + (136) - (137)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>34</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(139) 세 율</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>35</TD>
												   <TD CLASS="TD61"><SELECT NAME=txtData STYLE="Width: 100%" tag="25X8Z" onChange="vbscript:SetHeadReCalc()"><OPTION VALUE="" VAL="0" VIEW=""></OPTION></SELECT</TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(140) 산 출 세 액</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>36</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100% AutoCalc="No"></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(141) 감 면 세 액</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>37</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(142) 차 감 세 액 [(140) - (141)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>38</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(143) 공 제 세 액</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>39</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(144) 가 산 세 액</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>40</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(145) 가감계 [(142) - (143) + (144)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>41</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" ROWSPAN=3 ALIGN=CENTER>기납부<br>세액</TD>
												   <TD CLASS="TD51">(146) 수 시 부 과 세 액</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>42</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51">(147) (<INPUT name=txtW43_NM STYLE="WIDTH: 50%" tag="25" maxlength=20>) 세 액</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>43</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51">(148) 계 [(143)+(147)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>44</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(149) 차 감 납 부 할 세 액<br>[(145) - (148)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>45</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											</TABLE>
										</TD>
									</TR>
									<TR>
									     <TD HEIGHT=2></TD>
									</TR>
								   <TR>
										<TD>
											<TABLE <%=LR_SPACE_TYPE_20%> border="1" width="100%">
											 <TR>
												   <TD CLASS="TD51" width="5%" ROWSPAN=9 ALIGN=CENTER>(6)<br>세<br>액<br>계</TD>
												   <TD CLASS="TD51" width="60%" COLSPAN=2>(150) 차 감 납 부 할 세 액 계<br>[(134) + (149)]</TD>
												   <TD CLASS="TD61" width="5%" ALIGN=CENTER>46</TD>
												   <TD CLASS="TD61" width="30%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(151) 사실과 다른 회계처리<br>경 정 세 액 공 제</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>57</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(152) 분 납 세 액 계 산 범 위 액<br>[(150) - (124) - (133) - (144) + (131) - (151)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>47</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" width="10%" ROWSPAN=3 ALIGN=CENTER>분납할<br>세액</TD>
												   <TD CLASS="TD51">(153) 현 금 납 부</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>48</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X22" width = 100% noevent></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51">(154) 물 납</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>49</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51">(155) 계 [(153) + (154)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>50</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" ROWSPAN=3 ALIGN=CENTER>차감<br>납부<br>세액</TD>
												   <TD CLASS="TD51">(156) 현 금 납 부</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>51</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X22" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51">(157) 물 납</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>52</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51">(158) 계 [(156) + (157)]<br>[(158) = (150) - (151) - (155)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>53</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>

											</TABLE>
										</TD>
									</TR>
									<TR>
									     <TD HEIGHT=2></TD>
									</TR>
								   <TR>
										<TD height=100>
											<TABLE <%=LR_SPACE_TYPE_20%> border="1" width="100%">
											 <TR>
												   <TD width="100%" HEIGHT=100%>&nbsp;</TD>
											</TR>
											</TABLE>
										</TD>
									</TR>
								</TABLE>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex=-1></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" style="display:'none'" tabindex=-1></TEXTAREA>
<INPUT TYPE=HIDDEN TABINDEX=-1 NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN TABINDEX=-1 NAME="txtHeadMode" tag="24">
<INPUT TYPE=HIDDEN TABINDEX=-1 NAME="txtW14" tag="24"><INPUT TYPE=HIDDEN TABINDEX=-1 NAME="txtW35" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<input type="hidden" name="uname" TABINDEX="-1">
	<input type="hidden" name="dbname" TABINDEX="-1">
	<input type="hidden" name="filename" TABINDEX="-1">
	<input type="hidden" name="strUrl" TABINDEX="-1">
	<input type="hidden" name="date" TABINDEX="-1">
</FORM>
</BODY>
</HTML>

