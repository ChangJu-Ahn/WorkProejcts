<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 법인세 조정 
'*  3. Program ID           : W8103MA1
'*  4. Program Name         : W8103MA1.asp
'*  5. Program Desc         : 제58호 법인세중간예납신고납부계산서 
'*  6. Modified date(First) : 2005/01/28
'*  7. Modified date(Last)  : 2006/01/27
'*  8. Modifier (First)     : 최영태 
'*  9. Modifier (Last)      : HJO 
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
<STYLE>
	.CHECKBOX {
		BORDER: 0;
	}

</STYLE>
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

Const BIZ_MNU_ID		= "W8103MA1"
Const BIZ_PGM_ID		= "W8103mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_REF_PGM_ID	= "W8103mb2.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID		= "W8109OA1"

' -- 그리드 컬럼 정의 

Dim C_W1	' 그리드 행(디비의 열)
Dim C_W2
Dim C_W3
Dim C_W4_1
Dim C_W4_2
Dim C_W5_1
Dim C_W5_2
Dim C_W6_1
Dim C_W6_2
Dim C_W7
Dim C_W8
Dim C_W9
Dim C_W10_1

Dim C_W01	
Dim C_W02	
Dim C_W03	
Dim C_W04	
Dim C_W05	
Dim C_W06	
Dim C_W07	
Dim C_W09	
Dim C_W10	
Dim C_W11	
Dim C_W12	
Dim C_W13	
Dim C_W14	
Dim C_W15	

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

Dim C_W51	

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 
Dim lgCurrGrid, lgvspdData(2)
Dim	lgFISC_START_DT, lgFISC_END_DT, lgMonGap, lgW2018

Dim IsRunEvents	' ㅠㅠ 무한이벤트반복을 막음 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	C_W1	= 0	
	C_W2	= 1
	C_W3	= 2
	C_W4_1	= 3
	C_W4_2	= 4
	C_W5_1	= 5
	C_W5_2	= 6
	C_W6_1	= 7
	C_W6_2	= 8
	C_W7	= 9
	C_W8	= 10
	C_W9	= 11
	C_W10_1	= 12
	
	C_W01	= 13
	C_W02	= 14
	C_W03	= 15
	C_W04	= 16
	C_W05	= 17
	C_W06	= 18
	C_W07	= 19
	C_W09	= 20
	C_W10	= 21
	C_W11	= 22
	C_W12	= 23
	C_W13	= 24
	C_W14	= 25
	C_W15	= 26

	C_W31	= 27
	C_W32	= 28
	C_W33	= 29
	C_W34	= 30
	C_W35	= 31
	C_W36	= 32
	C_W37	= 33
	C_W38	= 34
	C_W39	= 35
	C_W40	= 36
	C_W41	= 37

	C_W51	= 38
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
	Call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))

	Call InitSpreadPosVariables

 	call CommonQueryRs("REFERENCE_1"," ufn_TB_Configuration('W2018','" & C_REVISION_YM & "') ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
           
    lgW2018 = Split(lgF0 , chr(11))
    	
	IsRunEvents = True	' OBJECT 에 값넣을때 이벤트가 발생하는것을 막음 
	
	Call CommonQueryRs("MINOR_CD,MINOR_NM"," dbo.ufn_TB_MINOR('W1065', '" & C_REVISION_YM & "') ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtData(C_W1) ,lgF0  ,lgF1  ,Chr(11))
    
 	Call CommonQueryRs("MINOR_CD,MINOR_NM"," dbo.ufn_TB_MINOR('W1064', '" & C_REVISION_YM & "') ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtData(C_W2) ,lgF0  ,lgF1  ,Chr(11))
 
    IsRunEvents = False
    
    
End Sub

Sub InitSpreadSheet()

	Call AppendNumberPlace("6","3","2")
	Call AppendNumberPlace("7","2","0")
	
End Sub


'============================================  그리드 함수  ====================================

Sub InitData()

	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
    
	Call GetFISC_DATE

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
			
	IntRetCD = CommonQueryRs("W1, W2"," dbo.ufn_TB_58_GetRef('" & sCoCd & "','" & sFiscYear & "','" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD = True Then
		arrW1		= Split(lgF0, chr(11))
		arrW2		= Split(lgF1, chr(11))
		iMaxRows	= UBound(arrW1)

		For iRow = 0 To iMaxRows -1
			sTmp = "frm1.txtData(C_W" & arrW1(iRow) & ").Text = """ & CStr(arrW2(iRow)) & """"	
			Execute sTmp	' -- 변수에 들어 있는 명령을 실행한다.  ** ufn_TB_3_GetRef의 W1필드의 값을 코드로 하지 않고, 코드의 배열인덱스값으로 하면 이렇게 안해도 됨.
		Next

	End If
	
	Call SetHeadReCalc()
End Function

' 헤더 재계산 
Sub SetHeadReCalc()	
	Dim dblSum, dblW4(100)	

	If IsRunEvents Then Exit Sub	' 아래 .vlaue = 에서 이벤트가 발생해 재귀함수로 가는걸 막는다.
	
	IsRunEvents = True
	
	With frm1
		dblW4(C_W01) = UNICDbl(.txtData(C_W01).value)
		dblW4(C_W02) = UNICDbl(.txtData(C_W02).value)
		dblW4(C_W03) = UNICDbl(.txtData(C_W03).value)
		
		dblW4(C_W04) = dblW4(C_W01) - dblW4(C_W02) + dblW4(C_W03)
		.txtData(C_W04).value = dblW4(C_W04)	' (104) 확정세액 
		
		dblW4(C_W05) = UNICDbl(.txtData(C_W05).value)
		dblW4(C_W06) = UNICDbl(.txtData(C_W06).value)
		
		dblW4(C_W07) = dblW4(C_W04) - dblW4(C_W05) - dblW4(C_W06)
		.txtData(C_W07).value = dblW4(C_W07)	' (107) 차감세액 
		
		If dblW4(C_W07) < 0 Then
			dblW4(C_W09) = 0
		Else
			dblW4(C_W09) = dblW4(C_W07) * 6 / lgMonGap 
		End If
		.txtData(C_W09).value = dblW4(C_W09)	' (108) 중간예납세액 
			
		dblW4(C_W10) = UNICDbl(.txtData(C_W10).value)
		
		dblW4(C_W11) = dblW4(C_W09) - dblW4(C_W10) 
		.txtData(C_W11).value = dblW4(C_W11)	' (110) 차감중간예납세액 

		dblW4(C_W12) = UNICDbl(.txtData(C_W12).value)
		
		dblW4(C_W13) = dblW4(C_W11) + dblW4(C_W12) 
		.txtData(C_W13).value = dblW4(C_W13)	' (112) 가산세액 
		
		If dblW4(C_W11) <= 10000000 Then
			dblW4(C_W14) = 0
		ElseIf dblW4(C_W11) > 10000000 And dblW4(C_W11) <= 20000000 Then
			dblW4(C_W14) = dblW4(C_W11) - 10000000
		ElseIf dblW4(C_W11) > 20000000 Then
			dblW4(C_W14) = dblW4(C_W11) * 0.5
		End If
		.txtData(C_W14).value = dblW4(C_W14)	' (113) 분납세액		
		
		dblW4(C_W15) = dblW4(C_W13) - dblW4(C_W14) 
		.txtData(C_W15).value = dblW4(C_W15)	' (114) 납부세액 
		
		dblW4(C_W31) = UNICDbl(.txtData(C_W31).value)
		If dblW4(C_W31) < 0 Then
			dblW4(C_W32) = 0
		ElseIf (dblW4(C_W31) * 12 / 6) <= 100000000 Then
			dblW4(C_W32) = lgW2018(0) ' 이상세율 
		Else
			dblW4(C_W32) = lgW2018(1) ' 초과세율 
		End If
		.txtData(C_W32).Text = (dblW4(C_W32) * 100) & "%"	' (116) 세율 
		
		If  (dblW4(C_W31) * 12 / 6) > 100000000 Then
			dblW4(C_W33) = ((dblW4(C_W31) * 12 / 6) - 100000000) * (dblW4(C_W32) * 6/12) + ((100000000 * lgW2018(0)) * 6/12)
		ElseIf (dblW4(C_W31) * 12 / 6) <= 100000000 Then
			dblW4(C_W33) = ((dblW4(C_W31) * 12 / 6) * lgW2018(0)) * 12 / 6
		End If
		.txtData(C_W33).Text = dblW4(C_W33)	' (116) 세율 
		
		' 33코드 계산추가 
		'dblW4(C_W33) = UNICDbl(.txtData(C_W33).value)
		dblW4(C_W34) = UNICDbl(.txtData(C_W34).value)
		dblW4(C_W35) = UNICDbl(.txtData(C_W35).value)
		dblW4(C_W36) = UNICDbl(.txtData(C_W36).value)
		
		dblW4(C_W37) = dblW4(C_W33) - dblW4(C_W34) - dblW4(C_W35) - dblW4(C_W36) 
		.txtData(C_W37).value = dblW4(C_W37)	' (121) 중간예납세액 
		
		dblW4(C_W38) = UNICDbl(.txtData(C_W38).value)
		
		dblW4(C_W39) = dblW4(C_W37) + dblW4(C_W38) 
		.txtData(C_W39).value = dblW4(C_W39)	' (123) 납부할세액계 
		
		If dblW4(C_W39) <= 10000000 Then
			dblW4(C_W40) = 0
		ElseIf dblW4(C_W39) > 10000000 And dblW4(C_W39) <= 20000000 Then
			dblW4(C_W40) = dblW4(C_W39) - 10000000
		ElseIf dblW4(C_W39) > 20000000 Then
			dblW4(C_W40) = dblW4(C_W39) * 0.5
		End If
		.txtData(C_W40).value = dblW4(C_W40)	' (125) 납부세액 

		dblW4(C_W41) = dblW4(C_W39) - dblW4(C_W40) 
		.txtData(C_W41).value = dblW4(C_W41)	' (125) 납부할세액 
					
	End With

	lgBlnFlgChgValue= True ' 변경여부 
	IsRunEvents = False	' 이벤트 발생금지를 해제함 
End Sub

Sub GetFISC_DATE()	' 법인의 조회조건에 만족하는 당기시작,종료일을 가져온다.
	Dim sFiscYear, sRepType, sCoCd, ret, datFISC_START_DT, datFISC_END_DT, iRet, sRepTypeNm
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	sRepTypeNm	= frm1.cboREP_TYPE.options(frm1.cboREP_TYPE.selectedIndex).text
	
	iRet = CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If iRet = False Then
		Call DisplayMsgBox("WC0037", parent.VB_INFORMATION, sFiscYear & "년", sRepTypeNm)	
		frm1.cboREP_TYPE.value = "1"
		Exit Sub
	End If
	
	' 법인 기간은 필수입력 
	lgFISC_START_DT = CDate(lgF0)
	lgFISC_END_DT = CDate(lgF1)

	With frm1
		
	IsRunEvents = True

		.txtData(C_W5_1).Text	= lgFISC_START_DT
		.txtData(C_W5_2).Text	= lgFISC_END_DT
		'.txtData(C_W4).Text		= "6"
	
		' 직전사업연도월수 구함 
		sFiscYear	= UNIFormatDate(CDate(frm1.txtFISC_YEAR.text)-1)
	
		ret = CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		If ret Then
			datFISC_START_DT = CDate(lgF0)
			datFISC_END_DT = CDate(lgF1)
			lgMonGap = DateDiff("m", datFISC_START_DT, datFISC_END_DT)+1
		Else
			lgMonGap = 12
		End If

		.txtData(C_W7).Text	= lgMonGap
		
	IsRunEvents = False
	
	End With
	
End Sub

'====================================== 탭 함수 =========================================

'============================================  조회조건 함수  ====================================


'============================================  폼 함수  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         

    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100100000000111")										<%'버튼 툴바 제어 %>
	  
	' 변경한곳 
	Call InitComboBox
	
	IsRunEvents = True
	
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
	'Call ggoOper.FormatDate(frm1.txtW2 , parent.gDateFormat,3)
	
	Call InitData 
	'
     
    If frm1.cboREP_TYPE.value <> "2" Then
		Call DisplayMsgBox("W80004", parent.VB_INFORMATION, "", "X")  
		Call SetToolbar("1100000000000111")	
    End If
    
    IsRunEvents = False
	
	Call FncQuery
	   
    'Call ChangeCombo(2,False)
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

Sub ChkW3(Byval Index)	
	Dim i
	With frm1
		IsRunEvents = True

		For i = 0 To 2
			.txtW3(i).checked = False
		Next
		.txtW3(Index).checked = True
		.txtData(C_W3).value  = Index+1
		
		IsRunEvents = False
		lgBlnFlgChgValue = True
	End With
End Sub

Sub ChkW4_1(Byval Index)
	Dim i
	With frm1
		IsRunEvents = True
		For i = 0 To 4
			.txtW4_1(i).checked = False
		Next
		.txtW4_1(Index).checked = True
		.txtData(C_W4_1).value  = Index+1	
		
		IsRunEvents = False
		lgBlnFlgChgValue = True
	End With
End Sub

Sub ChkW10_1(Byval Index)
	Dim i
	With frm1
		IsRunEvents = True

		For i = 0 To 1
			.txtW10_1(i).checked = False
		Next
		.txtW10_1(Index).checked = True
		.txtData(C_W10_1).value  = Index+1
						
		IsRunEvents = False
		lgBlnFlgChgValue = True
	End With
End Sub

Sub ChangeCombo(strWhere,strTag)
	lgBlnFlgChgValue = strTag
	Dim i
	
	with frm1
		select case strWhere
			Case "2"
		
				If .txtData(1).value=1 then 
				
					For i = 13 To 26
						Select Case i
							Case  13,14,15,17,18,20,21,23,25
								Call ggoOper.setreqAttr(.txtData(i),"D") 'N,R								

						End Select
						
					Next
					For i=27 to 37
						Select Case i
							Case 27,30,31,32,34,36
								Call ggoOper.setreqAttr(.txtData(i),"Q")
						End Select
						.txtData(i).Text=""
					Next
				ElseIf .txtData(1).value=2 Then
				
					For i = 13 To 26
						Select Case i
							Case  13,14,15,17,18,20,21,23,25
								Call ggoOper.setreqAttr(.txtData(i),"Q") 'N,R
						End Select		
						.txtData(i).Text=""		
					Next
					For i=27 to 37
						Select Case i
							Case 27,30,31,32,34,36
								Call ggoOper.setreqAttr(.txtData(i),"D") 'N,R

						End Select 
					Next			
				End If
		End Select
	End with
	
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
    Call InitData                              
    															
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
	    
	If Not chkField(Document, "2") Then                             '⊙: Check contents area
	   Exit Function
	End If

	With frm1
		If Not .txtW3(0).checked And Not .txtW3(1).checked And Not .txtW3(2).checked Then
			Call DisplayMsgBox("WC0030", "X", "(3) 법인구분", "X")                          <%'No data changed!!%>
			Exit Function
		End If
		If Not .txtW4_1(0).checked And Not .txtW4_1(1).checked And Not .txtW4_1(2).checked And Not .txtW4_1(3).checked And Not .txtW4_1(4).checked Then
			Call DisplayMsgBox("WC0030", "X", "(4) 종류별 구분", "X")                          <%'No data changed!!%>
			Exit Function
		End If
		If Not .txtW10_1(0).checked And Not .txtW10_1(1).checked Then
			Call DisplayMsgBox("WC0030", "X", "(7) 신고납부 방법", "X")                          <%'No data changed!!%>
			Exit Function
		End If
	End With
	
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
    IsRunEvents = True
    
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")
    Call InitVariables
    Call InitData

    Call SetToolbar("1100100000000111")

	frm1.txtCO_CD.focus

	IsRunEvents = False
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
	
	With frm1
		.txtW3(UNICDbl(.txtData(C_W3).value)-1).checked = True
		.txtW4_1(UNICDbl(.txtData(C_W4_1).value)-1).checked = True
		.txtW10_1(UNICDbl(.txtData(C_W10_1).value)-1).checked = True
	End With
	
	' 세무정보 조사 : 컨펌되면 락된다.
	Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	'1 컨펌체크 : 그리드 락 
	If wgConfirmFlg = "N" Then
		Call SetToolbar("1101100000000111")			
	Else
		
		'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
		Call SetToolbar("1100000000000111")										<%'버튼 툴바 제어 %>
	End If
	Call ChangeCombo(2,False)
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
	
		For i = C_W1 To C_W51	
			Select Case i
				Case C_W4_1
					strVal = strVal & .txtData(i).Value & Parent.gColSep& .txtData(i).Value  & Parent.gColSep	
				Case C_W4_2						
				Case C_W5_1, C_W5_2, C_W6_1, C_W6_2, C_W8, C_W32
					strVal = strVal & .txtData(i).Text & Parent.gColSep				
				Case Else
					strVal = strVal & .txtData(i).Value & Parent.gColSep
			End Select
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
    SetHeadReCalc();
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
									<TD CLASS="TD6"><script language =javascript src='./js/w8103ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
						<TABLE <%=LR_SPACE_TYPE_60%> border="0" height=100% width="100%">
						   <TR>
								<TD>1. 중간 예납 기간</TD>
						   </TR>
						   <TR>
								<TD width="100%">
								<TABLE <%=LR_SPACE_TYPE_60%> border="1" height=100% width="100%">
									 <TR>
										   <TD CLASS="TD51" width=18%>(1) 종료별 구분</TD>
										   <TD CLASS="TD61" COLSPAN=2 width=20%><SELECT id="txtData" name=txtData STYLE="WIDTH: 200"  tag="23" onChange="vbscript:ChangeCombo 1,True"></SELECT></TD>
										   <TD CLASS="TD51" COLSPAN=2 width=20%>(2) 세액계산기준</TD>
										   <TD CLASS="TD61" COLSPAN=2 width=42%><SELECT id="txtData" name=txtData STYLE="WIDTH: 200"  tag="23" onChange="vbscript:ChangeCombo 2,True"></SELECT></TD>
									 </TR>
									 <TR>
										   <TD CLASS="TD51" ROWSPAN=2>(3) 법인구분</TD>
										   <TD CLASS="TD61" ROWSPAN=2 COLSPAN=2><INPUT TYPE=HIDDEN ID="txtData" name="txtData">
										   <TABLE>
											<TR>
												<TD ALIGN=CENTER><INPUT TYPE=CHECKBOX CLASS="CHECKBOX" NAME=txtW3 VALUE="1" ID=txtW3_1 BORDER=0 onclick="vbscript:ChkW3(0)"><LABEL FOR="txtW3_1">1. 내국</LABEL></TD>
												<TD ALIGN=CENTER><INPUT TYPE=CHECKBOX CLASS="CHECKBOX" NAME=txtW3 VALUE="2" ID=txtW3_2 BORDER=0 onclick="vbscript:ChkW3(1)"><LABEL FOR="txtW3_2">2. 외국</LABEL></TD>
												<TD ALIGN=CENTER><INPUT TYPE=CHECKBOX CLASS="CHECKBOX" NAME=txtW3 VALUE="3" ID=txtW3_3 BORDER=0 onclick="vbscript:ChkW3(2)"><LABEL FOR="txtW3_3">3. 외투</LABEL></TD>
											</TR>
										   </TABLE>
										   </TD>
										   <TD CLASS="TD51" ROWSPAN=2 COLSPAN=2>(4) 종류별 구분</TD>
										   <TD CLASS="TD61" ALIGN=CENTER width=15%>영리법인</TD>
										   <TD CLASS="TD61" ALIGN=CENTER width=27%>비영리법인</TD>
									</TR>
									<TR>
										   <TD CLASS="TD61"><INPUT TYPE=HIDDEN ID="txtData" name="txtData">
										   <TABLE <%=LR_SPACE_TYPE_20%>>
											<TR>
												<TD ALIGN=CENTER><INPUT TYPE=CHECKBOX CLASS="CHECKBOX" NAME=txtW4_1 VALUE="1" ID=txtW4_11 BORDER=0 onclick="vbscript:ChkW4_1(0)"><LABEL FOR="txtW4_11">1. 중소</LABEL></TD>
												<TD ALIGN=CENTER><INPUT TYPE=CHECKBOX CLASS="CHECKBOX" NAME=txtW4_1 VALUE="2" ID=txtW4_12 BORDER=0 onclick="vbscript:ChkW4_1(1)"><LABEL FOR="txtW4_12">2. 일반</LABEL></TD>
											</TR>
										   </TABLE>
										   </TD>
										   <TD CLASS="TD61"><INPUT TYPE=HIDDEN ID="txtData" name="txtData">
										   <TABLE <%=LR_SPACE_TYPE_20%>>
											<TR>
												<TD ALIGN=CENTER><INPUT TYPE=CHECKBOX CLASS="CHECKBOX" NAME=txtW4_1 VALUE="3" ID=txtW4_13 BORDER=0 onclick="vbscript:ChkW4_1(2)"><LABEL FOR="txtW4_13">3. 당기순이익</LABEL></TD>
												<TD ALIGN=CENTER><INPUT TYPE=CHECKBOX CLASS="CHECKBOX" NAME=txtW4_1 VALUE="4" ID=txtW4_14 BORDER=0 onclick="vbscript:ChkW4_1(3)"><LABEL FOR="txtW4_14">4. 중소</LABEL></TD>
												<TD ALIGN=CENTER><INPUT TYPE=CHECKBOX CLASS="CHECKBOX" NAME=txtW4_1 VALUE="5" ID=txtW4_15 BORDER=0 onclick="vbscript:ChkW4_1(4)"><LABEL FOR="txtW4_15">5. 일반</LABEL></TD>
											</TR>
										   </TABLE>
										   </TD>
									 </TR>
									 <TR>
										   <TD CLASS="TD51">(5) 사업연도</TD>
										   <TD CLASS="TD61" COLSPAN=2><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script> ~ <script language =javascript src='./js/w8103ma1_txtData_N515311688.js'></script></TD>
										   <TD CLASS="TD51" COLSPAN=2>(6) 예납기간</TD>
										   <TD CLASS="TD61" COLSPAN=2><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script> ~ <script language =javascript src='./js/w8103ma1_txtData_N287269365.js'></script></TD>
									 </TR>
									 <TR>
										   <TD CLASS="TD51" width=18%>(7) 직전사업연도월수</TD>
										   <TD CLASS="TD61" width=10%><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script> 개월</TD>
										   <TD CLASS="TD51" COLSPAN=2 width=15%>(8) 신고일</TD>
										   <TD CLASS="TD61" width=15%><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
										   <TD CLASS="TD51" width=20%>(9) 수입금액</TD>
										   <TD CLASS="TD61" width=42%><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									 </TR>
									 <TR>
										   <TD CLASS="TD51">(10) 신고납부 방법</TD>
										   <TD CLASS="TD61" COLSPAN=6><INPUT TYPE=HIDDEN ID="txtData" name="txtData">
										   <TABLE <%=LR_SPACE_TYPE_20%> >
											<TR>
												<TD ALIGN=CENTER><INPUT TYPE=CHECKBOX CLASS="CHECKBOX" NAME=txtW10_1 VALUE="1" ID=txtW10_11 BORDER=0 onclick="vbscript:ChkW10_1(0)"><LABEL FOR="txtW10_11">1. 정 기 신 고</LABEL></TD>
												<TD ALIGN=CENTER><INPUT TYPE=CHECKBOX CLASS="CHECKBOX" NAME=txtW10_1 VALUE="2" ID=txtW10_12 BORDER=0 onclick="vbscript:ChkW10_1(1)"><LABEL FOR="txtW10_12">2. 기 한 후 신 고</LABEL></TD>
											</TR>
										   </TABLE>
										   </TD>
									 </TR>
								</TABLE>
								</TD>
						   </TR>
						   <TR>
								<TD>2. 신고 및 납부 세액 계산</TD>
						   </TR>
						   <TR>
								<TD width="100%">
									<TABLE <%=LR_SPACE_TYPE_60%> border="1" height=100% width="100%">
									 <TR>
										   <TD CLASS="TD61" COLSPAN=3 ALIGN=CENTER>구 분</TD>
										   <TD CLASS="TD61" COLSPAN=2 ALIGN=CENTER>법 인 세</TD>
									 </TR>
									 <TR>
										   <TD CLASS="TD51" width="5%" ROWSPAN=14 ALIGN=CENTER>
										   <TABLE <%=LR_SPACE_TYPE_20%>>
											<TR>
												<TD ALIGN=CENTER>(1)<br><br>직<br>전<br>사<br>업<br>연<br>도<br>법<br>인<br>세<br>기<br>준</TD>
												<TD ALIGN=CENTER>법<br><br>제<br>6<br>3<br>호<br><br>제<br>1<br>항</TD>
											</TR>
										   </TABLE></TD>
										   <TD CLASS="TD51" width="5%" ROWSPAN=7 ALIGN=CENTER>
										   <TABLE <%=LR_SPACE_TYPE_20%>>
											<TR>
												<TD ALIGN=CENTER>직<br>전<br>사<br>업<br>연<br>도</TD>
												<TD ALIGN=CENTER>법<br><br><br>인<br><br><br><br>세</TD>
											</TR>
										   </TABLE></TD>
										   <TD CLASS="TD51" width="45%">(101) 산 출 세 액</TD>
										   <TD CLASS="TD61" width="5%" ALIGN=CENTER>01</TD>
										   <TD CLASS="TD61" width="30%"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
										   <TD CLASS="TD51">(102) 공 제 감 면 세 액</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>02</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>										   
										   <TD CLASS="TD51">(103) 가 산 세 액</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>03</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51">(104) 확 정 세 액 [(101) - (102) + (103)]</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>04</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51">(105) 수 시 부 과 세 액</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>05</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>									  
									<TR>
									       <TD CLASS="TD51">(106) 원 천 납 부 세 액</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>06</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>	
									<TR>
									       <TD CLASS="TD51">(107) 차 감 세 액 [(104) - (105) - (106)]</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>07</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>
									       <TABLE <%=LR_SPACE_TYPE_20%> border="0">
												<TR>
													<TD ALIGN=CENTER WIDTH=30%>(108) 중간예납세액</TD>
													<TD ALIGN=CETER WIDTH=50%>
													<TABLE <%=LR_SPACE_TYPE_20%> border="0">
														<TR>
															<TD ROWSPAN=3 WIDTH=30% ALIGN=RIGHT>[(107) X&nbsp;&nbsp;</TD>
															<TD ALIGN=CENTER>6</TD>
															<TD ROWSPAN=3 WIDTH=10%>]</TD>
														</TR>
														<TR>
															<TD HEIGHT=1 BGCOLOR=BLACK></TD>
														</TR>
														<TR>
															<TD ALIGN=CENTER>직전사업연도월수</TD>
														</TR>
													</TABLE>	
													</TD>
													<TD WIDTH=20%>&nbsp;</TD>
												</TR>
											</TABLE></TD>															
										   <TD CLASS="TD61" ALIGN=CENTER>09</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(109) 임 시 투 자 세 액 공 제</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>10</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(110) 차 감 중 간 예 납 세 액 [(108) - (109)]</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>11</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>											
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(111) 가 산 세 액 </TD>
										   <TD CLASS="TD61" ALIGN=CENTER>12</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>											
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(112) 납 부 할 세 액 계 [(110) + (111)]</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>13</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>											
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(113) 분 납 세 액</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>14</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>											
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(114) 납 부 세 액 [(112) - (113)]</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>15</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>											
									<TR>
									     <TD HEIGHT=5></TD>
									</TR>
									 <TR>
										   <TD CLASS="TD61" COLSPAN=3 ALIGN=CENTER>구 분</TD>
										   <TD CLASS="TD61" COLSPAN=2 ALIGN=CENTER>법 인 세</TD>
									 </TR>
									<TR>
										   <TD CLASS="TD51" width="5%" ROWSPAN=11 ALIGN=CENTER>
										   <TABLE <%=LR_SPACE_TYPE_60%>>
											<TR>
												<TD ALIGN=CENTER>(2)<br><br>자<br>기<br>계<br>산<br>기<br>준</TD>
												<TD ALIGN=CENTER>법<br><br>제<br>6<br>3<br>호<br><br>제<br>4<br>항</TD>
											</TR>
										   </TABLE></TD>
									       <TD CLASS="TD51" COLSPAN=2>(115) 과 세 표 준</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>31</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>												 											 
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(116)세 율</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>32</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>											 
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(117) 산 출 세 액</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>33</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(118) 공 제 감 면 세 액</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>34</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(119) 수 시 부 과 세 액</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>35</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(120) 원 천 납 부 세 액</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>36</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(121) 중 간 예 납 세 액 [(117) - (118) - (119) - (120)]</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>37</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(122) 가 산 세 액</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>38</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(123) 납 부 할 세 액 계 [(121) + (122)]</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>39</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(124)분 납 세 액</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>40</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(125) 납 부 세 액 [(123) - (124)]</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>41</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									     <TD HEIGHT=5></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=3>(126) 실 납 부 세 액 [(114) 또는 (125)중 납부한 세액]</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>51</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" style="display:'none'"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHeadMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtW14" tag="24"><INPUT TYPE=HIDDEN NAME="txtW35" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

