<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 법인세 조정 
'*  3. Program ID           : WB107MA1
'*  4. Program Name         : WB107MA1.asp
'*  5. Program Desc         : 제51호 중소기업 기준검토표 
'*  6. Modified date(First) : 2005/02/14
'*  7. Modified date(Last)  : 2005/02/14
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
<STYLE>
	.RADIO {
		BORDER: 0
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

Const BIZ_MNU_ID		= "WB107MA1"
Const BIZ_PGM_ID		= "WB107mb1.asp"											 '☆: 비지니스 로직 ASP명 
Const BIZ_REF_PGM_ID	= "WB107mb2.asp"											 '☆: 비지니스 로직 ASP명 
Const EBR_RPT_ID = "WB107OA1"

' -- 그리드 컬럼 정의 
Dim C_W01	
Dim C_W04	
Dim C_W07	
Dim C_W02	
Dim C_W05	
Dim C_W08
Dim C_W06	
Dim C_W09
Dim C_W_SUM	
Dim C_W19	
Dim C_W10	
Dim C_W11	
Dim C_W12
Dim C_W13	
Dim C_W14	
Dim C_W15	
Dim C_W16	
Dim C_W17	
Dim C_W20	
Dim C_W21
Dim C_W18
Dim C_W22
Dim C_W23

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2

Dim IsRunEvents	' ㅠㅠ 무한이벤트반복을 막음 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	C_W01		= 0	' HTML상의 순서 
	C_W04		= 1
	C_W07		= 2	
	C_W02		= 3
	C_W05		= 4
	C_W08		= 5
	C_W06		= 6
	C_W09		= 7
	C_W_SUM		= 8
	C_W19		= 9
	C_W10		= 11
	C_W11		= 12
	C_W12		= 13
	C_W13		= 14
	C_W14		= 15
	C_W15		= 16
	C_W16		= 17
	C_W17		= 18
	C_W20		= 19
	C_W21		= 20
	C_W18		= 21 
	C_W22		= 22
	C_W23		= 10
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

	Call AppendNumberPlace("6","5","1")
	Call AppendNumberPlace("7","5","0")
	Call AppendNumberPlace("8","4","0")
	
End Sub


'============================================  그리드 함수  ====================================

Sub InitData()
	With frm1
	
	.txtFISC_YEAR.text = "<%=wgFISC_YEAR%>"
    .txtCO_CD.value = "<%=wgCO_CD%>"
    .txtCO_NM.value = "<%=wgCO_NM%>"
    .cboREP_TYPE.value = "<%=wgREP_TYPE%>"
    
    .txtW19(0).checked = true
    .txtW20(0).checked = true
    .txtW21(0).checked = true
    .txtW22(0).checked = true
    .txtW23(0).checked = true
    
    Call GetCompanyInfo
    
    Call InitVariables
    End With
End Sub

Sub InitSpreadComboBox()

End Sub

'============================== 레퍼런스 함수  ========================================

Function GetRef()	' 금액가져오기 링크 클릭시 

	Call window.open("WB107MA2.txt", BIZ_MNU_ID, _
	"Width=600px,Height=450px,center= Yes,status=yes,toolbar=no,menubar=no,location=no,scrollbars=yes")

End Function

' 헤더 재계산 
Sub SetHeadReCalc()	
	Dim dblSum, dblW07, dblW08, dblW09
	
	If IsRunEvents Then Exit Sub	' 아래 .vlaue = 에서 이벤트가 발생해 재귀함수로 가는걸 막는다.
	
	IsRunEvents = True
	
	With frm1
		dblW07 = UNICDbl(.txtData(C_W07).value)
		dblW08 = UNICDbl(.txtData(C_W08).value)
		dblW09 = UNICDbl(.txtData(C_W09).value)
		dblSum = dblw07 + dblW08 + dblW09
		.txtData(C_W_SUM).value = dblSum
	End With

	lgBlnFlgChgValue= True ' 변경여부 
	IsRunEvents = False	' 이벤트 발생금지를 해제함 
End Sub

Function  OpenPopUp(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strCd, strCode
	
	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 1
			strCode = frm1.txtData(C_W04).value
		Case 2
			strCode = frm1.txtData(C_W05).value
	End Select
	
	arrParam(0) = "표준소득율"								' 팝업 명칭 
	arrParam(1) = "tb_std_income_rate" 								' TABLE 명칭 
	arrParam(2) = Trim(strCode)										' Code Condition
	arrParam(3) = ""												' Name Cindition
	arrParam(4) = ""										' Where Condition
	arrParam(5) = "표준소득율"									' 조건필드의 라벨 명칭 
            
	arrField(0) = "STD_INCM_RT_CD"									' Field명(0)
	arrField(1) = "BUSNSECT_NM"									' Field명(1)
	arrField(2) = "DETAIL_NM"									' Field명(1)
	arrField(3) = "FULL_DETAIL_NM"									' Field명(1)
			
	arrHeader(0) = " 번호"									' Header명(0)
	arrHeader(1) = "업태"									' Header명(1)
	arrHeader(2) = "업종"									' Header명(1)
	arrHeader(3) = "업종상세"									' Header명(1)
	
	IsOpenPop = True
			
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=750px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then	    
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If
End Function

Function SetPopup(Byval arrRet,Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1
				.txtData(C_W04).value = arrRet(0)
			Case 2
				.txtData(C_W05).value = arrRet(0)
		End Select
	End With
	
	lgBlnFlgChgValue = True

End Function

Sub QueryRadio()
	' -- 쿼리후 히든의 값을 참조로 라디오버튼을 재처리한다.
	With frm1
		If .txtData(C_W19).value = "1" Then
			.txtW19(0).checked = true
		Else
			.txtW19(1).checked = true
		End If
		
		If .txtData(C_W20).value = "1" Then
			.txtW20(0).checked = true
		Else
			.txtW20(1).checked = true
		End If
		
		If .txtData(C_W21).value = "1" Then
			.txtW21(0).checked = true
		Else
			.txtW21(1).checked = true
		End If
		
		If .txtData(C_W22).value = "1" Then
			.txtW22(0).checked = true
		Else
			.txtW22(1).checked = true
		End If
		
		If .txtData(C_W23).value = "1" Then
			.txtW23(0).checked = true
		Else
			.txtW23(1).checked = true
		End If
		
	End With
End Sub

Sub SaveRadio()
	' -- 세이브시 현재 선택된 라디오버튼을 히든으로 처리한다.
	With frm1
		If .txtW19(0).checked = true Then
			.txtData(C_W19).value = "1"
		Else
			.txtData(C_W19).value = "2"
		End If
		
		If .txtW20(0).checked = true Then
			.txtData(C_W20).value = "1"
		Else
			.txtData(C_W20).value = "2"
		End If
		
		If .txtW21(0).checked = true Then
			.txtData(C_W21).value = "1"
		Else
			.txtData(C_W21).value = "2"
		End If
		
		If .txtW22(0).checked = true Then
			.txtData(C_W22).value = "1"
		Else
			.txtData(C_W22).value = "2"
		End If
		
		.txtW23(0).checked = false
		.txtW23(1).checked = false
			
		' -- 적정여부 
		If .txtData(C_W19).value = "1" And .txtData(C_W20).value = "1" And _
			.txtData(C_W21).value = "1" And .txtData(C_W22).value = "1" Then
			.txtW23(0).checked = true
			.txtData(C_W23).value = "1"
		ElseIf .txtData(C_W20).value = "2" And .txtData(C_W22).value = "1" Then
			.txtW23(0).checked = true
			.txtData(C_W23).value = "1"
		Else
			.txtW23(1).checked = true
			.txtData(C_W23).value = "2"
		End If
	
	End With
End Sub

Sub GetCompanyInfo()	' 요청법인의 조회조건에 만족하는 당기시작,종료일을 가져온다.
	Dim sFiscYear, sRepType, sCoCd, sFISC_START_DT, sFISC_END_DT, datMonCnt, i, datNow
	
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	call CommonQueryRs("IND_CLASS, HOME_TAX_MAIN_IND"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    With frm1
		IsRunEvents = True
		If lgF0 <> "" Then 
			.txtData(C_W01).value = Replace(lgF0, chr(11), "")
			.txtData(C_W04).value = Replace(lgF1, chr(11), "")
		Else
			.txtData(C_W01).value = ""
			.txtData(C_W04).value = ""
		End If
		IsRunEvents = False
	End With
End Sub

Sub RadioClicked()
	lgBlnFlgChgValue = True
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
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
    Call ggoOper.FormatDate(frm1.txtData(C_W18), parent.gDateFormat,3)
	'Call ggoOper.FormatDate(frm1.txtW2 , parent.gDateFormat,3)
	
	Call InitData 
	'
    Call fncQuery() 
    
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
	Call GetCompanyInfo
End Sub

Sub txtFISC_YEAR_Change()
	Call GetCompanyInfo
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

    Call SetToolbar("1111100000000011")

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
	' 세무정보 조사 : 컨펌되면 락된다.
	Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	'1 컨펌체크 : 그리드 락 
	If wgConfirmFlg = "N" Then
		lgIntFlgMode = parent.OPMD_UMODE
	
		' 세율 코드 환경값을 디비의 값과 비교함 
		With frm1

		End With
		Call SetToolbar("1101100000000111")										<%'버튼 툴바 제어 %>
	Else
		
		'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
		Call SetToolbar("1100000000000111")										<%'버튼 툴바 제어 %>
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
    
    Call SaveRadio	' -- 라디오버튼 처리 
    
	With frm1
	
		For i = C_W01 To C_W22	
			If i = C_W18 Then
				strVal = strVal & .txtData(i).text & Parent.gColSep
			Else
				strVal = strVal & .txtData(i).value & Parent.gColSep
			End If
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
	if (this.WithEvent == "1") {
		SetHeadReCalc();
	} else if (this.WithEvent == "2") {
		RadioClicked();
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
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;
						<a href="vbscript:GetRef">중소기업기본법시행령 별표1</A>  
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
									<TD CLASS="TD6"><script language =javascript src='./js/wb107ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
						<TABLE <%=LR_SPACE_TYPE_20%> border="0" height=100% width="100%">
						   <TR>
								<TD>
									<TABLE <%=LR_SPACE_TYPE_20%> border="1" height=100% width="100%">
									 <TR HEIGHT=25>
										   <TD CLASS="TD51" width="10%" COLSPAN=2>(1) 요 건</TD>
										   <TD CLASS="TD51" width="52%">(2) 검 토 내 용</TD>
										   <TD CLASS="TD51" width="8%">(3)적합여부</TD>
										   <TD CLASS="TD51" width="8%">(4)적정여부</TD>
									</TR>
									<TR>
									       <TD CLASS="TD61" width="10%" COLSPAN=2 ALIGN=CENTER title = "제조업(조세특례제한법시행규칙제2조제1항의 의제제조업 포함), 광업, 건설업, 엔지니어링사업, 물류사업, 해운업에 의한선박관리업, 운수업중 여객운송업, 어업, 도매업, 소매업, 전기통신업, 연구 및 개발업, 방송업, 정보처리 및 기타컴퓨터운영관련업, 자동차정비업, 의료업, 폐기물처리업, 폐수처리업, 분뇨등관련영업, 작물재배업, 축산업, 과학및기술서비스업, 포장및충전업, 영화산업, 공연산업, 전문디자인업, 뉴스제공업, 광고업, 무역전시산업, 직업기술분야학원, 관광사업(카지노, 관광유흥음식점업 및 외국인 전용유흥음식점을 제외), 노인복지법에 의한 노인복지시설운영업, 토양정화업" >(101)<br> 해 당<br>사 업</TD>
										   <TD CLASS="TD61">
										   <TABLE <%=LR_SPACE_TYPE_20%> border="1" height=90% width="90%">
											<TR>
												<TD CLASS="TD51" width="30%">업태별&nbsp;&nbsp;&nbsp;&nbsp;＼&nbsp;&nbsp;&nbsp;&nbsp;구분</TD>
												<TD CLASS="TD51" width="30%">기준경비율코드</TD>
												<TD CLASS="TD51" width="40%">사업수입금액</TD>
											</TR>
											<TR>
												<TD CLASS="TD61">(01) (<INPUT TYPE=text id="txtData" name=txtData size=15 maxlength=10 tag="25X" WithEvent="2">)업</TD>
												<TD CLASS="TD61">(04) <INPUT TYPE=text id="txtData" name=txtData size=10 maxlength=6 tag="25X"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopUp(1)"></TD>
												<TD CLASS="TD61">(07) <script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script></TD>
											</TR>
											<TR>
												<TD CLASS="TD61">(02) (<INPUT TYPE=text id="txtData" name=txtData size=15 maxlength=10 tag="25X" WithEvent="2">)업</TD>
												<TD CLASS="TD61">(05) <INPUT TYPE=text id="txtData" name=txtData size=10 maxlength=6 tag="25X"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopUp(2)"></TD>
												<TD CLASS="TD61">(08) <script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script></TD>
											</TR>
											<TR>
												<TD CLASS="TD61">(03) 기타사업</TD>
												<TD CLASS="TD61">(06) <INPUT TYPE=text id="txtData" name=txtData size=10 maxlength=6 tag="25X" WithEvent="2"></TD>
												<TD CLASS="TD61">(09) <script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script></TD>
											</TR>
											<TR>
												<TD CLASS="TD61" ALIGN=CENTER>계</TD>
												<TD CLASS="TD61">&nbsp;</TD>
												<TD CLASS="TD61"><script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script></TD>
											</TR>
										   </TABLE>
										   </TD>
										   <TD CLASS="TD61" ALIGN=CENTER VALIGN=MIDDLE>(19)<br> <INPUT TYPE=RADIO NAME=txtW19 ID=txtW19 tag="21" CLASS="RADIO" onclick="RadioClicked()"><br>적합<br>(Y)<br><br><INPUT TYPE=RADIO NAME=txtW19 ID=txtW19 tag="21" CLASS="RADIO" onclick="RadioClicked()"><br>부적합<br>(N)
										   <INPUT TYPE=HIDDEN NAME="txtData" name=txtData tag="24"></TD>
										   <TD CLASS="TD61" ALIGN=CENTER VALIGN=MIDDLE ROWSPAN=4>(23)<br> <INPUT TYPE=RADIO NAME=txtW23 ID=txtW23 tag="21" CLASS="RADIO" onclick="RadioClicked()"><br>적<br>(Y)<br><br><br><INPUT TYPE=RADIO NAME=txtW23 ID=txtW23 tag="21" CLASS="RADIO" onclick="RadioClicked()"><br>부<br>(N)
										   <INPUT TYPE=HIDDEN NAME="txtData" name=txtData tag="23"></TD>
									</TR>
									<TR>										   
										   <TD CLASS="TD61" ALIGN=CENTER COLSPAN=2 title = "○ 아래 요건 ①,②를 동시에 충족할 것                  ① 상시 사용 종업원수ㆍ자본금ㆍ매출액 중 하나 이상이 중소기업기본법시행령 별표1의 규모기준 이내일것 ②졸업제도(상시종업원수 1천명 미만, 자기자본 1천억 미만, 매출액 1천억 미만)이내일 것" >(102) 종업원수<br>·자본금<br>·매출액<br>·자기자본·자산<br>기준 </TD>
										   <TD CLASS="TD61">
										   <TABLE <%=LR_SPACE_TYPE_20%> border="0" height=100% width="100%">
											<TR>
												<TD CLASS="TD61" COLSPAN=4>&nbsp;가. 상시 종업원수(연평균인원)</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" COLSPAN=2 width=30%>(1) 당 회사(10)</TD>
												<TD CLASS="TD61" width=*>(<script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script>명)</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" COLSPAN=3>(2) 중소기업기본법시행령 별표1의</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" >규모기준(11)</TD>
												<TD CLASS="TD61" >(<script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script>명)미만</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" COLSPAN=4>&nbsp;나. 자본금</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" COLSPAN=2>(1) 당 회사(12)</TD>
												<TD CLASS="TD61">(<script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script>억)</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" COLSPAN=3>(2) 중소기업기본법시행령 별표1의</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" >규모기준(13)</TD>
												<TD CLASS="TD61" >(<script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script>억)이하</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" COLSPAN=4>&nbsp;다. 매출액</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" COLSPAN=2>(1) 당 회사(14)</TD>
												<TD CLASS="TD61">(<script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script>억)</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" COLSPAN=3>(2) 중소기업기본법시행령 별표1의</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" width=10>&nbsp;</TD>
												<TD CLASS="TD61" >규모기준(15)</TD>
												<TD CLASS="TD61" >(<script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script>억)이하</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" COLSPAN=3>&nbsp;라. 자기자본(16)</TD>
												<TD CLASS="TD61" >(<script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script>억)</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" COLSPAN=4>&nbsp;마. 상장.협회등록법인의 경우</TD>
											</TR>
											<TR>
												<TD CLASS="TD61" width=10></TD>
												<TD CLASS="TD61" width=10></TD>
												<TD CLASS="TD61" >규모기준(17)</TD>
												<TD CLASS="TD61" >(<script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script>억)</TD>
											</TR>
										   </TABLE>
										   </TD>
										   <TD CLASS="TD61" ALIGN=CENTER VALIGN=MIDDLE>(20)<br> <INPUT TYPE=RADIO NAME=txtW20 ID=txtW20 tag="21" CLASS="RADIO" onclick="RadioClicked()"><br>적합<br>(Y)<br><br><INPUT TYPE=RADIO NAME=txtW20 ID=txtW20 tag="21" CLASS="RADIO" onclick="RadioClicked()"><br>부적합<br>(N)
										   <INPUT TYPE=HIDDEN NAME="txtData" name=txtData tag="24"></TD>
											</TR>
											<TR>
												<TD CLASS="TD61"  ALIGN=CENTER COLSPAN=2 title = "중소기업기본법시행령 제3조 제2호의 기준에 의한 독립성 충족">(103) 소유·<br>경영의<br>독립성  </TD>
												<TD CLASS="TD61">&nbsp;○자산총액 5,000억원 이상인 법인이 발행주식의 30%이상 소유하고 있는 법인이 아닐 것<BR>
												&nbsp;○독점규제 및 공정거래에 관한 법률에 의한 상호출자제한기업집단에 속하지 않을것</TD>
												<TD CLASS="TD61" ALIGN=CENTER VALIGN=MIDDLE>(21)<INPUT TYPE=RADIO NAME=txtW21 ID=txtW21 tag="21" CLASS="RADIO" onclick="RadioClicked()">적합(Y)<br><INPUT TYPE=RADIO NAME=txtW21 ID=txtW21 tag="21" CLASS="RADIO" onclick="RadioClicked()">부적합(N)
												<INPUT TYPE=HIDDEN NAME="txtData" name=txtData tag="24"></TD>
											</TR>
											<TR>
												<TD CLASS="TD61"  ALIGN=CENTER COLSPAN=2 title = "01.1이후 개시연도 중 중소기업 기준을 규모가 초과하는 경우당해 초과연도와 그 후 3년간 중소기업으로 보고 그 후에는 매년마다 판단">(104) 유예기간</TD>
												<TD CLASS="TD61">&nbsp;○ 초과연도(18) (<script language =javascript src='./js/wb107ma1_txtData_txtData.js'></script>)년 
												&nbsp;&nbsp;* 2001이후</TD>
												<TD CLASS="TD61" ALIGN=CENTER VALIGN=MIDDLE>(22)<INPUT TYPE=RADIO NAME=txtW22 ID=txtW22 tag="21" CLASS="RADIO" onclick="RadioClicked()">적합(Y)<br><INPUT TYPE=RADIO NAME=txtW22 ID=txtW22 tag="21" CLASS="RADIO" onclick="RadioClicked()">부적합(N)
												<INPUT TYPE=HIDDEN NAME="txtData" name=txtData tag="24"></TD>
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
</FORM>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<input type="hidden" name="uname" TABINDEX="-1">
	<input type="hidden" name="dbname" TABINDEX="-1">
	<input type="hidden" name="filename" TABINDEX="-1">
	<input type="hidden" name="strUrl" TABINDEX="-1">
	<input type="hidden" name="date" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

