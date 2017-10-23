<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Exchange reference 
'*  3. Program ID           : A5954RA1
'*  4. Program Name         : 환율참조팝업 
'*  5. Program Desc         : Popup of Exchange
'*  6. Component List       : DB agent
'*  7. Modified date(First) : 2002.05.06
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : Jang Yoon Ki
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs">					</SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "AcctCtrl.vbs">							</SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

'****************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************

Const BIZ_PGM_ID 		= "a5954rb1.asp"                              '☆: Biz Logic ASP Name
Const STD_PGM_ID		= "F5954RA101"
Const MOV_PGM_ID		= "F5954RA102"
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

Const C_MaxKey          = 3                                           '☆: key count of SpreadSheet


'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop 
Dim lgMaxFieldCount 
Dim lgCookValue 
Dim IsOpenPop  
Dim lgSaveRow 
Dim CPGM_ID
Dim arrReturn
Dim arrParent
Dim arrParam
DIm txtStdDt		'결산 년월					
DIm txtStdYYMM		'결산 일자 
Dim ChcMnDt			'변동/고정환율 선택FG
	
'------ Set Parameters from Parent ASP -----------------------------------------------------------------------
arrParent		= window.dialogArguments
Set PopupParent = arrParent(0)
arrParam		= arrParent(1)

'txtStdYYMM = PopupParent.frm1.fpdtWk_yymm.Text
'txtStdDt = PopupParent.frm1.fpdtWk_yymmdd.Text

txtStdYYMM	= "<%=Request("txtStdYYMM")	%>"		'결산 년월 
	
txtStdDt	= "<%=Request("txtStdDt")	%>"		'결산 일자 
	
'top.document.title = "환율 참조 팜업"
top.document.title = PopupParent.gActivePRAspName
	
ChcMnDt = ""

'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------


 '==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    
    lgStrPrevKey     = ""
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1
    lgSaveRow        = 0

	Redim arrReturn(0,0)
	Self.Returnvalue = arrReturn
	
End Sub


Sub SetDefaultVal()
'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------

	Dim strYear, strMonth, strDay
	Dim strYearYM, strMonthYM, strDayYM
	 
	
	Call ExtractDateFrom(UNIConvDate(txtStdDt),PopupParent.gServerDateFormat,PopupParent.gServerDateType,strYear,strMonth,strDay)	
	
	frm1.txtStdDt.Year = strYear
	frm1.txtStdDt.Month = strMonth
	frm1.txtStdDt.Day = strDay
	
		
	Call ExtractDateFrom(txtStdYYMM,PopupParent.gDateFormatYYYYMM,PopupParent.gComDateType,strYearYM,strMonthYM,strDayYM)	
	
	frm1.txtStdYYMM.Year = strYearYM
	frm1.txtStdYYMM.Month = strMonthYM
	
	Call ggoOper.FormatDate(frm1.txtStdDt, PopupParent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtStdYYMM, PopupParent.gDateFormat, 2)

'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------

End Sub

'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call LoadInfTB19029A("Q", "A","NOCOOKIE","RA") %>
End Sub


'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  이 부분에서 컬럼 추가하고 데이타 전송이 일어나야 합니다.   							=
'========================================================================================================
	
Function OKClick()
		
	Dim intColCnt, intRowCnt, intInsRow
		
	if frm1.vspdData.ActiveRow > 0 Then 			
		
		intInsRow = 0
		with frm1
		
			Redim arrReturn(.vspdData.SelModeSelCount - 1, .vspdData.MaxCols - 1)
			
			For intRowCnt = 0 To .vspdData.MaxRows - 1
					
				.vspdData.Row = intRowCnt + 1

				If .vspdData.SelModeSelected Then
					
					.vspdData.Col = GetKeyPos("A",1)
					arrReturn(intInsRow, GetKeyPos("A",1) -1) = .vspdData.Text
					.vspdData.Col = GetKeyPos("A",2)
					arrReturn(intInsRow, GetKeyPos("A",2) -1) = .vspdData.Text
					.vspdData.Col = GetKeyPos("A",3)
					arrReturn(intInsRow, GetKeyPos("A",3) -1) = .vspdData.Text

					intInsRow = intInsRow + 1

				End IF
			Next
		End With
	End if			
		
	Self.Returnvalue = arrReturn
	Self.Close()
	
					
End Function

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================

Function CancelClick()
	Self.Close()			
End Function

'=========================================  2.3.3 Mouse Pointer 처리 함수 ===============================
'========================================================================================================

Function MousePointer(pstr1)
	Select case UCase(pstr1)
	case "PON"
		window.document.search.style.cursor = "wait"
	case "POFF"
		window.document.search.style.cursor = ""
	End Select
End Function


'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
    frm1.vspdData.OperationMode = 5
    Call SetZAdoSpreadSheet("F5954RA101", "S", "A", "V20021108", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X")
    Call SetSpreadLock()      
End Sub

Sub InitSpreadSheet1()    
    frm1.vspdData.OperationMode = 5
    Call SetZAdoSpreadSheet("F5954RA102", "S", "A", "V20021108", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X")
    Call SetSpreadLock()
End Sub


'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
	 ggoSpread.SpreadLockWithOddEvenRowColor()
    .vspdData.ReDraw = True

    End With
End Sub

'===========================================================================
' Function Name : OpenSortPopup
' Function Desc : OpenOrderBy Reference Popup
'===========================================================================
Function OpenSortPopup()
   
   	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True	
	
	If ChcMnDt = "Std" Then
			arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")
		ElseIf ChcMnDt = "Mov" Then
			arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")
		End If
		
	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
		Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))		
		Call InitVariables
		If ChcMnDt = "Std" Then
			Call InitSpreadSheet()			
		ElseIf ChcMnDt = "Mov" Then
			Call InitSpreadSheet1()			
		End If
		       
	End If

End Function

 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029														'⊙: Load table , B_numeric_format	
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	if frm1.Rb_Fg1.checked = True then
		
		txtStdDtTit.innerHTML = ""
		txtStdYYMMTit.innerHTML = "적용기준년월"
		
		Call ElementVisible(frm1.txtStdDt, 0)
		
		Call ElementVisible(frm1.txtStdYYMM, 1)
	

	end if
    
	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	ChcMnDt = "Std"	
	Call FncQuery()

'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
'	Call ElementVisible(frm1.txtDummy, 0)	'InVisible
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
End Sub

Sub Form_Load1()
    Call LoadInfTB19029		
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                  
	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet1()	
	ChcMnDt = "Mov"
	Call FncQuery()
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
'	Call ElementVisible(frm1.txtDummy, 0)	'InVisible
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub


'==========================================================================================
'   Event Name : DblClick
'   Event Desc :
'==========================================================================================
Sub txtStdDt_DblClick(Button)
	if Button = 1 then
		frm1.txtStdDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtStdDt.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : KeyPress
'   Event Desc :
'==========================================================================================

Sub txtStdDt_KeyPress(KeyAscii)
	On Error Resume Next

	If KeyAscii = 13 Then 
		Call DbQuery
	ElseIf KeyAscii = 27 Then
		Call CancelClick
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

   	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
    	If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End if
    	End If
    End If
    
     '----------  Coding part  -------------------------------------------------------------   
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Dim ii
    gMouseClickStatus = "SPC"   
    
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
End Sub


'======================================================================================================
'   Event Name : vspdData_KeyPress
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function


'======================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.MaxRows > 0 Then
		If Frm1.vspdData.ActiveRow = Row Or Frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub

 '*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* 
Function FncQuery() 
Dim IntRetCD
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If
	
    '-----------------------
    'Query function call area
    '-----------------------
    
	If DbQuery = False Then Exit Function	

    FncQuery = True		
End Function


 '*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 

	Dim strVal
	Dim txtStdYYMMYear, txtStdYYMMMonth
	
    DbQuery = False
    
    Err.Clear     
          
	Call LayerShowHide(1)
    
    
    With frm1
	
	txtStdYYMMYear = .txtStdYYMM.year
    txtStdYYMMMonth = .txtStdYYMM.Month
    
    If Len(txtStdYYMMMonth) = 1 Then
		txtStdYYMMMonth = "0" & txtStdYYMMMonth
	End if
    		
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------
	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		If ChcMnDt = "Std" Then
			strVal = BIZ_PGM_ID & "?txtStdYYMM=" & Trim(.hStdYYMM.Text)
		ElseIf ChcMnDt = "Mov" Then
			strVal = BIZ_PGM_ID & "?txtStdDt=" & Trim(.hStdDt.Text)
		End If
	Else 
		If ChcMnDt = "Std" Then
			strVal = BIZ_PGM_ID & "?txtStdYYMM=" & Trim(txtStdYYMMYear & txtStdYYMMMonth)
		ElseIf ChcMnDt = "Mov" Then
			strVal = BIZ_PGM_ID & "?txtStdDt=" & Trim(.txtStdDt.Text)
		End If
		
	End If
	
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------]
	If frm1.Rb_Fg1.checked = True Then	 '차입번호별 		
			strVal = strVal & "&txtPgmId="		 & STD_PGM_ID			
	ElseIf frm1.Rb_Fg2.checked = True Then	 '지급일자별 
			strVal = strVal & "&txtPgmId="		 & MOV_PGM_ID			
	End If	
        strVal = strVal & "&lgPageNo="       & lgPageNo         
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")         
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
	        			
       Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
   
    End With
	
    DbQuery = True
        

End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk()														'☆: 조회 성공후 실행로직 

    '-----------------------
    'Reset variables area
    '-----------------------
    lgBlnFlgChgValue = False                                                 'Indicates that no value changed
	lgIntFlgMode = PopupParent.OPMD_UMODE
	lgSaveRow        = 1
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
		frm1.vspdData.Row = 1
		frm1.vspdData.SelModeSelected = true
				
	End If
	
End Function



Function Rb_Fg1_OnClick() 
	if frm1.Rb_Fg1.checked = True then
		
		txtStdDtTit.innerHTML = ""
		txtStdYYMMTit.innerHTML = "적용기준년월"
		
		Call ElementVisible(frm1.txtStdDt, 0)
		
		Call ElementVisible(frm1.txtStdYYMM, 1)
	

	end if
End Function

Function Rb_Fg2_OnClick() 

	if frm1.Rb_Fg2.checked = True then
		
		txtStdDtTit.innerHTML = "적용기준일자"
		txtStdYYMMTit.innerHTML = ""
		
		Call ElementVisible(frm1.txtStdDt, 1)

		Call ElementVisible(frm1.txtStdYYMM, 0)
	

	end if
	
	call Form_load1() 
End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5 NOWRAP>환율</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Fg ID=Rb_Fg1 VALUE="Y" Checked Onclick="Form_load()"><LABEL FOR=Rb_Fg1>고정환율</LABEL>&nbsp;</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Fg ID=Rb_Fg2 VALUE="N"><LABEL FOR=Rb_Fg2>변동환율</LABEL>&nbsp;</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 ID="txtStdDtTit" NOWRAP>적용기준일자</TD>
						<TD CLASS=TD6 NOWRAP Colspan=2><script language =javascript src='./js/a5954ra1_txtStdDt_txtStdDt.js'></script></TD>						
					</TR>
					<TR>
						<TD CLASS=TD5 ID="txtStdYYMMTit" NOWRAP>적용기준년월</TD>
						<TD CLASS=TD6 NOWRAP Colspan=2><script language =javascript src='./js/a5954ra1_txtStdYYMM_txtStdYYMM.js'></script></TD>						
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=100%>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR HEIGHT=100%>
					<TD WIDTH=100%>
						<script language =javascript src='./js/a5954ra1_vspdData_vspdData.js'></script>
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
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
					                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
									 <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hStdDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hStdYYMM" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtPgmId" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
