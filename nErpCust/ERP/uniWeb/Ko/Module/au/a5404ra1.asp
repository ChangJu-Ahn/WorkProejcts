<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2000/12/09
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Byun Jee Hyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'********************************************************************************************* -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'#########################################################################################################
'												1. 선 언 부 
'#########################################################################################################

'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--
'=========================================  1.1.1 Style Sheet  ===========================================
'========================================================================================================= -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
'=====================================  1.1.2 공통 Include   =============================================
'========================================================================================================= -->

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ag/AcctCtrl.vbs">           </SCRIPT>    
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance

'****************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Dim lgBlnFlgChgValue                                        '☜: Variable is for Dirty flag            
Dim lgStrPrevKey                                            '☜: Next Key tag                          
Dim lgSortKey                                               '☜: Sort상태 저장변수                      
Dim lgIsOpenPop                                             '☜: Popup status                           

Dim lgSelectList                                            '☜: SpreadSheet의 초기  위치정보관련 변수 
Dim lgSelectListDT                                          '☜: SpreadSheet의 초기  위치정보관련 변수 

Dim lgMark

Dim IsOpenPop 
Dim lgParentsPgmID


'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "a5404rb1.asp"
Const C_SHEETMAXROWS    = 16                                   '☆: Spread sheet에서 보여지는 row
Dim lsPoNo                                                 '☆: Jump시 Cookie로 보낼 Grid value
Const C_MaxKey          = 1					                          '☆: SpreadSheet의 키의 갯수 

Dim arrReturn
Dim arrParent
Dim arrParam

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인					

	 '------ Set Parameters from Parent ASP ------ 
	arrParent		= window.dialogArguments
	Set PopupParent = arrParent(0)
	arrParam		= arrParent(1)
	
	top.document.title = "회계전표팝업"

'========================================================================================================= 
Sub InitVariables()
    Redim arrReturn(0)

    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1
	lgAuthorityFlag = arrParam(4)                          '권한관리 추가 
    
	Self.Returnvalue = arrReturn
	
	' 권한관리 추가 
	If UBound(arrParam) > 5 Then
		lgAuthBizAreaCd		= arrParam(5)
		lgInternalCd		= arrParam(6)
		lgSubInternalCd		= arrParam(7)
		lgAuthUsrID			= arrParam(8)
	End If	
End Sub


'========================================================================================================= 
Sub SetDefaultVal()
	Dim strYear, strMonth, strDay
	Dim StartDate,EndDate
	    
    Call	ExtractDateFrom("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gServerDateType, strYear, strMonth, strDay)

	StartDate= UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, "01")		'☆: 초기화면에 뿌려지는 시작 날짜 
	EndDate= UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, strDay)		'☆: 초기화면에 뿌려지는 마지막 날짜 


	frm1.txtfrgldt.Text	= StartDate
	frm1.txttogldt.Text	= EndDate
End Sub

Function OpenDept(Byval pStrCode)
	Dim iArrRet
	Dim iArrParam(8)
	Dim IntRetCD	
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("DeptPopupDtA2")

	If IsOpenPop = True Then Exit Function
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "DeptPopupDtA2", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	iArrParam(0) = pStrCode									'  Code Condition
   	iArrParam(1) = frm1.txtfrgldt.Text
	'iArrParam(2) = PopupParent.lgUsrIntCd								' 자료권한 Condition  
	' 권한관리 추가 
	iArrParam(5) = lgAuthBizAreaCd
	iArrParam(6) = lgInternalCd
	iArrParam(7) = lgSubInternalCd
	iArrParam(8) = lgAuthUsrID

	iArrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, iArrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	If iArrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	Else
		frm1.txtDeptCd.focus
		frm1.txtDeptCd.value = iArrRet(0)
		frm1.txtDeptNm.value = iArrRet(1)
	End If	
End Function


'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A","COOKIE","RA") %>                                '☆: 

End Sub


'========================================================================================================
	
	Function OKClick()
		
		Dim intColCnt, intRowCnt, intInsRow
		
		if frm1.vspdData.ActiveRow > 0 Then 			
		
			intInsRow = 0

			Redim arrReturn(1)
			
			For intRowCnt = 0 To frm1.vspdData.MaxRows - 1
			
				frm1.vspdData.Row = intRowCnt + 1
			
				If frm1.vspdData.SelModeSelected Then
				   frm1.vspdData.Col = 1
				   arrReturn(intColCnt) = frm1.vspdData.Text										
				   intInsRow = intInsRow + 1					
				End IF
			Next
			
		End if			
		
		Self.Returnvalue = arrReturn
		Self.Close()
					
	End Function


'========================================================================================================

	Function CancelClick()
		Self.Close()			
	End Function

'========================================================================================================

	Function MousePointer(pstr1)
	      Select case UCase(pstr1)
	            case "PON"
					window.document.search.style.cursor = "wait"
	            case "POFF"
					window.document.search.style.cursor = ""
	      End Select
	End Function


'==========================================================================================================
Sub InitSpreadSheet()
    
    frm1.vspdData.OperationMode = 3
       
       Call SetZAdoSpreadSheet("A5104RA1","S","A","V20030206",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	   Call SetSpreadLock 

End Sub


'=========================================================================================================
Sub SetSpreadLock()
     ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub


 
'===========================================================================
Function PopZAdoConfigGrid()
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function



'===========================================================================
Function OpenOrderBy()

	Dim arrRet
	Dim arrParam
	Dim TInf(5)
	Dim ii
	
	On Error Resume Next
	
	ReDim arrParam(PopupParent.C_MaxSelList * 2 - 1 )

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
    TInf(0) = PopupParent.gMethodText    
  
	For ii = 0 to PopupParent.C_MaxSelList * 2 - 1 Step 2
      arrParam(ii + 0 ) = lgPopUpR(ii / 2  , 0)
      arrParam(ii + 1 ) = lgPopUpR(ii / 2  , 1)
    Next  
  
	arrRet = window.showModalDialog("../../ComAsp/ADOGrpSortPopup.asp",Array(lgSortFieldCD,lgSortFieldNm,arrParam,TInf),"dialogWidth=420px; dialogHeight=250px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "0" Then
		If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
	
	   For ii = 0 to PopupParent.C_MaxSelList * 2 - 1 Step 2
           lgPopUpR(ii / 2 ,0) = arrRet(ii + 1)  
           lgPopUpR(ii / 2 ,1) = arrRet(ii + 2)
       Next    
	   
       Call InitVariables
       Call InitSpreadSheet
   End If
End Function


 '++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

 '==========================================   CookiePage()  ======================================
'	Name : CookiePage()
'	Description : JUMP시 Load화면으로 조건부로 Value
'==================================================================================================== 
Sub CookiePage(ByVal Kubun)

	Select Case Kubun		
		Case "FORM_LOAD"
			lgParentsPgmID = ReadCookie("PGMID")			
			Call WriteCookie("PGMID", "")			
		Case Else			
	End Select
End Sub

'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)

    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    
	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call CookiePage("FORM_LOAD")
    
End Sub


'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'==========================================================================================

Sub txtfrgldt_DblClick(Button)
	if Button = 1 then
		frm1.txtfrgldt.Action = 7
	End if
End Sub

'==========================================================================================
Sub txttogldt_DblClick(Button)
	if Button = 1 then
		frm1.txttogldt.Action = 7
	End if
End Sub


'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop)
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then	'☜: 재쿼리 체크'
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			DbQuery
		End If
   End if
    
End Sub


'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
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
    
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	If Row < 1 Then Exit Sub

'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)
End Sub

Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function



Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.MaxRows > 0 Then
		If Frm1.vspdData.ActiveRow = Row Or Frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub

Sub txtfrgldt_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call Fncquery()
    End if
End Sub

Sub txttogldt_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call Fncquery()
    End if
End Sub


'==========================================================================================
Sub txtFrGlNo_OnKeyPress()	
	If window.event.keycode = 39 then	'Single quotation mark 입력불가 
		window.event.keycode = 0	
	End If
End Sub


'==========================================================================================
Sub txtFrGlNo_onpaste()	
	Dim iStrGlNo 	
	iStrGlNo = window.clipboardData.getData("Text")
	iStrGlNo = RePlace(iStrGlNo, "'", "")
	Call window.clipboardData.setData("text",iStrGlNo)		
End Sub


'==========================================================================================
Sub txtToGlNo_OnKeyPress()	
	If window.event.keycode = 39 then	'Single quotation mark 입력불가 
		window.event.keycode = 0	
	End If
End Sub


'==========================================================================================
Sub txtToGlNo_onpaste()	
	Dim iStrGlNo 	
	iStrGlNo = window.clipboardData.getData("Text")
	iStrGlNo = RePlace(iStrGlNo, "'", "")
	Call window.clipboardData.setData("text",iStrGlNo)		
End Sub

 
'********************************************************************************************************* 
Function FncQuery() 
	Dim IntRetCD
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
   

    '-----------------------
    'Erase contents area
    '-----------------------
    'Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If

	If CompareDateByFormat(frm1.txtFrGlDt.text,frm1.txtToGlDt.text,frm1.txtFrGlDt.Alt,frm1.txtToGlDt.Alt, _
                        "970025",frm1.txtFrGlDt.UserDefinedFormat,PopupParent.gComDateType,True) = False Then		
		Exit Function    
    End If

    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData

    
    '-----------------------
    'Query function call area
    '-----------------------
    if  DbQuery	= False Then														'☜: Query db data
		Exit Function
    End If

    FncQuery = True		
End Function


'========================================================================================

Function FncPrint() 
    Call popupparent.FncPrint()
End Function


'========================================================================================

Function FncExcel() 
	Call popupparent.FncExport(PopupParent.C_MULTI)
End Function


'========================================================================================

Function FncFind() 
    Call popupparent.FncFind(PopupParent.C_MULTI , False)                                     '☜:화면 유형, Tab 유무 
End Function


'========================================================================================

Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", PopupParent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function

 
'========================================================================================

Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear           

	Call LayerShowHide(1)
    
    With frm1
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------
		strVal = BIZ_PGM_ID & "?txtfrgldt=" & Trim(.txtfrgldt.Text)
		strVal = strVal & "&txttogldt=" & Trim(.txttogldt.Text)
		strVal = strVal & "&txtfrglno=" & Trim(.txtfrGlNo.value)
		strVal = strVal & "&txttoglno=" & Trim(.txttoGlNo.value)
		strVal = strVal & "&txtdeptcd=" & Trim(.txtdeptcd.value)
				

'--------------- 개발자 coding part(실행로직,End)------------------------------------------------
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag
'       strVal = strVal & "&lgMaxCount="     & CStr(C_SHEETMAXROWS_D)            '☜: 한번에 가져올수 있는 데이타 건수 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderbyList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		strVal = strVal & "&lgAuthorityFlag="   & EnCoding(lgAuthorityFlag)            '권한관리 추가		
		
		' 권한관리 추가 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인		
		
        Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True


End Function


'========================================================================================

Function DbQueryOk()														'☆: 조회 성공후 실행로직 

    '-----------------------
    'Reset variables area
    '-----------------------
    Dim iCnt
    lgBlnFlgChgValue = True                                                 'Indicates that no value changed
	If frm1.vspdData.MaxRows > 0  Then
	   frm1.vspdData.focus
	End If
	
	Select Case lgParentsPgmID
	
		Case "A5104MA1"	
			For iCnt = 1 To frm1.vspdData.MaxRows
				frm1.vspdData.Row = iCnt		
				frm1.vspdData.Col = 7				
				If Trim(UCase(frm1.vspdData.Text)) <> "GL" then		
					frm1.vspdData.Col = -1
					frm1.vspdData.ForeColor = &H00808080& 'RGB(192, 192, 192)
				End If
			Next
		Case "A5126MA1"
			For iCnt = 1 To frm1.vspdData.MaxRows
				frm1.vspdData.Row = iCnt		
				frm1.vspdData.Col = 7				
				If Trim(UCase(frm1.vspdData.Text)) = "GL" then		
					frm1.vspdData.Col = -1
					frm1.vspdData.ForeColor = &H00808080& 'RGB(192, 192, 192)
				End If
			Next	
		
	End Select
	
End Function


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->
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
						<TD CLASS=TD5 NOWRAP>전표일자</TD>
						<TD CLASS=TD6 NOWRAP>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime1 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtfrgldt CLASSID=<%=gCLSIDFPDT%> ALT="시작일자" tag="11"> </OBJECT>');</SCRIPT>&nbsp;~&nbsp;
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txttogldt CLASSID=<%=gCLSIDFPDT%> ALT="종료일자" tag="11"> </OBJECT>');</SCRIPT>
						</TD>												
						<TD CLASS=TD5 NOWRAP>전표번호</TD>				
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtfrGlNo" SIZE=12 MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag="1" ALT="전표번호">&nbsp;~&nbsp;
											 <INPUT TYPE="Text" NAME="txttoGlNo" SIZE=12 MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag="1" ALT="전표번호"></TD>
					</TR>
					<TR>				
						<TD CLASS=TD5 NOWRAP>부서코드</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeptCd" ALT="부서코드" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: left" tag ="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(frm1.txtDeptCd.Value)">&nbsp;
						                     <INPUT NAME="txtDeptNm" ALT="부서명"   MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN: left" tag ="14XXXU"></TD>
						<TD CLASS=TD5 NOWRAP></TD>
						<TD CLASS=TD6 NOWRAP></TD>				
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
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% id=vspdData tag="2"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"><PARAM NAME="ReDraw" VALUE="0"> <PARAM NAME="FontSize" VALUE="10"> </OBJECT>');</SCRIPT>
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
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>&nbsp;<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


  