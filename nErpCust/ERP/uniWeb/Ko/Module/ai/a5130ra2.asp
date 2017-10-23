<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Exchange reference 
'*  3. Program ID           : A5954RA1
'*  4. Program Name         : 환율참조팝업
'*  5. Program Desc         : Popup of Exchange
'*  6. Component List       : DB agent
'*  7. Modified date(First) : 2002.05.06
'*  8. Modified date(Last)  : 2002.05.06
'*  9. Modifier (First)     : Jang Yoon Ki
'* 10. Modifier (Last)      : Jang Yoon Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE></TITLE>
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

Const BIZ_PGM_ID 		= "a5130rb2.asp"                              '☆: Biz Logic ASP Name
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


top.document.title = PopupParent.gActivePRAspName
	
ChcMnDt = ""

'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------


'========================================================================================
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


'========================================================================================
Sub SetDefaultVal()
	frm1.txtAcctCd.value  = arrParam(0) '계정코드
	frm1.txtStdYYYY.value = Trim(arrParam(1))
	frm1.txtStdDt.value = Trim(arrParam(2))
	frm1.txtBizAreaCd.value = Trim(arrParam(3))	
End Sub

'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call LoadInfTB19029A("Q", "A","NOCOOKIE","RA") %>
End Sub


'========================================================================================
Function OKClick()
		
	'Dim intColCnt, intRowCnt, intInsRow
		
	'if frm1.vspdData.ActiveRow > 0 Then 			
		
	'	intInsRow = 0
	'End if			
		
	'Self.Returnvalue = arrReturn
	'Self.Close()
	
					
End Function

'========================================================================================
Function CancelClick()
	Self.Close()			
End Function

'========================================================================================
Function MousePointer(pstr1)
	Select case UCase(pstr1)
	case "PON"
		window.document.search.style.cursor = "wait"
	case "POFF"
		window.document.search.style.cursor = ""
	End Select
End Function


'========================================================================================
Sub InitSpreadSheet()
    frm1.vspdData.OperationMode = 5
    Call SetZAdoSpreadSheet("A5130RA2","S","A","V20021211",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock("A")
End Sub



'========================================================================================
Sub SetSpreadLock(s)
    With frm1
    
    .vspdData.ReDraw = False
	 ggoSpread.SpreadLockWithOddEvenRowColor()
    .vspdData.ReDraw = True

    End With
End Sub

'========================================================================================
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

'========================================================================================
Sub Form_Load()

    Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")
	Call InitVariables
	Call SetDefaultVal
	Call InitSpreadSheet()
	Call FncQuery()
End Sub

'========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub


'========================================================================================
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
	ElseIf KeyAscii = 27 Then
		Call CancelClick
	End If
End Sub

'========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

   	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgPageNo <> "" Then
			If DbQuery = False Then
				Exit Sub
			End if
    	End If
    End If
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


'========================================================================================
Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function


'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.MaxRows > 0 Then
		If Frm1.vspdData.ActiveRow = Row Or Frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub

'========================================================================================
Function FncQuery() 
	Dim IntRetCD
    FncQuery = False 

    Err.Clear

    Call ggoOper.ClearField(Document, "2")
    Call InitVariables

    If Not chkField(Document, "1") Then
       Exit Function
    End If

     If Trim(frm1.txtAcctCd.value) = "" or Trim(frm1.txtStdYYYY.value) = "" Then
		Call DisplayMsgBox("110100", "X", "X", "X")
		Call CancelClick()
		Exit Function
    End If

	If DbQuery = False Then Exit Function

    FncQuery = True
End Function


'========================================================================================
Function DbQuery() 
	Dim strVal
	Dim txtStdYYMMYear, txtStdYYMMMonth

    DbQuery = False
    Err.Clear

	Call LayerShowHide(1)
    With frm1
		strVal = BIZ_PGM_ID & "?txtStdYYYY=" & Trim(.txtStdYYYY.Value)
		strVal = strVal & "&txtStdDt=" & Trim(.txtStdDt.Value)
		strVal = strVal & "&txtBizAreaCd=" & Trim(.txtBizAreaCd.Value)
		strVal = strVal & "&txtBizAreaCd_Alt=" & Trim(.txtBizAreaCd.Alt)
		strVal = strVal & "&txtAcctCd=" & Trim(.txtAcctCd.Value)
	'===================================================================
		strVal = strVal & "&lgPageNo="   & lgPageNo                      '☜: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
	        			
       Call RunMyBizASP(MyBizASP, strVal)
    End With
    DbQuery = True
End Function


'========================================================================================
Function DbQueryOk()
    lgBlnFlgChgValue = False
	lgIntFlgMode = PopupParent.OPMD_UMODE
	lgSaveRow        = 1

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
		frm1.vspdData.Row = 1
		frm1.vspdData.SelModeSelected = true
	End If
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
						<TD CLASS=TD5 ID="txtStdDtTit" NOWRAP>{{회계년도}}</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtStdYYYY" SIZE=10 MAXLENGTH=10 tag="14XXXU" ALT="{{회계년도}}" STYLE="TEXT-ALIGN:left"></TD>						
						<TD CLASS="TD5" NOWRAP>{{계정코드}}</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtAcctCd" SIZE=10 MAXLENGTH=10 tag="14XXXU" ALT="{{계정코드}}" STYLE="TEXT-ALIGN:left">&nbsp;<INPUT TYPE=TEXT NAME="txtAcctCdNm" SIZE=20 tag="24X" ALT="{{계정코드명}}" STYLE="TEXT-ALIGN: Left">
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 ID="txtStdYYMMTit" NOWRAP>{{월}}</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtStdDt" SIZE=10 MAXLENGTH=10 tag="14XXXU" ALT="{{월}}" STYLE="TEXT-ALIGN:left"></TD>						
						<TD CLASS="TD5" NOWRAP>{{사업장코드}}</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="14XXXU" ALT="{{사업장코드}}" STYLE="TEXT-ALIGN:left">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=20 tag="24X" ALT="{{사업장명}}" STYLE="TEXT-ALIGN: Left">
						</TD>
						</TD>
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
						<script language =javascript src='./js/a5130ra2_vspdData_vspdData.js'></script>
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
					<TD ALIGN=RIGHT> &nbsp;
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
