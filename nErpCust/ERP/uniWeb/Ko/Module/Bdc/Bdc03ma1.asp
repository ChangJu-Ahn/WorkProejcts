<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : BDC
'*  2. Function Name        : 
'*  3. Program ID           : BDC04MA1
'*  4. Program Name         : BDC 업무등록 
'*  5. Program Desc         : BDC 마스터 데이타와 컴포넌트 정보 입력 
'*  6. Component List       : BDC001
'*  7. Modified date(First) : 2005/01/20
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kweon, SoonTae
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'=======================================================================================================-->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit
'☜: indicates that All variables must be declared in advance 
<!-- #Include file="../../inc/lgvariables.inc" -->
'==========================================================================================================
' 상수 및 변수 선언 
'----------------------------------------------------------------------------------------------------------
Const BIZ_PGM_ID	 = "BDC03MB1.asp"
Const C_PROCESSID	 = 1
Const C_PROCESS_NAME = 2
Const C_USE_FLAG	 = 3
Const C_TRAN_FLAG	 = 4
Const C_RUN_TIME	 = 5
Const C_START_ROW	 = 6
Const C_UPDATE_ID	 = 7
Const C_UPDATE_DT	 = 8
Const C_EXCEL		 = 9

Dim IsOpenPop

'==========================================================================================================
' 페이지 로드가 완료되면 자동으로 호출되는 함수.
' 초기화 루틴을 이곳에 집중시켜 주어야 함.
' ../../inc/incCliMAMain.vbs 파일에 이 함수를 호출 하도록 하는 모듈이 있슴 
'----------------------------------------------------------------------------------------------------------
Sub Form_Load()
	Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
    Call InitSpreadSheet
    Call InitVariables
    Call InitComboBox
    Call InitGridComboBox
    Call SetToolbar("11000000000111")
    
    frm1.txtProcId.focus
	Set gActiveElement = document.activeElement
End Sub

'==========================================================================================================
' 페이지를 떠날때 호출된다.
' 이 함수가 없으면 트리 메뉴를 이용해 다른 페이지로의 이동이 안된다.
'----------------------------------------------------------------------------------------------------------
Function FncExit()
	Dim IntRetCD

	FncExit = False

    ggoSpread.Source = frm1.vspdData

	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then

		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X","X")  '⊙: Will you destory previous data

		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    FncExit = True
End Function

'==========================================================================================================
' 시스템에 설정된 화폐단위, 언어코드, 등등등의 설정값을 초기화 하는 함수.
' ../../inc/incCliVariables.vbs 과 ../../ComAsp/LoadInfTB19029.asp  파일에 종속적이다.
'----------------------------------------------------------------------------------------------------------
Sub LoadInfTB19029()
<% Call loadInfTB19029A("I", "*","NOCOOKIE", "MA") %>
End Sub

'==========================================================================================================
' 스프레드 초기화 함수 
' 프로그램에 따라 사용자들이 조정해 주어야 하는 부분 
'----------------------------------------------------------------------------------------------------------
Sub InitSpreadSheet()
    With frm1.vspdData
        ggoSpread.Source = frm1.vspdData
        Call ggoSpread.Spreadinit					'("V20021124",,Parent.gAllowDragDropSpread)

        .ReDraw = False
        .MaxCols = C_EXCEL + 1
        .MaxRows = 0

        ggoSpread.SSSetEdit  C_PROCESSID,    "업무코드", 11,  , , 15, 2
        ggoSpread.SSSetEdit  C_PROCESS_NAME, "업 무 명", 30,  , , 40
        ggoSpread.SSSetCombo C_USE_FLAG,     "사용", 6, 2, False
        ggoSpread.SSSetCombo C_TRAN_FLAG,    "TRN", 6, 2, False
        ggoSpread.SSSetEdit  C_RUN_TIME,	 "시간", 6,   , , 5, 2
        ggoSpread.SSSetEdit  C_START_ROW,    "시작행", 8,   , , 1, 2
        ggoSpread.SSSetEdit  C_UPDATE_ID,    "등 록 자", 12,  , , 12, 2
        ggoSpread.SSSetEdit  C_UPDATE_DT,    "수정일시", 16,  , , 20, 2
		ggoSpread.SSSetButton C_EXCEL

        ggoSpread.SSSetProtected C_PROCESSID,    -1
        ggoSpread.SpreadLock     C_PROCESSID,    -1, C_PROCESSID       'khy200307
        ggoSpread.SSSetRequired  C_PROCESS_NAME, -1
        ggoSpread.SSSetRequired  C_USE_FLAG,     -1
        ggoSpread.SSSetRequired  C_TRAN_FLAG,    -1
        ggoSpread.SSSetRequired  C_RUN_TIME,     -1
        ggoSpread.SSSetRequired  C_START_ROW,	 -1
        ggoSpread.SpreadLock     C_UPDATE_ID,    -1, C_UPDATE_DT       'khy200307
        .ReDraw = True

        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    End With
End Sub

'==========================================================================================================
' 광역 변수들을 초기화 시킨다.
'----------------------------------------------------------------------------------------------------------
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE
    lgIntGrpCount = 0
    lgStrPrevKey = ""
    lgLngCurRows = 0
    lgSortKey = 1    
End Sub

'==========================================================================================================
' 스프레드시트 이외의 콤보박스들을 초기화 한다.
'----------------------------------------------------------------------------------------------------------
Sub InitComboBox()
End Sub

'==========================================================================================================
' 스프레드 시트의 콤보박스의 값을 초기화 한다.
'----------------------------------------------------------------------------------------------------------
Sub InitGridComboBox()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SetCombo "Y" & vbTab & "N", C_USE_FLAG
    ggoSpread.SetCombo "Y" & vbTab & "N", C_TRAN_FLAG
End Sub

'==========================================================================================================
' 업무코드 참조 팝업 창을 생성시킨다.
'----------------------------------------------------------------------------------------------------------
Function OpenPopup(Byval StrCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	arrParam(0) = "업무팝업"				' 팝업 명칭 
	arrParam(1) = "B_BDC_MASTER"			' TABLE 명칭 
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = ""						' Code Condition
	arrParam(5) = "업무"

	arrField(0) = "PROCESS_ID"				' Field명(0)
	arrField(1) = "PROCESS_NAME"			' Field명(1)

	arrHeader(0) = "업무코드"				' Header명(0)
	arrHeader(1) = "업 무 명"				' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
	                                Array(arrParam, arrField, arrHeader), _
		                            "dialogWidth=420px; dialogHeight=450px; " & _
		                            "center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtProcId.focus
		Exit Function
	Else
		frm1.txtProcId.focus
		frm1.txtProcId.value = arrRet(0)
		frm1.txtProcNm.value = arrRet(1)
	End If
End Function

'==========================================================================================================
' 메뉴바의 조회 버튼을 눌렀을때 호출되는 메세지 핸들러이다.
' 전달인수:
'----------------------------------------------------------------------------------------------------------
Function FncQuery()
    Dim IntRetCD 
    FncQuery = False
    Err.Clear

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
          Exit Function
        End If
    End If
    
    If Not chkField(Document, "1") Then
		Exit Function
    End If
    
    Call ggoSpread.ClearSpreadData()
    Call InitVariables
    
    If DbQuery = False Then
       Exit Function
    End If
       
    FncQuery = True
End Function

'==========================================================================================================
' 메뉴바의 조회 버튼을 눌렀을때 호출되는 메세지 핸들러이다.
' 전달인수:
'----------------------------------------------------------------------------------------------------------
Function DbQuery() 
    Dim strVal    
    Dim IntRetCD

    DbQuery = False

    Call LayerShowHide(1)
    With frm1
        strVal = BIZ_PGM_ID & _
                "?txtMode=" & Parent.UID_M0001 & _
                "&txtProcId=" & Trim(.txtProcId.value) & _
                "&txtMaxRows=" & .vspdData.MaxRows & _
                "&cboUseYN=" & Trim(.hUseYN.value) & _
                "&lgStrPrevKey=" & lgStrPrevKey
        Call RunMyBizASP(MyBizASP, strVal)
    End With

    DbQuery = True
End Function

'==========================================================================================================
' 조회 작업이 완료 되었을 때 자식 프레임에 의해 호출된다.
' 전달인수:
'----------------------------------------------------------------------------------------------------------
Function DbQueryOk()
    lgIntFlgMode = Parent.OPMD_UMODE

    Call ggoOper.LockField(Document, "Q")
    Call SetToolbar("11000000000111")
End Function

'==========================================================================================================
' 메뉴바의 저장 버튼을 눌렀을때 호출되는 메세지 핸들러이다.
' 전달인수:
' 참고: 사용자가 입력한 값을 저장한다.
'----------------------------------------------------------------------------------------------------------
Function FncSave() 
    Dim IntRetCD
    FncSave = False

    ggoSpread.Source = frm1.vspdData
   If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")
        Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then
       Exit Function
    End If
    
    If DbSave = False Then
       Exit Function
    End If

    FncSave = True
End Function

'==========================================================================================================
' FncSave 함수에 의해서 호출되는 함수로 사용자가 작성한 데이타를 가공하여 비지니스 로직이 있는 프로그램에 
' 전달해 준다.
' 전달인수:
'----------------------------------------------------------------------------------------------------------
Function DbSave() 
    Dim lRow
    Dim lGrpCnt
    Dim retVal
    Dim boolCheck
    Dim lStartRow
    Dim lEndRow
    Dim lRestGrpCnt
    Dim strVal, strDel
    Dim iColSep, iRowSep
    Dim IntRetCD

    iColSep = parent.gColSep
    iRowSep = parent.gRowSep

    DbSave = False

    Call LayerShowHide(1)

    With frm1
        .txtMode.value = Parent.UID_M0002
        .txtUpdtUserId.value = Parent.gUsrID
        .txtInsrtUserId.value = Parent.gUsrID

        lGrpCnt = 1

        strVal = ""
        strDel = ""

        For lRow = 1 To .vspdData.MaxRows
            Select Case GetSpreadText(.vspdData, 0, lRow, "X", "X")
                Case ggoSpread.UpdateFlag
                    strVal = strVal & "U" & iColSep & lRow & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_PROCESSID,	lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_PROCESS_NAME, lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_USE_FLAG,		lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_TRAN_FLAG,	lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_RUN_TIME,		lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_START_ROW,	lRow, "X", "X")) & iRowSep
                    lGrpCnt = lGrpCnt + 1
            End Select
        Next
    
        .txtMaxRows.value = lGrpCnt-1
        .txtSpread.value = strDel & strVal
        
        Call ExecMyBizASP(frm1, BIZ_PGM_ID)
    End With

    DbSave = True
End Function

'==========================================================================================================
' 저장 작업이 완료한 차일드 페이지에서 호출하는 메소드 이다.
' 전달인수:
' 참    고: 사용자 변경 사항이 있을 경우 정말 떠날것인지를 물어본다.
'----------------------------------------------------------------------------------------------------------
Function DbSaveOk()
    Call InitVariables
    frm1.vspdData.MaxRows = 0
    Call MainQuery()
End Function

'==========================================================================================================
' 현재 작업 페이지를 떠날때 호출되는 메세지 핸들러이다.
' 전달인수:
' 참    고: 사용자 변경 사항이 있을 경우 정말 떠날것인지를 물어본다.
'----------------------------------------------------------------------------------------------------------
Function FncExit()
    Dim IntRetCD
    FncExit = False
    
    ggoSpread.Source = frm1.vspdData    
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If

    FncExit = True
End Function

'==========================================================================================================
' 네번째 스프레드의 버튼셀이 클릭되었을때의 메세지 핸들러이다.
' 전달인수:
'  Col		  : 클릭된 열 
'  Row		  : 클릭된 행 
'  ButtonDown : 항상 0 임(무시해도 됨)
'
'  참고: vspdData2의 현제 행에 필드정보를 보낸다.
'----------------------------------------------------------------------------------------------------------
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
    Dim arrRet
	Dim arFieldInfo(3)
	Dim szProcessID
    
    szProcessID = GetSpreadText(frm1.vspdData, C_PROCESSID, Row, "X", "X")
	Call CommonQueryRs(" FIELD_ID, SHEET_NO, FIELD_SEQ, FIELD_NAME ", _
					   " B_BDC_FIELD ", _
					   " PROCESS_ID = '" & szProcessID & "' ORDER BY SHEET_NO, FIELD_SEQ", _
					   lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

	arFieldInfo(0) = lgF0
	arFieldInfo(1) = lgF1
	arFieldInfo(2) = lgF2
	arFieldInfo(3) = lgF3

    ExcelBrokerControl.CreateExcel(arFieldInfo)
End Sub

'==========================================================================================================
' 사용자가 레코드의 셀을 수정했을 경우 발생하는 메세지 핸들러 
' 전달인수:
'  Col : 수정된 열 번호 
'  Row : 수정된 행 번호 
'----------------------------------------------------------------------------------------------------------
Sub vspdData_Change(ByVal Col , ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    Call SetToolbar("11001000000111")
End Sub

'==========================================================================================================
' 라디오 버튼 1 이벤트 핸들러 
'----------------------------------------------------------------------------------------------------------
Function rdoCfmAll_OnClick()
	frm1.hUseYN.value = frm1.rdoCfmAll.value
End Function

'==========================================================================================================
' 라디오 버튼 2 이벤트 핸들러 
'----------------------------------------------------------------------------------------------------------
Function rdoCfmYes_OnClick()
	frm1.hUseYN.value = frm1.rdoCfmYes.value
End Function

'==========================================================================================================
' 라디오 버튼 3 이벤트 핸들러 
'----------------------------------------------------------------------------------------------------------
Function rdoCfmNo_OnClick()
	frm1.hUseYN.value = frm1.rdoCfmNo.value
End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
<!--
<OBJECT CLASSID="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
	<PARAM NAME="LPKPath" VALUE="../../Control/ExcelBroker.lpk">
</OBJECT>
-->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME="frm1" TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>업무조회/관리</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
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
            <TABLE CLASS="BasicTB" CELLSPACING=0>
                <TR>
                    <TD HEIGHT=5 WIDTH=100%></TD>
                </TR>
                <TR>
                    <TD HEIGHT=20 WIDTH=100%>
                        <FIELDSET CLASS="CLSFLD">
                            <TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
                                <TR>
									<TD CLASS="TD5" NOWRAP>업무</TD>
									<TD CLASS="TD6" NOWRAP>
									    <INPUT NAME="txtProcID" MAXLENGTH="20" SIZE=20 ALT ="업무코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProcID" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(txtProcID.value, 0)">
										<INPUT NAME="txtProcNm" MAXLENGTH="40" SIZE=40 ALT ="업 무 명" tag="14X"  STYLE="TEXT-ALIGN:left"></TD>
                                    <TD CLASS="TD5">사용여부</TD>
                                    <TD CLASS="TD6">
										<INPUT TYPE=radio CLASS="RADIO" NAME="rdoStatusflag" ID="rdoCfmAll" VALUE="" TAG = "11X" CHECKED>
											<LABEL FOR="rdoCfmAll">전체</LABEL>&nbsp;&nbsp;
										<INPUT type=radio CLASS="RADIO" NAME="rdoStatusflag" ID="rdoCfmYes" VALUE="Y" TAG = "11X">
											<LABEL FOR="rdoCfmYes">사용</LABEL>&nbsp;&nbsp;
										<INPUT TYPE=radio CLASS = "RADIO" NAME="rdoStatusflag" ID="rdoCfmNo" VALUE="N" TAG = "11X">
											<LABEL FOR="rdoCfmNo">미사용</LABEL>
                                    </TD>
                                </TR>
                            </TABLE>
                        </FIELDSET>
                    </TD>
                </TR>
                <TR>
                    <TD WIDTH=100% HEIGHT=* valign=top>
                        <TABLE WIDTH="100%" HEIGHT="100%">
                            <TR>
                                <TD HEIGHT="100%">
                                    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>

                                </TD>
                            </TR>
                        </TABLE>
                    </TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>  
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>>
            <IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hUseYN" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
    <IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO NORESIZE FRAMESPACING=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
<OBJECT ID="ExcelBrokerControl"
		CLASSID="CLSID:3894EE93-0291-4D97-8423-FAE813587B6E"
		CODEBASE="../../Control/ExcelBroker.CAB#version=1,1,0,64"
		WIDTH=0	HEIGHT=0>
</OBJECT>
</BODY>
</HTML>
