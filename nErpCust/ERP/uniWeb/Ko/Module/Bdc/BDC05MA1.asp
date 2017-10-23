<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : DBC
'*  2. Function Name        : 배치작업 상세조회 
'*  3. Program ID           : BDC05MA1.ASP
'*  4. Program Name         : 작업상세조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2005.02.07
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kweon, Soon Tae
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'********************************************************************************************-->
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
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit
'☜: indicates that All variables must be declared in advance 
<!-- #Include file="../../inc/lgvariables.inc" -->
'==========================================================================================================
' 상수 및 변수 선언 
'----------------------------------------------------------------------------------------------------------
Const BIZ_PGM_ID = "BDC05MB1.ASP"										'☆: 조회 비지니스 로직 ASP명 
Const BIZ_PGM_JUMP_ID = "BDC04MA1"

Dim lgRetFlag
Dim IsOpenPop
Dim iColSep, iRowSep
Dim lgOldRow

Dim C_SP1_SEQ
Dim C_SP1_TIM
Dim C_SP1_RES
Dim C_SP1_COM
Dim C_SP1_MTH

Dim strMode

Dim szCurMth
Dim szCurPrm
Dim szCurJon

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
	Call SetToolbar("1100000000000111")
	
	If parent.ReadCookie("txtJobId") <> "" Then

		Call SetCookieVal
	End If
	frm1.txtJobID.focus
	
	
End Sub

'=========================================================================================================
Function FncCancel()
    ggoSpread.EditUndo
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
	
	Call InitSpreadPosVariables()

	With frm1.vspdData
        .ReDraw = False
		.RowHeadersShow = True
		.MaxCols = C_SP1_MTH
        .MaxRows = 0

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20041121", , parent.gAllowDragDropSpread
		
		Call GetSpreadColumnPos()
		
		ggoSpread.SSSetEdit   C_SP1_SEQ,  "순서", 6, , , 3
		ggoSpread.SSSetEdit   C_SP1_TIM,  "처리시간", 20, , , 26
		ggoSpread.SSSetEdit   C_SP1_RES,  "결과", 6, , , 6
		ggoSpread.SSSetEdit   C_SP1_COM,  "오류모듈", 20, , , 40
		ggoSpread.SSSetEdit   C_SP1_MTH,  "오류설명", 40, , , 200
		
		ggoSpread.SSSetSplit2(1)
		.ReDraw = True
	End With
	
	Call SetSpreadLock()
	
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()	
	
	' Grid 1(vspdData) - Operation 
	C_SP1_SEQ = 1
	C_SP1_TIM = 2
	C_SP1_RES = 3
	C_SP1_COM = 4
	C_SP1_MTH = 5

End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos()
 	Dim iCurColumnPos

 	ggoSpread.Source = frm1.vspdData
 			
 	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 	
 	C_SP1_SEQ = iCurColumnPos(1)
	C_SP1_TIM = iCurColumnPos(2)
	C_SP1_RES = iCurColumnPos(3)
	C_SP1_COM = iCurColumnPos(4)
	C_SP1_MTH = iCurColumnPos(5)
	
End Sub				


'==========================================================================================================
' 광역 변수들을 초기화 시킨다.
'----------------------------------------------------------------------------------------------------------
Sub InitVariables()
	Dim i, j

    lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

'========================================  2.2.1 SetCookieVal()  ======================================
'	Name : SetCookieVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=================================================================================================== 
Sub SetCookieVal()
   	
	frm1.txtJobID.value	= ReadCookie("txtJobId")
    call dbQuery()
	WriteCookie "txtJobId", ""
		
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
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
	
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey			'Sort in Descending
 			lgSortKey = 1
 		End If
 		
 		lgOldRow = Row
 		
	Else
 		'------ Developer Coding part (Start)
 		If lgOldRow <> Row Then		
			frm1.vspdData.Row = row
		
			lgOldRow = Row
		
		End If		
	 	'------ Developer Coding part (End)
	
 	End If
	
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub


'=========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)        
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
  
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos()
End Sub 

'==========================================================================================================
' 업무코드 참조 팝업 창을 생성시킨다.
'----------------------------------------------------------------------------------------------------------
Function OpenPopup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	arrParam(0) = "작업팝업"				' 팝업 명칭 
	arrParam(1) = "B_BDC_JOBS"				' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtJobID.Value)
	arrParam(3) = ""
	arrParam(4) = " job_state= " & Filtervar("D", "''", "S")						' Code Condition
	arrParam(5) = "업무"

	arrField(0) = "JOB_ID"					' Field명(0)
	arrField(1) = "JOB_TITLE"				' Field명(1)

	arrHeader(0) = "작업코드"				' Header명(0)
	arrHeader(1) = "작업명"				' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
	                                Array(arrParam, arrField, arrHeader), _
		                            "dialogWidth=420px; dialogHeight=450px; " & _
		                            "center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtJobID.focus
		Exit Function
	Else
		frm1.txtJobID.value = arrRet(0)
		frm1.txtJobNm.value = arrRet(1)
		frm1.txtJobID.focus
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

    Call ggoSpread.ClearSpreadData()
    Call InitVariables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then
		Exit Function
    End If

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
                "&txtJobId=" & Trim(.txtJobId.value) & _
                "&txtMaxRows=" & .vspdData.MaxRows & _
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
End Function

'==========================================================================================================
' 메뉴바의 저장 버튼을 눌렀을때 호출되는 메세지 핸들러이다.
' 전달인수:
' 참고: 사용자가 입력한 값을 저장한다.
'----------------------------------------------------------------------------------------------------------
Function FncSave() 
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



'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    
	Dim LngRow
	 
    ggoSpread.Source = gActiveSpdSheet

    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()  
    
	Call ggoSpread.ReOrderingSpreadData

End Sub 

Function JumpJobRun()

    Dim IntRetCd, strVal
    
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	WriteCookie "txtJobId", UCase(Trim(frm1.txtJobID.value))
	
	PgmJump(BIZ_PGM_JUMP_ID)
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><IMG SRC="../../../CShared/image/table/seltab_up_left.gif" WIDTH="9" HEIGHT="23"></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="center" CLASS="CLSMTAB"><FONT COLOR=white>작업상세조회</FONT></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="right"><IMG SRC="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH="*" align=right>&nbsp;</td>
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
									<TD CLASS="TD5" NOWRAP>작업코드</TD>
									<TD CLASS="TD656" NOWRAP>
									    <INPUT NAME="txtJobID" MAXLENGTH="18" SIZE=18 ALT ="작업코드" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProcID" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup()">
										<INPUT NAME="txtJobNm" MAXLENGTH="80" SIZE=50 ALT ="작업명" tag="14X"  STYLE="TEXT-ALIGN:left"></TD>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="50%">
								<TD WIDTH="100%" colspan=8>
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
	<TR HEIGHT=20>   
        <TD WIDTH=100%>   
            <TABLE <%=LR_SPACE_TYPE_30%>>   
                <TR>   
                    <TD>	
						<TABLE <%=LR_SPACE_TYPE_30%>>
							<TR>
							  <TD WIDTH=10>&nbsp;</TD>
							  <TD align="left">
							  </TD>
							  <TD WIDTH=* Align=right><A href="vbscript:JumpJobRun">작업관리</A> </TD>
							  <TD WIDTH=10>&nbsp;</TD>
							</TR>
						</TABLE>
					</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
</HTML>

