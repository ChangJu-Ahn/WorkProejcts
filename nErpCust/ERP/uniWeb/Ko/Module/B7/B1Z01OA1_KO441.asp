<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Yim Young Ju
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript" SRC = "../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_ID = "B1Z01OB1_KO441.asp"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgQueryFlag				 '--- 1:New Query 0:Continuous Query 
Dim ArrParent

Dim arrParam				'--- First Parameter Group 
ReDim arrParam(5)
Dim arrReturn				'--- Return Parameter Group 

Dim IsOpenPop   
Dim gDepart       

'------ Set Parameters from Parent ASP ------ 
ArrParent = window.dialogArguments
Set PopupParent = ArrParent(0)
arrParam(0) = ArrParent(1)
arrParam(1) = ArrParent(2)
arrParam(2) = ArrParent(3)
arrParam(3) = ArrParent(4) '자신이 소속된 부서코드
arrParam(4) = ArrParent(5)

top.document.title = PopupParent.gActivePRAspName
'--------------------------------------------- 

Function InitVariables()
	lgSortKey = 1                            '⊙: initializes sort direction
	lgQueryFlag = "1"
End Function


Dim C_QueryID
Dim C_QueryNM
Dim C_DeptCD
Dim C_DeptNM
Dim C_SELECT
Dim C_FROM
Dim C_WHERE
Dim C_ETC
Dim C_REMARK

Sub initSpreadPosVariables()  

	C_QueryID		=	1
	C_QueryNM		=	2
	C_DeptCD		=	3
	C_DeptNM		=	4
	C_SELECT		=	5
	C_FROM			=	6
	C_WHERE			=	7
	C_ETC			=	8
	C_REMARK		=	9

End Sub

Sub SetDefaultVal()
	
	gDepart = arrParam(3)
	frm1.txtQueryID.Value = arrParam(2)
	Self.Returnvalue = Array("")
		
End Sub

Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "Q","NOCOOKIE","PA") %>
End Sub

Sub InitSpreadSheet()
	Call initSpreadPosVariables()    

	ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20040518",,PopupParent.gAllowDragDropSpread
	
	With vspdData
		.ReDraw = False
		
		.MaxCols = C_REMARK
		.MaxRows = 0
	End With
	
	Call GetSpreadColumnPos("A")
	
	With ggoSpread
		
		.SSSetEdit C_QueryID,"Query 번호", 15
		.SSSetEdit C_QueryNM,"Query 명", 20
		.SSSetEdit C_DeptCD,"부서", 10
		.SSSetEdit C_DeptNM,"부서명", 15
		.SSSetEdit C_SELECT,"SELECT 절", 20
		.SSSetEdit C_FROM,"FROM 절",20
		.SSSetEdit C_WHERE,"WHERE 절",20
		.SSSetEdit C_ETC,"기타 절",20
		.SSSetEdit C_REMARK,"비고",20
		
	End With
		
	Call SetSpreadLock
	
	vspdData.ReDraw = True
End Sub

Sub SetSpreadLock()	
    ggoSpread.Source = vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = vspdData
            
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_QueryID = iCurColumnPos(1)
			C_QueryNM = iCurColumnPos(2)
			C_DeptCD = iCurColumnPos(3)
			C_DeptNM = iCurColumnPos(4)
			C_SELECT = iCurColumnPos(5)
			C_FROM = iCurColumnPos(6)
			C_WHERE = iCurColumnPos(7)
			C_ETC = iCurColumnPos(8)
			C_REMARK = iCurColumnPos(9)

	End Select    
End Sub

Function OKClick()
	Dim intColCnt, iCurColumnPos
	
	If vspdData.ActiveRow > 0 Then	
		Redim arrReturn(vspdData.MaxCols - 1)
	
		ggoSpread.Source = vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		vspdData.Row = vspdData.ActiveRow 
				
		For intColCnt = 0 To vspdData.MaxCols - 1
			vspdData.Col = iCurColumnPos(CInt(intColCnt + 1))
			arrReturn(intColCnt) = vspdData.Text
		Next

		Self.Returnvalue = arrReturn
	End If
	
	Self.Close()
End Function

Function CancelClick()
	On Error Resume Next
	Redim arrReturn(0)
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

Sub Form_Load()

	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	
	Call InitVariables
	Call SetDefaultVal()

	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec)
		
	Call InitSpreadSheet()
	Call InitComboBox()
	Call FncQuery()
	
	If frm1.txtQueryID.Value = "" Then
		frm1.txtQueryID.focus
    End If
    
	Set gActiveElement = document.activeElement 
	
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("B0020", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    
    Call SetCombo2(frm1.cboRole_type,iCodeArr, iNameArr,Chr(11))                  ''''''''DB에서 불러 condition에서
End Sub

Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	arrParam(0) = "쿼리번호팝업"
	arrParam(1) = "B_QUERY_COMMAND A"
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = ""
	arrParam(5) = "Query 번호"
	
	arrField(0) = "A.QUERY_ID"
	arrField(1) = "A.QUERY_NM"
			    
	arrHeader(0) = "Query 번호"
	arrHeader(1) = "Query 명"
		
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtQueryID.focus
		Exit Function
	End If
	
	frm1.txtQueryID.value = arrRet(0)	
	frm1.txtQueryID.focus
			

End Function


Sub txtToInspReqDt_DblClick(Button)
    If Button = 1 Then
        txtToInspReqDt.Action = 7
    End If
End Sub

Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")

	gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = vspdData

    If vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
        ggoSpread.Source = vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
   	
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then              ' 타이틀 cell을 dblclick했거나....
	   Exit Function
	End If
	
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick()
		End If
	End If
End Function

Function vspdData_KeyPress(KeyAscii)
	If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	
	'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음
	If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			If DBQuery = False Then
				Exit Sub
			End If
		End If
	End If

End Sub


Sub txtToInspReqDt_KeyPress(KeyAscii)
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End If
End Sub

Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    vspdData.Redraw = False
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
	vspdData.Redraw = True
End Sub

Function FncQuery()
	FncQuery = False
   	
   	vspdData.MaxRows = 0
	lgQueryFlag = "1"
	lgStrPrevKey = ""

	If Not chkField(Document, "1") Then
		Exit Function
	End If
	
	if DbQuery = false then
		Exit Function
	End if

	FncQuery = True
End Function

Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub


'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtDept_cd.value			            '  Code Condition
	End If
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  

	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(PopupParent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDept_cd.focus
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
	
	lgBlnFlgChgValue = True
			
End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)
		
	With frm1
		Select Case iWhere
		     Case "0"
               .txtDept_cd.value = arrRet(0)
               .txtDept_nm.value = arrRet(1)
               .txtDept_cd.focus
        End Select
	End With
End Function      

Function DbQuery()
	Dim strVal
	Dim txtMaxRows
	
	DbQuery = False 	

	'Show Processing Bar
    Call LayerShowHide(1)  

	txtMaxRows = vspdData.MaxRows
	
	strVal = BIZ_PGM_ID & "?QueryFlag=" & lgQueryFlag _
			& "&txtQueryID=" & Trim(frm1.txtQueryID.Value) _
			& "&txtQueryNm=" & Trim(frm1.txtQueryNm.Value) _
			& "&txtDept_cd=" & Trim(frm1.txtDept_cd.Value) _
			& "&gDepart=" & gDepart _
			& "&cboRole_type=" & frm1.cboRole_type.value 
				
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동
	
	DbQuery = True 
	
End Function

Function DbQueryOk()								'☆: 조회 성공후 실행로직
	lgQueryFlag = "0"
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR HEIGHT=*>
		<TD  WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%>></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									
									<TD CLASS="TD5" NOWRAP>Query 번호</TD>
									<TD CLASS="TD6" NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtQueryID" SIZE="20" MAXLENGTH="18" ALT="Query 번호" TAG="11XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCurCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtQueryID.value, 0)">&nbsp;<INPUT NAME="txtQueryNm" ALT="Query 명" MAXLENGTH="100" SIZE=60 tag="11XXXU"></TD>
								</TR>
								
								<TR>
									<TD CLASS=TD5 NOWRAP>부서</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept_cd" ALT="부서코드" TYPE="Text" SiZE=10 MAXLENGTH=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenDept(0)">
									<INPUT NAME="txtDept_nm" ALT="부서코드명" TYPE="Text" SiZE=20 MAXLENGTH=40  tag="14">
									</TD>
									
									<TD CLASS=TD5 NOWRAP>권한</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboRole_type" ALT="권한" CLASS ="cbonormal" TAG="11XXXU"></SELECT></TD>
									
									</TD>
									
								</TR>
								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
</FORM>				
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=*  WIDTH=100% VALIGN=TOP>						
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD>
									<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="24" TITLE="SPREAD" id=OBJECT1>
										<PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0">
									</OBJECT>
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
        			<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
						<IMG SRC="../../../CShared/image/query_d.gif" Style="CURSOR: hand" ALT="Search" NAME="search" OnClick="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT>
					<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
					<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>


<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>  
