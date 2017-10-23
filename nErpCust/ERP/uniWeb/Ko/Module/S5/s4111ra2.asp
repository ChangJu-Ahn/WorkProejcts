<%@ LANGUAGE="VBSCRIPT" %>
<%
'************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S4111RA2.ASP
'*  4. Program Name         : 운송정보참조 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/17
'*  8. Modified date(Last)  : 2002/12/17
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : SON BUM YEOL
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>운송정보참조</TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit
<!-- #Include file="../../inc/lgvariables.inc" --> 
	
Dim gblnWinEvent
Dim arrReturn
Dim lgIsOpenPop

Dim arrParent
Dim PopupParent

arrParent = window.dialogArguments
Set PopupParent  = ArrParent(0)
top.document.title = PopupParent.gActivePRAspName

Const BIZ_PGM_ID        = "s4111rb2.asp"
Const C_MaxKey          = 5

'=========================================
Sub InitVariables()
         
    lgBlnFlgChgValue = False                               
    lgStrPrevKey     = ""                                  
    lgSortKey        = 1
    lgPageNo         = ""
	lgIntFlgMode = PopupParent.OPMD_CMODE	
     Redim arrReturn(0)        
     Self.Returnvalue = arrReturn     
     
End Sub

'=========================================
Sub SetDefaultVal()
End Sub

'=========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "*", "NOCOOKIE", "PA") %>
End Sub

'=========================================
Sub InitSpreadSheet()
    '그룹일 경우 C_GROUP_DBAGENT 
	Call SetZAdoSpreadSheet("s4111ra2","S","A","V20021211", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock 
	

End Sub

'=========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()	
End Sub

'========================================
Function OKClick()

	Dim intColCnt
		
	If frm1.vspdData.ActiveRow > 0 Then	
		
		Redim arrReturn(C_MaxKey - 1)
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
			
		For intColCnt = 0 To C_MaxKey -1
			frm1.vspdData.Col = GetKeyPos("A",intColCnt + 1)		
			arrReturn(intColCnt) = frm1.vspdData.Text
		Next	
					
	End If
		
	Self.Returnvalue = arrReturn
	Self.Close()
	
End Function

'=========================================
Function CancelClick()
	Self.Close()
End Function

'=========================================
Function OpenTransCo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "운송회사"							
	arrParam(1) = "B_MAJOR A , B_MINOR B"						
	arrParam(2) = ""										
	arrParam(3) = ""									
	arrParam(4) = " A.MAJOR_CD = B.MAJOR_CD AND B.MAJOR_CD = " & FilterVar("B9031", "''", "S") & " "				
	arrParam(5) = "운송회사"							

	arrField(0) = "B.MINOR_CD"								
	arrField(1) = "B.MINOR_NM"								

	arrHeader(0) = "순번"							
	arrHeader(1) = "운송회사명"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	frm1.txtTransCo.focus
	
	If arrRet(0) <> "" Then
		Call SetTransCo(arrRet)
	End If
End Function

'=========================================
Function OpenVehicleNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "차량번호"							
	arrParam(1) = "B_MAJOR A , B_MINOR B"						
	arrParam(2) = ""			
	arrParam(3) = ""									
	arrParam(4) = " A.MAJOR_CD = B.MAJOR_CD AND B.MAJOR_CD = " & FilterVar("B9032", "''", "S") & " "				
	arrParam(5) = "차량번호"							

	arrField(0) = "B.MINOR_CD"								
	arrField(1) = "B.MINOR_NM"								

	arrHeader(0) = "순번"							
	arrHeader(1) = "차량번호"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	frm1.txtVehicleNo.focus
	
	If arrRet(0) <> "" Then
		Call SetVehicleNo(arrRet)
	End If
End Function

'=========================================
Function SetTransCo(arrRet)
	frm1.txtTransCo.value = arrRet(1)
End Function

'=========================================
Function SetVehicleNo(arrRet)
	frm1.txtVehicleNo.value = arrRet(1)
End Function

'========================================
Function OpenSortPopup()
	
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
	End If
End Function

'=========================================
Sub Form_Load()
    Call LoadInfTB19029			
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitVariables			    
	Call SetDefaultVal
	Call InitSpreadSheet()
	Call FncQuery()
End Sub

'=========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'=========================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'=========================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

'=========================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then 	Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	  '☜: 재쿼리 체크	
    	If lgPageNo <> "" Then                   '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
		    If CheckRunningBizProcess Then Exit Sub
			Call DbQuery
    	End If
	End If
End Sub

'=========================================
Sub txtTransCo_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'=========================================
Sub txtVehicleNo_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'=========================================
Sub txtSender_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub


'=========================================
Sub txtDriver_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'=========================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function
	
'=========================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
		
	If Row = 0 Or frm1.vspdData.MaxRows = 0 Then 
        Exit Function
	End If
		
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function
	
'=========================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'=========================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'=========================================
Function FncQuery() 
	
	FncQuery = False                                                        
    
    Err.Clear                                                               

    Call ggoOper.ClearField(Document, "2")									
    Call InitVariables 														
    
    Call DbQuery															'☜: Query db data

    FncQuery = True		
End Function

'=========================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
    
    With frm1
       If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001							
		    strVal = strVal & "&txtTransCo=" & Trim(frm1.txtHTransCo.value)				
			strVal = strVal & "&txtSender=" & Trim(frm1.txtHSender.value)						    
		    strVal = strVal & "&txtVehicleNo=" & Trim(frm1.txtHVehicleNo.value)				
			strVal = strVal & "&txtDriver=" & Trim(frm1.txtHDriver.value)						    

	   Else
		    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001							
		    strVal = strVal & "&txtTransCo=" & Trim(frm1.txtTransCo.value)				
			strVal = strVal & "&txtSender=" & Trim(frm1.txtSender.value)						    
		    strVal = strVal & "&txtVehicleNo=" & Trim(frm1.txtVehicleNo.value)				
			strVal = strVal & "&txtDriver=" & Trim(frm1.txtDriver.value)						    

       End if   
			strVal = strVal & "&lgPageNo="       & lgPageNo                '☜: Next key tag
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
			
       Call RunMyBizASP(MyBizASP, strVal)										
			
	End With

	DbQuery = True

End Function

'=========================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
    
	lgIntFlgMode = PopupParent.OPMD_UMODE
    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus
    Else
       frm1.txtTransCo.focus	
    End if

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
	<TABLE <%=LR_SPACE_TYPE_20%>>
		<TR>
			<TD <%=HEIGHT_TYPE_02%> WIDTH=100%>
				<FIELDSET CLASS="CLSFLD">
					<TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
							<TD CLASS=TD5>운송회사</TD>
							<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtTransCo" SIZE=20 MAXLENGTH=50 TAG="11XXXX" ALT="운송회사"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransCo" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenTransCo()"></TD>
							<TD CLASS=TD5 NOWRAP>인계자명</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSender" ALT="인계자명" TYPE="Text" MAXLENGTH="50" SIZE= 20 tag="11XXXX"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5>차량번호</TD>
							<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtVehicleNo" SIZE=20 MAXLENGTH=20 TAG="11XXXX" ALT="차량번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVehicleNo" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenVehicleNo()"></TD>
							<TD CLASS=TD5 NOWRAP>운전자명</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDriver" ALT="운전자명" TYPE="Text" MAXLENGTH="50" SIZE= 20 tag="11XXXX"></TD>
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
					<TR>
						<TD HEIGHT="100%" NOWRAP>
							<script language =javascript src='./js/s4111ra2_vaSpread_vspdData.js'></script>
						</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD <%=HEIGHT_TYPE_01%>></TD>
		</TR>
		<TR HEIGHT="20">
			<TD WIDTH="100%">
				<TABLE <%=LR_SPACE_TYPE_30%>>
					<TR>
						<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
						<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG>
	                    <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" OnClick="OpenSortPopup()" ></IMG>
						</TD>
						<TD WIDTH=30% ALIGN=RIGHT>
						<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
						<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
						</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0  TABINDEX ="-1"></IFRAME></TD>
		</TR>
	</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">

<INPUT TYPE=HIDDEN NAME="txtHTransCo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHSender" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHVehicleNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHDriver" TAG="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>


