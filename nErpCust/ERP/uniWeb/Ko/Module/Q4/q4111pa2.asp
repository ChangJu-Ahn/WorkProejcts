<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q4111PA2
'*  4. Program Name         : 
'*  5. Program Desc         : 검사 Release 팝업 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
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
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Dim ArrParent

Dim arrParam1
Dim arrParam2				

Dim arrReturn				'--- Return Parameter Group 


Dim IsOpenPop          

Dim lgPlantCd

<!-- #Include file="../../inc/lgvariables.inc" -->	

'------ Set Parameters from Parent ASP ------ 
ArrParent = window.dialogArguments
Set PopupParent = ArrParent(0)

arrParam1 = arrParent(1)
arrParam2 = arrParent(2)

top.document.title = PopupParent.gActivePRAspName
'--------------------------------------------- 

Function InitVariables()
	
End Function

Sub SetDefaultVal()
	
	txtReleaseDt.Text = arrParam1(0)
	txtGoodsQty.Text = arrParam1(1)
	txtDefectivesQty.Text = arrParam1(2)
	txtGoodsSLCd.value = arrParam1(3)
	txtGoodsSLNm.value = arrParam1(4)
	txtDefectivesSLCd.value = arrParam1(5)
	txtDefectivesSLNm.value = arrParam1(6)
	
	Call ProtectFields(arrParam2(0), arrParam2(1), arrParam2(2), arrParam2(3), arrParam2(4))
	lgPlantCd = arrParam2(5)
	
	Self.Returnvalue = Array("")
End Sub

Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","PA") %>
End Sub

Sub ProtectFields(Byval strIFYesNo, Byval strInspClassCd, Byval strReceivingInspType, Byval strAutoPR, Byval strAutoST)
	If strIFYesNo = "N" Then
		'자체 검사의뢰인 경우는 양품창고/불량품창고 선택 사항 
		If UNICDbl(txtGoodsQty.Text) > 0 Then
			Call ggoOper.SetReqAttr(txtGoodsSLCd, "D")
		Else
			txtGoodsSLCd.value = ""
			txtGoodsSLNm.value = ""
			Call ggoOper.SetReqAttr(txtGoodsSLCd, "Q")
		End If
					
		If UNICDbl(txtDefectivesQty.Text) > 0 Then
			Call ggoOper.SetReqAttr(txtDefectivesSLCd, "D")
		Else
			txtDefectivesSLCd.value = ""
			txtDefectivesSLNm.value = ""
			Call ggoOper.SetReqAttr(txtDefectivesSLCd, "Q")
		End If
		
	Else
		Select Case strInspClassCd
			Case "R"
				If strReceivingInspType = "A" Then
					If strAutoST = "Y" then
						If UNICDbl(txtGoodsQty.Text) > 0 Then
							Call ggoOper.SetReqAttr(txtGoodsSLCd, "N")
						Else
							txtGoodsSLCd.value = ""
							txtGoodsSLNm.value = ""
							Call ggoOper.SetReqAttr(txtGoodsSLCd, "Q")
						End If
					
						If UNICDbl(txtDefectivesQty.Text) > 0 Then
							Call ggoOper.SetReqAttr(txtDefectivesSLCd, "N")
						Else
							txtDefectivesSLCd.value = ""
							txtDefectivesSLNm.value = ""
							Call ggoOper.SetReqAttr(txtDefectivesSLCd, "Q")
						End If
					Else
						txtGoodsSLCd.value = ""
						txtGoodsSLNm.value = ""
						txtDefectivesSLCd.value = ""
						txtDefectivesSLNm.value = ""
						Call ggoOper.SetReqAttr(txtGoodsSLCd, "Q")
						Call ggoOper.SetReqAttr(txtDefectivesSLCd, "Q")
					End If					
				ElseIf strReceivingInspType = "B" Then
					If strAutoPR = "Y" then
						If UNICDbl(txtGoodsQty.Text) > 0 Then
							Call ggoOper.SetReqAttr(txtGoodsSLCd, "N")
						Else
							txtGoodsSLCd.value = ""
							txtGoodsSLNm.value = ""
							Call ggoOper.SetReqAttr(txtGoodsSLCd, "Q")
						End If
					Else
						txtGoodsSLCd.value = ""
						txtGoodsSLNm.value = ""
						Call ggoOper.SetReqAttr(txtGoodsSLCd, "Q")
					End If					
					
					txtDefectivesSLCd.value = ""
					txtDefectivesSLNm.value = ""
					Call ggoOper.SetReqAttr(txtDefectivesSLCd, "Q")
				Else
					txtGoodsSLCd.value = ""
					txtGoodsSLNm.value = ""
					txtDefectivesSLCd.value = ""
					txtDefectivesSLNm.value = ""
					Call ggoOper.SetReqAttr(txtGoodsSLCd, "Q")
					Call ggoOper.SetReqAttr(txtDefectivesSLCd, "Q")
				End If
				
			Case "P"
				txtGoodsSLCd.value = ""
				txtGoodsSLNm.value = ""
				txtDefectivesSLCd.value = ""
				txtDefectivesSLNm.value = ""
				Call ggoOper.SetReqAttr(txtGoodsSLCd, "Q")
				Call ggoOper.SetReqAttr(txtDefectivesSLCd, "Q")
			Case "F"
				If UNICDbl(txtGoodsQty.Text) > 0 Then
					Call ggoOper.SetReqAttr(txtGoodsSLCd, "N")
				Else
					txtGoodsSLCd.value = ""
					txtGoodsSLNm.value = ""
					Call ggoOper.SetReqAttr(txtGoodsSLCd, "Q")
				End If
					
				If UNICDbl(txtDefectivesQty.Text) > 0 Then
					Call ggoOper.SetReqAttr(txtDefectivesSLCd, "N")
				Else
					txtDefectivesSLCd.value = ""
					txtDefectivesSLNm.value = ""
					Call ggoOper.SetReqAttr(txtDefectivesSLCd, "Q")
				End If
					
			Case "S"
				txtGoodsSLCd.value = ""
				txtGoodsSLNm.value = ""
				txtDefectivesSLCd.value = ""
				txtDefectivesSLNm.value = ""
				Call ggoOper.SetReqAttr(txtGoodsSLCd, "Q")
				Call ggoOper.SetReqAttr(txtDefectivesSLCd, "Q")
				
			Case Else
				txtGoodsSLCd.value = ""
				txtGoodsSLNm.value = ""
				txtDefectivesSLCd.value = ""
				txtDefectivesSLNm.value = ""
				Call ggoOper.SetReqAttr(txtGoodsSLCd, "Q")
				Call ggoOper.SetReqAttr(txtDefectivesSLCd, "Q")
		End Select
	End If		
End Sub

Sub OpenGoodsSL()
	Dim strCode
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Sub

	If UCase(txtGoodsSLCd.ClassName) = UCase(PopupParent.UCN_PROTECTED)  Then
		Exit Sub
	End If
	
	IsOpenPop = True
	strCode = Trim(txtGoodsSLCd.Value)
	arrParam(0) = "양품창고팝업"	
	arrParam(1) = "B_Storage_Location"				
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD =  " & FilterVar(lgPlantCd , "''", "S") & " AND SL_TYPE <> " & FilterVar("E", "''", "S") & " "    ' Where Condition
	arrParam(5) = "창고"			
	
    arrField(0) = "SL_CD"	
    arrField(1) = "SL_NM"	
    
    arrHeader(0) = "창고코드"		
    arrHeader(1) = "창고명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWIDTH=420px; dialogHEIGHT=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	txtGoodsSLCd.Focus	
	If arrRet(0) = "" Then
		Exit Sub
	Else
		txtGoodsSLCd.value = arrRet(0)   
		txtGoodsSLNm.value = arrRet(1)   
		lgBlnFlgChgValue = True 
		txtGoodsSLCd.Focus 
	End If	
	
	Set gActiveElement = document.activeElement
End Sub

Sub OpenDefectivesSL()
	Dim strCode
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Sub

	If UCase(txtDefectivesSLCd.ClassName) = UCase(PopupParent.UCN_PROTECTED)  Then
		Exit Sub
	End If
	
	IsOpenPop = True
	strCode = Trim(txtDefectivesSLCd.Value)
	arrParam(0) = "불량품창고팝업"	
	arrParam(1) = "B_Storage_Location"				
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD =  " & FilterVar(lgPlantCd , "''", "S") & " AND SL_TYPE <> " & FilterVar("E", "''", "S") & " "    ' Where Condition
	arrParam(5) = "창고"			
	
    arrField(0) = "SL_CD"	
    arrField(1) = "SL_NM"	
    
    arrHeader(0) = "창고코드"		
    arrHeader(1) = "창고명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWIDTH=420px; dialogHEIGHT=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	txtDefectivesSLCd.Focus
	If arrRet(0) = "" Then
		Exit Sub
	Else
		txtDefectivesSLCd.value = arrRet(0)   
		txtDefectivesSLNm.value = arrRet(1)  
		txtDefectivesSLCd.Focus
		lgBlnFlgChgValue = True 
	End If	
	
	Set gActiveElement = document.activeElement
End Sub

Function OKClick()
	
	On Error Resume Next
	'-----------------------
	'Precheck area
	'-----------------------
	If lgBlnFlgChgValue = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		Exit Function
	End If
    
	If Not chkField(Document, "2") Then                                  '⊙: Check contents area
    	Exit Function
    End If
    
	Redim arrReturn(4)
	
	arrReturn(0) = txtReleaseDt.Text
	arrReturn(1) = txtGoodsSLCd.value
	arrReturn(2) = txtGoodsSLNm.value
	arrReturn(3) = txtDefectivesSLCd.value
	arrReturn(4) = txtDefectivesSLNm.value
			
	Self.Returnvalue = arrReturn
	
	Self.Close()
End Function

Function CancelClick()
	On Error Resume Next
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
	Call LoadInfTB19029
	
	Call ggoOper.LockField(Document, "N")                                   	'⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec)
	
	Call SetDefaultVal()
	Call InitVariables
	
	txtReleaseDt.focus
	Set gActiveElement = document.activeElement 
End Sub

Sub Form_QueryUnload(Cancel, UnloadMode)
	
End Sub

Sub txtReleaseDt_DblClick(Button)
    If Button = 1 Then
        txtReleaseDt.Action = 7
    End If
End Sub

Sub txtReleaseDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtReleaseDt_KeyPress(KeyAscii)
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call OKClick()
	End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR HEIGHT=*>
		<TD  WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%>></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100% VALIGN=TOP>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>양품수</TD>        
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q4111pa2_txtGoodsQty_txtGoodsQty.js'></script>
									</TD>
							    	<TD CLASS="TD5" NOWRAP>불량품수</TD>        
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q4111pa2_txtDefectivesQty_txtDefectivesQty.js'></script>
									</TD>
							    </TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>Release일</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q4111pa2_txtReleaseDt_txtReleaseDt.js'></script>
									</TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
												
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>양품창고</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtGoodsSLCd" SIZE="10" MAXLENGTH="7" ALT="양품창고" TAG="22XXXU"><IMG ALIGN=top HEIGHT=20 NAME=btntxtGoodsSLPopup ONCLICK=vbscript:OpenGoodsSL() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtGoodsSLNm" TAG="24">
									</TD>
							    	<TD CLASS="TD5" NOWRAP>불량품창고</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtDefectivesSLCd" SIZE="10" MAXLENGTH="7" ALT="불량품창고" TAG="22XXXU"><IMG ALIGN=top HEIGHT=20 NAME=btntxtDefectivesSLPopup ONCLICK=vbscript:OpenDefectivesSL() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtDefectivesSLNm" TAG="24">
									</TD>
							    </TR>
							</TABLE>
						</FIELDSET>
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
					<TD WIDTH=90% ALIGN=RIGHT>
						<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
						<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>  
