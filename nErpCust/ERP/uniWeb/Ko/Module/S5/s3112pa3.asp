<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 수주관리 
'*  3. Program ID           : s3112pa1
'*  4. Program Name         : 품목팝업 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2003/06/11
'*  9. Modifier (First)     : Kim Hyungsuk
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18  Date 표준적용 
'*                            2002/12/26 UI 성능향상 강준구 
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>품목</TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgIsOpenPop

Dim arrReturn					
Dim arrParent
Dim arrParam

ArrParent = window.dialogArguments
Set PopupParent  = ArrParent(0)
top.document.title = PopupParent.gActivePRAspName

Const BIZ_PGM_ID        = "s3112pb3.asp"
Const C_MaxKey          = 4                                    '☆☆☆☆: Max key value

'=========================================
Sub InitVariables()
    lgBlnFlgChgValue = False                               
    lgStrPrevKey     = ""                                  
    lgSortKey        = 1
	
	Redim arrReturn(0)
	Self.Returnvalue = arrReturn
End Sub

'=========================================
Sub SetDefaultVal()

	arrParam = arrParent(1)

	frm1.txtItem.value = arrParam(0)
	frm1.txtPlant.value = arrParam(1) 
	frm1.txtSLCd.value = arrParam(2) 
End Sub

'=========================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "PA") %>
End Sub

'=========================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S3112PA3","S","A","V20021106", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
    Call SetSpreadLock 
    
    frm1.vspdData.OperationMode = 3
End Sub

'=========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'=========================================
Function OpenJnlItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "품목계정"								
	arrParam(1) = "A_JNL_ITEM"									
	arrParam(2) = Trim(frm1.txtJnlItem.value)					
	arrParam(3) = ""											
	arrParam(4) = "JNL_TYPE = " & FilterVar("IA", "''", "S") & ""								
	arrParam(5) = "품목계정"								

	arrField(0) = "JNL_CD"										
	arrField(1) = "JNL_NM"										

	arrHeader(0) = "품목계정"								
	arrHeader(1) = "품목계정명"								

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	frm1.txtJnlItem.focus
	
	If arrRet(0) <> "" Then
		frm1.txtJnlItem.value = arrRet(0)
		frm1.txtJnlItemNm.value = arrRet(1)
	End If
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
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call FncQuery()
End Sub

'=========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.ActiveRow > 0 Then	Call OKClick
End Function

'=======================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess Then Exit Sub
			Call DbQuery
		End If
   End if
    
End Sub

'========================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then   'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function

'========================================
Function OKClick()
		
	Redim arrReturn(3)
	If frm1.vspdData.ActiveRow > 0 Then				
		
		frm1.vspdData.Row = frm1.vspdData.ActiveRow

		frm1.vspdData.Col = GetKeyPos("A",1) ' 1
		arrReturn(0) = frm1.vspdData.Text
			
		frm1.vspdData.Col = GetKeyPos("A",2) ' 2
		arrReturn(1) = frm1.vspdData.Text

		frm1.vspdData.Col = GetKeyPos("A",3) ' 9
		arrReturn(2) = frm1.vspdData.Text
			
		frm1.vspdData.Col = GetKeyPos("A",4) '10
		arrReturn(3) = frm1.vspdData.Text
			
		Self.Returnvalue = arrReturn
	End If

	Self.Close()
End Function

'========================================
Function CancelClick()
	Self.Close()
End Function

'========================================
Function FncQuery() 

    FncQuery = False                                                        
    
    Err.Clear                                                               

    Call ggoOper.ClearField(Document, "2")									
    Call InitVariables 														
    
    Call DbQuery															'☜: Query db data

    FncQuery = True		
End Function

'=====================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               

	If LayerShowHide(1) = False Then
		Exit Function
	End If
    
    With frm1

		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					
		strVal = strVal & "&txtItem=" & Trim(frm1.txtItem.value)		
		strVal = strVal & "&txtItemNm=" & Trim(frm1.txtItemNm.value)	
		strVal = strVal & "&txtJnlItem=" & Trim(frm1.txtJnlItem.value)
		strVal = strVal & "&txtPlant=" & Trim(frm1.txtPlant.value)
		strVal = strVal & "&txtSLCD=" & Trim(frm1.txtSLCD.value)
		
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
  
        Call RunMyBizASP(MyBizASP, strVal)										
    End With
    
    DbQuery = True
End Function

'=====================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
	Else
		frm1.txtItem.focus
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>	
						<TD CLASS=TD5 NOWRAP>품목</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT NAME="txtItem" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="11XXXU" ALT="품목"></TD>
						<TD CLASS=TD5 NOWRAP>품목계정</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT NAME="txtJnlItem" TYPE="Text" MAXLENGTH="20" SIZE=10 tag="11XXXU" ALT="품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnJnlItem" align=top TYPE="BUTTON" OnClick="vbscript:OpenJnlItem">&nbsp;
							<INPUT NAME="txtJnlItemNm" TYPE="Text" SIZE=20 tag="24">
						</TD>
					</TR>	
					<TR>
						<TD CLASS=TD5 NOWRAP>품목명</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemNm" TYPE="Text" SIZE=30 MAXLENGTH="50" ALT="품목명" tag="11"></TD>
						<TD CLASS="TD5" NOWRAP>공장</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtPlant" ALT="공장" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="14XXXU">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT NAME="txtPlantNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
					</TR>	
					<TR>
						<TD CLASS=TD5 NOWRAP>창고</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT NAME="txtSLCd" TYPE="Text" MAXLENGTH="20" SIZE=5 tag="14XXXU" ALT="창고">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT NAME="txtSLNm" TYPE="Text" SIZE=20 tag="24">
						</TD>
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
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT="100%">
						<script language =javascript src='./js/s3112pa3_vaSpread_vspdData.js'></script>
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
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
	                    <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" OnClick="OpenSortPopup()" ></IMG>
					</TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX ="-1"></IFRAME></TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
