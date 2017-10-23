<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : Sales																		*
'*  2. Function Name        : 								     										*
'*  3. Program ID           : S1511RA1																	*
'*  4. Program Name         : 품목참조										         					*
'*  5. Program Desc         : 품목그룹별품먹구성비등록을 위한 품목참조									*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2002/05/08																*
'*  8. Modified date(Last)  : 																			*
'*  9. Modifier (First)     : Cho inkuk																	*
'* 10. Modifier (Last)      : 																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              :
'=======================================================================================================
%>
<HTML>
<HEAD>
<!--TITLE>품목참조</TITLE-->
<TITLE></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBScript">

Option Explicit		

Const BIZ_PGM_ID 		= "s1511rb1.asp"                              '☆: Biz Logic ASP Name

Const C_MaxKey          = 4                                           '☆: key count of SpreadSheet

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim arrParent
ArrParent = window.dialogArguments
Set PopupParent  = ArrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim lgSelectList                   
Dim lgSelectListDT                 

Dim lgSortFieldNm                  
Dim lgSortFieldCD                  

Dim lgPopUpR                       

Dim lgKeyPos                       
Dim lgKeyPosVal                    
Dim lgCookValue 

Dim IsOpenPop  
Dim gblnWinEvent

Dim arrReturn										<% '--- Return Parameter Group %>
Dim arrParam	

Dim lgIsOpenPop                                          

'========================================================================================================
Function InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1
			
	gblnWinEvent = False
	ReDim arrReturn(0,0)
	Self.Returnvalue = arrReturn
End Function

'========================================================================================================
	Sub SetDefaultVal()

		Dim arrTemp
		
		arrTemp = Split(ArrParent(1), PopupParent.gColSep)

		frm1.txtItemGroup.Value = arrTemp(0)
		frm1.txtItemGroupNm.value = arrTemp(1)		

	End Sub

'========================================================================================================
	Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
		'------ Developer Coding part (Start ) -------------------------------------------------------------- 

		<% Call loadInfTB19029A("Q", "S", "NOCOOKIE", "RA") %>                                '☆: 

		'------ Developer Coding part (End )   -------------------------------------------------------------- 
	End Sub

'========================================================================================================
	Sub InitSpreadSheet()
		Call SetZAdoSpreadSheet("S1511RA1","S","A","V20021210", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
		With frm1.vspdData
			ggoSpread.Source = frm1.vspdData
			.OperationMode = 5
			Call SetSpreadLock 
		End With

	End Sub


'========================================================================================================
	Sub SetSpreadLock()
	    With frm1
	    .vspdData.ReDraw = False
		'------ Developer Coding part (Start ) -------------------------------------------------------------- 
		ggoSpread.SpreadLockWithOddEvenRowColor()
		'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	    .vspdData.ReDraw = True

	    End With
	End Sub


'========================================================================================================
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

'========================================================================================================
	Function OKClick()
	
		Dim intColCnt, intRowCnt, intInsRow

		If frm1.vspdData.SelModeSelCount > 0 Then 

			intInsRow = 0

			Redim arrReturn(frm1.vspdData.SelModeSelCount, frm1.vspdData.MaxCols)

			For intRowCnt = 1 To frm1.vspdData.MaxRows

				frm1.vspdData.Row = intRowCnt

				If frm1.vspdData.SelModeSelected Then
					For intColCnt = 0 To frm1.vspdData.MaxCols - 2
						frm1.vspdData.Col = Getkeypos("A",intColCnt + 1)						
						arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
					Next

					intInsRow = intInsRow + 1

				End IF
			Next
		End if			
		
		Self.Returnvalue = arrReturn
		Self.Close()
	End Function	

'========================================================================================================
	Function CancelClick()
		Redim arrReturn(1,1)
		arrReturn(0,0) = ""
		Self.Returnvalue = arrReturn
		Self.Close()
	End Function

'========================================================================================================
	Sub Form_Load()
		Call LoadInfTB19029				                                           '⊙: Load table , B_numeric_format
   
		'------ Developer Coding part (Start ) -------------------------------------------------------------- 
		Call ggoOper.LockField(Document, "N")						<% '⊙: Lock  Suitable  Field %>

		Call InitVariables														    '⊙: Initializes local global variables
		Call SetDefaultVal	
		Call InitSpreadSheet()
		Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/zpConfig.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
		Call FncQuery()
	End Sub

'========================================================================================================
	Function vspdData_DblClick(ByVal Col, ByVal Row)
	    If Row = 0 Then  Exit Function
		If frm1.vspdData.ActiveRow > 0 Then  Call OKClick
	End Function

'========================================================================================================
    Function vspdData_KeyPress(KeyAscii)
         On Error Resume Next
         If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
            Call OKClick()
         ElseIf KeyAscii = 27 Then
            Call CancelClick()
         End If
    End Function

'========================================================================================================
	Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
    If OldLeft <> NewLeft Then Exit Sub

		If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
			If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
				Call DbQuery()
			End If
		End If
	End Sub	


'========================================================================================
Function FncQuery() 
    
    FncQuery = False                                                 
    
    Err.Clear 
    
	Call ggoOper.ClearField(Document, "2")							
	Call InitVariables												

	If Not chkField(Document, "1") Then				
		Exit Function
	End If

    If DbQuery = False Then Exit Function							

    FncQuery = True									
        
End Function

'********************************************************************************************************
	Function DbQuery()
		Err.Clear															<%'☜: Protect system from crashing%>

		DbQuery = False														<%'⊙: Processing is NG%>

		If LayerShowHide(1) = False Then
			Exit Function
		End If

		Dim strVal

		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				<%'☜: 비지니스 처리 ASP의 상태 %>
			strVal = strVal & "&txtItemGroup=" & Trim(frm1.txtItemGroup.value)		<%'☆: 조회 조건 데이타 %>
			strVal = strVal & "&txtItem=" & Trim(frm1.HItem.value)				
			strVal = strVal & "&txtItemNm=" & Trim(frm1.HItemNm.value)
			strVal = strVal & "&txtItemSpec=" & Trim(frm1.txtHItemSpec.value)
			
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				<%'☜: 비지니스 처리 ASP의 상태 %>
			strVal = strVal & "&txtItemGroup=" & Trim(frm1.txtItemGroup.value)		<%'☆: 조회 조건 데이타 %>
			strVal = strVal & "&txtItem=" & Trim(frm1.txtItem.Value)				
			strVal = strVal & "&txtItemNm=" & Trim(frm1.txtItemNm.Value)
			strVal = strVal & "&txtItemSpec=" & Trim(frm1.txtItemSpec.value)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		End If

    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal =     strVal & "&lgPageNo="       & lgPageNo                          '☜: Next key tag
		strVal =     strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal =     strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
       	strVal =     strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

		Call RunMyBizASP(MyBizASP, strVal)									<%'☜: 비지니스 ASP 를 가동 %>

		DbQuery = True														<%'⊙: Processing is NG%>
	End Function

'========================================================================================
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then				 
		frm1.vspdData.Focus
		'frm1.vspdData.Row = 1	
		'frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtItem.focus
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
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
						<TD CLASS="TD5" NOWRAP>품목그룹</TD>
						<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtItemGroup" SIZE=10 MAXLENGTH=10 TAG="14" ALT="품목그룹">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=20 TAG="14"></TD>
						<TD CLASS="TD5" NOWRAP>품목</TD>
						<TD CLASS="TD6"><INPUT NAME="txtItem" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=18 tag="11XXXU">&nbsp;<INPUT NAME="txtItemNm" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="11XXXU"></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>규격</TD>
						<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtItemSpec" SIZE=33 MAXLENGTH=50 TAG="11" ALT="규격"></TD>
						<TD CLASS="TD5" NOWRAP></TD>
						<TD CLASS="TD6" NOWRAP></TD>
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
						<script language =javascript src='./js/s1511ra1_vaSpread_vspdData.js'></script>
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
					<TD>&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"     ONCLICK="FncQuery()"     ></IMG>
									<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  OnClick="OpenSortPopup()"></IMG></TD>
					<TD ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"        ONCLICK="OkClick()"      ></IMG>
							        <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"    ONCLICK="CancelClick()"  ></IMG></TD>				
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="HItem" tag="24">
<INPUT TYPE=HIDDEN NAME="HItemNm" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHItemSpec" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>