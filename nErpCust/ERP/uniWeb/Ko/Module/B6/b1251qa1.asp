<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : b1251qa1
'*  4. Program Name         : 구매그룹조회 
'*  5. Program Desc         : 구매그룹조회 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/06/08
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'******************************************  1.1 Inc 선언   **********************************************-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance

'****************************************  1.2 Global 변수/상수 선언  ***********************************
Const BIZ_PGM_ID 	= "b1251qb1.asp"
Const BIZ_PGM_JUMP_ID 	= "b1251ma1"

Const C_MaxKey          = 12                                           '☆: key count of SpreadSheet

<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
Dim IsOpenPop  
Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 
Dim lgIsOpenPop

'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================
Function FncSplitColumn()
    
  If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Function

'========================================================================================
Function CookiePage(Byval Kubun)
	Dim strTemp, arrVal
	Dim IntRetCD

	If Kubun = 1 Then	   
		if frm1.vspdData.ActiveRow  > 0 then		
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.Col = GetKeyPos("A",1) 'C_GroupCd
			WriteCookie "GroupCd" , frm1.vspdData.Text
			frm1.vspdData.Col = GetKeyPos("A",2) 'C_GroupNm
			WriteCookie "GroupNm" , frm1.vspdData.Text
		end if
	
		Call PgmJump(BIZ_PGM_JUMP_ID)
	else 
	    If ReadCookie("Kubun") = "Y" then 
			frm1.txtOrgCd.value		= ReadCookie ("OrgCd")
			frm1.txtGroupCd.value	= ReadCookie ("GroupCd")
			
		    WriteCookie "OrgCd" , ""
	    	WriteCookie "Kubun" , ""
			WriteCookie "GroupCd" , ""
	    	
	    	Call MainQuery()
	    	
	   	End if
	End IF
End Function

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
	lgStrPrevKey     = ""								   'initializes Previous Key
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = Parent.OPMD_CMODE                          'Indicates that current mode is Create mode
                
    gblnWinEvent = False
End Sub

Sub SetDefaultVal()

End Sub
 
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
End Sub

'==========================================  2.2.3 InitSpreadSheet()  ===================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("B1251QA1","S","A","V20030331",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Call SetSpreadLock("A")      
End Sub

'============================================ 2.2.4 SetSpreadLock()  ====================================
Sub SetSpreadLock(ByVal pOpt)
    If pOpt = "A" Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
	Else
	
	End If
End Sub

'------------------------------------  PopZAdoConfigGrid()  ----------------------------------------------
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenSortPopup("A")
End Sub

'========================================================================================================
Function OpenSortPopup(ByVal pSpdNo)
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'------------------------------------------  OpenORG()  -------------------------------------------------
Function OpenORG()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매조직"						<%' 팝업 명칭 %>
	arrParam(1) = "B_PUR_ORG"						<%' TABLE 명칭 %>
	
	arrParam(2) = UCase(Trim(frm1.txtORGCd.Value)) 	<%' Code Condition%>
'	arrParam(3) = Trim(frm1.txtORGNm.Value)	<%' Name Cindition%>
	
	arrParam(4) = ""							<%' Where Condition%>
	arrParam(5) = "구매조직"							<%' TextBox 명칭 %>
	
    arrField(0) = "Pur_Org"					<%' Field명(0)%>
    arrField(1) = "Pur_Org_NM"					<%' Field명(1)%>
    
    arrHeader(0) = "구매조직"						<%' Header명(0)%>
    arrHeader(1) = "구매조직명"						<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtORGCd.focus
		Exit Function
	Else
		frm1.txtORGCd.value = arrRet(0)
		frm1.txtORGNm.value = arrRet(1)
		frm1.txtORGCd.focus
	End If	
End Function


'------------------------------------------  OpenBA()  -------------------------------------------------
Function OpenBA()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사업장"	
	arrParam(1) = "B_BIZ_AREA"
	arrParam(2) = UCase(Trim(frm1.txtBaCd.Value))
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "사업장"			
	
    arrField(0) = "BIZ_AREA_CD"	
    arrField(1) = "BIZ_AREA_NM"	
    
    arrHeader(0) = "사업장"		
    arrHeader(1) = "사업장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBaCd.focus
		Exit Function
	Else
		frm1.txtBACd.Value = arrRet(0)
		frm1.txtBANm.Value = arrRet(1)
		frm1.txtBaCd.focus
	End If	
	
End Function

'------------------------------------------  OpenGroup()  -------------------------------------------------
Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"						<%' 팝업 명칭 %>
	arrParam(1) = "B_PUR_GRP A, B_PUR_ORG B "						<%' TABLE 명칭 %>
	
	arrParam(2) = UCase(Trim(frm1.txtGroupCd.Value))	<%' Code Condition%>
	arrParam(3) = Trim(frm1.txtGroupNm.Value)	<%' Name Cindition%>
	
	arrParam(4) = "A.PUR_ORG = B.PUR_ORG AND A.PUR_ORG >= " & FilterVar(UCase(frm1.txtORGCd.Value), "''", "S") & " " <%' Where Condition%>
	arrParam(5) = "구매그룹"							<%' TextBox 명칭 %>
	
    arrField(0) = "A.Pur_Grp"					<%' Field명(0)%>
    arrField(1) = "A.Pur_Grp_NM"					<%' Field명(1)%>
    arrField(2) = "A.Pur_Org"					<%' Field명(0)%>
    arrField(3) = "B.Pur_Org_NM"					<%' Field명(1)%>
    
    arrHeader(0) = "구매그룹"						<%' Header명(0)%>
    arrHeader(1) = "구매그룹명"						<%' Header명(1)%>
    arrHeader(2) = "구매조직"						<%' Header명(0)%>
    arrHeader(3) = "구매조직명"						<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtGroupCd.focus
		Exit Function
	Else
		frm1.txtGroupCd.value = arrRet(0)
		frm1.txtGroupNm.value = arrRet(1)
		frm1.txtOrgCd.value = arrRet(2)
		frm1.txtOrgNm.value = arrRet(3)
		frm1.txtGroupCd.focus
	End If	
End Function

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
	Call LoadInfTB19029				                                           '⊙: Load table , B_numeric_format
'    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")                                       '⊙: Lock  Suitable  Field

    Call InitVariables
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("11000000000011")							'⊙: 버튼 툴바 제어 

	frm1.txtORGCd.focus
    Set gActiveElement = document.activeElement 
    Call CookiePage(0)
End Sub
 
'********************************************************************************************************* %>
Sub vspdData_Click(ByVal Col, ByVal Row)
    Set gActiveSpdSheet = frm1.vspdData
    SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"  
    If Row = 0 Then
    
    	frm1.vspdData.OperationMode = 0
    
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
    else
    	frm1.vspdData.OperationMode = 0
    End if
End Sub

'==========================================================================================
Sub vspdData_DragDropBlock(ByVal Col , ByVal Row , ByVal Col2 , ByVal Row2 , ByVal NewCol , ByVal NewRow , ByVal NewCol2 , ByVal NewRow2 , ByVal Overwrite , Action , DataOnly , Cancel )
    
    Row = 0: Row2 = -1: NewRow = 0
    ggoSpread.SwapRange Col, Row, Col2, Row2, NewCol, NewRow, Cancel
    
End Sub

'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If
    

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub

'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        <%'⊙: Processing is NG%>
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>

    '-----------------------
    'Check condition area
    '----------------------- 
'    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
'       Exit Function
'    End If
    
    '-----------------------
    'Erase contents area
    '----------------------- 
'    Call ggoOper.ClearField(Document, "2")										<%'⊙: Clear Contents  Field%>
	frm1.vspdData.maxrows = 0
    Call InitVariables															<%'⊙: Initializes local global variables%>

    '-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery																<%'☜: Query db data%>

    FncQuery = True																<%'⊙: Processing is OK%>
        
End Function

'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    Call ggoOper.ClearField(Document, "1")                                         '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                         '⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
    Call InitVariables                                                      '⊙: Initializes local global variables
    Call SetDefaultVal
    frm1.vspdData.MaxRows = 0
    Call SetToolbar("11100000000000")
    FncNew = True                                                           '⊙: Processing is OK

End Function

'========================================================================================
Function FncCancel() 
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function

'========================================================================================
Function FncPrint()
	ggoSpread.Source = frm1.vspdData 
    Call parent.FncPrint()
End Function

'========================================================================================
Function FncExcel()
	ggoSpread.Source = frm1.vspdData  
    Call parent.FncExport(Parent.C_Multi)												<%'☜: 화면 유형 %>
End Function

'========================================================================================
Function FncFind() 
	ggoSpread.Source = frm1.vspdData 
    Call parent.FncFind(Parent.C_Multi , False)                                     <%'☜:화면 유형, Tab 유무 %>
End Function

'========================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
	Call LayerShowHide(1)
    
    With frm1

	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtORGCd=" & Trim(.txtORGCd.value)
		strVal = strVal & "&txtGroupCd=" & Trim(.txtGroupCd.value)
		strVal = strVal & "&txtBACd=" & .txtBACd.value
		if .rdoUseflg1.checked = true  then
			strVal = strVal & "&rdoUseflg=" & "A"
		elseif .rdoUseflg2.checked = true  then
			strVal = strVal & "&rdoUseflg=" & "Y"
		elseif .rdoUseflg3.checked = true  then
			strVal = strVal & "&rdoUseflg=" & "N"
		end if 		
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag

        strVal =     strVal & "&lgPageNo="       & lgPageNo                          '☜: Next key tag
		strVal =     strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
         
        strVal =     strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal =     strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
    
		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True

End Function
 
'========================================================================================
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>
	
	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode

    Call ggoOper.LockField(Document, "Q")									<%'⊙: This function lock the suitable field%>
	Call SetToolbar("11000000000111")
	frm1.vspdData.focus
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
</HEAD>
<!--
'#########################################################################################################
'       					6. Tag부 
'######################################################################################################### 
-->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="*" align=right></td>
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
								<tr>
									<TD CLASS="TD5">구매조직</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ALT="구매조직" NAME="txtORGCd" SIZE=10 MAXLENGTH=4  tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenORG()">
													<INPUT TYPE=TEXT ALT="구매조직" ID="txtORGNm" NAME="arrCond" tag="14X"></TD>
									<TD CLASS="TD5">구매그룹</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ALT="구매그룹" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4   tag="1XNXXU" onchange= "vbscript:ChkGroup()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()">
													<INPUT TYPE=TEXT ALT="구매그룹" ID="txtGroupNm" NAME="arrCond" tag="14X"></TD>
								</tr>
								<tr>
									<TD class="TD5">사업장</TD>
									<TD class="TD6"><INPUT TYPE=TEXT ALT="사업장" NAME="txtBACd" SIZE=10 MAXLENGTH=10  tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBa()">
													<INPUT TYPE=TEXT ALT="사업장" ID="txtBANm" NAME="arrCond" tag="14X"></TD>
									<TD CLASS="TD5">사용여부</TD>
									<TD CLASS="TD6"><INPUT TYPE=radio ALT="사용여부" NAME="rdoUseflg" class="radio" Value="A" id="rdoUseflg1" checked tag="1X">
													<label for="rdoUseflg1">전체</label>
													<INPUT TYPE=radio ALT="사용여부" NAME="rdoUseflg" class="radio" Value="Y" id="rdoUseflg2" tag="1X">
													<label for="rdoUseflg2">예</label>
													<INPUT TYPE=radio ALT="사용여부" NAME="rdoUseflg" class="radio" Value="N" id="rdoUseflg3" tag="1X">
													<label for="rdoUseflg3">아니오</label></TD>
								</tr>
							</TABLE>
						</FIELDSET>
					</TD>
			</TR>
			<TR>
				<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>
			<TR>
				<TD WIDTH=100% valign=top>
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
		</TABLE></TD>
	</TR>
	
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
		<td WIDTH="100%">
			<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<td WIDTH="*" align="right"><a ONCLICK="VBSCRIPT:CookiePage(1)">구매그룹등록</a></td>
					<td WIDTH="10"></td>
				</tr>
			</table>
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
