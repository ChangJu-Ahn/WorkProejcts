<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m1311qa1
'*  4. Program Name         : 외주pl
'*  5. Program Desc         :
'*  6. Modified date(First) : MHJ
'*  7. Modified date(Last)  : Kim Jin Ha
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : 2003-06-02
'* 10. Comment              :
'* 11. Common Coding Guide  :      
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'********************************************************************************************************* !-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<!--'==========================================  1.1.1 Style Sheet  ============================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ===========================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit					

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgIsOpenPop                                           
Dim lgSortKey_A                                            
Dim lgPageNo1                                       
Dim lgSortKey_B                                           
Dim lgKeyPos                                           
Dim lgKeyPosVal                                            
Dim	lgTopLeft
Dim IscookieSplit 
Dim lgSaveRow                                           
Dim Query_Msg_Flg

Const BIZ_PGM_ID 		= "m1311qb1.asp"  
Const BIZ_PGM_ID1       = "m1311qb2.asp"
Const BIZ_PGM_JUMP_ID 	= "m1311ma1"
Const C_MaxKey			  = 11			
'===================================================================================================================================
Function setCookie()

	if GetKeyPosVal("A",1) <> "" then		
		WriteCookie "m1311Supplier", Trim(GetKeyPosVal("A",1))
		WriteCookie "m1311Plant", Trim(GetKeyPosVal("A",3))
		WriteCookie "m1311Item", Trim(GetKeyPosVal("A",5))
	end if
	
	Call PgmJump(BIZ_PGM_JUMP_ID)

End Function
'===================================================================================================================================
Sub InitVariables()

	lgBlnFlgChgValue = False                               'Indicates that no value changed

    lgPageNo   = ""                                  'initializes Previous Key for spreadsheet #1
    lgSortKey_A      = 1
    lgPageNo1   = ""                                  'initializes Previous Key for spreadsheet #2
    lgSortKey_B      = 1

	Query_Msg_Flg		= False
    lgIntFlgMode = parent.OPMD_CMODE 
    lgPageNo         = ""
    lgPageNo1        = ""
End Sub
'===================================================================================================================================
Sub SetDefaultVal()
	frm1.txtPlantCd.value=parent.gPlant
	frm1.txtPlantNm.value=parent.gPlantNm
End Sub
'===================================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
End Sub
'===================================================================================================================================
Sub InitSpreadSheet()
    
    Call SetZAdoSpreadSheet("M1311QA101","S","A","V20030329", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
    Call SetZAdoSpreadSheet("M1311QA102","S","B","V20030329", Parent.C_SORT_DBAGENT, frm1.vspdData2, C_MaxKey, "X", "X" )
    Call SetSpreadLock("A") 
    Call SetSpreadLock("B") 
End Sub
'===================================================================================================================================
Sub SetSpreadLock(ByVal pOpt)
    If pOpt = "A" Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
    Else
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.SpreadLockWithOddEvenRowColor()
    End If   
End Sub
'===================================================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Or UCase(frm1.txtPlantCd.ClassName)=UCase(parent.UCN_PROTECTED) Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "공장"	
	arrParam(1) = "B_Plant"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "Plant_CD"	
    arrField(1) = "Plant_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)
		frm1.txtPlantNm.Value= arrRet(1)
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
	End If	
	
End Function
'===================================================================================================================================
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	
	If lgIsOpenPop = True Then Exit Function
	if UCase(frm1.txtItemCd.ClassName) = UCase(parent.UCN_PROTECTED) then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End if
		
	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 
	arrField(2) = 3 ' -- Spec
	
	
    iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam,arrField, arrHeader), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtItemCd.Value  = arrRet(0)		
		frm1.txtItemNm.Value  = arrRet(1)	
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement
	End If	
End Function
'===================================================================================================================================
Function OpenSppl()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "외주처"					
	arrParam(1) = "B_Biz_Partner"				
	arrParam(2) = Trim(frm1.txtSpplCd.Value)		
'	arrParam(3) = Trim(frm1.txtSpplNm.Value)		
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "외주처"					
	
    arrField(0) = "BP_CD"						
    arrField(1) = "BP_NM"						
    
    arrHeader(0) = "외주처"					
    arrHeader(1) = "외주처명"				
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSpplCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtSpplCd.Value = arrRet(0)
		frm1.txtSpplNm.Value = arrRet(1)
		frm1.txtSpplCd.focus
		Set gActiveElement = document.activeElement
	End If	
End Function
'===================================================================================================================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If

	Call OpenOrderByPopup(gActiveSpdSheet.Id)
End Sub
'===================================================================================================================================
Function OpenOrderByPopup(ByVal pSpdNo)

	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
    arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(pSpdNo), gMethodText),"dialogWidth=" & parent.GROUPW_WIDTH & "px; dialogHeight=" & parent.GROUPW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")
	lgIsOpenPop = False

	If arrRet(0) = "X" Then
		If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo, arrRet(0), arrRet(1))
       Call InitVariables
       Call InitSpreadSheet
   End If
End Function
'===================================================================================================================================
Sub Form_Load()

    Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call AppendNumberPlace("6","5","4")
	Call InitSpreadSheet()
	Call SetToolbar("11000000000011")											
	
	frm1.txtSpplCd.focus
	Set gActiveElement = document.activeElement
End Sub
'===================================================================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub
'===================================================================================================================================
Sub vspdData2_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub   
'===================================================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub
'===================================================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub
'===================================================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
		Call vspdData_Click(NewCol, NewRow)
		frm1.vspdData2.MaxRows = 0
		Call DbQuery("2", NewRow)
    End If
End Sub
'===================================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Dim ii
    
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    SetPopupMenuItemInf("00000000001")

    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey_A = 1 Then
            ggoSpread.SSSort, lgSortKey_A
            lgSortKey_A = 2
        Else
            ggoSpread.SSSort, lgSortKey_A
            lgSortKey_A = 1
        End If    
        Exit Sub
    End If
    
    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)
    
     lgPageNo1   = ""     
     lgSortKey_B      = 1
End Sub
'===================================================================================================================================
Sub vspdData2_Click(ByVal Col,  ByVal Row)

	gMouseClickStatus = "SP2C"
	Set gActiveSpdSheet = frm1.vspdData2
	SetPopupMenuItemInf("00000000001")

    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey_B = 1 Then
            ggoSpread.SSSort, lgSortKey_B
            lgSortKey_B = 2
        Else
            ggoSpread.SSSort, lgSortKey_B
            lgSortKey_B = 1
        End If    
        Exit Sub
    End If
End Sub
'===================================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If
    

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)
			If DbQuery("1", 0) = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If
End Sub
'===================================================================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If
    

	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo1 <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)
			If DbQuery("2", frm1.vspdData.ActiveRow) = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If
End Sub
'===================================================================================================================================
Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear     

    Call ggoOper.ClearField(Document, "2")						
    Call InitVariables 											
    
    If Not chkField(Document, "1") Then							
       Exit Function
    End If
	
    If DbQuery("1", 0) = False then Exit Function    							

    FncQuery = True		
End Function
'===================================================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
    Set gActiveElement = document.activeElement
End Function
'===================================================================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
	Set gActiveElement = document.activeElement
End Function
'===================================================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)   
    Set gActiveElement = document.activeElement                         
End Function
'===================================================================================================================================
Function FncExit()
    FncExit = True
    Set gActiveElement = document.activeElement
End Function
'===================================================================================================================================
Function DbQuery(iOpt, currRow) 
	Dim strVal
	Dim strCfmFlg
    DbQuery = False
    
    If iOpt <> "1" and frm1.vspdData.MaxRows < 1 Then Exit Function

    Err.Clear                                                       
	If LayerShowHide(1) = False Then Exit Function
	
	With frm1
		
        If iOpt = "1" Then
			If lgIntFlgMode = parent.OPMD_UMODE Then

				strVal = BIZ_PGM_ID & "?txtSpplCd=" & Trim(.hdnSpplCd.value)
			    strVal = strVal & "&txtPlantCd=" & Trim(.hdnPlantCd.value)
			    strVal = strVal & "&txtItemCd=" & Trim(.hdnItemCd.value)
			Else
					
				strVal = BIZ_PGM_ID & "?txtSpplCd=" & Trim(.txtSpplCd.value)
			    strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
			    strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
			End if
                strVal = strVal & "&lgPageNo="   & lgPageNo                      '☜: Next key tag
			    strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
			    strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			    strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
        Else
        	frm1.vspddata.Row = currRow
        	frm1.vspddata.Col = GetKeyPos("A",10)
		
			strVal = BIZ_PGM_ID1 & "?txtPLNo="	 & frm1.vspddata.text
			strVal = strVal & "&lgPageNo1="   & lgPageNo1                      '☜: Next key tag
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("B")
			strVal = strVal & "&lgTailList="     & Space(1) & MakeSQLGroupOrderByList("B") 
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("B"))
        End If   
		
		if Query_Msg_Flg = false then
			strVal = strVal & "&Query_Msg_Flg=" & "F"
		else
			strVal = strVal & "&Query_Msg_Flg=" & "T"
		end if
		
		Call RunMyBizASP(MyBizASP, strVal)							
    End With
    
    DbQuery = True
    Call SetToolbar("1100000000011111")								

End Function
'===================================================================================================================================
Function DbQueryOk( iOpt)												

  	lgBlnFlgChgValue = False
    lgSaveRow        = 1
    lgIntFlgMode = Parent.OPMD_UMODE
	If iOpt = 1 Then
		If lgTopLeft <> "Y" Then
			Call vspdData_Click(1, 1)
			Call DbQuery("2", 1)
		End If
		lgTopLeft = "N"
		frm1.vspdData.focus
	Else
		Query_Msg_Flg = true
		frm1.vspdData.focus
	End If							                                     '⊙: This function lock the suitable field
End Function
'===================================================================================================================================
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>외주P/L</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
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
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>외주처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="외주처" NAME="txtSpplCd"  SIZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSpplCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSppl()" >
														   <INPUT TYPE=TEXT NAME="txtSpplNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>					   
								</TR>					   
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
													<INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm" SIZE=20 tag="14x"></TD>
									<TD CLASS="TD5" NOWRAP>모품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="모품목" NAME="txtItemCd" SIZE=10 MAXLENGTH=18 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">
													<INPUT TYPE=TEXT ALT="모품목" NAME="txtItemNm" SIZE=20 tag="14x"></TD>
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
									<script language =javascript src='./js/m1311qa1_A_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/m1311qa1_B_vspdData2.js'></script>
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
    <TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH="*" ALIGN="RIGHT"><a ONCLICK="VBSCRIPT:setCookie()">외주P/L등록</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnSpplCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="24">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
