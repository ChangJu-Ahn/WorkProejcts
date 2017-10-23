<%@ LANGUAGE="VBSCRIPT" %>
<%
'************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : B1255MA8
'*  4. Program Name         : 영업조직조회 
'*  5. Program Desc         : 영업조직조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/03/21
'*  8. Modified date(Last)  : 2002/04/21
'*  9. Modifier (First)     : Mr Cho
'* 10. Modifier (Last)      : Park in sik
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit				

<!-- #Include file="../../inc/lgvariables.inc" --> 
Dim lgIsOpenPop                                              
Dim lgMark                                                  

Const BIZ_PGM_ID = "b1255mb8.asp"	
Const BIZ_PGM_JUMP_ID = "b1255ma1"											
Const C_MaxKey          = 2                                    '☆☆☆☆: Max key value

Dim IsOpenPop

Dim lsConcd
Dim lsConNm

'=============================================================================================================
Sub InitVariables()	
	lgIntFlgMode = parent.OPMD_CMODE                       
    lgStrPrevKey = ""
	lgPageNo     = ""       
End Sub

'=============================================================================================================
Sub SetDefaultVal()	
	frm1.txtSales_Org.focus
End Sub
	
'=============================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub

'=============================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("B1255MA8","S","A","V20021106", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
    Call SetSpreadLock 
End Sub
		
'=============================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'=============================================================================================================
Sub SetSpreadColor(ByVal lRow)
End Sub

'=============================================================================================================
Function OpenSorgCode(ByVal iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(1) = "B_SALES_ORG"						

	Select Case iWhere
	Case 1
		arrParam(2) = Trim(frm1.txtSales_Org.Value)				
		arrParam(4) = ""										
		arrParam(5) = "영업조직"							

		arrHeader(0) = "영업조직"							
		arrHeader(1) = "영업조직명"							

	Case 2
		arrParam(2) = Trim(frm1.txtUpper_Sales_Org.Value)					
		arrParam(4) = ""										
		arrParam(5) = "상위영업조직"						

		arrHeader(0) = "상위영업조직"						
		arrHeader(1) = "상위영업조직명"						

	End Select

    arrField(0) = "SALES_ORG"						
    arrField(1) = "SALES_ORG_NM"					

	arrParam(0) = arrParam(5)						
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	Select Case iWhere
	Case 1
		frm1.txtSales_Org.focus
	Case 2
		frm1.txtUpper_Sales_Org.focus
	End Select	

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSorgCode(arrRet,iWhere)
	End If	
	
End Function

'=============================================================================================================
Function PopZAdoConfigGrid()
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'=============================================================================================================
Function SetSorgCode(Byval arrRet,ByVal iWhere)

	Select Case iWhere
	Case 1
		frm1.txtSales_Org.value = arrRet(0) 
		frm1.txtSales_Org_nm.value = arrRet(1)   

	Case 2
		frm1.txtUpper_Sales_Org.value = arrRet(0) 
		frm1.txtUpper_Sales_OrgNm.value = arrRet(1)   

	End Select
	
End Function

'=============================================================================================================
Function CookiePage(ByVal Kubun)

	On Error Resume Next

	Const CookieSplit = 4877						
	
	Dim strTemp, arrVal

	Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

	If Kubun = 1 Then

		WriteCookie CookieSplit , lsConcd & parent.gRowSep & lsConnm

	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function
		
		arrVal = Split(strTemp, parent.gRowSep)

		If arrVal(0) = "" then Exit Function

		frm1.txtSales_Org.value =  arrVal(0)
		frm1.txtSales_Org_nm.value =  arrVal(1)
		frm1.txtUpper_Sales_Org.value =  arrVal(2)
		frm1.txtUpper_Sales_OrgNm.value =  arrVal(3)
		frm1.txtlvl.value =  arrVal(4)

		Select Case arrVal(5) 
		Case "Y"
			frm1.txtRadio.value = frm1.rdoUsage_flag2.value
			frm1.rdoUsage_flag2.checked = True
		Case "N"
			frm1.txtRadio.value = frm1.rdoUsage_flag3.value
			frm1.rdoUsage_flag3.checked = True
		Case Else
			frm1.txtRadio.value = frm1.rdoUsage_flag1.value
			frm1.rdoUsage_flag1.checked = True
		End Select
		
		if Err.number <> 0 then
			Err.Clear
			WriteCookie CookieSplit , ""
			exit function
		end if
		
		FncQuery()
		
		WriteCookie CookieSplit , ""

	End IF


End Function

'=============================================================================================================
Function NumericCheck()

	Dim objEl, KeyCode
	
	Set objEl = window.event.srcElement
	KeyCode = window.event.keycode

	Select Case KeyCode
    Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
	Case Else
		window.event.keycode = 0
	End Select

End Function

'=============================================================================================================
Sub Form_Load()
	
	Call LoadInfTB19029														
	Call ggoOper.LockField(Document, "N")                                   
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	
	Call InitVariables														    
	Call SetDefaultVal	
	Call InitSpreadSheet()

    Call SetToolBar("1100000000001111")										
	Call CookiePage(0)
	
    frm1.txtSales_Org.focus	
    
End Sub

'=============================================================================================================
Sub txtlvl_onKeyPress()
	Call NumericCheck()
End Sub


'=============================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("00000000001")
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
		Exit Sub
	End If

	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col			'Sort In Assending
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort In Desending
			lgSortKey = 1
		End If
		Exit Sub
	End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	If Row < 1 Then Exit Sub
	frm1.vspdData.Row = Row
	frm1.vspdData.Col = GetKeyPos("A",1) ' 2
	lsConcd=frm1.vspdData.Text
    
	frm1.vspdData.Row = Row
	frm1.vspdData.Col = GetKeyPos("A",2) ' 3
	lsConnm=frm1.vspdData.Text    
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 

End Sub

'=============================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

'=============================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	lgBlnFlgChgValue = True
	
End Sub

'=============================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
		If CheckRunningBizProcess = True Then Exit Sub	
    	If lgPageNo <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           Call DBQuery
    	End If
    End If    
End Sub

'===============================================================================================================
Function FncQuery() 

    FncQuery = False                                                        
    
    Err.Clear                                                               

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")		
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    Call InitVariables 														

	If frm1.rdoUsage_flag1.checked = True Then
		frm1.txtRadio.value = frm1.rdoUsage_flag1.value 
	ElseIf frm1.rdoUsage_flag2.checked = True Then
		frm1.txtRadio.value = frm1.rdoUsage_flag2.value 
	ElseIf frm1.rdoUsage_flag3.checked = True Then
		frm1.txtRadio.value = frm1.rdoUsage_flag3.value 
	End If

    Call DbQuery																'☜: Query db data

    FncQuery = True																'⊙: Processing is OK

End Function

'===============================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'===============================================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function

'===============================================================================================================
Function FncFind() 
	Call parent.FncFind(parent.C_MULTI, False)
End Function

'===============================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'===============================================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")   		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'===============================================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim StrNextKey      

    DbQuery = False
    
    	
	If   LayerShowHide(1) = False Then
             Exit Function 
    End If

    
    Err.Clear                                                               

	Dim strVal
    
    With frm1	
		If lgIntFlgMode = parent.OPMD_UMODE Then 
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			
			strVal = strVal & "&txtSales_Org=" & Trim(.HSales_Org.value)
			strVal = strVal & "&txtRadio=" & Trim(.HRadio.value)
			strVal = strVal & "&txtUpper_Sales_Org=" & Trim(.HUpper_Sales_Org.value)
			strVal = strVal & "&txtlvl=" & Trim(.Hlvl.value)
		else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			
			strVal = strVal & "&txtSales_Org=" & Trim(.txtSales_Org.value)
			strVal = strVal & "&txtRadio=" & Trim(.txtRadio.value)
			strVal = strVal & "&txtUpper_Sales_Org=" & Trim(.txtUpper_Sales_Org.value)
			strVal = strVal & "&txtlvl=" & Trim(.txtlvl.value)
		end if		
			strVal = strVal & "&lgPageNo="       & lgPageNo                '☜: Next key tag			
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True

End Function

'===============================================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 

	lgIntFlgMode = parent.OPMD_UMODE                   'Indicates that current mode is Update mode
    lgBlnFlgChgValue = False   
  
	Call ggoOper.LockField(Document, "Q")
	Call SetToolBar("1100000000011111")	

    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus
    End if  	

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
									<TD CLASS="TD5" NOWRAP>영업조직</TD>
									<TD CLASS="TD6" NOWRAP >
										<input NAME="txtSales_Org" TYPE="Text" MAXLENGTH="4" tag="11XXXU" ALT = "영업조직코드" size="10"><img SRC="../../../CShared/image/btnPopup.gif" NAME="ImgSaleOrgCode" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenSorgCode 1"> 
										<input NAME="txtSales_Org_nm" TYPE="Text" MAXLENGTH="30" tag="14XXX" size="25"></TD>
									<TD CLASS="TD5" NOWRAP>상위영업조직</TD>
									<TD CLASS="TD6"><input NAME="txtUpper_Sales_Org" TYPE="Text" MAXLENGTH="4" tag="11XXXU" size="10"><img SRC="../../../CShared/image/btnPopup.gif" NAME="btnSales_org" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenSorgCode 2">&nbsp;<input NAME="txtUpper_Sales_OrgNm" TYPE="Text" MAXLENGTH="30" tag="24" size="25"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>조직레벨</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtlvl" TYPE="Text" MAXLENGTH="2" tag="11XXXU" size="10"></TD>
									<TD CLASS="TD5" NOWRAP>사용여부</TD>
									<TD CLASS="TD6" NOWRAP style="TEXT-ALIGN:center;">
										<input type=radio CLASS="RADIO" name="rdoUsage_flag" id="rdoUsage_flag1" value="" tag = "11XXX" checked>
											<label for="rdoUsage_flag1">전체</label>&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoUsage_flag" id="rdoUsage_flag2" value="Y" tag = "11XXX">
											<label for="rdoUsage_flag2">사용</label>&nbsp;
										<input type=radio CLASS="RADIO" name="rdoUsage_flag" id="rdoUsage_flag3" value="N" tag = "11XXX">
											<label for="rdoUsage_flag3">미사용</label></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD HEIGHT="100%">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">영업조직등록</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> 
		                FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadio" tag="24">
<INPUT TYPE=HIDDEN NAME="HSales_Org" tag="24">
<INPUT TYPE=HIDDEN NAME="HRadio" tag="24">
<INPUT TYPE=HIDDEN NAME="HUpper_Sales_Org" tag="24">
<INPUT TYPE=HIDDEN NAME="Hlvl" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1" ></iframe>
</DIV>
</BODY>
</HTML>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        

