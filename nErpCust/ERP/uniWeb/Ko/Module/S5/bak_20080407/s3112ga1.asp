<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S3112GA1
'*  4. Program Name         : 미출고집계조회 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001/12/19
'*  8. Modified date(Last)  : 2003/06/11
'*  9. Modifier (First)     : Kim Hyungsuk
'* 10. Modifier (Last)      : Hwang Seongbae
'**********************************************************************************************
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgIsOpenPop                                            
Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, Parent.gServerDateFormat, Parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, Parent.gDateFormat)

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "s3112gb1.asp"
Const C_MaxKey          = 1		                               '☆☆☆☆: Max key value

'========================================
Sub InitVariables()
    lgPageNo         = ""
    lgIntFlgMode     = parent.OPMD_CMODE
    lgStrPrevKey     = ""                                  
    lgSortKey        = 1

    Call SetToolbar("11000000000011")

End Sub

'========================================= 
Sub SetDefaultVal()
	frm1.txtDlvyFromDt.text = StartDate
	frm1.txtDlvyToDt.text = EndDate

	frm1.txtSalesGroup.focus	  
	
End Sub

'=========================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "QA") %>
End Sub

'==========================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S3112GA1","G","A","V20051106", Parent.C_GROUP_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
    Call SetSpreadLock 
End Sub

'=========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================
Function OpenConSItemDC(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Select Case iWhere
	Case 0
		arrParam(1) = "B_BIZ_PARTNER"						
		arrParam(2) = Trim(frm1.txtconBp_cd.Value)			
		arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"				
		arrParam(5) = "납품처"							
	
		arrField(0) = "BP_CD"								
		arrField(1) = "BP_NM"								
    
		arrHeader(0) = "납품처"							
		arrHeader(1) = "납품처명"
		
		frm1.txtconBp_cd.focus						

	Case 2
		arrParam(1) = "B_SALES_GRP"							
		arrParam(2) = Trim(frm1.txtSalesGroup.Value)		
		arrParam(4) = ""									
		arrParam(5) = "영업그룹"						
	
		arrField(0) = "SALES_GRP"							
		arrField(1) = "SALES_GRP_NM"							
    
		arrHeader(0) = "영업그룹"						
		arrHeader(1) = "영업그룹명"							

		frm1.txtSalesGroup.focus
		
	Case 3
		arrParam(1) = "B_ITEM"								
		arrParam(2) = Trim(frm1.txtItem_cd.Value)			
		arrParam(4) = "PHANTOM_FLG = " & FilterVar("N", "''", "S") & " "									
		arrParam(5) = "품목"							
	
		arrField(0) = "ITEM_CD"								
		arrField(1) = "ITEM_NM"								
		arrField(2) = "SPEC"								
    
		arrHeader(0) = "품목"							
		arrHeader(1) = "품목명"
		arrHeader(2) = "규격"
		
		frm1.txtItem_cd.focus
									
	Case 4
		arrParam(1) = "B_PLANT"								
		arrParam(2) = Trim(frm1.txtPlant.value)				
		arrParam(4) = ""									
		arrParam(5) = "공장"							
	
		arrField(0) = "PLANT_CD"							
		arrField(1) = "PLANT_NM"							
    
		arrHeader(0) = "공장"							
		arrHeader(1) = "공장명"
		
		frm1.txtPlant.focus
	Case 5
		'2002-10-07 s3135pa1.asp 추가 
		Dim strRet, iCalledAspName
		
		Dim arrTNParam(5), i

		For i = 0 to UBound(arrTNParam)
			arrTNParam(i) = ""
		Next	

		'20021227 kangjungu dynamic popup
		iCalledAspName = AskPRAspName("s3135pa1")	
		if Trim(iCalledAspName) = "" then
			IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3135pa1", "x")
			lgIsOpenPop = False
			exit Function
		end if

		strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrTNParam), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		lgIsOpenPop = False

		If strRet = "" Then
			Exit Function
		Else
			frm1.txtTrackingNo.value = strRet 
		End If		
		
		frm1.txtTrackingNo.focus
		Exit Function				
		
	End Select

	arrParam(0) = arrParam(5)								
	
	Select Case iWhere
	Case 3
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Case Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select

	lgIsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetConSItemDC(arrRet, iWhere)
	End If	
	
End Function

'========================================
Function PopZAdoConfigGrid()
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOGroupPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.GROUPW_WIDTH & "px; dialogHeight=" & parent.GROUPW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If

End Function

'========================================
Function SetConSItemDC(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
		Case 0
			.txtconBp_cd.value = arrRet(0) 
			.txtconBp_Nm.value = arrRet(1)   
		Case 2
			.txtSalesGroup.value = arrRet(0) 
			.txtSalesGroupNm.value = arrRet(1)   
		Case 3
			.txtItem_cd.value = arrRet(0) 
			.txtItem_Nm.value = arrRet(1)   
		Case 4
			.txtPlant.value = arrRet(0) 
			.txtPlantNm.value = arrRet(1)   
		End Select
	End With
End Function

'==================================================================
Sub Form_Load()

    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
End Sub

'========================================
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

End Sub

'=======================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'=======================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then Exit Sub

	If Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData,NewTop) Then	    
		If CheckRunningBizProcess = True Then Exit Sub
    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           Call DbQuery
    	End If
    End If
   
End Sub

'========================================
Sub txtDlvyFromDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtDlvyFromDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtDlvyFromDt.Focus
	End If
End Sub

'========================================
Sub txtDlvyToDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtDlvyToDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtDlvyToDt.Focus
	End If
End Sub

'========================================
Sub txtDlvyFromDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'========================================
Sub txtDlvyToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'========================================
Function FncQuery() 

    FncQuery = False                                                        
    
    Err.Clear                                                 
    
	If ValidDateCheck(Frm1.txtDlvyFromDt, Frm1.txtDlvyToDt) = False Then Exit Function

    Call ggoOper.ClearField(Document, "2")									
    Call InitVariables 														
    
    Call DbQuery															'☜: Query db data

    FncQuery = True		
End Function

'=====================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'=====================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function

'=====================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                                     <%'☜:화면 유형, Tab 유무 %>
End Function

'=====================================================
Function FncExit()
    
    On Error Resume Next                                                          
    Err.Clear                                                                     

    FncExit = True                                                             
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
   
    	If lgIntFlgMode = parent.OPMD_UMODE Then  
    	
    	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									

        	strVal = strVal & "&txtconBp_cd=" & Trim(.HtxtconBp_cd.value)
    		strVal = strVal & "&txtSalesGroup=" & Trim(.HtxtSalesGroup.value)
    		strVal = strVal & "&txtItem_cd=" & Trim(.HtxtItem_cd.value)
    		strVal = strVal & "&txtPlant=" & Trim(.HtxtPlant.value)
    		strVal = strVal & "&txtDlvyFromDt=" & Trim(.HtxtDlvyFromDt.value)
    		strVal = strVal & "&txtDlvyToDt=" & Trim(.HtxtDlvyToDt.value)
    		strVal = strVal & "&txtTrackingNO=" & Trim(.HtxtTrackingNo.value)
    	    
            strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag

        Else
    		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			
    		strVal = strVal & "&txtconBp_cd=" & Trim(.txtconBp_cd.value)
    		strVal = strVal & "&txtSalesGroup=" & Trim(.txtSalesGroup.value)
    		strVal = strVal & "&txtItem_cd=" & Trim(.txtItem_cd.value)
    		strVal = strVal & "&txtPlant=" & Trim(.txtPlant.value)
    		strVal = strVal & "&txtDlvyFromDt=" & Trim(.txtDlvyFromDt.text)
    		strVal = strVal & "&txtDlvyToDt=" & Trim(.txtDlvyToDt.text)
    		strVal = strVal & "&txtTrackingNO=" & Trim(.txtTrackingNo.value)

            strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag
        
        End If 
    		
	End With

		strVal = strVal & "&lgPageNo="       & lgPageNo                          '☜: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
    Call RunMyBizASP(MyBizASP, strVal)										
    
    DbQuery = True
End Function

'=====================================================
Function DbQueryOk()											'☆: 조회 성공후 실행로직 
	lgIntFlgMode     = parent.OPMD_UMODE						'⊙: Indicates that current mode is Update mode

	Call SetToolbar("11000000000111")							'⊙: 버튼 툴바 제어 

    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus
    End if  	

End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>미출고집계조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH="*">&nbsp;</td>
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
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 2">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14"></TD>
									<TD CLASS="TD5" NOWRAP>납품처</TD>
									<TD CLASS="TD6"><INPUT NAME="txtconBp_cd" ALT="납품처" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 0">&nbsp;<INPUT NAME="txtconBp_nm" TYPE="Text" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6"><INPUT NAME="txtItem_cd" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnStoRo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 3">&nbsp;<INPUT NAME="txtItem_Nm" TYPE="Text" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlant" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" Onclick="vbscript:OpenConSItemDC 4">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 TAG="14"></TD>
								</TR>	
								<TR>
									<TD CLASS=TD5 NOWRAP>납기일</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtDlvyFromDt" CLASS=FPDTYYYYMMDD tag="11X1" Title="FPDATETIME" ALT="납기시작일"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtDlvyToDt" CLASS=FPDTYYYYMMDD tag="11X1" Title="FPDATETIME" ALT="납기종료일"></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>Tracking No</TD>
									<TD CLASS="TD6"><INPUT NAME="txtTrackingNo" ALT="Tracking No" TYPE="Text" MAXLENGTH=25 SiZE=30 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConSItemDC 5"></TD>	
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
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> 
		                FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">

<INPUT TYPE=HIDDEN NAME="HtxtconBp_cd" tag="24"> 
<INPUT TYPE=HIDDEN NAME="HtxtSalesGroup" tag="24">
<INPUT TYPE=HIDDEN NAME="HtxtItem_cd" tag="24">
<INPUT TYPE=HIDDEN NAME="HtxtPlant" tag="24">
<INPUT TYPE=HIDDEN NAME="HtxtDlvyFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="HtxtDlvyToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="HtxtTrackingNo" tag="24">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 TABINDEX="-1" src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
