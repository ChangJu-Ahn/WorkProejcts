<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 수주관리 
'*  3. Program ID           : S3114MA8
'*  4. Program Name         : 회답납기조회 
'*  5. Program Desc         : 회답납기조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/03/24
'*  8. Modified date(Last)  : 2002/03/06
'*  9. Modifier (First)     : Cho song hyon 
'* 10. Modifier (Last)      : Ahn Tae Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/08/10 : 4th 화면 Layout 수정 
'*                            -2001/12/18 : Date 표준적용 
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	
<!-- #Include file="../../inc/lgvariables.inc" --> 

Dim lgMark                                                  
Dim lgIsOpenPop  
Dim IscookieSplit 
Dim ArrParam(7)

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID = "s3114mb8.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_JUMP_ID = "s3112ma1"		

Const C_MaxKey          = 1                                    '☆☆☆☆: Max key value										'☆: JUMP시 비지니스 로직 ASP명 

Const C_SoSeq = 1			'수주순번 
Const C_SoSChoNo = 2		'SCHO NO
Const C_ItemCd = 3			'품목코드 
Const C_ItemNm = 4			'품목 
Const C_TrackingNo = 5		'Tracking No
Const C_ReqDt = 6			'납기일자Form_QueryUnLoad
Const C_PromiseDt = 7		'출하가능일자 
Const C_CfmSoQty = 8		'확정수주량 
Const C_CfmBonusSoQty = 9	'확정수주할증량 
'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------

Dim IsOpenPop 

'============================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE        'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                'Indicates that no value changed     
    lgPageNo = ""                           'initializes Previous Key
    lgLngCurRows = 0                        'initializes Deleted Rows Count
	lgSortKey = 1
End Sub

'============================================================================================================
Sub SetDefaultVal()	
End Sub

'============================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "MA") %>
End Sub

'============================================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("S3114QA8","S","A","V20030712", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )    
    Call SetSpreadLock        
End Sub

'============================================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
	ggoSpread.SpreadLockWithOddEvenRowColor()
    .vspdData.ReDraw = True
    End With
End Sub

'============================================================================================================
Sub SetSpreadColor(ByVal lRow)    
End Sub

'============================================================================================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenOrderByPopup("A")
End Sub

'============================================================================================================
Sub OpenOrderByPopup(ByVal pSpdNo)
	Dim arrRet
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Sub
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Sub
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Sub

'============================================================================================================
Function OpenSoNo()
	Dim iCalledAspName
	Dim strRet

	If IsOpenPop = True Then Exit Function
			
	'20021227 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("s3111pa1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3111pa1", "x")
		IsOpenPop = False
		exit Function
	end if
	IsOpenPop = True
		
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, ""), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	frm1.txtSoNo.focus

	If strRet = "" Then
		Exit Function
	Else
		frm1.txtSoNo.value = strRet 
	End If	

End Function


'============================================================================================================
Function OpenItem()
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(1) = "b_item"									
	arrParam(2) = Trim(frm1.txtItemCd.Value)			
	arrParam(3) = ""                             			
	arrParam(4) = "PHANTOM_FLG = " & FilterVar("N", "''", "S") & " "										
	arrParam(5) = "품목"								
	
	arrField(0) = "Item_cd"									
	arrField(1) = "Item_nm"									
    arrField(2) = "SPEC"	
    
	arrHeader(0) = "품목"								
	arrHeader(1) = "품목명"								
    arrHeader(2) = "규격"			
    
	arrParam(0) = arrParam(5)								
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtItemCd.focus
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtItemCd.value = arrRet(0)
		frm1.txtItemNm.value = arrRet(1)
	End If	
	
End Function


'============================================================================================================

Function CookiePage(ByVal Kubun)

	On Error Resume Next

	Const CookieSplit = 4877						<%'Cookie Split String : CookiePage Function Use%>
	Dim strTemp, arrVal

	If Kubun = 1 Then

		WriteCookie CookieSplit , frm1.txtSoNo.value
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, parent.gRowSep)

		If arrVal(0) = "" Then Exit Function
		
		frm1.txtSoNo.value =  arrVal(0)

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call FncQuery()
			
		WriteCookie CookieSplit , ""
		
	End If

End Function

'============================================================================================================
Sub Form_Load()
    Call LoadInfTB19029														
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                          
    
    '----------  Coding part  -------------------------------------------------------------
	Call InitVariables                                                      
	Call SetDefaultVal
	Call InitSpreadSheet
    Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어 
	Call CookiePage(0)
	
    frm1.txtSoNo.focus
End Sub

'============================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
 
 	Call SetPopupMenuItemInf("00000000001")
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
	
    If Row <= 0 Then
       
       ggoSpread.Source = frm1.vspdData
       If lgSortKey = 1 Then	
			ggoSpread.SSSort Col				'Sort in ascending
			lgSortKey = 2
	   Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in descending
			lgSortKey = 1
       End If
       
       Exit Sub
    End If            
    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)   
    
End Sub

'============================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'============================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

'============================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	'추가 
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	lgBlnFlgChgValue = True
	'추가 
End Sub

'============================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           Call DBQuery
    	End If
    End If    
End Sub

'============================================================================================================
Function FncQuery() 

	Dim IntRetCD

    FncQuery = False                                                        
    
    Err.Clear                                                               

    Call ggoOper.ClearField(Document, "2")									
    Call InitVariables 														
    
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

    Call DbQuery																

    FncQuery = True																'⊙: Processing is OK

End Function

'============================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'============================================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function

'============================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                                     
End Function


'============================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'============================================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    FncExit = True
End Function

'============================================================================================================  
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               
	Call LayerShowHide(1)
    
    With frm1
    
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			
		strVal = strVal & "&txtSoNo=" & Trim(.txtSoNo.value)				
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)

        strVal = strVal & "&lgPageNo="   & lgPageNo                      '☜: Next key tag
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
      
	    Call RunMyBizASP(MyBizASP, strVal)										
    End With
    
    DbQuery = True

End Function

'============================================================================================================
Function DbQueryOk()														

	lgIntFlgMode = parent.OPMD_UMODE	
	lgBlnFlgChgValue = False

    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	Call SetToolbar("11000000000111")

    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus
    Else
       'frm1.txtSoNo.focus	
    End if 

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>회답납기조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>
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
									<TD CLASS=TD5 NOWRAP>수주번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoNo" ALT="수주번호" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSONo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSoNo()"></TD>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemCd" ALT="품목" TYPE="Text" MAXLENGTH="18" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSONo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">&nbsp;<INPUT NAME="txtItemNm" TYPE="Text" SIZE=25 tag="14"></TD>
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
										<script language =javascript src='./js/s3114ma8_I518194901_vspdData.js'></script>
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
					<TD WIDTH="*" Align=Right><A HREF="VBSCRIPT:CookiePage(1)">수주내역등록</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> 
		            FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 TABINDEX="-1" src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
