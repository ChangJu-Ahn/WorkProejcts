<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : s3135pa1
'*  4. Program Name         : Tracking No(수주진행별조회)
'*  5. Program Desc         : Tracking No(수주진행별조회)
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/02/16
'*  8. Modified date(Last)  : 2002/04/12
'*  9. Modifier (First)     : Choinkuk		
'* 10. Modifier (Last)      : Choinkuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE></TITLE>

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

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

Const BIZ_PGM_ID 		= "s3135pb1.asp"                              
Const C_MaxKey          = 3		                                        
             
Dim lgIsOpenPop                      
Dim IscookieSplit 
Dim IsOpenPop  
Dim gblnWinEvent											
Dim arrReturn												
Dim strParam

Dim arrParent
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

'======================================================================================================================
Function InitVariables()
	lgStrPrevKey     = ""								   'initializes Previous Key
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE              'Indicates that current mode is Create mode
    lgSortKey        = 1   
        
    gblnWinEvent = False

	arrReturn = ""
		
    Self.Returnvalue = arrReturn     
End Function

'======================================================================================================================
Sub SetDefaultVal()

	Dim arrParam

	arrParam = arrParent(1)		
		
	If Len(arrParam(0)) then
		frm1.txtPtnBpCd.value = arrParam(0)
	End If
			
	If Len(arrParam(1)) then
		frm1.txtSalesGrp.value = arrParam(1)
	End If

	If Len(arrParam(2)) then
		frm1.txtPlant.value = arrParam(2)
	End If

	If Len(arrParam(3)) then
		frm1.txtItem.value = arrParam(3)
	End If

	If Len(arrParam(4)) then
		frm1.txtSoNo.value = arrParam(4)
	End If

	If Len(arrParam(5)) then
		strParam = arrParam(5)	
	Else
		strParam = ""
	End If

	Dim i
		
	i = UBound(arrParam)

	If i > 5 Then 
		If arrParam(6) = "M" And arrParam(2) <> "" Then
			Call ggoOper.SetReqAttr(frm1.txtPlant, "Q")
		End If
	End If		

	frm1.txtFromDt.text = StartDate
	frm1.txtToDt.text	= EndDate

End Sub

'=============================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "PA") %>
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'=============================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S3135MA1","S","A","V20021106", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )    
    Call SetSpreadLock 	  	      
End Sub

'=============================================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	ggoSpread.SpreadLockWithOddEvenRowColor()
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    .vspdData.ReDraw = True
    .vspdData.OperationMode = 5

    End With
End Sub	

'=============================================================================================================
Function OKClick()
		
	Dim arrReturn
	If frm1.vspdData.ActiveRow > 0 Then				
		
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = GetKeyPos("A",3)
		arrReturn = frm1.vspdData.Text

		Self.Returnvalue = arrReturn
	End If

	Self.Close()

End Function

'=============================================================================================================
Function CancelClick()
	arrReturn = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'=============================================================================================================
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

'=============================================================================================================
Function OpenConSItemDC(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True
	
	Select Case iWhere
	Case 0		
		arrParam(0) = "수주번호"										' TextBox 명칭 
		arrParam(1) = "S_SO_HDR SO, B_BIZ_PARTNER BP"						' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtSoNo.value)								' Code Condition
		arrParam(4) = "SO.SOLD_TO_PARTY=BP.BP_CD AND SO.CFM_FLAG = " & FilterVar("Y", "''", "S") & " "		' Where Condition
		arrParam(5) = "수주번호"										' TextBox 명칭 
			
		arrField(0) = "SO.SO_NO"											' Field명(0)
		arrField(1) = "BP.BP_NM"											' Field명(1)
    
		arrHeader(0) = "수주번호"										' Header명(0)
		arrHeader(1) = "주문처"											' Header명(1)
		
	Case 1
		arrParam(0) = "주문처"		
		arrParam(1) = "B_BIZ_PARTNER"
		arrParam(2) = Trim(frm1.txtPtnBpCd.value)		
		arrParam(3) = ""					
		arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") AND usage_flag = " & FilterVar("Y", "''", "S") & " "		
		arrParam(5) = "주문처"		
			
		arrField(0) = "BP_CD"							
		arrField(1) = "BP_NM"				
		
	    arrHeader(0) = "주문처"							
	    arrHeader(1) = "주문처명"						

	Case 2

		arrParam(0) = "영업그룹"						
		arrParam(1) = "B_SALES_GRP"							
		arrParam(2) = Trim(frm1.txtSalesGrp.value)		
		arrParam(3) = ""									
		arrParam(4) = ""									
		arrParam(5) = "영업그룹"						

		arrField(0) = "SALES_GRP"							
		arrField(1) = "SALES_GRP_NM"						

		arrHeader(0) = "영업그룹"						
		arrHeader(1) = "영업그룹명"						
	    
    
	Case 3

		If UCase(frm1.txtPlant.className) = PopupParent.UCN_PROTECTED Then 
			gblnWinEvent = False			
			Exit Function
		End IF
		
		arrParam(0) = "공장"				
		arrParam(1) = "B_PLANT"							
		arrParam(2) = Trim(frm1.txtPlant.value)		
		arrParam(4) = ""							
		arrParam(5) = "공장"				
		
		arrField(0) = "PLANT_CD"				
		arrField(1) = "PLANT_NM"				
	    
		arrHeader(0) = "공장"					
		arrHeader(1) = "공장명"				

	Case 4
		arrParam(0) = "품목"							
		arrParam(1) = "B_ITEM"								
		arrParam(2) = Trim(frm1.txtItem.value)				
		arrParam(3) = ""									
		arrParam(4) = ""									
		arrParam(5) = "품목"							
		
		arrField(0) = "ITEM_CD"								
		arrField(1) = "ITEM_NM"								
		
		arrHeader(0) = "품목"							
		arrHeader(1) = "품목명"							
		
	End Select


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConSItemDC(arrRet, iWhere)
	End If	
	
End Function

'=============================================================================================================
Function SetConSItemDC(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
		Case 0
			.txtSoNo.value = arrRet(0) 		
			.txtSoNo.focus
		Case 1
			.txtPtnBpCd.value = arrRet(0) 
			.txtPtnBpNm.value = arrRet(1)   
			.txtPtnBpCd.focus
		Case 2
			.txtSalesGrp.value = arrRet(0)
			.txtSalesGrpNm.value = arrRet(1)  
			.txtSalesGrp.focus
		Case 3
			.txtPlant.value = arrRet(0) 
			.txtPlantNm.value = arrRet(1) 
			.txtPlant.focus 		 
		Case 4
			.txtItem.value = arrRet(0) 
			.txtItemNm.value = arrRet(1)  
			.txtItem.focus		 
		End Select
	End With
End Function


'=============================================================================================================
Sub Form_Load()
    Call LoadInfTB19029													
     
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,PopupParent.ggStrMinPart,PopupParent.ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,PopupParent.ggStrMinPart,PopupParent.ggStrMaxPart)
	
	Call ggoOper.LockField(Document, "N")                             
    
	Call InitVariables											  
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()

End Sub

'=============================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then Exit Function
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function

'=============================================================================================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function

'=============================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub    

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    	
	   	If lgPageNo <> "" Then
           Call DBQuery          
    	End If
    End If    
End Sub

'=============================================================================================================
Sub txtFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromDt.Action = 7		
		Call SetFocusToDocument("P")
		frm1.txtFromDt.Focus
	End If
End Sub

Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtToDt.Focus
	End If
End Sub

'=============================================================================================================
Sub txtFromDt_KeyPress(KeyAscii)
	On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub

Sub txtToDt_KeyPress(KeyAscii)
	On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub

'=============================================================================================================
Function FncQuery() 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               

	If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt) = False Then Exit Function
   
    Call ggoOper.ClearField(Document, "2")	         						
    Call InitVariables   

	If DbQuery = False Then Exit Function									

    FncQuery = True		
    
End Function

'=============================================================================================================
Function DbQuery() 

	Err.Clear														
	DbQuery = False													
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
    
    With frm1
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then		
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
			strVal = strVal & "&txtSoNo=" & Trim(.txtHSoNo.value)
			strVal = strVal & "&txtPtnBpCd=" & Trim(.txtHPtnBpCd.value)
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtHSalesGrp.value)
			strVal = strVal & "&txtPlant=" & Trim(.txtHPlant.value)
			strVal = strVal & "&txtItem=" & Trim(.txtHItem.value)
			strVal = strVal & "&txtFromDt=" & Trim(.txtHFromDt.value)
			strVal = strVal & "&txtToDt=" & Trim(.txtHToDt.value)
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey
			strVal = strVal & "&strParam="   & strParam

        Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
			strVal = strVal & "&txtSoNo=" & Trim(.txtSoNo.value)
			strVal = strVal & "&txtPtnBpCd=" & Trim(.txtPtnBpCd.value)
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtSalesGrp.value)
			strVal = strVal & "&txtPlant=" & Trim(.txtPlant.value)
			strVal = strVal & "&txtItem=" & Trim(.txtItem.value)
			strVal = strVal & "&txtFromDt=" & Trim(.txtFromDt.text)
			strVal = strVal & "&txtToDt=" & Trim(.txtToDt.text)
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey
			strVal = strVal & "&strParam="   & strParam
			
		End If				
			
        strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag         
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
  
        Call RunMyBizASP(MyBizASP, strVal)		    						
        
    End With
    
    DbQuery = True    

End Function

'=============================================================================================================
Function DbQueryOk()	    												

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
						<TD CLASS=TD5 NOWRAP>주문처</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtPtnBpCd" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="주문처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoRef" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 1">&nbsp;<INPUT TYPE=TEXT NAME="txtPtnBpNm" SIZE=20 TAG="14"></TD>
						<TD CLASS=TD5 NOWRAP>영업그룹</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGrp" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoRef" align=top TYPE="BUTTON" ONCLICK="Vbscript:OpenConSItemDC 2">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGrpNm" SIZE=20 TAG="14">
						</TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>공장</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtPlant" ALT="공장" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoRef" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConSItemDC 3">&nbsp;<INPUT NAME="txtPlantNm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="24"></TD>
						<TD CLASS=TD5 NOWRAP>품목</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtItem" ALT="품목" TYPE="Text" MAXLENGTH=18 SIZE=10 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoRef" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConSItemDC 4">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 TAG="14"></TD>
					</TR>					
					<TR>
						<TD CLASS=TD5 NOWRAP>수주일</TD>
						<TD CLASS=TD6>
							<TABLE CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD>
										<script language =javascript src='./js/s3135pa1_fpDateTime1_txtFromDt.js'></script>
									</TD>
									<TD>
										&nbsp;~&nbsp;
									</TD>
									<TD>
										<script language =javascript src='./js/s3135pa1_fpDateTime2_txtToDt.js'></script>
									</TD>
								</TR>
							</TABLE>
						</TD>					
						<TD CLASS=TD5 NOWRAP>수주번호</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtSoNo" SIZE=34 MAXLENGTH=18 TAG="11XXXU" ALT="S/O번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoRef" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 0">&nbsp;</TD>
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
						<script language =javascript src='./js/s3135pa1_vaSpread_vspdData.js'></script>
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
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
					                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
                                     <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtHSoNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPtnBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHSalesGrp" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPlant" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHItem" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHToDt" tag="24">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
