<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S3312QA1
'*  4. Program Name         : ��ǰ��Ȳ(�ŷ�ó) cf.s7/s5117qa1
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/06/25
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kwakeunkyoung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<Script Language="VBScript">
Option Explicit

Const BIZ_PGM_ID 		= "S3312QB1.asp" 
Const C_MaxKey          = 10				                         

<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================
'=										User-defind Variables
'========================================
Dim lgIsOpenPop                                          
Dim IsOpenPop  

Dim lgSaveRow 

<% 
   BaseDate     = GetSvrDate                                                        
%>  

Dim FromDateOfDB
Dim ToDateOfDB

FromDateOfDB	= UNIConvDateAToB(UniDateAdd("m",-1,"<%=BaseDate%>",parent.gServerDateFormat), parent.gServerDateFormat,parent.gDateFormat)
ToDateOfDB		= UNIConvDateAToB(UniDateAdd("m", 0,"<%=BaseDate%>",parent.gServerDateFormat), parent.gServerDateFormat,parent.gDateFormat)

Dim lgStartRow
Dim lgEndRow

Const C_PopSoldToParty	= 1
Const C_PopSalesGrp		= 2
Const C_PopItemCd		= 3


'========================================	
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                     
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0    
End Sub


'========================================
Sub SetDefaultVal()
	With frm1
		.txtConFromDt.Text	= cstr(FromDateOfDB)
		.txtConToDt.Text	= cstr(ToDateOfDB)
	End With
End Sub							
	
									
'========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

	<% Call loadInfTB19029A("Q", "S","NOCOOKIE","QA") %>                                 
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>

End Sub


'========================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("S3312QA1","S","A", "V20030625", parent.C_SORT_DBAGENT,frm1.vspdData,C_MaxKey, "X", "X")
    Call SetSpreadLock()       
End Sub


'========================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub


'========================================
Sub Form_Load()
    Call LoadInfTB19029				                                           


    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")

	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolBar("1100000000001111")										
        
    Frm1.txtConFromDt.Focus

End Sub


'========================================
Function FncQuery() 
    On Error Resume Next                                                      
    Err.Clear                                                                    

	If ValidDateCheck(frm1.txtConFromDt, frm1.txtConToDt) = False Then Exit Function

    FncQuery = False                                                             
    
    Call ggoOper.ClearField(Document, "2")									     
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.ClearSpreadData()
    
    Call InitVariables 														    

	frm1.txt_TOTAL_RET_QTY.text = 0
	frm1.txt_TOTAL_RET_AMT.text = 0
    
    If Not chkField(Document, "1") Then								              
       Exit Function
    End If

    If DbQuery = False Then 
       Exit Function
    End If   

    If Err.number = 0 Then
       FncQuery = True                                                           
    End If   

    Set gActiveElement = document.ActiveElement      
End Function
	

'========================================
Function FncPrint()
    On Error Resume Next                                                         
    Err.Clear                                                                   

    FncPrint = False                                                             
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                        

    If Err.number = 0 Then
       FncPrint = True                                                           
    End If   

    Set gActiveElement = document.ActiveElement   
End Function


'========================================
Function FncExcel() 
    On Error Resume Next                                                        
    Err.Clear                                                                    

    FncExcel = False                                                              

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	Call Parent.FncExport(parent.C_MULTI)
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then
       FncExcel = True                                                            
    End If   

    Set gActiveElement = document.ActiveElement   
End Function


'========================================
Function FncFind() 
    On Error Resume Next                                                         
    Err.Clear                                                                    

    FncFind = False                                                             

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	Call Parent.FncFind(parent.C_MULTI, True)
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then
       FncFind = True                                                            
    End If   

    Set gActiveElement = document.ActiveElement   
End Function


'========================================
Function FncExit()
    On Error Resume Next                                                         
    Err.Clear                                                                    

    FncExit = False                                                              
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then
       FncExit = True                                                            
    End If   

    Set gActiveElement = document.ActiveElement  
End Function


'========================================
Function DbQuery() 
    On Error Resume Next                                                        
    Err.Clear                                                                    

    DbQuery = False
    
	Call LayerShowHide(1)

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    With frm1
        If lgIntFlgMode  <> parent.OPMD_UMODE Then								
			
			.txtHConFromDt.value	= .txtConFromDt.text
			.txtHConToDt.value		= .txtConToDt.text	

			.txtHSoldToParty.value	= .txtSoldToParty.value
			.txtHSalesGrp.value		= .txtSalesGrp.value
			.txtHItemCd.value		= .txtItemCd.value
				
			.txtHlgSelectListDT.value	= GetSQLSelectListDataType("A") 
			.txtHlgTailList.value		= MakeSQLGroupOrderByList("A")
			.txtHlgSelectList.value		= EnCoding(GetSQLSelectList("A"))
        End If    
        
        .txtHlgPageNo.value	= lgPageNo
        
        lgStartRow = .vspdData.MaxRows + 1										
    End With
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    If Err.number = 0 Then
       DbQuery = True																
    End If   

    Set gActiveElement = document.ActiveElement   
End Function


'========================================
Function DbQueryOk()	
    On Error Resume Next                                                         
    Err.Clear                                                                     

	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE										 
    lgSaveRow        = 1

    Call SetToolBar("11000000000111")
    
    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus
    End if  
        
    Set gActiveElement = document.ActiveElement   
End Function


'========================================
Function OpenConPopup(ByVal pvIntWhere)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	With frm1
		Select Case pvIntWhere
			' �ֹ�ó 
			Case C_PopSoldToParty												
				iArrParam(1) = "dbo.b_biz_partner BP"						' TABLE ��Ī 
				iArrParam(2) = Trim(.txtSoldToParty.value)					' Code Condition
				iArrParam(3) = ""											' Name Cindition
				iArrParam(4) = "BP.bp_type IN (" & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ") AND BP.usage_flag = " & FilterVar("Y", "''", "S") & " "		' Where Condition
					
				iArrField(0) = "ED15" & Parent.gColSep & "BP.bp_cd"			' Select Column
				iArrField(1) = "ED30" & Parent.gColSep & "BP.bp_nm"
				    
				iArrHeader(0) = .txtSoldtoParty.Alt							' Spread Title�� 
				iArrHeader(1) = .txtSoldtoPartyNm.Alt
	
				.txtSoldToParty.focus
				
			' ǰ�� 
			Case C_PopItemCd
				iArrParam(1) = "b_item"
				iArrParam(2) = Trim(.txtItemCd.value)
				iArrParam(3) = ""
				iArrParam(4) = ""

				iArrField(0) = "ED15" & Parent.gColSep & "Item_Cd"
				iArrField(1) = "ED30" & Parent.gColSep & "Item_Nm"

				iArrHeader(0) = .txtItemCd.Alt
				iArrHeader(1) = .txtItemNm.Alt
				
				.txtItemCd.focus
				
			' �����׷� 
			Case C_PopSalesGrp												
				iArrParam(1) = "dbo.B_SALES_GRP"
				iArrParam(2) = Trim(.txtSalesGrp.value)
				iArrParam(3) = ""
				iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
				
				iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"
				iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"
    
			    iArrHeader(0) = .txtSalesGrp.Alt
			    iArrHeader(1) = .txtSalesGrpNm.Alt
			    
			    .txtSalesGrp.focus

		End Select
	End With
 
	iArrParam(0) = iArrHeader(0)							' �˾� Title
	iArrParam(5) = iArrHeader(0)							' ��ȸ���� ��Ī 
	
	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If iArrRet(0) <> "" Then
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	

End Function


'========================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)
	SetConPopup = False

	With frm1
		Select Case pvIntWhere
			Case C_PopSoldToParty
				.txtSoldToParty.value = pvArrRet(0) 
				.txtSoldToPartyNm.value = pvArrRet(1)   

			Case C_PopItemCd
				.txtItemCd.value = pvArrRet(0) 
				.txtItemNm.value = pvArrRet(1)   
				
			Case C_PopSalesGrp
				.txtSalesGrp.value = pvArrRet(0) 
				.txtSalesGrpNm.value = pvArrRet(1)
		End Select
	End With

	SetConPopup = True		
	
End Function


'===============================================
Sub PopZAdoConfigGrid()
  If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
     Exit Sub
  End If

  Call OpenOrderBy("A")  
End Sub


'========================================
Sub OpenOrderBy(ByVal pvPsdNo)
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Sub
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(pvPsdNo),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then												' Means that nothing is happened!!!
	   Exit Sub
	Else
	   Call ggoSpread.SaveXMLData(pvPsdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Sub


'========================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
    
    If Frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    

    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)		

End Sub


'========================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)	    
    If Row <= 0 Then
	

    End If	
End Sub


'========================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

End Sub


'========================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    

'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub

	If CheckRunningBizProcess = True Then Exit Sub
    
	If Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           Call DbQuery
    	End If
    End If    
End Sub

'========================================
Sub txtConFromDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtConFromDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtConFromDt.Focus
	End If
End Sub

'========================================
Sub txtConToDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtConToDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtConToDt.Focus
	End If
End Sub

'========================================
Sub txtConFromDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'========================================
Sub txtConToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub


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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��ǰ��Ȳ(�ŷ�ó)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="30" align=right></td>
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
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>������</TD>									
									<TD CLASS="TD6" NOWRAP>							        
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
											<script language =javascript src='./js/s3312qa1_OBJECT1_txtConFromDt.js'></script>
											</TD>
											<TD>
											&nbsp;~&nbsp;
											</TD>
											<TD>
											<script language =javascript src='./js/s3312qa1_OBJECT2_txtConToDt.js'></script>
											</TD>
										</TR>
									</TABLE>							        
							        </TD>
									<TD CLASS=TD5 NOWRAP>�����׷�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGrp" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="�����׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopSalesGrp">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGrpNm" SIZE=20 TAG="14" ALT="�����׷��"></TD>							        
								</TR>	
								<TR>
									<TD CLASS="TD5" NOWRAP>�ŷ�ó</TD>
									<TD CLASS="TD6"><INPUT TYPE="Text" NAME="txtSoldToParty" SiZE=10 MAXLENGTH=10 tag="11XXXU" ALT="�ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoldToParty" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopSoldToParty">&nbsp;<INPUT TYPE="Text" NAME="txtSoldToPartyNm" SIZE=20 tag="14" ALT="�ŷ�ó��"></TD>
									<TD CLASS="TD5" NOWRAP>ǰ��</TD>
									<TD CLASS="TD6"><INPUT NAME="txtItemCd" ALT="ǰ��" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopItemCd">&nbsp;<INPUT NAME="txtItemNm" TYPE="Text" SIZE=20 tag="14" ALT="ǰ���"></TD>
								</TR>								
							</TABLE>
						</FIELDSET>
					</TD>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
<!-- SUM START -->
				<TR>
				  <TD  HEIGHT=3></TD>
				</TR>    
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>�����հ�</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/s3312qa1_fpDoubleSingle2_txt_TOTAL_RET_QTY.js'></script>
									</TD>
									<TD CLASS="TD5" NOWRAP>�ݾ��հ�</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/s3312qa1_fpDoubleSingle2_txt_TOTAL_RET_AMT.js'></script>
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
<!-- SUM END -->
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/s3312qa1_vspdData_vspdData.js'></script>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtHConFromDt"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConToDt"		tag="24" TABINDEX="-1">
				
<INPUT TYPE=HIDDEN NAME="txtHSoldToParty"	tag="24" TABINDEX="-1"> 
<INPUT TYPE=HIDDEN NAME="txtHSalesGrp"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHItemCd"		tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHlgPageNo"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgSelectListDT"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgTailList"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgSelectList"		tag="24" TABINDEX="-1">				

</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
