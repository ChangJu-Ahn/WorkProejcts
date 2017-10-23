<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M1211QA1
'*  4. Program Name         : ǰ�񺰰���ó��ȸ 
'*  5. Program Desc         : ǰ�񺰰���ó��ȸ 
'*  6. Component List       : 
'*  7. Modified date(First) : 2001/01/08
'*  8. Modified date(Last)  : 2003/05/26
'*  9. Modifier (First)     : Min, Hak-jun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<!--'==========================================  1.1.1 Style Sheet  ========================================
'=======================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'=======================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit

'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'********************************************************************************************************
Const BIZ_PGM_ID		= "m1211qb1.asp"									
Const BIZ_PGM_JUMP_ID1	= "m1211ma1"
Const BIZ_PGM_JUMP_ID2	= "m1211ma2"

<!--'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================!-->
Const C_MaxKey          = 22            
                                     
<!-- '==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= !-->
<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim lgIsOpenPop                                          

'==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode	 = parent.OPMD_CMODE        
    lgBlnFlgChgValue = False         
    lgStrPrevKey	 = ""                
    lgSortKey        = 1
    lgPageNo         = ""    
End Sub

'==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ���� 
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ���� 
'=========================================================================================================
Sub SetDefaultVal()
	frm1.txtPlantCd.value=parent.gPlant
	frm1.txtPlantNm.value=parent.gPlantNm	
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement
End Sub

'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
End Sub


'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("M1211QA1","S","A","V20030615",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock("A") 
End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock(ByVal pOpt)
    If pOpt = "A" Then
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
	Else
	
	End If
End Sub

'------------------------------------  PopZAdoConfigGrid()  ----------------------------------------------
'	Name : PopZAdoConfigGrid()
'	Description : Group Condition PopUp
'---------------------------------------------------------------------------------------------------------
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenOrderByPopup("A")
End Sub

'===========================================================================
' Function Name : OpenOrderByPopup
' Function Desc : OpenOrderByPopup Reference Popup
'===========================================================================
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

<!-- '------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- !-->
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "����"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "�����"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)		
		frm1.txtPlantNm.value= arrRet(1)	
		frm1.txtPlantCd.focus
	End If	
	frm1.txtItemCd.value=""
	frm1.txtItemNm.value=""
End Function

'===================================================================================================================================
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD

	If lgIsOpenPop = True Then Exit Function

	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if

	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' ��12!MO"�� ���� -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- ��¥ 
	arrParam(5) = ""		'-- ����(b_item_by_plant a, b_item b: and ���� ����)

	arrField(0) = 1 ' -- ǰ���ڵ� 
	arrField(1) = 2 ' -- ǰ��� 

	iCalledAspName = AskPRAspName("B1B11PA3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		lgIsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value	= arrRet(0)
		frm1.txtItemNm.Value	= arrRet(1)
		frm1.txtItemCd.focus
	End If
End Function



'------------------------------------------  OpenSupplier()  -------------------------------------------------
'	Name : OpenSupplier()
'	Description :
'---------------------------------------------------------------------------------------------------------
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Or UCase(frm1.txtSupplierCd.ClassName) = UCase(parent.UCN_PROTECTED) Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "����ó"			
	arrParam(1) = "B_Biz_Partner"		
	
	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
	arrParam(3) = ""	
	
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "		
	arrParam(5) = "����ó"			
	
	arrField(0) = "BP_Cd"				
	arrField(1) = "BP_NM"				

	arrHeader(0) = "����ó"			
	arrHeader(1) = "����ó��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus
		Exit Function
	Else
		frm1.txtSupplierCd.Value    = arrRet(0)		
		frm1.txtSupplierNm.Value    = arrRet(1)	
		frm1.txtSupplierCd.focus
	End If	
End Function

'==========================================   CookiePage()  ======================================
'	Name : CookiePage()
'	Description : JUMP�� Loadȭ������ ���Ǻη� Value
'====================================================================================================
Function WriteCookiePage(ByVal Kubun)
	Dim strTemp, arrVal
	Dim IntRetCD	
    
	With frm1.vspdData
		If .ActiveRow > 0 Then
			Call WriteCookie("m1211qa1_plantcd" , frm1.txtPlantCd.Value)
			Call WriteCookie("m1211qa1_itemcd" , Trim(GetSpreadText(frm1.vspdData,GetKeyPos("A",1),.ActiveRow ,"X","X")))
			Call WriteCookie("m1211qa1_suppliercd" , Trim(GetSpreadText(frm1.vspdData,GetKeyPos("A",4),.ActiveRow ,"X","X")))
		End If
		If Kubun = 1 Then
			Call PgmJump(BIZ_PGM_JUMP_ID1)	
		Else
			Call PgmJump(BIZ_PGM_JUMP_ID2)	
		End If					
	End With
End Function
'====================================================================================================
Sub ReadCookiePage()
	if Trim(ReadCookie("m1211ma1_plantcd")) = "" then Exit Sub
	
	frm1.txtPlantCd.Value	 = ReadCookie("m1211ma1_plantcd")
	frm1.txtItemCd.Value	 = ReadCookie("m1211ma1_itemcd")
	frm1.txtSupplierCd.Value = ReadCookie("m1211ma1_suppliercd")
	
	Call MainQuery()
	
	Call WriteCookie("m1211ma1_plantcd","")
	Call WriteCookie("m1211ma1_itemcd","")
	Call WriteCookie("m1211ma1_suppliercd","")
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029
'    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")   
    
	Call InitVariables						
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call SetToolbar("1100000000001111")		
    Call ReadCookiePage()	
    frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Set gActiveSpdSheet = frm1.vspdData
    SetPopupMenuItemInf("00000000001")
	
	gMouseClickStatus = "SPC"   
	   
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    
    Call SetSpreadColumnValue("A",Frm1.vspdData, Col, Row)  
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgPageNo <> "" Then							
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If			
			Call DisableToolBar(parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If    
End Sub

<!--
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
-->
Function FncQuery() 
    Dim IntRetCD 
    Err.Clear      
    
    FncQuery = False                                        

    '-----------------------
    'Erase contents area
    '-----------------------
	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
			
	Call InitVariables    														
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then
		Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function
       
    FncQuery = True	    
End Function


<!--
'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
-->
Function FncPrint()
	ggoSpread.Source = frm1.vspdData
    Call parent.FncPrint()
End Function

<!--
'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
-->
Function FncExcel()
	ggoSpread.Source = frm1.vspdData 
    Call parent.FncExport(parent.C_Multi)			
End Function 
		
<!--
'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
-->
Function FncFind() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncFind(parent.C_Multi , False)    
End Function

<!--
'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
-->
Function FncExit()    
    FncExit = True
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'*********************************************************************************************************
<!--
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
-->
Function DbQuery() 
	Dim strVal
    Err.Clear                               

    DbQuery = False
    
    if LayerShowHide(1) = False then
       Exit Function
    end if

    With frm1
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
	    strVal = strVal & "&txtPlantCd=" & .hdnPlant.value
	    strVal = strVal & "&txtItemCd=" & .hdnItem.value
	    strVal = strVal & "&txtSupplierCd=" & .hdnSupplier.value
		strVal = strVal & "&rdoUseflg=" & .hdnflg.Value
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	Else
	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
	    strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
	    strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
	    strVal = strVal & "&txtSupplierCd=" & Trim(.txtSupplierCd.value)
	    
		if .rdoUseflg(0).checked=true then
			strVal = strVal & "&rdoUseflg=" & Trim(.rdoUseflg(0).value)
	    elseif .rdoUseflg(1).checked=true then
			strVal = strVal & "&rdoUseflg=" & Trim(.rdoUseflg(1).value)
		else 
			strVal = strVal & "&rdoUseflg=" & Trim(.rdoUseflg(2).value)
		end if 
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    End If
        strVal = strVal & "&lgPageNo="   & lgPageNo         
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
  	    strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
	
	Call RunMyBizASP(MyBizASP, strVal)		
    
    End With
    
    DbQuery = True
   
    Call SetToolbar("1100000000011111")									
End Function

<!--
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
-->
Function DbQueryOk()						
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE				
    Call SetToolbar("1100000000011111")		
    Call ggoOper.LockField(Document, "Q")	
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspddata.focus
	Else
		frm1.txtPlantCd.focus
	End If
	Set gActiveElement = document.activeElement
End Function


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!--
'#########################################################################################################
'       					6. Tag�� 
'	���: Tag�κ� ���� 
	' �Է� �ʵ��� ��� MaxLength=? �� ��� 
	' CLASS="required" required  : �ش� Element�� Style �� Default Attribute 
		' Normal Field�϶��� ������� ���� 
		' Required Field�϶��� required�� �߰��Ͻʽÿ�.
		' Protected Field�϶��� protected�� �߰��Ͻʽÿ�.
			' Protected Field�ϰ�� ReadOnly �� TabIndex=-1 �� ǥ���� 
	' Select Type�� ��쿡�� className�� ralargeCB�� ���� width="153", rqmiddleCB�� ���� width="90"
	' Text-Transform : uppercase  : ǥ�Ⱑ �빮�ڷ� �� �ؽ�Ʈ 
	' ���� �ʵ��� ��� 3���� Attribute ( DDecPoint DPointer DDataFormat ) �� ��� 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ǰ�񺰰���ó</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH="*" align=right></td>
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
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="����" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant() " OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
														   <INPUT TYPE=TEXT ALT="����" name="txtPlantNm" tag="14X"></TD>
									<TD CLASS="TD5" NOWRAP>ǰ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="ǰ��" NAME="txtItemCd" SIZE=10 MAXLENGTH=18 STYLE="text-transform:uppercase" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">
														   <INPUT TYPE=TEXT ALT="ǰ��" NAME="txtItemNm" SIZE=20 CLASS=protected readonly=true tag="14X" tabindex = -1></TD>
														   
								</TR>
								<tr>
									<TD CLASS="TD5" NOWRAP>����ó</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="����ó" NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
														   <INPUT TYPE=TEXT ALT="����ó" name="txtSupplierNm" SIZE=20 tag="14X"></TD>
									<TD CLASS="TD5" NOWRAP>��뿩��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio ALT="��뿩��" class="radio" NAME="rdoUseflg" ID="rdoUseflg1" checked value = "A" tag="1X">
														   <label for="rdoUseflg1">��ü</label>
														   <INPUT TYPE=radio ALT="��뿩��" class="radio" NAME="rdoUseflg" ID="rdoUseflg2" value = "Y" tag="1X">
														   <label for="rdoUseflg2">��</label>
														   <INPUT TYPE=radio ALT="��뿩��" class="radio" NAME="rdoUseflg" ID="rdoUseflg3" value = "N" tag="1X">
														   <label for="rdoUseflg3">�ƴϿ�</label></TD>
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
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
					<TD WIDTH=10>&nbsp;</TD>
					<td WIDTH="*" align="right"><a href="VBSCRIPT:WriteCookiePage(1)">ǰ�񺰰���ó���</a> | <a href="VBSCRIPT:WriteCookiePage(2)">����ó����к���</a></td>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
			</table>
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPlant" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItem" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSupplier" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFlg" tag="24">
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
