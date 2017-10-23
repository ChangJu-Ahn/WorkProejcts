<%@ LANGUAGE="VBSCRIPT" %>

<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ����ä�ǰ��� 
'*  3. Program ID           : S5114MA2
'*  4. Program Name         : ����ä���ϰ�Ȯ�� 
'*  5. Program Desc         :
'*  6. Comproxy List        : S51115BatchArProcessSvr
'*  7. Modified date(First) : 2000/04/20
'*  8. Modified date(Last)  : 2003/07/03
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/04/20 : 3rd ȭ�� layout & ASP Coding
'*                            -2000/08/11 : 4th ȭ�� layout
'*                            -2001/12/19 : Date ǥ������ 
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             

'==========================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim iDBSYSDate
Dim EndDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

Const BIZ_PGM_ID = "s5114mb2.asp"

 Dim C_Select			
 Dim C_PostFlag		
 Dim C_BillNo			
 Dim C_BillDate		
 Dim C_SoldToPartyCd	
 Dim C_SoldToPartyNm	
 Dim C_Cur			
 Dim C_BillAmt			
 Dim C_BillVatAmt		
 Dim C_IncomeAmt		
 Dim C_BillToPartyCd	
 Dim C_BillToPartyNm	
 Dim C_TransTypeCd		
 Dim C_TransTypeNm		
 Dim C_BizBpCd			
 Dim C_BizBpNm
 Dim C_SalesGrpCd
 Dim C_SalesGrpNm

Dim IsOpenPop						' Popup
Dim lgLngMaxRows

'========================================
Sub InitSpreadPosVariables()
	C_Select		= 1	
	C_PostFlag		= 2	
	C_BillNo		= 3	
	C_BillDate		= 4	
	C_SoldToPartyCd = 5	
	C_SoldToPartyNm = 6	
	C_Cur			= 7	
	C_BillAmt		= 8	 
	C_BillVatAmt	= 9	
	C_IncomeAmt		= 10	
	C_BillToPartyCd = 11	
	C_BillToPartyNm	= 12	
	C_TransTypeCd	= 13	
	C_TransTypeNm	= 14	
	C_BizBpCd		= 15	
	C_BizBpNm       = 16
	C_SalesGrpCd	= 17
	C_SalesGrpNm	= 18
End Sub

'========================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE            
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                           

    lgStrPrevKey = ""
    lgLngCurRows = 0      

    Call SetToolbar("11000000000011")

End Sub

'=========================================
Sub SetDefaultVal()
	frm1.txtReqDateFrom.Text = UNIGetFirstDay(EndDate, Parent.gDateFormat)
	frm1.txtReqDateTo.Text = EndDate
	frm1.txtReqDateFrom.focus

	lgBlnFlgChgValue = False
End Sub

'========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
    <% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()

	With frm1.vspdData
		
		ggoSpread.Source = frm1.vspdData
		'patch version
		ggoSpread.Spreadinit "V20030701",,parent.gAllowDragDropSpread    
		.ReDraw = False
	    
	    .MaxRows = 0 : .MaxCols = 0	    
	    .MaxCols = C_SalesGrpNm + 1	
	    
	    Call GetSpreadColumnPos("A")
       
		ggoSpread.SSSetCheck	C_Select,		"����",				6,,,true
	    ggoSpread.SSSetEdit		C_PostFlag,		"Ȯ������",			10,2,,,2
		ggoSpread.SSSetEdit		C_BillNo,		"����ä�ǹ�ȣ",		18,,,,2
	    ggoSpread.SSSetDate		C_BillDate,		"����ä����",		10,2,parent.gDateFormat
	    ggoSpread.SSSetEdit		C_SoldToPartyCd,"�ֹ�ó",			15,,,,2
	    ggoSpread.SSSetEdit		C_SoldToPartyNm,"�ֹ�ó��",			25
	    ggoSpread.SSSetEdit		C_Cur,			"ȭ��",				10, 2
		ggoSpread.SSSetFloat	C_BillAmt,		"����ä�Ǿ�",		15,"A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C_BillVatAmt,	"����ä��VAT�ݾ�",	15,"A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C_IncomeAmt,	"���ݾ�",			15,"A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	    ggoSpread.SSSetEdit		C_BillToPartyCd,"����ó",			15,,,,2
	    ggoSpread.SSSetEdit		C_BillToPartyNm,"����ó��",			25
	    ggoSpread.SSSetEdit		C_TransTypeCd,	"����ä������",		15,,,,2
	    ggoSpread.SSSetEdit		C_TransTypeNm,	"����ä�����¸�",	25
	    ggoSpread.SSSetEdit		C_BizBpCd,		"���ݽŰ�����",	20,,,,2
	    ggoSpread.SSSetEdit		C_BizBpNm,		"���ݽŰ������", 30
	    
	    ggoSpread.SSSetEdit		C_SalesGrpCd,	"�����׷�",	10,,,,2	    
	    ggoSpread.SSSetEdit		C_SalesGrpNm,	"�����׷��", 15
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)				'��: ������Ʈ�� ��� Hidden Column
   		.ReDraw = true

		Call SetSpreadLock()

    End With

End Sub

'========================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
						            
                 C_Select			= iCurColumnPos(1)     
                 C_PostFlag			= iCurColumnPos(2)     	
                 C_BillNo			= iCurColumnPos(3)     	
                 C_BillDate			= iCurColumnPos(4)     
                 C_SoldToPartyCd	= iCurColumnPos(5)     
                 C_SoldToPartyNm	= iCurColumnPos(6)     	
                 C_Cur				= iCurColumnPos(7)     	
                 C_BillAmt			= iCurColumnPos(8)      
                 C_BillVatAmt		= iCurColumnPos(9)     	
                 C_IncomeAmt		= iCurColumnPos(10)     
                 C_BillToPartyCd	= iCurColumnPos(11)     
                 C_BillToPartyNm	= iCurColumnPos(12)     
                 C_TransTypeCd		= iCurColumnPos(13)     	
                 C_TransTypeNm		= iCurColumnPos(14)     
                 C_BizBpCd			= iCurColumnPos(15)     
                 C_BizBpNm			= iCurColumnPos(16)     
                 C_SalesGrpCd		= iCurColumnPos(17) 
				 C_SalesGrpNm		= iCurColumnPos(18) 
                       
    End Select    
End Sub

'========================================
Sub SetSpreadLock()
	Dim GCol

	ggoSpread.Source = frm1.vspdData
			
	frm1.vspdData.ReDraw = False
			
	For GCol = C_PostFlag To C_BizBpNm
		ggoSpread.SpreadLock GCol, -1
	Next

	frm1.vspdData.ReDraw = True
End Sub

' ���� �߻��� �ش� ��ġ�� Focus�̵� 
'=========================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow

           If Not Frm1.vspdData.ColHidden Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
       Next
    End If   
End Sub

'=========================================
Function OpenConPopUp(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	Case 1					 '����ä������ 
		arrParam(1) = "s_bill_type_config"								
		arrParam(2) = Trim(frm1.txtBillTypeCd.value)					
		arrParam(3) = ""												
		arrParam(4) = ""										
		arrParam(5) = "����ä������"									
	
		arrField(0) = "bill_type"										
		arrField(1) = "bill_type_nm"										
    
		arrHeader(0) = "����ä������"									
		arrHeader(1) = "����ä�����¸�"
		
		frm1.txtBillTypeCd.focus									

	Case 2					 '����ó 
		arrParam(1) = "B_BIZ_PARTNER"									
		arrParam(2) = Trim(frm1.txtBillToPartyCd.value)					
		arrParam(3) = ""												
		arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"							
		arrParam(5) = "����ó"										
		
		arrField(0) = "BP_CD"											
		arrField(1) = "BP_NM"											
		
		arrHeader(0) = "����ó"										
		arrHeader(1) = "����ó��"									

		frm1.txtBillToPartyCd.focus									

	Case 3					 '���ݽŰ����� 
		arrParam(1) = "B_TAX_BIZ_AREA"										
		arrParam(2) = Trim(frm1.txtTaxBizAreaCd.value)					
		arrParam(3) = ""												
		arrParam(4) = ""												
		arrParam(5) = "���ݽŰ�����"								
					
		arrField(0) = "TAX_BIZ_AREA_CD"										
		arrField(1) = "TAX_BIZ_AREA_NM"										
				    
		arrHeader(0) = "�����"										
		arrHeader(1) = "������"									
		
		frm1.txtTaxBizAreaCd.focus									

	Case 4
		arrParam(1) = "B_BIZ_PARTNER"									
		arrParam(2) = Trim(frm1.txtSoldToPartyCd.value)					
		arrParam(3) = ""												
		arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"							
		arrParam(5) = "�ֹ�ó"										
		
		arrField(0) = "BP_CD"											
		arrField(1) = "BP_NM"											
					
		arrHeader(0) = "�ֹ�ó"										
		arrHeader(1) = "�ֹ�ó��"									

		frm1.txtSoldToPartyCd.focus	
		
	Case 11
		arrParam(1) = "B_SALES_GRP"									
		arrParam(2) = Trim(frm1.txtSalesGrpCd.value)					
		arrParam(3) = ""												
		arrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & ""							
		arrParam(5) = "�����׷�"
		
		arrField(0) = "SALES_GRP"											
		arrField(1) = "SALES_GRP_NM"											
					
		arrHeader(0) = "�����׷�"
		arrHeader(1) = "�����׷��"									

		frm1.txtSalesGrpCd.focus									

	Case 12
		arrParam(1) = "B_SALES_ORG"									
		arrParam(2) = Trim(frm1.txtSalesOrgCd.value)					
		arrParam(3) = ""												
		arrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & ""							
		arrParam(5) = "�����׷�"
		
		arrField(0) = "SALES_ORG"											
		arrField(1) = "SALES_ORG_NM"											
					
		arrHeader(0) = "��������"
		arrHeader(1) = "����������"

		frm1.txtSalesOrgCd.focus
									
	End Select

	arrParam(0) = arrParam(5)											' �˾� ��Ī 

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then	Call SetConPopUp(arrRet,iWhere)
	
End Function

'=========================================
Function SetConPopUp(Byval arrRet,Byval iWhere)

	Select Case iWhere
	Case 1
		frm1.txtBillTypeCd.value = arrRet(0)
		frm1.txtBillTypeNm.value = arrRet(1) 
	Case 2
		frm1.txtBillToPartyCd.value = arrRet(0)
		frm1.txtBillToPartyNm.value = arrRet(1) 
	Case 3
		frm1.txtTaxBizAreaCd.value = arrRet(0)
		frm1.txtTaxBizAreaNm.value = arrRet(1) 
	Case 4
		frm1.txtSoldToPartyCd.value = arrRet(0)
		frm1.txtSoldToPartyNm.value = arrRet(1) 
	Case 11
		frm1.txtSalesGrpCd.value = arrRet(0)
		frm1.txtSalesGrpNm.value = arrRet(1) 
	Case 12
		frm1.txtSalesOrgCd.value = arrRet(0)
		frm1.txtSalesOrgNm.value = arrRet(1) 	
	End Select

End Function

'========================================
Sub Form_Load()

    Call LoadInfTB19029														
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   

	Call InitSpreadSheet
	Call SetDefaultVal	
	Call InitVariables														
End Sub

'==========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'==========================================
Sub txtReqDateFrom_DblClick(Button)
	If Button = 1 Then
		frm1.txtReqDateFrom.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtReqDateFrom.Focus
	End If
End Sub

'==========================================
Sub txtReqDateTo_DblClick(Button)
	If Button = 1 Then
		frm1.txtReqDateTo.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtReqDateTo.Focus
	End If
End Sub

'==========================================
Sub txtReqDateFrom_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'==========================================
Sub txtReqDateTo_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

' ��ü���� 
'========================================
Sub chkSelectAll_onClick()
	If frm1.vspdData.MaxRows = 0 Then Exit Sub

	ggoSpread.Source = frm1.vspdData	
	With frm1.vspdData
		.Row = 1			:	.Row2 = .MaxRows
		
		' ��ü���� 
		If frm1.chkSelectAll.checked Then
			' Row Header ����(����)
			.Col = 0			:	.Col2 = 0
			.Clip = Replace(.Clip, vbCrLf, ggoSpread.UpdateFlag & vbCrLf)
			
			' ���ù�ư�� ���ÿ��� ���� 
			.Col = C_Select		:	.Col2 = C_Select
			.Clip = Replace(.Clip, "0", "1")
			
		' ��ü���� ��� 
		Else
			' Row Header ����(����)
			.Col = 0			:	.Col2 = 0
			.Clip = Replace(.Clip, ggoSpread.UpdateFlag, "")

			.Col = C_SELECT		:	.Col2 = C_SELECT
			.Clip = Replace(.Clip, "1", "0")
		End if
	End With

	' Active Cell ����	
	Call SetActiveCell(frm1.vspdData,C_Select, 1,"M","X","X")
End Sub

'==========================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

    On Error Resume Next
	If lgIntFlgMode = parent.OPMD_CMODE Then Exit Sub

	If Col = C_Select And Row > 0 Then
	    Select Case ButtonDown
	    Case 0
			Call FncCancel
	    Case 1
			ggoSpread.Source = frm1.vspdData
			ggoSpread.UpdateRow Row
	    End Select
    End If

End Sub

'==========================================
Sub vspdData_Click(ByVal Col , ByVal Row)
	Call SetPopupMenuItemInf("1101111111")
    
	gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col			'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		Exit Sub
	End If 
     
End Sub

'========================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'==========================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'========================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    If Col <= C_Select Or NewCol <= C_Select Then
        Cancel = True
        Exit Sub
    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then Exit Sub
	    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then
		If CheckRunningBizProcess Then Exit Sub
		    
		Call DisableToolBar(Parent.TBC_QUERY)
		Call DBQuery
	End If
End Sub

'========================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear

	If ValidDateCheck(frm1.txtReqDateFrom, frm1.txtReqDateTo) = False Then Exit Function

	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										
    Call InitVariables															

	frm1.chkSelectAll.checked = False
	
	If frm1.rdoPostFlagN.checked = True Then
		frm1.txtPostFlag.value = frm1.rdoPostFlagN.value
	ElseIf frm1.rdoPostFlagY.checked = True Then
		frm1.txtPostFlag.value = frm1.rdoPostFlagY.value
	End If

    Call DbQuery																

    FncQuery = True																
        
End Function

'========================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    
	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") 
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                      
    Call ggoOper.LockField(Document, "N")                                       
    Call SetDefaultVal
    Call InitVariables															

    FncNew = True																

End Function

'========================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear
    
	ggoSpread.Source = frm1.vspdData	
	If Not ggoSpread.SSCheckChange Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
	    Exit Function
    End If
    
    If ggoSpread.SSDefaultCheck = False Then Exit Function

    CAll DbSave
    
    FncSave = True                                                          
    
End Function

'========================================
Function FncCancel() 
    On Error Resume Next
	If frm1.vspdData.MaxRows < 1 Then Exit Function
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo
    
    Call FormatSpreadCellByCurrency(frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow)
End Function

'========================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'========================================
Function FncExcel() 
	Call parent.FncExport(parent.C_SINGLEMULTI)
End Function

'========================================
Function FncFind() 
	Call parent.FncFind(parent.C_SINGLEMULTI, False)
End Function

'========================================
Sub FncSplitColumn()    
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub

'========================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")   '�� �ٲ�κ� 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================
Function DbQuery() 

    Err.Clear
    
	If LayerShowHide(1) = False Then Exit Function 
    
    DbQuery = False                                                         
    
    Dim iStrVal

	With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			iStrVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									
			iStrVal = iStrVal & "&txtSoldToPartyCd=" & Trim(.HSoldToParty.value)			
			iStrVal = iStrVal & "&txtBillToPartyCd=" & Trim(.HBillToParty.value)			
			iStrVal = iStrVal & "&txtReqDateFrom=" & Trim(.HReqDateFrom.value)
			iStrVal = iStrVal & "&txtReqDateTo=" & Trim(.HReqDateTo.value)
			iStrVal = iStrVal & "&txtBillTypeCd=" & Trim(.HBillTypeCd.value)
			iStrVal = iStrVal & "&txtTaxBizAreaCd=" & Trim(.HTaxBizAreaCd.value)
			iStrVal = iStrVal & "&txtPostFlag=" & Trim(.HPostFlag.value)
			iStrVal = iStrVal & "&txtSalesGrpCd=" & Trim(.HSalesGrpCd.value)
			iStrVal = iStrVal & "&txtSalesOrgCd=" & Trim(.HSalesOrgCd.value)
			iStrVal = iStrVal & "&lgStrPrevKey=" & lgStrPrevKey
		Else
			iStrVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									
			iStrVal = iStrVal & "&txtSoldToPartyCd=" & Trim(.txtSoldToPartyCd.value)			
			iStrVal = iStrVal & "&txtBillToPartyCd=" & Trim(.txtBillToPartyCd.value)			
			iStrVal = iStrVal & "&txtReqDateFrom=" & Trim(.txtReqDateFrom.Text)
			iStrVal = iStrVal & "&txtReqDateTo=" & Trim(.txtReqDateTo.Text)
			iStrVal = iStrVal & "&txtBillTypeCd=" & Trim(.txtBillTypeCd.value)
			iStrVal = iStrVal & "&txtTaxBizAreaCd=" & Trim(.txtTaxBizAreaCd.value)
			iStrVal = iStrVal & "&txtPostFlag=" & Trim(.txtPostFlag.value)			
			iStrVal = iStrVal & "&txtSalesGrpCd=" & Trim(.txtSalesGrpCd.value)
			iStrVal = iStrVal & "&txtSalesOrgCd=" & Trim(.txtSalesOrgCd.value)
			iStrVal = iStrVal & "&lgStrPrevKey=" & lgStrPrevKey
		End if
	
		' �ϰ���ȸ���� 
		If frm1.chkBatchQuery.checked Then
			iStrVal = iStrVal & "&txtBatchQuery=Y"
		Else
			iStrVal = iStrVal & "&txtBatchQuery=N"
		End If

		lgLngMaxRows = .vspdData.MaxRows
		iStrVal = iStrVal & "&txtMaxRows=" & lgLngMaxRows

	End With
	
	Call RunMyBizASP(MyBizASP, iStrVal)												
	
    DbQuery = True																	

End Function

'========================================
Function DbQueryOk()														
    lgIntFlgMode = parent.OPMD_UMODE												
	lgBlnFlgChgValue = False
	Call SetToolbar("11001001000111")
	
    Call FormatSpreadCellByCurrency(lgLngMaxRows, frm1.vspdData.MaxRows)
    
	frm1.vspdData.focus
End Function

'========================================
Function DbSave() 

    Err.Clear																
 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal,strDel
	Dim intInsrtCnt
			
	If Not LayerShowHide(1) Then Exit Function 

    DbSave = False                                                          
    
	With frm1
		.txtMode.value = parent.UID_M0002
    
		lGrpCnt = 1
    
		strVal = ""
		strDel = ""
		intInsrtCnt = 1
		
		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = C_Select

			If .vspdData.Text = 1 Then
				strVal = strVal & lRow & parent.gColSep

				.vspdData.Col = C_BillNo
				strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep				
						
				lGrpCnt = lGrpCnt + 1
				intInsrtCnt = intInsrtCnt + 1
			End If
		Next
	
		.txtSpread.value =  strVal		

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	
	End With
	
    DbSave = True                                                           
    
End Function

'========================================
Function DbSaveOk()

    Call InitVariables
	frm1.vspdData.MaxRows = 0
    Call MainQuery()

End Function

' ȭ�󺰷� Cell Formating�� �缳���Ѵ�.
Sub FormatSpreadCellByCurrency(ByVal pvLngStartRow, ByVal pvLngEndRow)
	frm1.vspdData.Redraw = False
	Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,pvLngStartRow,pvLngEndRow,C_Cur,C_BillAmt,"A" ,"I","X","X")         
	Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,pvLngStartRow,pvLngEndRow,C_Cur,C_BillVatAmt,"A" ,"I","X","X")         
	Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,pvLngStartRow,pvLngEndRow,C_Cur,C_IncomeAmt,"A" ,"I","X","X")         
	frm1.vspdData.Redraw = True
End Sub

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����ä���ϰ�Ȯ��</font></td>
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
									<TD CLASS="TD5" NOWRAP>����ä����</TD>
									<TD CLASS="TD6" NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtReqDateFrom" Alt="����ä�ǽ�����" CLASS="FPDTYYYYMMDD" tag="12X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
												<TD>
													&nbsp;~&nbsp;
												</TD>
												<TD>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtReqDateTo" Alt="����ä��������" CLASS="FPDTYYYYMMDD" tag="12X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
											</TR>
										</TABLE>
									</TD>
									<TD CLASS=TD5 NOWRAP>�ֹ�ó</TD>
									<TD CLASS=TD6><INPUT NAME="txtSoldToPartyCd" TYPE="Text" Alt="�ֹ�ó" MAXLENGTH=10 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenPopUp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp 4">&nbsp;<INPUT NAME="txtSoldToPartyNm" TYPE="Text" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>����ä������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBillTypeCd" TYPE="Text" MAXLENGTH="20" SIZE=10 Alt="����ä������" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenPopUp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp 1">&nbsp;<INPUT NAME="txtBillTypeNm" TYPE="Text" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>����ó</TD>
									<TD CLASS="TD6"><INPUT NAME="txtBillToPartyCd" TYPE="Text" Alt="����ó" MAXLENGTH=10 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenPopUp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp 2">&nbsp;<INPUT NAME="txtBillToPartyNm" TYPE="Text" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�����׷�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesGrpCd" TYPE="Text" MAXLENGTH="4" SIZE=10 Alt="�����׷�" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenPopUp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp 11">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>��������</TD>
									<TD CLASS="TD6"><INPUT NAME="txtSalesOrgCd" TYPE="Text" Alt="��������" MAXLENGTH=4 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenPopUp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp 12">&nbsp;<INPUT NAME="txtSalesOrgNm" TYPE="Text" SIZE=25 tag="14"></TD>
								</TR>
								<TR>	
									<TD CLASS=TD5 NOWRAP>���ݽŰ�����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxBizAreaCd" TYPE="Text" Alt="���ݽŰ�����" MAXLENGTH="10" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp 3">&nbsp;<INPUT NAME="txtTaxBizAreaNm" TYPE="Text" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>Ȯ������</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoPostFlag" id="rdoPostFlagY" value="Y" tag = "11">
											<label for="rdoPostFlagY">Ȯ��</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoPostFlag" id="rdoPostFlagN" value="N" tag = "11" checked>
											<label for="rdoPostFlagN">��Ȯ��</label></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�ϰ���ȸ</TD>
									<TD CLASS=TD6>
										<INPUT TYPE=CHECKBOX NAME="chkBatchQuery" ID="chkBatchQuery" tag="11" Class="Check">
									</TD>
									<TD CLASS=TD5 NOWRAP>��ü����</TD>
									<TD CLASS=TD6 >
										<INPUT TYPE=CHECKBOX NAME="chkSelectAll" ID="chkSelectAll" tag="11" Class="Check">
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  TABINDEX= -1></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= -1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX= -1>
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA>

<INPUT TYPE=HIDDEN NAME="txtPostDate" tag="24" TABINDEX= -1>
<INPUT TYPE=HIDDEN NAME="txtPostFlag" tag="24" TABINDEX= -1>

<INPUT TYPE=HIDDEN NAME="HSoldToParty" tag="24" TABINDEX= -1>
<INPUT TYPE=HIDDEN NAME="HBillToParty" tag="24" TABINDEX= -1>
<INPUT TYPE=HIDDEN NAME="HReqDateFrom" tag="24" TABINDEX= -1>
<INPUT TYPE=HIDDEN NAME="HReqDateTo" tag="24" TABINDEX= -1>
<INPUT TYPE=HIDDEN NAME="HPostFlag" tag="24" TABINDEX= -1>
<INPUT TYPE=HIDDEN NAME="HBillTypeCd" tag="24" TABINDEX= -1>
<INPUT TYPE=HIDDEN NAME="HTaxBizAreaCd" tag="24" TABINDEX= -1>
<INPUT TYPE=HIDDEN NAME="HSalesGrpCd" tag="24" TABINDEX= -1>
<INPUT TYPE=HIDDEN NAME="HSalesOrgCd" tag="24" TABINDEX= -1>

</FORM>
  <DIV ID="MousePT" NAME="MousePT">
  	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
  </DIV>
</BODY>
</HTML>
