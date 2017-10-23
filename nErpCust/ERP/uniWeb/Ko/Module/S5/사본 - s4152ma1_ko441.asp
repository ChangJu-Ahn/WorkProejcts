<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : s4152ma1_ko441
'*  4. Program Name         : ���ݾ׻���������ȸ������(KO441)
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2008/07/29
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              : 
'=======================================================================================================
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                 '��: indicates that All variables must be declared in advance

Dim prDBSYSDate

Dim EndDate ,StartDate
Dim lgStrComDateType		'Company Date Type�� ����(��� Mask�� �����.)

prDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAToB(prDBSYSDate ,parent.gServerDateFormat,parent.gDateFormat)               'Convert DB date type to Company

StartDate = UniDateAdd("m", -1, EndDate,parent.gDateFormat)


Const BIZ_PGM_ID      = "s4152mb1_ko441.asp"            '��: �����Ͻ� ���� ASP�� 

Dim C_Item_Cd
Dim C_Item_Nm
Dim C_Plant_Nm
Dim C_Gi_Dt
Dim C_Pl_No
Dim C_In_Lot_No
Dim C_Cust_Lot_No
Dim C_Gi_Type
Dim C_Gi_Qty
Dim C_Gi_Unit
Dim C_Gi_Price
Dim C_Gi_Amt
Dim C_Gi_Amt_Loc
Dim C_xch_rate
Dim C_Cur
Dim C_Type_Nm
Dim C_Price1
Dim C_Amt1
Dim C_Price2
Dim C_Amt2
Dim C_Price3
Dim C_Amt3
Dim C_Price4
Dim C_Amt4
Dim C_Price5
Dim C_Amt5
Dim C_Price6
Dim C_Amt6
Dim C_Price7
Dim C_Amt7
Dim C_Price8
Dim C_Amt8
Dim C_Price9
Dim C_Amt9
Dim C_Price10
Dim C_Amt10
Dim C_Price11
Dim C_Amt11
Dim C_Price12
Dim C_Amt12
Dim C_Price13
Dim C_Amt13
Dim C_Price14
Dim C_Amt14
Dim C_Price15
Dim C_Amt15
Dim C_Po_No
Dim C_Pgm_nm
Dim C_Trans_Time
Dim C_OutType_Sub


<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim gblnWinEvent

'========================================================================================================
Sub initSpreadPosVariables()  
    Dim i 
	i = 1

	C_Item_Cd             = i : i = i + 1
	C_Item_Nm             = i : i = i + 1
	C_Plant_Nm            = i : i = i + 1
	C_Gi_Dt			      = i : i = i + 1
	C_Pl_No               = i : i = i + 1
	C_In_Lot_No           = i : i = i + 1
	C_Cust_Lot_No         = i : i = i + 1
	C_Gi_Type             = i : i = i + 1
	C_Gi_Qty		      = i : i = i + 1
	C_Gi_Unit			  = i : i = i + 1
	C_Gi_Price			  = i : i = i + 1
	C_Gi_Amt			  = i : i = i + 1 
	C_Gi_Amt_Loc		  = i : i = i + 1
	C_xch_rate  		  = i : i = i + 1
	C_Cur                 = i : i = i + 1
	C_Type_Nm			  = i : i = i + 1
	C_Price1			  = i : i = i + 1 
	C_Amt1				  = i : i = i + 1
	C_Price2			  = i : i = i + 1 
	C_Amt2				  = i : i = i + 1
	C_Price3			  = i : i = i + 1 
	C_Amt3				  = i : i = i + 1
	C_Price4			  = i : i = i + 1 
	C_Amt4				  = i : i = i + 1
	C_Price5			  = i : i = i + 1 
	C_Amt5				  = i : i = i + 1
	C_Price6			  = i : i = i + 1 
	C_Amt6				  = i : i = i + 1
	C_Price7			  = i : i = i + 1 
	C_Amt7				  = i : i = i + 1
	C_Price8			  = i : i = i + 1 
	C_Amt8				  = i : i = i + 1
	C_Price9			  = i : i = i + 1 
	C_Amt9				  = i : i = i + 1
	C_Price10			  = i : i = i + 1 
	C_Amt10				  = i : i = i + 1
	C_Price11			  = i : i = i + 1 
	C_Amt11				  = i : i = i + 1
	C_Price12			  = i : i = i + 1 
	C_Amt12				  = i : i = i + 1
	C_Price13			  = i : i = i + 1 
	C_Amt13				  = i : i = i + 1
	C_Price14			  = i : i = i + 1 
	C_Amt14				  = i : i = i + 1
	C_Price15			  = i : i = i + 1 
	C_Amt15				  = i : i = i + 1
	C_Po_No               = i : i = i + 1
	C_Pgm_nm              = i : i = i + 1
	C_Trans_Time          = i : i = i + 1
	C_OutType_Sub         = i : i = i + 1

End Sub

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count

End Sub

'========================================================================================================= 
Sub SetDefaultVal()
 frm1.txtconBp_cd.focus 
 lgBlnFlgChgValue = False
End Sub


'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub


'========================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	ggoSpread.Source = frm1.vspdData
 	With frm1.vspdData

       ggoSpread.Spreadinit "V20050503",,parent.gAllowDragDropSpread    

       .MaxCols   = C_OutType_Sub														' ��:��: Add 1 to Maxcols
	   .Col       = .MaxCols        : .ColHidden = True
       .MaxRows = 0                                                                  ' ��: Clear spreadsheet data 
      Call GetSpreadColumnPos("A")
	   .ReDraw = false
 			      
		ggoSpread.SSSetEdit     C_Item_Cd,              "ǰ���ڵ�" ,12, 0,,10,2
		ggoSpread.SSSetEdit     C_Item_Nm,              "ǰ���", 18, 0 
		ggoSpread.SSSetEdit     C_Plant_Nm,             "����", 10, 0 
        ggoSpread.SSSetDate     C_Gi_Dt,				"������", 10,2, parent.gDateFormat   'Lock->Unlock/ Date
		ggoSpread.SSSetEdit     C_Pl_No,                "����P/L No" ,12, 0
		ggoSpread.SSSetEdit     C_In_Lot_No,            "��LOT No" ,12, 0
		ggoSpread.SSSetEdit     C_Cust_Lot_No,          "�԰�LOT No", 12, 0  
		ggoSpread.SSSetEdit     C_Gi_Type,              "���TYPE", 10, 0  
		ggoSpread.SSSetFloat    C_Gi_Qty,				"���",10, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetEdit     C_Gi_Unit,              "������", 8, 0 
		ggoSpread.SSSetFloat    C_Gi_Price,				"���ܰ�",10, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Gi_Amt,				"���ݾ�",10, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Gi_Amt_Loc,			"����ڱ��ݾ�",10, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_xch_rate,				"ȯ��", 7, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetEdit     C_Cur,                  "ȯ��", 5, 0,,3,2
		ggoSpread.SSSetFloat    C_Price1,				"���ܰ�1",10, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Amt1,				    "���ݾ�1",10, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Price2,				"���ܰ�2",10, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Amt2,				    "���ݾ�2",10, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Price3,				"���ܰ�3",10, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Amt3,				    "���ݾ�3",10, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Price4,				"���ܰ�4",10, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Amt4,				    "���ݾ�4",10, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Price5,				"���ܰ�5",10, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Amt5,				    "���ݾ�5",10, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Price6,				"���ܰ�6",10, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Amt6,				    "���ݾ�6",10, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Price7,				"���ܰ�7",10, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Amt7,				    "���ݾ�7",10, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Price8,				"���ܰ�8",10, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Amt8,				    "���ݾ�8",10, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Price9,				"���ܰ�9",10, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Amt9,				    "���ݾ�9",10, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Price10,				"���ܰ�10",10, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Amt10,			    "���ݾ�10",10, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Price11,				"���ܰ�11",10, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Amt11,			    "���ݾ�11",10, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Price12,				"���ܰ�12",10, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Amt12,			    "���ݾ�12",10, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Price13,				"���ܰ�13",10, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Amt13,			    "���ݾ�13",10, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Price14,				"���ܰ�14",10, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Amt14,			    "���ݾ�14",10, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Price15,				"���ܰ�15",10, Parent.ggUnitCostNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetFloat    C_Amt15,			    "���ݾ�15",10, Parent.ggAmtOfMoneyNo , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetEdit     C_Po_No,                "P/O No", 14, 0
		ggoSpread.SSSetEdit     C_Pgm_nm,               "PGM��", 14, 0
		ggoSpread.SSSetEdit     C_Trans_Time,           "Trans_Time", 14, 0
		ggoSpread.SSSetEdit     C_OutType_Sub,          "OutType_Sub", 14, 0
		     
		.ReDraw = true   

	'	call ggoSpread.MakePairsColumn(C_Item_Cd,C_Item_Cd_Popup)

		Call ggoSpread.SSSetColHidden(C_Trans_Time,C_OutType_Sub,True)
	'	Call ggoSpread.SSSetColHidden(C_xch_rate,C_xch_rate,True)

		Call SetSpreadLock 
	    
	End With
    
End Sub

'========================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    
	ggoSpread.spreadlock		C_Item_Cd, -1, C_Cur
'	ggoSpread.SSSetRequired		C_Price,  -1, C_Price
	ggoSpread.spreadlock		C_Po_No, -1, C_Pgm_nm
    
    .vspdData.ReDraw = True

    End With
End Sub

'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    
    With frm1
    
    .vspdData.ReDraw = False
    
  '  ggoSpread.SSSetRequired    C_Item_Cd,              pvStartRow, pvEndRow
    ggoSpread.SSSetProtected    C_Item_Cd  ,            pvStartRow, pvEndRow
    ggoSpread.SSSetProtected    C_Item_Nm  ,            pvStartRow, pvEndRow
    ggoSpread.SSSetProtected    C_Gi_Dt,         pvStartRow, pvEndRow
    ggoSpread.SSSetProtected    C_Pl_No,              pvStartRow, pvEndRow
    ggoSpread.SSSetProtected    C_In_Lot_No  ,            pvStartRow, pvEndRow
    ggoSpread.SSSetProtected    C_Cust_Lot_No  ,            pvStartRow, pvEndRow
    ggoSpread.SSSetProtected    C_Gi_Type  ,         pvStartRow, pvEndRow
    ggoSpread.SSSetProtected    C_Gi_Qty,       pvStartRow, pvEndRow
    ggoSpread.SSSetProtected    C_Gi_Unit,       pvStartRow, pvEndRow
    ggoSpread.SSSetProtected    C_Gi_Price,        pvStartRow, pvEndRow
	ggoSpread.SSSetProtected    C_Gi_Amt,       pvStartRow, pvEndRow
	ggoSpread.SSSetProtected    C_Gi_Amt_Loc,       pvStartRow, pvEndRow
	ggoSpread.SSSetProtected    C_Cur,                pvStartRow, pvEndRow
		
    .vspdData.ReDraw = True
    
    End With

End Sub

'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
	Dim i 
	i = 1
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_Item_Cd             = iCurColumnPos(i) : i = i + 1
			C_Item_Nm             = iCurColumnPos(i) : i = i + 1
			C_Plant_Nm            = iCurColumnPos(i) : i = i + 1
			C_Gi_Dt		          = iCurColumnPos(i) : i = i + 1
			C_Pl_No               = iCurColumnPos(i) : i = i + 1
			C_In_Lot_No           = iCurColumnPos(i) : i = i + 1
			C_Cust_Lot_No         = iCurColumnPos(i) : i = i + 1
			C_Gi_Type             = iCurColumnPos(i) : i = i + 1
			C_Gi_Qty		      = iCurColumnPos(i) : i = i + 1
			C_Gi_Unit             = iCurColumnPos(i) : i = i + 1
			C_Gi_Price			  = iCurColumnPos(i) : i = i + 1
			C_Gi_Amt			  = iCurColumnPos(i) : i = i + 1
			C_Gi_Amt_Loc		  = iCurColumnPos(i) : i = i + 1
			C_xch_rate            = iCurColumnPos(i) : i = i + 1
			C_Cur                 = iCurColumnPos(i) : i = i + 1
			C_Price1			  = iCurColumnPos(i) : i = i + 1
			C_Amt1			      = iCurColumnPos(i) : i = i + 1
			C_Price2			  = iCurColumnPos(i) : i = i + 1
			C_Amt2			      = iCurColumnPos(i) : i = i + 1
			C_Price3			  = iCurColumnPos(i) : i = i + 1
			C_Amt3			      = iCurColumnPos(i) : i = i + 1
			C_Price4			  = iCurColumnPos(i) : i = i + 1 
			C_Amt4				  = iCurColumnPos(i) : i = i + 1
			C_Price5			  = iCurColumnPos(i) : i = i + 1 
			C_Amt5				  = iCurColumnPos(i) : i = i + 1
			C_Price6			  = iCurColumnPos(i) : i = i + 1 
			C_Amt6				  = iCurColumnPos(i) : i = i + 1
			C_Price7			  = iCurColumnPos(i) : i = i + 1 
			C_Amt7				  = iCurColumnPos(i) : i = i + 1
			C_Price8			  = iCurColumnPos(i) : i = i + 1 
			C_Amt8				  = iCurColumnPos(i) : i = i + 1
			C_Price9			  = iCurColumnPos(i) : i = i + 1 
			C_Amt9				  = iCurColumnPos(i) : i = i + 1
			C_Price10			  = iCurColumnPos(i) : i = i + 1 
			C_Amt10				  = iCurColumnPos(i) : i = i + 1
			C_Price11			  = iCurColumnPos(i) : i = i + 1 
			C_Amt11				  = iCurColumnPos(i) : i = i + 1
			C_Price12			  = iCurColumnPos(i) : i = i + 1 
			C_Amt12				  = iCurColumnPos(i) : i = i + 1
			C_Price13			  = iCurColumnPos(i) : i = i + 1 
			C_Amt13				  = iCurColumnPos(i) : i = i + 1
			C_Price14			  = iCurColumnPos(i) : i = i + 1 
			C_Amt14				  = iCurColumnPos(i) : i = i + 1
			C_Price15			  = iCurColumnPos(i) : i = i + 1 
			C_Amt15				  = iCurColumnPos(i) : i = i + 1
			C_Po_No               = iCurColumnPos(i) : i = i + 1
			C_Pgm_nm              = iCurColumnPos(i) : i = i + 1
			C_Trans_Time          = iCurColumnPos(i) : i = i + 1
			C_OutType_Sub         = iCurColumnPos(i) : i = i + 1
						
    End Select    
End Sub
'========================== 2.2.6 InitSpreadComboBox()  =====================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'=========================================================================================
Sub InitSpreadComboBox()
    Dim strCD
    Dim strVal		    

    
End Sub
'===========================================================================
Function OpenConSItemDC(Byval strCode, Byval iWhere)

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If gblnWinEvent = True Then Exit Function

 gblnWinEvent = True

 Select Case iWhere
 Case 5
  arrParam(1) = "B_BIZ_PARTNER"      <%' TABLE ��Ī %>
  arrParam(2) = Trim(strCode) <%' Code Condition%>
  arrParam(3) = ""                                    <%' Name Cindition%>
  arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"             <%' Where Condition%>
  arrParam(5) = "��"          <%' TextBox ��Ī %>
 
  arrField(0) = "BP_CD"           <%' Field��(0)%>
  arrField(1) = "BP_NM"           <%' Field��(1)%>
    
  arrHeader(0) = "��"          <%' Header��(0)%>
  arrHeader(1) = "����"            <%' Header��(1)%>
  frm1.txtconBp_cd.focus 
 Case 0
  arrParam(1) = "b_item"                           <%' TABLE ��Ī %>
  arrParam(2) = Trim(strCode)  <%' Code Condition%>
  arrParam(3) = ""                           <%' Name Cindition%>
  arrParam(4) = ""         <%' Where Condition%>
  arrParam(5) = "ǰ��"       <%' TextBox ��Ī %>
 
  arrField(0) = "item_cd"        <%' Field��(0)%>
  arrField(1) = "item_nm"        <%' Field��(1)%>
  arrField(2) = "spec"        <%' Field��(1)%>
    
  arrHeader(0) = "ǰ��"       <%' Header��(0)%>
  arrHeader(1) = "ǰ���"       <%' Header��(1)%> 
  arrHeader(2) = "�԰�"       <%' Header��(1)%>  
 Case 1
  arrParam(1) = "B_USER_DEFINED_MINOR"                           <%' TABLE ��Ī %>
  arrParam(2) = Trim(strCode)  <%' Code Condition%>
  arrParam(3) = ""                           <%' Name Cindition%>
  arrParam(4) = "UD_MAJOR_CD='ZZ002'"         <%' Where Condition%>
  arrParam(5) = "���TYPE"       <%' TextBox ��Ī %>
 
  arrField(0) = "UD_MINOR_CD"        <%' Field��(0)%>
  arrField(1) = "UD_MINOR_NM"        <%' Field��(1)%>
    
  arrHeader(0) = "���TYPE"       <%' Header��(0)%>
  arrHeader(1) = "���TYPE��"       <%' Header��(1)%> 
 Case 4					'���� 
	arrParam(1) = "B_PLANT"								
	arrParam(2) = Trim(strCode)				
	arrParam(4) = ""									
	arrParam(5) = "����"							

	arrField(0) = "PLANT_CD"							
	arrField(1) = "PLANT_NM"							

	arrHeader(0) = "����"							
	arrHeader(1) = "�����"							
	
	frm1.txtPlantCode.focus

 End Select
    
    arrParam(3) = "" 
 arrParam(0) = arrParam(5)        <%' �˾� ��Ī %>

	Select Case iWhere
		Case 0
			arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
		Case Else
           arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
           "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select

 gblnWinEvent = False

 If arrRet(0) = "" Then
	 Exit Function
 Else
	 With frm1
	  Select Case iWhere
	  Case 0
	   .txtconItem_cd.value = arrRet(0) 
	   .txtconItem_nm.value = arrRet(1)   
	  Case 1
	   .txtconOutType.value = arrRet(0) 
	   .txtconOutTypeNm.value = arrRet(1)   
	  Case 4
		.txtPlantCode.value = arrRet(0) 
		.txtPlantName.value = arrRet(1)   
	  Case 5
	   .txtconBp_cd.value = arrRet(0) 
	   .txtconBp_Nm.value = arrRet(1)    
	  End Select
	 End With
 End If 
 
End Function

'===========================================================================
 Function  OpenBp_cd(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
  
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "��"       <%' �˾� ��Ī %>
  arrParam(1) = "B_Biz_Partner"           <%' TABLE ��Ī %>
  arrParam(2) = Trim(strCode)       <%' Code Condition%>
  arrParam(3) = ""         <%' Name Cindition%>
  arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"         <%' Where Condition%>
  arrParam(5) = "��"       <%' TextBox ��Ī %>

  arrField(0) = "Bp_cd"        <%' Field��(0)%>
  arrField(1) = "Bp_nm"        <%' Field��(1)%>

  arrHeader(0) = "��"       <%' Header��(0)%>
  arrHeader(1) = "����"       <%' Header��(1)%>

  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
	 With frm1
	  .vspdData.Col = C_Pl_No
	  .vspdData.Text = arrRet(0)
	  .vspdData.Col = C_Cust_Lot_No
	  .vspdData.Text = arrRet(1)
	 End With
  End If
 End Function


'===========================================================================
 Function  OpenItem_cd(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "ǰ��"       <%' �˾� ��Ī %>
  arrParam(1) = "B_ITEM"           <%' TABLE ��Ī %>
  arrParam(2) = Trim(strCode)       <%' Code Condition%>
  arrParam(3) = ""         <%' Name Cindition%>
  arrParam(4) = ""         <%' Where Condition%>
  arrParam(5) = "ǰ��"       <%' TextBox ��Ī %>

  arrField(0) = "Item_cd"        <%' Field��(0)%>
  arrField(1) = "Item_nm"        <%' Field��(1)%>  
	arrField(2) = "Spec"	        <%' Field��(2)%>
	'arrField(3) = "HH" & parent.gColSep & "BasiC_Output_rate"	        <%' Field��(3)%>

  arrHeader(0) = "ǰ��"       <%' Header��(0)%>
  arrHeader(1) = "ǰ���"       <%' Header��(1)%>
	arrHeader(2) = "�԰�"       <%' Header��(2)%>
	'arrHeader(3) = "����"       <%' Header��(3)%>
	   
  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
	 With frm1
	  .vspdData.Col = C_Item_cd
	  .vspdData.Text = arrRet(0)
	  .vspdData.Col = C_Item_Nm
	  .vspdData.Text = arrRet(1)
	 End With

  End If
 End Function

'===========================================================================
 Function  OpenCur(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "ȭ��"       <%' �˾� ��Ī %>
  arrParam(1) = "B_CURRENCY"      <%' TABLE ��Ī %>
  arrParam(2) = Trim(strCode)       <%' Code Condition%>
  arrParam(3) = ""         <%' Name Cindition%>
  arrParam(4) = ""         <%' Where Condition%>
  arrParam(5) = "ȭ��"       <%' TextBox ��Ī %>

  arrField(0) = "CURRENCY"        <%' Field��(0)%>
  arrField(1) = "CURRENCY_DESC"        <%' Field��(1)%>

  arrHeader(0) = "ȭ��"       <%' Header��(0)%>
  arrHeader(1) = "ȭ���"      <%' Header��(1)%>

  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
	 With frm1
	  .vspdData.Col = C_Cur
	  .vspdData.Text = arrRet(0) 
	 End With
'	 Call vspdData_Change(C_Cur,frm1.vspdData.ActiveRow)

  End If
 End Function


'===========================================================================
 Function  OpenTester(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "TESTER��"       <%' �˾� ��Ī %>
  arrParam(1) = "b_user_defined_minor"      <%' TABLE ��Ī %>
  arrParam(2) = Trim(strCode)       <%' Code Condition%>
  arrParam(3) = ""         <%' Name Cindition%>
  arrParam(4) = " ud_major_cd= 'zz001' "         <%' Where Condition%>
  arrParam(5) = "TESTER��"       <%' TextBox ��Ī %>

  arrField(0) = "ud_minor_cd"        <%' Field��(0)%>
  arrField(1) = "ud_minor_nm"        <%' Field��(1)%>

  arrHeader(0) = "TESTER��"       <%' Header��(0)%>
  arrHeader(1) = "TESTER���"      <%' Header��(1)%>

  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
	 With frm1
	  .vspdData.Col = C_Gi_Type
	  .vspdData.Text = arrRet(0) 
	  .vspdData.Col = C_Gi_Unit
	  .vspdData.Text = arrRet(1) 
	 End With

  End If
 End Function


'========================================================================================================= 
Sub Form_Load()
 
	 Call LoadInfTB19029()
	 Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)   '��: Format Contents  Field
	 Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)   '��: Format Contents  Field
	 
	 Call ggoOper.LockField(Document, "N")                                          '��: Lock  Suitable  Field

	 Call InitSpreadSheet 

	 Call SetDefaultVal
	 Call InitVariables     

	Call InitSpreadComboBox
	Call SetToolBar("1100000000001111")          '��: ��ư ���� ���� 
 
End Sub

'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData 
 
		ggoSpread.Source = frm1.vspdData

		If Row > 0 Then
			Select Case Col
			End Select
		 
			Call SetActiveCell(frm1.vspdData,Col-1,Row,"M","X","X")  
		End If

	End With
End Sub

'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	gMouseClickStatus = "SPC"   

	Set gActiveSpdSheet = frm1.vspdData
	' Context �޴��� �Է�, ����, ������ �Է�, ��� 
	Call SetPopupMenuItemInf(Mid(gToolBarBit, 6, 2) + "0" + Mid(gToolBarBit, 8, 1) & "111111")
	    
	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
		Exit Sub
	End If
	   	    
	If Row <= 0 Then
	    ggoSpread.Source = frm1.vspdData
	    If lgSortKey = 1 Then
	        ggoSpread.SSSort Col				'Sort in Ascending
	        lgSortKey = 2
	    Else
	        ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
	        lgSortKey = 1
	    End If
		 Exit Sub     
	End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	frm1.vspdData.Row = Row
	'---frm1.vspdData.Col = C_MajorCd
		
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
'==========================================================================================
'   Event Name :vspddata_ComboSelChange                                                          
'   Event Desc :Combo Change Event                                                                           
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
		
		.Row = Row
        Select Case Col
		End Select
	End With

   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			.Row = intRow


    	Next	
	End With
End Sub

'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub

'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
   Dim iDx
   Dim IntRetCd
   Dim iDblPrice, iDblAmt, iXch_rate

   ggoSpread.Source = frm1.vspdData
   Frm1.vspdData.Row = Row
   Frm1.vspdData.Col = Col
   
   iDblPrice = 0
   iDblAmt = 0

   With frm1.vspdData

        Select Case Col
			Case C_Price1, C_Price2, C_Price3, C_Price4, C_Price5, C_Price6, C_Price7, C_Price8, C_Price9, C_Price10, C_Price11, C_Price12,C_Price13,C_Price14,C_Price15
                .Col = C_Price1:  iDblPrice = iDblPrice + UNICDbl(.Text)
                .Col = C_Price2:  iDblPrice = iDblPrice + UNICDbl(.Text)
                .Col = C_Price3:  iDblPrice = iDblPrice + UNICDbl(.Text)
                .Col = C_Price4:  iDblPrice = iDblPrice + UNICDbl(.Text)
                .Col = C_Price5:  iDblPrice = iDblPrice + UNICDbl(.Text)
                .Col = C_Price6:  iDblPrice = iDblPrice + UNICDbl(.Text)
                .Col = C_Price7:  iDblPrice = iDblPrice + UNICDbl(.Text)
                .Col = C_Price8:  iDblPrice = iDblPrice + UNICDbl(.Text)
                .Col = C_Price9:  iDblPrice = iDblPrice + UNICDbl(.Text)
                .Col = C_Price10: iDblPrice = iDblPrice + UNICDbl(.Text)
                .Col = C_Price11: iDblPrice = iDblPrice + UNICDbl(.Text)
                .Col = C_Price12: iDblPrice = iDblPrice + UNICDbl(.Text)
                .Col = C_Price13: iDblPrice = iDblPrice + UNICDbl(.Text)
                .Col = C_Price14: iDblPrice = iDblPrice + UNICDbl(.Text)
                .Col = C_Price15: iDblPrice = iDblPrice + UNICDbl(.Text)

                ' ������ �� ���� 
                .Col = C_Gi_Price: .Text = iDblPrice
                
                .Col = C_Amt1:  iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt2:  iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt3:  iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt4:  iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt5:  iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt6:  iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt7:  iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt8:  iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt9:  iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt10: iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt11: iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt12: iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt13: iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt14: iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt15: iDblAmt = iDblAmt + UNICDbl(.Text)

                .Col = C_Gi_Amt: .Text = iDblAmt
                
                .Col = C_xch_rate: iXch_rate = UNICDbl(.Text)
                
                .Col = C_Gi_Amt_Loc: .Text = round(iDblAmt * iXch_rate,0)
                
			Case C_Amt1, C_Amt2, C_Amt3, C_Amt4, C_Amt5, C_Amt6, C_Amt7, C_Amt8, C_Amt9, C_Amt10, C_Amt11, C_Amt12,C_Amt13,C_Amt14,C_Amt15
                .Col = C_Amt1:  iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt2:  iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt3:  iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt4:  iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt5:  iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt6:  iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt7:  iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt8:  iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt9:  iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt10: iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt11: iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt12: iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt13: iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt14: iDblAmt = iDblAmt + UNICDbl(.Text)
                .Col = C_Amt15: iDblAmt = iDblAmt + UNICDbl(.Text)

                .Col = C_Gi_Amt: .Text = iDblAmt

                .Col = C_xch_rate: iXch_rate = UNICDbl(.Text)
                
                .Col = C_Gi_Amt_Loc: .Text = round(iDblAmt * iXch_rate,0)
	
	    End Select    
   End With 
   
   ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
Sub vspdData_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
    End Select
End Sub

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then Exit Sub
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then 
		If lgStrPrevKey <> "" Then  
			If CheckRunningBizProcess Then Exit Sub
			Call DisableToolBar(parent.TBC_QUERY)
			Call DbQuery()
		End if
	End if     
    
End Sub

<%
'==========================================================================================
'   Event Name : OCX_DbClick()
'   Event Desc : OCX_DbClick() �� Calendar Popup
'==========================================================================================
%>
Sub txtconFr_dt_DblClick(Button)
 If Button = 1 Then
  frm1.txtconFr_dt.Action = 7
  Call SetFocusToDocument("M")
  Frm1.txtconFr_dt.Focus
 End If
End Sub

Sub txtconTo_dt_DblClick(Button)
 If Button = 1 Then
  frm1.txtconTo_dt.Action = 7
  Call SetFocusToDocument("M")
  Frm1.txtconTo_dt.Focus
 End If
End Sub

<%
'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : ��ȸ���Ǻ��� OCX_KeyDown�� EnterKey�� ���� Query
'==========================================================================================
%>
Sub txtconFr_dt_KeyDown(KeyCode, Shift)

 If KeyCode = 13 Then Call MainQuery()

End Sub

Sub txtconTo_dt_KeyDown(KeyCode, Shift)

 If KeyCode = 13 Then Call MainQuery()

End Sub

Function txtconBp_cd_OnChange()
    txtconBp_cd_OnChange = true
    
    If  frm1.txtconBp_cd.value = "" Then
        frm1.txtconBp_nm.value = ""
        frm1.txtconBp_cd.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" BP_NM "," B_BIZ_PARTNER "," BP_CD =  " & FilterVar(frm1.txtconBp_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            Call  DisplayMsgBox("970000", "x","�ŷ�ó","x")

            frm1.txtconBp_nm.value = ""
	        frm1.txtconBp_cd.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtconBp_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If

End Function


Function txtPlantCode_OnChange()
    If  frm1.txtPlantCode.value <> "" Then
        if   CommonQueryRs(" plant_nm "," B_PLANT "," plant_cd =  " & FilterVar(frm1.txtPlantCode.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtPlantName.value = ""
            Call  DisplayMsgBox("970000", "x","�����ڵ�","x")
	        frm1.txtPlantCode.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtPlantName.value = Replace(lgF0, Chr(11), "")
	    End If
	else 
		 frm1.txtPlantName.value=""
    End If

End Function

Function txtconItem_cd_OnChange()
   txtconItem_cd_OnChange = true
    
    If  frm1.txtconItem_cd.value = "" Then
        frm1.txtconItem_nm.value = ""
        frm1.txtconItem_cd.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" ITEM_NM "," B_ITEM "," ITEM_CD = " & FilterVar(frm1.txtconItem_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            Call  DisplayMsgBox("970000", "x","ǰ���ڵ�","x")

            frm1.txtconItem_nm.value = ""
	        frm1.txtconItem_cd.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtconItem_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If

End Function

Function txtconOutType_OnChange()
    txtconOutType_OnChange = true
    
    If  frm1.txtconOutType.value = "" Then
        frm1.txtconBp_nm.value = ""
        frm1.txtconOutType.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" UD_MINOR_NM "," B_USER_DEFINED_MINOR "," UD_MAJOR_CD = " & FilterVar("ZZ002", "''", "S") & " AND UD_MINOR_CD =  " & FilterVar(frm1.txtconOutType.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            Call  DisplayMsgBox("970000", "x","���Type","x")

            frm1.txtconOutTypeNm.value = ""
	        frm1.txtconOutType.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtconOutTypeNm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If

End Function
'========================================================================================================
 Function FncQuery()
  Dim IntRetCD

  FncQuery = False             <% '��: Processing is NG %>

  Err.Clear               <% '��: Protect system from crashing %>

  <% '------ Check previous data area ------ %>
  ggoSpread.Source = frm1.vspdData  
  If ggoSpread.SSCheckChange = True and lgBlnFlgChgValue=true Then   
   IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")   <% '��: "Will you destory previous data" %>
   If IntRetCD = vbNo Then
    Exit Function
   End If
  End If

  <% '------ Erase contents area ------ %>
  Call ggoOper.ClearField(Document, "2")        <% '��: Clear Contents  Field %>
  Call InitVariables             <% '��: Initializes local global variables %>

  <% '------ Check condition area ------ %>
  If Not chkField(Document, "1") Then       <% '��: This function check indispensable field %>
   Exit Function
  End If

  <% '------ Query function call area ------ %>
  Call DbQuery()              <% '��: Query db data %>

  FncQuery = True              <% '��: Processing is OK %>
 End Function
 
'========================================================================================================
 Function FncNew()
  Dim IntRetCD 

  FncNew = False              <% '��: Protect system from crashing %>

  <% '------ Check previous data area ------ %>
  ggoSpread.Source = frm1.vspdData
  If ggoSpread.SSCheckChange = True Then
   IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")
   If IntRetCD = vbNo Then
    Exit Function
   End If
  End If

  <% '------ Erase condition area ----- %>
  <% '------ Erase contents area ------ %>
  Call ggoOper.ClearField(Document, "A")        <%'��: Clear Condition Field%>
  Call ggoOper.LockField(Document, "N")        <%'��: Lock  Suitable  Field%>
  Call SetDefaultVal
  Call SetToolBar("1100100100001111")          '��: ��ư ���� ���� 
  Call InitVariables             <%'��: Initializes local global variables%>

  FncNew = True              <%'��: Processing is OK%>

 End Function
 
 
'========================================================================================================
 Function FncSave()
  Dim IntRetCD
  
  FncSave = False                  <% '��: Processing is NG %>
  
  Err.Clear                   <% '��: Protect system from crashing %>
  
  <% '------ Precheck area ------ %>
  ggoSpread.Source = frm1.vspdData
  If ggoSpread.SSCheckChange = False Then        <% 'Check if there is retrived data %>   
      IntRetCD = DisplayMsgBox("900001","x","x","x")     <% '��: No data changed!! %>
      Exit Function
  End If
  
  <% '------ Check contents area ------ %>
  ggoSpread.Source = frm1.vspdData

  If Not chkField(Document, "2") Then  <% '��: Check contents area %>
   Exit Function
  End If

  If Not ggoSpread.SSDefaultCheck Then  <% '��: Check contents area %>
   Exit Function
  End If
  
  <% '------ Save function call area ------ %>
  Call DbSave                   <% '��: Save db data %>
  
  FncSave = True                  <% '��: Processing is OK %>
 End Function

'========================================================================================================
 Function FncCopy()
	frm1.vspdData.ReDraw = False
	  
	if frm1.vspdData.maxrows < 1 then exit function
	  
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.CopyRow
	SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

	frm1.vspdData.ReDraw = True
 End Function

'========================================================================================================
 Function FncCancel() 
   
  if frm1.vspdData.maxrows < 1 then exit function
  
  frm1.vspdData.ReDraw = False    
  ggoSpread.Source = frm1.vspdData
  ggoSpread.EditUndo              <%'��: Protect system from crashing%>
  frm1.vspdData.ReDraw = True
 End Function

'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
    FncInsertRow = False                                                         '��: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else
        imRow = AskSpdSheetAddRowCount()
        If imRow = "" Then
            Exit Function
        End If
    End If

	With frm1
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow ,imRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		lgBlnFlgChgValue = True
		.vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
		FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
 End Function

'========================================================================================================
 Function FncDeleteRow()
  Dim lDelRows
  Dim iDelRowCnt, i
  
  if frm1.vspdData.maxrows < 1 then exit function
 
  With frm1.vspdData 
   If .MaxRows = 0 Then
    Exit Function
   End If

   .focus
   ggoSpread.Source = frm1.vspdData

   lDelRows = ggoSpread.DeleteRow

   lgBlnFlgChgValue = True
  End With
 End Function

'========================================================================================================
 Function FncPrint()
     ggoSpread.Source = frm1.vspdData
  Call parent.FncPrint()             <%'��: Protect system from crashing%>
 End Function

'========================================================================================================
 Function FncExcel() 
  Call parent.FncExport(Parent.C_MULTI)
 End Function

'========================================================================================================
 Function FncFind() 
  Call parent.FncFind(Parent.C_MULTI, False)
 End Function

'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitSpreadComboBox()
	Call ggoSpread.ReOrderingSpreadData()
	Call SetSpreadColor(-1)
	Call InitData()

End Sub

'========================================================================================================
 Function FncExit()
  Dim IntRetCD

  FncExit = False

  ggoSpread.Source = frm1.vspdData
  If ggoSpread.SSCheckChange = True Then
   IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")   <%'��: "Will you destory previous data"%>
   If IntRetCD = vbNo Then
    Exit Function
   End If
  End If

  FncExit = True
 End Function


'========================================================================================================
 Function DbQuery()
  Err.Clear               <%'��: Protect system from crashing%>

  DbQuery = False              <%'��: Processing is NG%>

  Dim strVal
  
  If   LayerShowHide(1) = False Then
             Exit Function 
  End If  

  If lgIntFlgMode = Parent.OPMD_UMODE Then
   strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001                      <%'��: �����Ͻ� ó�� ASP�� ���� %>
   strVal = strVal & "&txtconBp_cd=" & Trim(frm1.txtHconBp_cd.value)        <%'��: ��ȸ ���� ����Ÿ %>
   strVal = strVal & "&txtPlantCode=" & Trim(frm1.txtHPlantCode.value)
   strVal = strVal & "&txtconItem_cd=" & Trim(frm1.txtHconItem_cd.value)        <%'��: ��ȸ ���� ����Ÿ %>
   strVal = strVal & "&txtconOutType=" & Trim(frm1.txtHconOutType.value)  
   strVal = strVal & "&txtconFr_dt=" & Trim(frm1.txtHconFr_dt.value)   
   strVal = strVal & "&txtconTo_dt=" & Trim(frm1.txtHconTo_dt.value)   
   strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
   strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows  
     
  Else
   strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001                      <%'��: �����Ͻ� ó�� ASP�� ���� %>
   strVal = strVal & "&txtconBp_cd=" & Trim(frm1.txtconBp_cd.value)        <%'��: ��ȸ ���� ����Ÿ %>
   strVal = strVal & "&txtPlantCode=" & Trim(frm1.txtPlantCode.value)
   strVal = strVal & "&txtconItem_cd=" & Trim(frm1.txtconItem_cd.value)   
   strVal = strVal & "&txtconOutType=" & Trim(frm1.txtconOutType.value)  
   strVal = strVal & "&txtconFr_dt=" & Trim(frm1.txtconFr_dt.text)   
   strVal = strVal & "&txtconTo_dt=" & Trim(frm1.txtconTo_dt.text)  
   strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
   strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
  End If

  Call RunMyBizASP(MyBizASP, strVal)         <%'��: �����Ͻ� ASP �� ���� %>
 
  DbQuery = True              <%'��: Processing is NG%>
  frm1.vspdData.Focus 
 End Function
 
'========================================================================================================
 Function DbSave() 
  Dim lRow
  Dim lGrpCnt
  Dim strVal, strDel 

  DbSave = False              <% '��: Processing is OK %>
    
  On Error Resume Next            <% '��: Protect system from crashing %>
  
  If   LayerShowHide(1) = False Then
             Exit Function 
        End If

  With frm1
   .txtMode.value = Parent.UID_M0002
   .txtUpdtUserId.value = Parent.gUsrID
   .txtInsrtUserId.value = Parent.gUsrID

   lGrpCnt = 1

   strVal = ""
   strDel = ""

 
   For lRow = 1 To .vspdData.MaxRows
    .vspdData.Row = lRow
    .vspdData.Col = 0            
 
    
    Select Case .vspdData.Text
     Case ggoSpread.InsertFlag       
										  strVal = strVal & "C" & parent.gColSep 
										  strVal = strVal & lRow & parent.gColSep
		.vspdData.Col = C_Pl_No		    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Trans_Time    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_OutType_Sub   : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep

		lGrpCnt = lGrpCnt + 1
  
     Case ggoSpread.UpdateFlag      
										  strVal = strVal & "U" & parent.gColSep 
										  strVal = strVal & lRow & parent.gColSep
		.vspdData.Col = C_Pl_No		    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Trans_Time    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_OutType_Sub   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Price1		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Price2		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Price3		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Price4		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Price5		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Price6		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Price7		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Price8		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Price9		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Price10		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Price11		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Price12		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Price13		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Price14		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Price15		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Amt1			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Amt2			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Amt3			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Amt4			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Amt5			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Amt6			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Amt7			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Amt8			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Amt9			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Amt10			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Amt11			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Amt12			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Amt13			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Amt14			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Amt15			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

		.vspdData.Col = C_Gi_Price		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Gi_Amt		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Gi_Amt_Loc	: strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep


    
		lGrpCnt = lGrpCnt + 1
 
     Case ggoSpread.DeleteFlag  
										  strDel = strDel & "D" & parent.gColSep 
										  strDel = strDel & lRow & parent.gColSep
		.vspdData.Col = C_Pl_No		    : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_Trans_Time    : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		.vspdData.Col = C_OutType_Sub   : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep

		lGrpCnt = lGrpCnt + 1
    End Select    
   Next

   .txtMaxRows.value = lGrpCnt-1
   .txtSpread.value =  strDel & strVal


   Call ExecMyBizASP(frm1, BIZ_PGM_ID)      <% '��: �����Ͻ� ASP �� ���� %>

  End With

  DbSave = True              <% '��: Processing is NG %>
 End Function
 
'========================================================================================================
 Function DbQueryOk()             <% '��: ��ȸ ������ ������� %>

  lgIntFlgMode = Parent.OPMD_UMODE           <% '��: Indicates that current mode is Update mode %>
  lgBlnFlgChgValue = False
  Call InitData()
  
  Call ggoOper.LockField(Document, "Q")        <% '��: This function lock the suitable field %>
  Call SetToolBar("1100100100011111") 
  frm1.vspdData.Focus     

 End Function
 
'========================================================================================================
 Function DbSaveOk()              <%'��: ���� ������ ���� ���� %>
  Call ggoOper.ClearField(Document, "2")
  Call InitVariables     
  Call MainQuery()
  
 End Function
 
</SCRIPT>
<!-- #Include file="../../inc/UNI2kCM.inc" --> 
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
     <TD CLASS="CLSLTAB">
      <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
       <TR>
        <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���ݾ׻���������ȸ������</font></td>
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
         <TD CLASS="TD5" NOWRAP>����</TD>
         <TD CLASS="TD6"><INPUT NAME="txtconBp_cd" ALT="����" TYPE="Text" MAXLENGTH=10 SiZE=12  tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconBp_cd.value,5">&nbsp;<INPUT NAME="txtconBp_nm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD>
         <TD CLASS="TD5" NOWRAP>�������</TD>
         <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtconFr_dt" CLASS=FPDTYYYYMMDD tag="11X1X" ALT="��������" Title="FPDATETIME"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
                                <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtconTo_dt" CLASS=FPDTYYYYMMDD tag="11X1X" ALT="���������" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>       
        </TR>
        <TR>
		 <TD CLASS="TD5" NOWRAP>����</TD>
		 <TD CLASS="TD6"><INPUT NAME="txtPlantCode" TYPE="Text" ALT="����" MAXLENGTH=10 SiZE=12 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSDN" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtPlantCode.value,4">&nbsp;<INPUT NAME="txtPlantName" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
         <TD CLASS="TD5" NOWRAP>ǰ��</TD>
         <TD CLASS="TD6"><INPUT NAME="txtconItem_cd" ALT="ǰ��" TYPE="Text" MAXLENGTH=18 SiZE=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconItem_cd.value,0">&nbsp;<INPUT NAME="txtconItem_nm" TYPE="Text" MAXLENGTH="40" SIZE=25 tag="14"></TD>
        </TR>
        <TR>
         <TD CLASS="TD5" NOWRAP>���Type</TD>
         <TD CLASS="TD6"><INPUT NAME="txtconOutType" ALT="���Type" TYPE="Text" MAXLENGTH=18 SiZE=12  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconOutType.value,1">&nbsp;<INPUT NAME="txtconOutTypeNm" TYPE="Text" MAXLENGTH="40" SIZE=25 tag="14"></TD>
         <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
         <TD CLASS="TD6" NOWRAP>&nbsp;</TD>
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
    <TR>
     <TD WIDTH=100% HEIGHT=50 valign=top>
       <TABLE <%=LR_SPACE_TYPE_20%>>
        <TR>
         <TD HEIGHT="100%">
          <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%>  FRAMEBORDER=0 SCROLLING=no noresize framespacing=0  TABINDEX = -1></IFRAME>
  </TD>
 </TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHconBp_cd" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHconItem_cd" tag="24" TABINDEX = -1> 
<INPUT TYPE=HIDDEN NAME="txtHconOutType" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHconFr_dt" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHconTo_dt" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHPlantCode" tag="24" TABINDEX = -1>
</FORM>

<DIV ID="MousePT" NAME="MousePT">
 <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>