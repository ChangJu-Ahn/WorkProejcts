
<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================

'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ��Ư��2ȣ��2����Ǿ��������������� 
'*  3. Program ID           : W6109MA1
'*  4. Program Name         : W6109MA1.asp
'*  5. Program Desc         : ��Ư��2ȣ��2����Ǿ��������������� 
'*  6. Modified date(First) : 2005/03/18
'*  7. Modified date(Last)  : 2005/03/18
'*  8. Modifier (First)     : ȫ���� 
'*  9. Modifier (Last)      : ȫ���� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%
'=========================  �α����� ������ �����ڵ带 ����ϱ� ����  ======================
    Call LoadBasisGlobalInf()
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  ���/���� ����  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->
Const BIZ_MNU_ID = "W6109MA1"											 '
Const BIZ_PGM_ID = "W6109MB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID = "W6109OA1"

Const C_SHEETMAXROWS = 100


Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 

dim lgblnEvents
Dim strW5_R
'============================================  �ʱ�ȭ �Լ�  ====================================`
Sub InitSpreadPosVariables()

    

    
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""                          'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
    lgRefMode = False
    lgblnEvents = False
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
	Dim strYear
	Dim strMonth
	Dim strInsurDt
	Dim stReturnrInsurDt

   lgKeyStream = UCASE(Frm1.txtCo_Cd.Value) &  parent.gColSep  
   lgKeyStream = lgKeyStream & (Frm1.txtFISC_YEAR.text) &  parent.gColSep ' 
   lgKeyStream = lgKeyStream & UCASE(Frm1.cboREP_TYPE.Value ) &  parent.gColSep '   
   lgKeyStream = lgKeyStream & strW3_R &  parent.gColSep ' 
    lgKeyStream = lgKeyStream &  (Frm1.txtW3_Rate.Value)   &  parent.gColSep ' 
      lgKeyStream = lgKeyStream & strW5_R &  parent.gColSep '  
   lgKeyStream = lgKeyStream &  (Frm1.txtW5_Rate.Value)   &  parent.gColSep '   


End Sub 

'============================================  �Ű��� �޺� �ڽ� ä���  ====================================

Sub InitComboBox()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub InitSpreadSheet()

  


			
		
    
End Sub

'============================== ���۷��� �Լ�  ========================================
Function GetRef()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	Dim arrW1 ,arrW2 ,  arrW3, arrW4, arrW5, arrW6, iRow, iCol
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' �������� ���� : �޽�����������.
	
	
	if wgConfirmFlg = "Y" then    'Ȯ���� 
	   Exit function
	end if   
	
	' �·ε�� ���۷����޽��� �����´�.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
    '����� ���α׷� 
   
	IntRetCD =  CommonQueryRs("W4","dbo.ufn_TB_JT2_2_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD = false Then	 
	   IntRetCD =   DisplayMsgBox("W60006", "x", "(120) ���⼼��"  , "X")  
	   Exit Function
	else   
	    frm1.txtw13.value = unicdbl(lgF0)
    end if
	 
	Call CalSum
     

end function
'============================================  �׸��� �Լ�  ====================================





'============================================  ��ȸ���� �Լ�  ====================================

'============================================  �� �Լ�  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                             <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet()                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100100000001111")										<%'��ư ���� ���� %>

	' �����Ѱ� 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
    Call SetDefaultVal()

	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	Call FncQuery
End Sub




Sub SetDefaultVal()
	dim arrTmp(1) 
	DIM strW1

	With frm1
	           
	call CommonQueryRs("REFERENCE_1, REFERENCE_2"," ufn_TB_Configuration('W4002', '" & C_REVISION_YM & "')", "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	    	
	arrTmp(0) 	=  Split(lgF0,chr(11)) 
	arrTmp(1) 	=  Split(lgF1,chr(11)) 
	
	.txtW12_GA_B_VAL.value =	arrTmp(0)(0)
	.txtW12_NA_B_VAL.value =	arrTmp(0)(1)

	.txtW12_GA_B_VIEW.value =	arrTmp(1)(0)
	.txtW12_NA_B_VIEW.value =	arrTmp(1)(1)
	
	call CommonQueryRs("REFERENCE_1, REFERENCE_2"," ufn_TB_Configuration('W4031', '" & C_REVISION_YM & "')","" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	    	
	frm1.txtW14_VAL.value 	=  Replace(lgF0,chr(11),"") 
	frm1.txtW14_VIEW.value 	=  Replace(lgF1,chr(11),"") 
	
	End With
End Sub


'============================================  �̺�Ʈ �Լ�  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub


Function CalSum()

	If lgblnEvents Then Exit Function	' -- �Ʒ� �� ���濡 ���� �̺�Ʈ ȣ���� ������ 
	lgblnEvents = True
	
	With frm1
		.txtW8_A_SUM.value = UNICDbl(.txtW8_A_1.value) + UNICDbl(.txtW8_A_2.value)
		.txtW8_B_SUM.value = UNICDbl(.txtW8_B_1.value) + UNICDbl(.txtW8_B_2.value)
		.txtW8_C_SUM.value = UNICDbl(.txtW8_C_1.value) + UNICDbl(.txtW8_C_2.value)
		.txtW8_D_SUM.value = UNICDbl(.txtW8_D_1.value) + UNICDbl(.txtW8_D_2.value)
		.txtW8_E_SUM.value = UNICDbl(.txtW8_E_1.value) + UNICDbl(.txtW8_E_2.value)
		.txtW8_F_SUM.value = UNICDbl(.txtW8_F_1.value) + UNICDbl(.txtW8_F_2.value)
		.txtW8_HAP_SUM.value = UNICDbl(.txtW8_A_SUM.value) + UNICDbl(.txtW8_B_SUM.value) + UNICDbl(.txtW8_C_SUM.value) + UNICDbl(.txtW8_D_SUM.value) + UNICDbl(.txtW8_E_SUM.value) + UNICDbl(.txtW8_F_SUM.value)
		.txtW8_HAP_1.value = UNICDbl(.txtW8_A_1.value) + UNICDbl(.txtW8_B_1.value) + UNICDbl(.txtW8_C_1.value) + UNICDbl(.txtW8_D_1.value) + UNICDbl(.txtW8_E_1.value) + UNICDbl(.txtW8_F_1.value)
		.txtW8_HAP_2.value = UNICDbl(.txtW8_A_2.value) + UNICDbl(.txtW8_B_2.value) + UNICDbl(.txtW8_C_2.value) + UNICDbl(.txtW8_D_2.value) + UNICDbl(.txtW8_E_2.value) + UNICDbl(.txtW8_F_2.value)
		
		.txtW12_GA_A.value = UNICDbl(.txtW8_HAP_1.value) - UNICDbl(.txtW11.value)
		.txtW12_GA_C.value = UNICDbl(.txtW12_GA_A.value) * UNICDbl(.txtW12_GA_B_VAL.value)
		
		.txtW12_NA_A.value = UNICDbl(.txtW8_HAP_2.value) - UNICDbl(.txtW11.value)
		.txtW12_NA_C.value = UNICDbl(.txtW12_NA_A.value) * UNICDbl(.txtW12_NA_B_VAL.value)
		
		.txtW12_HAP_C.value = UNICDbl(.txtW12_GA_C.value) + UNICDbl(.txtW12_NA_C.value)
		
		.txtW14.value = UNICDbl(.txtW13.value) * UNICDbl(.txtW14_VAL.value)
		
		If UNICDbl(.txtW12_HAP_C.value) < UNICDbl(.txtW14.value) Then
			.txtW15.value = .txtW12_HAP_C.value
		Else
			.txtW15.value = .txtW14.value
		End If
		
	End With
	
	lgBlnFlgChgValue = True
	lgblnEvents = False
end function

Function CheckMessage(ByVal Obj)
dim IntRetCD
    if  UNICDbl(Obj.value) < 0 then
        IntRetCD =  DisplayMsgBox("WC0006","x",Obj.alt,"x")  			
        Obj.value = 0
        Obj.focus
        Set gActiveElement = document.ActiveElement
        exit function	
    end if
    
end function

Sub txtW8_A_SUM_Change()  
    Call CalSum()
End Sub

Sub txtW8_A_1_Change()  
    Call CalSum()
End Sub

Sub txtW8_A_2_Change()  
    Call CalSum()
End Sub


Sub txtW8_B_SUM_Change()  
    Call CalSum()
End Sub

Sub txtW8_B_1_Change()  
    Call CalSum()
End Sub

Sub txtW8_B_2_Change()  
    Call CalSum()
End Sub

Sub txtW8_C_SUM_Change()  
    Call CalSum()
End Sub

Sub txtW8_C_1_Change()  
    Call CalSum()
End Sub

Sub txtW8_C_2_Change()  
    Call CalSum()
End Sub

Sub txtW8_D_SUM_Change()  
    Call CalSum()
End Sub

Sub txtW8_D_1_Change()  
    Call CalSum()
End Sub

Sub txtW8_D_2_Change()  
    Call CalSum()
End Sub

Sub txtW8_E_SUM_Change()  
    Call CalSum()
End Sub

Sub txtW8_E_1_Change()  
    Call CalSum()
End Sub

Sub txtW8_E_2_Change()  
    Call CalSum()
End Sub

Sub txtW8_F_SUM_Change()  
    Call CalSum()
End Sub

Sub txtW8_F_1_Change()  
    Call CalSum()
End Sub

Sub txtW8_F_2_Change()  
    Call CalSum()
End Sub

Sub txtW11_Change()  
    Call CalSum()
End Sub

'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	

End Sub


Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub




'============================================  �������� �Լ�  ====================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    

	If lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") 
		If IntRetCD = vbNo Then	Exit Function
    End If

     Call ggoOper.ClearField(Document, "2")
     'Call ggoOper.LockField(Document, "N")
     Call InitVariables               
	 Call SetDefaultVal()
	 
     Call SetToolbar("1100100000001111")          '��: ��ư ���� ���� 
     
     frm1.txtW8_A_1.focus
    FncNew = True                

End Function

Function FncQuery() 
    Dim IntRetCD 

    
    FncQuery = False															 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '��: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    'Call ggoOper.LockField(Document, "Q")
    Call  ggoOper.ClearField(Document, "2")										 '��: Clear Contents  Field
    
    If Not chkField(Document, "1") Then									         '��: This function check required field
       Exit Function
    End If

    Call InitVariables                                                           '��: Initializes local global variables
    'Call MakeKeyStream("Q")
    
	Call  DisableToolBar( parent.TBC_QUERY)
    If DbQuery = False Then
		Call  RestoreToolBar()
        Exit Function
    End If
              
    FncQuery = True  
    
End Function

Function FncSave() 
        
    FncSave = False                                                         
    dim IntRetCD
    
    
    

    Err.Clear                                                               <%'��: Protect system from crashing%>
    'On Error Resume Next                                                   <%'��: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    If lgBlnFlgChgValue = False Then
		IntRetCD =  DisplayMsgBox("900001","x","x","x")  					 '��: Data is changed.  Do you want to display it? 

			Exit Function

    End If
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    'Call ggoOper.LockField(Document, "N")
    'Call MakeKeyStream("Q")
    If DbSave = False Then Exit Function                                        '��: Save db data
  
    FncSave = True                                                          
    
End Function

Function FncCopy() 
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCopy = False                                                               '��: Processing is NG


	
    Set gActiveElement = document.ActiveElement   
	
End Function
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim IntRetCd

    FncDelete = False                                                             '��: Processing is NG
    
    
    <%  '-----------------------
    'Check previous data area
    '----------------------- %>

    If lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
    
    
    If lgIntFlgMode <>  parent.OPMD_UMODE  Then                                            'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")                                  '��: Please do Display first.
        Exit Function
    End If

    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		                  '��: Do you want to delete?
	If IntRetCD = vbNo Then
		Exit Function
	End If
    'Call MakeKeyStream("Q")
    If DbDelete= False Then
       Exit Function
    End If												                  '��: Delete db data

    FncDelete=  True                                                              '��: Processing is OK
End Function


Function FncCancel() 
                                           '��: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim uCountID

    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
	
    Set gActiveElement = document.ActiveElement   
    
End Function

Function FncDeleteRow() 
    Dim lDelRows

    
End Function

Function FncPrint()
    Call parent.FncPrint()                                                   '��: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)											 <%'��: ȭ�� ���� %>
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                      <%'��:ȭ�� ����, Tab ���� %>
End Function

Function FncExit()
Dim IntRetCD

	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")

		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'============================================  DB �＼�� �Լ�  ====================================
Function DbQuery() 

    Dim strVal
    Err.Clear                                                                    '��: Clear err status

    DbQuery = False                                                              '��: Processing is NG

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode="          &  parent.UID_M0001                       '��: Query
    'strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '��: Query Key
    'strVal = strVal     & "&txtPrevNext="      & ""	                             '��: Direction
    Call RunMyBizASP(MyBizASP, strVal)                                           '��:  Run biz logic

    DbQuery = True                                       
  
End Function

Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	 'Call ggoOper.LockField(Document, "N")
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE
    Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	'1 ����üũ 
	If wgConfirmFlg = "N" Then
	
		'2 ���ȯ�氪 , �ε��ȯ�氪 �� 
		 Call SetToolbar("1101100000011111")										<%'��ư ���� ���� %>
	Else
		 Call SetToolbar("1100100000011111")										<%'��ư ���� ���� %>
	End If
	
    lgBlnFlgChgValue = false
   
  
		
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 


 
    DbSave = False														         '��: Processing is NG
		
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

	With Frm1
		.txtMode.value        =  parent.UID_M0002                                        '��: Delete
		.txtFlgMode.value     = lgIntFlgMode
        '.txtKeyStream.Value   = lgKeyStream                                      '��: Save Key
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                            
End Function


Function DbSaveOk()													        <%' ���� ������ ���� ���� %>
	Call InitVariables
	

    Call MainQuery()
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '��: Clear err status

	DbDelete = False			                                                 '��: Processing is NG

    If LayerShowHide(1) = false Then
        Exit Function
    End If

	With Frm1
		.txtMode.value        =  parent.UID_M0003                                        '��: Delete
		.txtFlgMode.value     = lgIntFlgMode
        '.txtKeyStream.Value   = lgKeyStream                                      '��: Save Key
	End With

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	DbDelete = True                                                              '��: Processing is NG

End Function
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call MainNew()
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
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
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;
						<a href="vbscript:GetRef">�ݾ� �ҷ�����</A>  
					</TD>
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
									<TD CLASS="TD5">�������</TD>
									<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFISC_YEAR CLASS=FPDTYYYY title=FPDATETIME ALT="�������" tag="14X1" id=txtFISC_YEAR></OBJECT>');</SCRIPT>
									<TD CLASS="TD5">���θ�</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
									</TD>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">�Ű���</TD>
									<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE" ALT="�Ű���" STYLE="WIDTH: 50%" tag="14X1"></SELECT>
									</TD>
									<TD CLASS="TD5"></TD>
									<TD CLASS="TD6"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%> </TD>
				</TR>
				
				
				
				
					<TR>
					<TD valign=top >
					   
					    
									<TABLE width = 90% bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
										<TR>
											<TD CLASS="TD51" align =center width = 10% ROWSPAN=9>(8)<br>���<br>�ݾ�</TD>
											<TD CLASS="TD51" align =center width = 45% ROWSPAN=2 COLSPAN=2>�� ��</TD>
											<TD CLASS="TD51" align =center width = 15% ROWSPAN=2>�� ��</TD>
											<TD CLASS="TD51" align =center width = 30% COLSPAN=2>���ޱⰣ ��</TD>
										</TR>
										<TR>
											<TD CLASS="TD51" align =center width = 15%>30�� �̳�</TD>
											<TD CLASS="TD51" align =center width = 15%>31�� ~ 60��</TD>
										</TR>
										<TR>
											<TD CLASS="TD51" align =left width = 30% COLSPAN=2>ȯ���� �����ݾ�</TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_A_SUM" name=txtW8_A_SUM CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" width = 100% ></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_A_1" name=txtW8_A_1 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_A_2" name=txtW8_A_2 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" align =left width = 30% COLSPAN=2>�ǸŴ���߽��Ƿڼ� �����ݾ�</TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_B_SUM" name=txtW8_B_SUM CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" width = 100% ></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_B_1" name=txtW8_B_1 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_B_2" name=txtW8_B_2 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" align =left width = 30% COLSPAN=2>�����������ī�� ���ݾ�</TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_C_SUM" name=txtW8_C_SUM CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" width = 100% ></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_C_1" name=txtW8_C_1 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_C_2" name=txtW8_C_2 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" align =left width = 30% COLSPAN=2>�ܻ����ä�Ǵ㺸�������� �̿�ݾ�</TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_D_SUM" name=txtW8_D_SUM CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" width = 100% ></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_D_1" name=txtW8_D_1 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_D_2" name=txtW8_D_2 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" align =left width = 30% COLSPAN=2>���ŷ����� �̿�ݾ�</TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_E_SUM" name=txtW8_E_SUM CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" width = 100% ></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_E_1" name=txtW8_E_1 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_E_2" name=txtW8_E_2 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" align =left width = 30% COLSPAN=2>��Ʈ��ũ������ �̿�ݾ�</TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_F_SUM" name=txtW8_F_SUM CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" width = 100% ></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_F_1" name=txtW8_F_1 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_F_2" name=txtW8_F_2 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" align =center width = 30% COLSPAN=2>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ��</TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_HAP_SUM" name=txtW8_HAP_SUM CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" width = 100% ></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_HAP_1" name=txtW8_HAP_1 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" width = 100% ></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD51" align =center width = 15%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8_HAP_2" name=txtW8_HAP_2 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" width = 100% ></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" align =center COLSPAN=3>(11)��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��</TD>
											<TD CLASS="TD51" align =center COLSPAN=3><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW11" name=txtW11 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X26" width = 100% ></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" align =center ROWSPAN=4 COLSPAN=2 width=15%>(12)�����ݾ�</TD>
											<TD CLASS="TD51" align =center width=40%>�������ݾ�(a)</TD>
											<TD CLASS="TD51" align =center>������(b)</TD>
											<TD CLASS="TD51" align =center COLSPAN=2>�����ݾ�((a) X (b))</TD>
										</TR>
										<TR>
											<TD CLASS="TD51" align =center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW12_GA_A" name=txtW12_GA_A CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" width = 100% ></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD51" align =center><INPUT TYPE=Text NAME=txtW12_GA_B_VIEW tag="34X" STYLE="width: 100%; text-align: 'center'">
											<INPUT TYPE=HIDDEN NAME=txtW12_GA_B_VAL tag="35X26" ></TD>
											<TD CLASS="TD51" align =center COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW12_GA_C" name=txtW12_GA_C CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" width = 100% ></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" align =center><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW12_NA_A" name=txtW12_NA_A CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" width = 100% ></OBJECT>');</SCRIPT></TD>
											<TD CLASS="TD51" align =center><INPUT TYPE=Text NAME=txtW12_NA_B_VIEW tag="34X" STYLE="width: 100%; text-align: 'center'">
											<INPUT TYPE=HIDDEN NAME=txtW12_NA_B_VAL tag="35X26" ></TD>
											<TD CLASS="TD51" align =center COLSPAN=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW12_NA_C" name=txtW12_NA_C CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" width = 100% ></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" align =center>(12) �� ��</TD>
											<TD CLASS="TD51" align =center COLSPAN=3><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW12_HAP_C" name=txtW12_HAP_C CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" width = 100% ></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" align =center COLSPAN=3>(13)�� �� �� ��</TD>
											<TD CLASS="TD51" align =center COLSPAN=3><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW13" name=txtW13 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" width = 100% ></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" align =center COLSPAN=3>(14)�� �� �� ((13) X (<INPUT TYPE=TEXT NAME="txtW14_VIEW" tag="34X26" size=8 style="text-align: 'center'">)<INPUT TYPE=HIDDEN NAME=txtW14_VAL tag="35X26" ></TD>
											<TD CLASS="TD51" align =center COLSPAN=3><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW14" name=txtW14 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" width = 100% ></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" align =center COLSPAN=3>(15)�� �� �� �� ((12)�� (14)�� ���� �ݾ�)</TD>
											<TD CLASS="TD51" align =center COLSPAN=3><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW15" name=txtW15 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X26" width = 100% ></OBJECT>');</SCRIPT></TD>
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
    <TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE CLASS="TB3" CELLSPACING=0>
			    
		
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FNCBtnPrint('VIEW')" Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:FNCBtnPrint('PRINT')"   Flag=1>�μ�</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txt3w120" tag="24">

<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">

<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

