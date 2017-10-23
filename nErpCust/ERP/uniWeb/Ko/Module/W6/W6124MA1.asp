
<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>

<%'======================================================================================================

'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ��8-1 �������鼼�� ��꼭(1)
'*  3. Program ID           : W6124MA1
'*  4. Program Name         : W6124MA1.asp
'*  5. Program Desc         : ��8-1 �������鼼�� ��꼭(1)
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
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  ���/���� ����  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->
Const BIZ_MNU_ID = "W6124MA1"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID = "W6124MB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID = "W6124OA1"


Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 

dim strW1_R
dim strW5_R

'============================================  �ʱ�ȭ �Լ�  ====================================
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



End Sub 

'============================================  �Ű��� �޺� �ڽ� ä���  ====================================

Sub InitComboBox()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub InitSpreadSheet()
	
    
End Sub



'============================================  �׸��� �Լ�  ====================================





'============================================  ��ȸ���� �Լ�  ====================================

'============================================  �� �Լ�  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                             <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet()                                                    <%'Setup the Spread sheet%>
                                                     <%'Initializes local global variables%>
    
    Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>

	' �����Ѱ� 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
    Call SetDefaultVal()
    Call InitVariables  
     
    ' �������� üũȣ�� 
	Call FncQuery
  
End Sub




Sub SetDefaultVal()
dim strWhere 
DIM strW1
Dim sFiscYear, sRepType, sCoCd, iGap

	
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
  

End Sub



'============================================  ���������  ====================================




Function GetRef()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
    Dim arrW1 ,arrW2 ,  arrW3, arrW4, arrW5, arrW6, iRow, iCol
	Dim sMesg
	
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value


	' �������� ���� : �޽�����������.
	
	
	if wgConfirmFlg = "Y" then    'Ȯ���� 
	   Exit function
	end if   
	
	 wgRefDoc = GetDocRef(sCoCd,sFiscYear, sRepType, BIZ_MNU_ID)
	
	    sMesg = wgRefDoc & vbCrLf & vbCrLf
    call SelectColor(frm1.txtW3_1_A)
    call SelectColor(frm1.txtW3_1_C)  
    
    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"
     Call ggoOper.LockField(Document, "N")
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If

   
    call CommonQueryRs("w3A,w3C","dbo.ufn_TB_8_1_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
    if  replace(lgF1,chr(11),"")  = ""  or  unicdbl(lgF0) = 0  then
    
        IntRetCD = DisplayMsgBox("W60006", "x", "(120) ���⼼��"  , "X")          
        Exit Function
    end if    
     
       frm1.txtW3_1_A.value = unicdbl(lgF0) 
       frm1.txtW3_1_C.value =unicdbl(lgF1) 
      
      
      
      lgBlnFlgChgValue = TRUE
	
   
End Function

'============================================  �̺�Ʈ �Լ�  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub


Function Fn_SumCal()
         if unicdbl(frm1.txtW3_1_C.text)  = 0 then
            frm1.txtW4_1.text = 0
         else
            frm1.txtW4_1.text = frm1.txtW3_1_A.text * (frm1.txtW3_1_B.text/frm1.txtW3_1_C.text)
         end if
         
         if unicdbl(frm1.txtW3_2_C.text)  = 0 then
            frm1.txtW4_2.text = 0
         else
            frm1.txtW4_2.text = frm1.txtW3_2_A.text * (frm1.txtW3_2_B.text/frm1.txtW3_2_C.text)
         end if
         
         frm1.txtW4_SUM.text  = unicdbl( frm1.txtW4_1.text ) + unicdbl( frm1.txtW4_2.text )  + unicdbl( frm1.txtW4_3.text ) 
         
     

end function





Sub txtW3_1_A_CHANGE()
    lgBlnFlgChgValue = TRUE
    Call Fn_SumCal
End Sub


Sub txtW3_1_B_CHANGE()
    lgBlnFlgChgValue = TRUE
    Call Fn_SumCal
End Sub


Sub txtW3_1_C_CHANGE()
    lgBlnFlgChgValue = TRUE
    Call Fn_SumCal
End Sub



Sub txtW3_1_A_CHANGE()
    lgBlnFlgChgValue = TRUE
    Call Fn_SumCal
End Sub


Sub txtW3_2_B_CHANGE()
    lgBlnFlgChgValue = TRUE
    Call Fn_SumCal
End Sub


Sub txtW3_2_C_CHANGE()
    lgBlnFlgChgValue = TRUE
    Call Fn_SumCal
End Sub


Sub txtW1_3_OnCHANGE()
    lgBlnFlgChgValue = TRUE
End Sub


Sub txtW2_3_OnCHANGE()
    lgBlnFlgChgValue = TRUE
End Sub


Sub txtW3_3_OnCHANGE()
    lgBlnFlgChgValue = TRUE
End Sub

Sub txtW4_3_CHANGE()
    lgBlnFlgChgValue = TRUE
      Call Fn_SumCal
End Sub


Sub txtW5_1_OnCHANGE()
    lgBlnFlgChgValue = TRUE
End Sub

Sub txtW5_2_CHANGE()
    lgBlnFlgChgValue = TRUE
End Sub

Sub txtW5_4_CHANGE()
    lgBlnFlgChgValue = TRUE
End Sub

Sub txtw5_3_CHANGE()
    lgBlnFlgChgValue = TRUE
End Sub

Sub txtw5_4_GB_CHANGE()
    lgBlnFlgChgValue = TRUE
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


' ----------------------  ���� -------------------------
Function  Verification()


	Dim sFiscYear, sRepType, sCoCd, IntRetCD ,sMesg
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	
	Verification = False
    '  Call CommonQueryRs("w3A,w3C","dbo.ufn_TB_8_1_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
    'if  replace(lgF1,chr(11),"")  = ""  or  unicdbl(lgF0) = 0  then
    '    sMesg  = "���μ� ���� 3ȣ ���� ���μ�����ǥ�� �� ���� ������꼭���� (120) ���⼼���� ������ �ʾҽ��ϴ�" & vbCr
    '    sMesg  = sMesg & "�ش� �������� �̵��Ͽ� ���� ������ ����Ͽ� �ֽʽÿ�) "
    '    IntRetCD = DisplayMsgBox("x", "x", sMesg  , "X")          
    '    Exit Function
    'end if    

     'if unicdbl(frm1.txtW3_1_A.text) <> unicdbl(replace(lgF0,Chr(11),""))  then
     '   call SelectColor(frm1.txtW3_1_A)
     '   IntRetCD = DisplayMsgBox("WC0004", "x", "8��(1)" & frm1.txtW3_1_C.alt  ,   vbCr & "3ȣ ���� ���μ�����ǥ�� �� ���� ������꼭���� (120) ���⼼��")          
     '    Call ggoOper.LockField(Document, "N")
     '    Exit Function
     'end if
      
         
     ' if  unicdbl(frm1.txtW3_1_C.text) <> unicdbl(lgF1)  then
      '    call SelectColor(frm1.txtW3_1_C)
      '    IntRetCD = DisplayMsgBox("WC0004", "x", "8��(1)" & frm1.txtW3_1_C.alt  , vbCr &"3ȣ ���� ���μ�����ǥ�� �� ���� ������꼭���� (113) ����ǥ��")   
      '    Call ggoOper.LockField(Document, "N")
      '    Exit Function
      'end if
      
      '���ڰ� �и𺸴� Ŭ�� 
      if unicdbl(frm1.txtW3_1_B.value) > unicdbl(frm1.txtW3_1_C.value) then
          IntRetCD = DisplayMsgBox("WC0010", "x", "����ҵ�" , " ����ǥ��")  
          Exit Function  
      end if
      
       '���ڰ� �и𺸴� Ŭ�� 
      if unicdbl(frm1.txtW3_2_B.value) > unicdbl(frm1.txtW3_2_C.value) then
          IntRetCD = DisplayMsgBox("WC0010", "x", "��ǵ� ����� �ڻ갡��" , " ����� �ڻ��Ѿ�")  
          Exit Function    
      end if
      
	
	Verification = True	
End Function

'============================================  �������� �Լ�  ====================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    

	If lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") 
		If IntRetCD = vbNo Then	Exit Function
    End If

     Call ggoOper.ClearField(Document, "2")
     Call ggoOper.LockField(Document, "N")

     Call InitVariables               

     Call SetToolbar("1100100000000111")          '��: ��ư ���� ���� 
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
    Call ggoOper.LockField(Document, "Q")
    Call  ggoOper.ClearField(Document, "2")										 '��: Clear Contents  Field
    
    If Not chkField(Document, "1") Then									         '��: This function check required field
       Exit Function
    End If

    Call InitVariables                                                           '��: Initializes local global variables
    Call MakeKeyStream("Q")
    
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
    
    
    If Not chkField(Document, "2") Then									'��: This function check indispensable field
       Exit Function
    End If 

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
    Call ggoOper.LockField(Document, "N")
    Call Fn_SumCal
    
    
    if Verification = False then exit Function '���� 
    
    Call MakeKeyStream("Q")
   
   
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
    Call MakeKeyStream("Q")
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
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '��: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '��: Direction
    Call RunMyBizASP(MyBizASP, strVal)                                           '��:  Run biz logic

    DbQuery = True                                       
  
End Function

Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    dim IntRetCD

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	 Call ggoOper.LockField(Document, "N")
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE

    lgBlnFlgChgValue = false
 										<%'��ư ���� ���� %>
        Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	'1 ����üũ 
	If wgConfirmFlg = "Y" Then
	    Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>
	    
		
	Else
	   '2 ���ȯ�氪 , �ε��ȯ�氪 �� 
		 Call SetToolbar("1101100000000111")										<%'��ư ���� ���� %>
		
	End If
  
		
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
        .txtKeyStream.Value   = lgKeyStream                                      '��: Save Key
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
        .txtKeyStream.Value   = lgKeyStream                                      '��: Save Key
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
'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to
              Exit For
           End If

       Next

    End If
End Sub





'======================================================================================================
'   Event Name : txtW5_2_DblClick
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtW5_2_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtW5_2.Action = 7
		frm1.txtW5_2.focus
	End If
End Sub

'======================================================================================================
'   Event Name : txtW5_3_DblClick
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtW5_3_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtW5_3.Action = 7
		frm1.txtW5_3.focus
	End If
End Sub

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
					<TD >
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" idth="200" ><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;
						<a href="vbscript:GetRef">�ݾ� �ҷ�����</A>  
					</TD>
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
					<TD <%=HEIGHT_TYPE_03%> WIDTH=*> </TD>
				</TR>
					<TR>
					<TD WIDTH=800 valign=top  >
					   
									<TABLE width = 100% bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
									   
										
											<TR>
													<TD CLASS="TD51" align =center width = 20% colspan =2 >
														(1)����(�ٰ� �� ����)
													</TD>
												
													<TD CLASS="TD51" align =center width = 30% colspan =2  >
														(2)������ 
													</TD>
													<TD CLASS="TD51" align =center width = 35%  >
														(3)��곻�� 
													</TD>
													<TD CLASS="TD51" align =center width = 15% >
														(4)�������鼼�� 
													</TD>
												
												
											</TR>
										
										
											<TR>
											
											
										   
													<TD CLASS="TD51" align =left  colspan =2  >
														1)�����������Կ� ����<BR>
															&nbsp;&nbsp;���μ� ����(����Ư��<BR>
															&nbsp;&nbsp;���ѹ� ��20�� 2��)
													
													</TD>		
											
									
											       <TD CLASS="TD51" colspan =2 >
														<TABLE  CLASS="BasicTB" CELLSPACING=0 border="0">
																<TR>
																	<TD ALIGN=CENTER WIDTH=35%>(108)���⼼��</TD>
																	<TD ALIGN=CETER WIDTH=45%>
																	<TABLE  CLASS="BasicTB" CELLSPACING=0 border="0">
																		<TR>
																			<TD ROWSPAN=3 WIDTH=30% ALIGN=RIGHT>x&nbsp;&nbsp;&nbsp;</TD>
																			<TD ALIGN=CENTER>����ҵ�</TD>
																			<TD ROWSPAN=3 ></TD>
																		</TR>
																		<TR>
																			<TD HEIGHT=1 BGCOLOR=BLACK></TD>
																		</TR>
																		<TR>
																			<TD ALIGN=CENTER>����ǥ�رݾ�</TD>
																		</TR>
																	</TABLE>	
																	</TD>
																	<TD >&nbsp;</TD>
																</TR>
															</TABLE></TD>															
												  
											
											     <TD CLASS="TD51" >
														<TABLE  CLASS="BasicTB" CELLSPACING=0 border="0">
																<TR>
																	<TD ALIGN=CENTER WIDTH=35%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW3_1_A" name=txtW3_1_A CLASS=FPDS140 title=FPDOUBLESINGLE ALT="��곻��" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT> </TD>
																	<TD ALIGN=CETER WIDTH=45% colspan = 2>
																		<TABLE  CLASS="BasicTB" CELLSPACING=0 border="0">
																			<TR>
																				<TD ROWSPAN=3 WIDTH=30% ALIGN=RIGHT>x&nbsp;&nbsp;&nbsp;</TD>
																				<TD ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW3_1_B" name=txtW3_1_B CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X2Z" width = 100%></OBJECT>');</SCRIPT> </TD>
																			
																			</TR>
																			<TR>
																				<TD HEIGHT=1 BGCOLOR=BLACK></TD>
																			</TR>
																			<TR>
																				<TD ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW3_1_C" name=txtW3_1_C CLASS=FPDS140 title=FPDOUBLESINGLE ALT="��곻��" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT> </TD>
																			</TR>
																		</TABLE>	
																	</TD>
																	
																</TR>
															</TABLE></TD>															
												
												   <TD CLASS="TD51" > <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW4_1" name=txtW4_1 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT> </TD>
											
											
											
											
											</TR>
											
											
											
											<TR>
											
											
										   
													<TD CLASS="TD51" align =left  colspan =2  >
														 2)���ؼսǼ��װ���<BR>
															&nbsp;&nbsp;(���μ��� ��25��)
															
													</TD>
											
									
											       <TD CLASS="TD51" colspan =2 >
											       <TABLE  CLASS="BasicTB" CELLSPACING=0 border="0" >
														<TR>
															<TD ALIGN=CENTER WIDTH=35%>�̳��� �Ǵ�<BR>�����Ҽ���</TD>
															<TD ALIGN=CETER WIDTH=45%>
															<TABLE  CLASS="BasicTB" CELLSPACING=0 border="0">
																<TR>
																	<TD ROWSPAN=3 WIDTH=30% ALIGN=RIGHT>x&nbsp;&nbsp;&nbsp;</TD>
																	<TD ALIGN=CENTER>��ǵȻ����<BR>�ڻ갡��</TD>
																	<TD ROWSPAN=3 ></TD>
																</TR>
																<TR>
																	<TD HEIGHT=1 BGCOLOR=BLACK></TD>
																</TR>
																<TR>
																	<TD ALIGN=CENTER>�����<BR> �ڻ��Ѿ�</TD>
																</TR>
															</TABLE>	
															</TD>
															<TD >&nbsp;</TD>
														</TR>
													</TABLE></TD>															
												  
											
											     <TD CLASS="TD51" align =center valign=middle>
											       <TABLE  CLASS="BasicTB" CELLSPACING=0 border="0">
														<TR>
															<TD ALIGN=CENTER WIDTH=35%><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW3_2_A" name=txtW3_2_A CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X2Z" width = 100%></OBJECT>');</SCRIPT> </TD>
															<TD ALIGN=CETER WIDTH=45% colspan =2 valign=middle>
																<TABLE  CLASS="BasicTB" CELLSPACING=0 border="0">
																	<TR>
																		<TD ROWSPAN=3 WIDTH=30% ALIGN=RIGHT>x&nbsp;&nbsp;&nbsp;</TD>
																		<TD ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW3_2_B" name=txtW3_2_B CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X2Z" width = 100%></OBJECT>');</SCRIPT> </TD>
																	
																	</TR>
																	<TR>
																		<TD HEIGHT=1 BGCOLOR=BLACK></TD>
																	</TR>
																	<TR>
																		<TD ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW3_2_C" name=txtW3_2_C CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X2Z" width = 100%></OBJECT>');</SCRIPT> </TD>
																	</TR>
																</TABLE>	
															</TD>
														
														</TR>
													</TABLE>
													</TD>															
												
												   <TD CLASS="TD51" > <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW4_2" name=txtW4_2 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT> </TD>
											
											
											
											
											</TR>
											
										
											
										
										
										    <TR>
											    <TD CLASS="TD51" align =left colspan =2  >
												   3)<INPUT TYPE=TEXT NAME="txtw1_3"   tag="25"  maxlength=100 width = 100% >
											    </TD>
																		   
												<TD ALIGN=CENTER CLASS="TD51" colspan =2 ><INPUT TYPE=TEXT NAME="txtw2_3"   size =30 tag="25"  maxlength=100 width = 100% ></TD>
												<TD ALIGN=CENTER CLASS="TD51" ><INPUT TYPE=TEXT NAME="txtW3_3"   size =30 tag="25"  maxlength=100 width = 100% ></TD>
												<TD CLASS="TD51" > <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW4_3" name=txtW4_3 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X2Z" width = 100%></OBJECT>');</SCRIPT> </TD>
											
											</TR>
										   
										    <TR>
											    <TD CLASS="TD51" align =center colspan =2  >
												  �� 
											    </TD>
																		   
												<TD ALIGN=CENTER CLASS="TD51" colspan =2  ></TD>
												<TD ALIGN=CENTER CLASS="TD51" ></TD>
												<TD CLASS="TD51" > <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW4_SUM" name=txtW4_SUM CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100%></OBJECT>');</SCRIPT> </TD>
											
											</TR>
											
											<TR>
											
											    <TD CLASS="TD51" align =left colspan =6 >
												  (5)���� �߻����� 
											    </TD>
											
											</TR>
											<TR>
											    <TD CLASS="TD51" align =left  width = 10% >
												  1)���س��� 
											    </TD>
																		    
												<TD CLASS="TD51" align =left colspan =2 width = 15%  >
												   <INPUT TYPE=TEXT NAME="txtW5_1"   size =15 tag="25"  maxlength=100 width = 100% >
											    </TD>
											    <TD CLASS="TD51" align =center rowspan =3  >
												  4)�̳��� �Ǵ� <br>������ ���׸� 
											    </TD>
											    <TD CLASS="TD51" align =center  rowspan =2   >
												  ���� 
											    </TD>
											    <TD CLASS="TD51" align =center rowspan =2   >
												  ���μ� 
											    </TD>
														
											</TR>
											<TR>
											    <TD CLASS="TD51" align =left  >
												  2)���ع߻��� 
											    </TD>
																		   
												<TD CLASS="TD51" align =center colspan =2  >
												   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtw5_2 CLASS=FPDTYYYYMM title=FPDATETIME ALT="���ع߻���" tag="25X1" id=txtw5_2></OBJECT>');</SCRIPT>
											    </TD>
											   
											  
														
											</TR>
											<TR>
											    <TD CLASS="TD51" align =left  >
												  3)������û�� 
											    </TD>
																		   
												<TD CLASS="TD51" align =center  colspan =2  >
												   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtw5_3 CLASS=FPDTYYYYMM title=FPDATETIME ALT="������û��" tag="25X1" id=txtw5_3></OBJECT>');</SCRIPT>
											    </TD>
											     <TD CLASS="TD51" align =center  >
												  <INPUT TYPE=TEXT NAME="txtw5_4_GB"   size =30 tag="25"  maxlength=20 width = 100% >
											    </TD>
																		   
												<TD CLASS="TD51" align =center  >
												   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtw5_4" name=txtw5_4 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X2Z" width = 100%></OBJECT>');</SCRIPT> </OBJECT>
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

<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24" >
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" >



</FORM>
<DIV ID="MousePT" NAME="MousePT">

<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

