
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================

'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ��2ȣ�����Ư��������ǥ�ع׼��׽Ű� 
'*  3. Program ID           : W8113MA1
'*  4. Program Name         : W8113MA1.asp
'*  5. Program Desc         : ��2ȣ�����Ư��������ǥ�ع׼��׽Ű� 
'*  6. Modified date(First) : 2005/03/18
'*  7. Modified date(Last)  : 2005/03/18
'*  8. Modifier (First)     : ȫ���� 
'*  9. Modifier (Last)      : ȫ���� 
'* 10. Comment              : ���� : ufn_TB_2_GetRef
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
Const BIZ_MNU_ID = "W8113MA1"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID = "W8113MB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID = "W8105OA1"

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
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>

	' �����Ѱ� 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
    Call SetDefaultVal()

     
    ' �������� üũȣ�� 
	Call FncQuery
  
End Sub




Sub SetDefaultVal()
dim strWhere 
DIM strW1

    frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"


End Sub



Function GetRef()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	Dim sMesg
	Dim W1,W2,W134,W154,W150
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	
	
	if wgConfirmFlg = "Y" then    'Ȯ���� 
	   Exit function
	end if   
	
	' �·ε�� ���۷����޽��� �����´�.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf
	Call selectColor(frm1.txtW1)
    Call selectColor(frm1.txtW2)

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"
	Call ggoOper.LockField(Document, "N") 
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If


       ' W1 = 2ȣ ������  ��Ұ��� �ݾ��� �Է���.
       ' W2 = 12ȣ ������  ��Ұ��� �鼼���� �Է���.
	   ' W150 = (150)���������� ���� �� 


	call CommonQueryRs("W1,W2,W150","dbo.ufn_TB_2_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	     IF lgF0 = "" THEN EXIT Function 
	        W1	    = unicdbl(replace(lgF0, chr(11),""))		 
            W2      = unicdbl(replace(lgF1, chr(11),""))
            W150    = unicdbl(replace(lgF2, chr(11),"")) 
		
		
		   frm1.txtW1.value	    = W1 
           frm1.txtW2.value     = W2
 
           if  unicdbl(W150 ) >= 0  then
               Call ggoOper.SetReqAttr(frm1.txtW10_1 , "Q")
           else
               Call ggoOper.SetReqAttr(frm1.txtW10_1 , "D")
               
               frm1.txtW10_1.value = W150 * (-1)
           end if   
               
    
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


Function CalSum()
	Dim W1,W2,W134,W154,W150
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	
	'��� �� �� �� ��	�����������Ҽ���  < 0  �ΰ�� "0"�� �Է��ϰ�, �� ���� ���� �Ʒ��� ���� �Է���.
	'3ȣ������ (154)�� > 0 �� ��� 
	'[ �� �� 3ȣ������ (154) �� 3ȣ������ (134) ] �� ����Ͽ� �Է��ϰ� 
	'3ȣ������ (154)�� <= 0 �� ��� 
	'����   �����������Ҽ��� <= 5,000,000 �� ��� "0"�� �Է��ϰ� 
	'����   �����������Ҽ��� <= 10,000,000 �� ��� (�����������Ҽ���-5,000,000)�� ����Ͽ� �Է��ϰ� 
	'�� ���� ���� (�����������Ҽ���  �� 50%) �� ����Ͽ� �Է���.
    
    '*			 W154  --3ȣ���� 154
	'*			 W150  --3ȣ���� 150
	'*			 W134  --3ȣ���� 134
	Call CommonQueryRs("W134,W154,W150","dbo.ufn_TB_2_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

            W134    = unicdbl(replace(lgF0, chr(11),""))		
            W154    = unicdbl(replace(lgF1, chr(11),""))
            W150    = unicdbl(replace(lgF2, chr(11),"")) 
            

	
    frm1.txtw4.value =  (UNICDbl(frm1.txtw2.value ) + UNICDbl(frm1.txtw3.value))
  
    frm1.txtw6.value =  UNICDbl(frm1.txtw4.value) -  UNICDbl(frm1.txtw5.value)
    
    
    if  UNICDbl(frm1.txtw6.value) < 0 then
         frm1.txtw7.value = 0
    else
          if  unicdbl(W150) > 0 then
				if	UNICDbl(W134)  = 0 then
				      frm1.txtw7.value = 0
				else
						frm1.txtw7.value =  UNICDbl(frm1.txtw6.value) * (UNICDbl(W154)  / UNICDbl(W134) )
		
				end if   
		  else
		        if unicdbl(frm1.txtw6.value) <= 5000000 then
		                    frm1.txtw7.value = 0
		        elseif      unicdbl(frm1.txtw6.value) <= 10000000  and unicdbl(frm1.txtw6.value) > 5000000 then
		                    frm1.txtw7.value =  unicdbl(frm1.txtw6.value) - 5000000
		        else
						    frm1.txtw7.value  =unicdbl(frm1.txtw6.value)  * 0.5
		        
		        end if            
		  
		  end if		
              
    end if
     
     frm1.txtw8.value =  UNICDbl(frm1.txtw6.value) -  UNICDbl(frm1.txtw7.value)
     
     if UNICDbl(frm1.txtw8.value ) - UNICDbl(frm1.txtw10_2.value ) <= 0 then
        frm1.txtw9.value = 0
     else   
        frm1.txtw9.value = UNICDbl(frm1.txtw8.value ) - UNICDbl(frm1.txtw10_2.value ) 
     end if 
     

end function

Function CheckMessage(ByVal Obj)
dim IntRetCD
    if  UNICDbl(Obj.value) < 0 then
        Obj.value = 0
        Obj.focus

    end if
    
end function






Sub txtw1_Change()  
     
    lgBlnFlgChgValue  = True 
    Call CheckMessage(frm1.txtw1)
    Call CalSum() 
    
End Sub


Sub txtw2_Change()  
     
    lgBlnFlgChgValue  = True 
    Call CheckMessage(frm1.txtw2)
    Call CalSum() 
    
End Sub

Sub txtw3_Change()  
     
    lgBlnFlgChgValue  = True 

    Call CalSum() 
    
End Sub

Sub txtw4_Change()  
     
    lgBlnFlgChgValue  = True 

    Call CalSum() 
    
End Sub


Sub txtw5_Change()  
     
    lgBlnFlgChgValue  = True 

    Call CalSum() 
    
End Sub

Sub txtw6_Change()  
     
    lgBlnFlgChgValue  = True 

    frm1.txtw8.value =  UNICDbl(frm1.txtw6.value) -  UNICDbl(frm1.txtw7.value)
    
     if UNICDbl(frm1.txtw8.value ) - UNICDbl(frm1.txtw10_2.value ) <= 0 then
        frm1.txtw9.value = 0
     else   
        frm1.txtw9.value = UNICDbl(frm1.txtw8.value ) - UNICDbl(frm1.txtw10_2.value ) 
     end if 
    
End Sub
 
 
 Sub txtw7_Change()  
     
    lgBlnFlgChgValue  = True 

    Call txtw6_Change
   
    
End Sub


Sub txtW10_1_Change()  
     
    lgBlnFlgChgValue  = True 
    
End Sub



Sub txtw10_2_Change()  
     
    lgBlnFlgChgValue  = True 
    Call CheckMessage(frm1.txtw2)
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
	    Call SetToolbar("11001000000000111")										<%'��ư ���� ���� %>
	    
		
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
	Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w8113ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
					<TD WIDTH=520 valign=top  >
					   
					       <FIELDSET CLASS="CLSFLD">
									<TABLE bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
									   
									
										<TR>
											 <TD CLASS="TD51" align =LEFT width = 15% >
												(1)����ǥ�� 
											</TD>
											
										    <TD CLASS="TD61" align =LEFT   colspan = "2" >
												<script language =javascript src='./js/w8113ma1_txtW1_txtW1.js'></script>
											</TD>
											
										</TR>
										
										<TR>
											 <TD CLASS="TD51" align =LEFT width = 15%>
												(2)���⼼�� 
											</TD>
											 
										    <TD CLASS="TD61" align =LEFT  colspan = "2" >
												<script language =javascript src='./js/w8113ma1_txtW2_txtW2.js'></script>
											</TD>
											
										</TR>
										<TR>
											 <TD CLASS="TD51" align =LEFT width = 15% >
												(3)���꼼�� 
											</TD>
											
										    <TD CLASS="TD61" align =LEFT colspan = "2" >
												<script language =javascript src='./js/w8113ma1_txtW3_txtW3.js'></script>
											</TD>
											
										</TR>
										<TR>
											 <TD CLASS="TD51" align =LEFT width = 15% >
												(4)�Ѻδ㼼�� 
											</TD>
											
										    <TD CLASS="TD61" align =LEFT colspan = "2" >
												<script language =javascript src='./js/w8113ma1_txtW4_txtW4.js'></script>
											</TD>
											
										</TR>
										<TR>
											 <TD CLASS="TD51" align =LEFT width = 15%  >
												(5)�ⳳ�μ��� 
											</TD>
											
										    <TD CLASS="TD61" align =LEFT colspan = "2" >
												<script language =javascript src='./js/w8113ma1_txtW5_txtW5.js'></script>
											</TD>
											
										</TR>
										<TR>
											 <TD CLASS="TD51" align =LEFT width = 15%  >
												(6)���������Ҽ��� 
											</TD>
											
										    <TD CLASS="TD61" align =center width = 15% colspan = "2"  >
												<script language =javascript src='./js/w8113ma1_txtW6_txtW6.js'></script>
											</TD>
											
										</TR>
										<TR>
											 <TD CLASS="TD51" align =LEFT width = 15%>
												(7)�г��Ҽ��� 
											</TD>
											
										    <TD CLASS="TD61" align =LEFT  colspan = "2" >
												<script language =javascript src='./js/w8113ma1_txtW7_txtW7.js'></script>
											</TD>
											
										</TR>
										<TR>
											 <TD CLASS="TD51" align =LEFT  width = 15%>
												(8)�������μ��� 
											</TD>
											
										    <TD CLASS="TD61" align =LEFT colspan = "2" >
												<script language =javascript src='./js/w8113ma1_txtW8_txtW8.js'></script>
											</TD>
											
										</TR>
										<TR>
											 <TD CLASS="TD51" align =LEFT width = 15% >
												(9)����� ���μ��� 
											</TD>
											
										    <TD CLASS="TD61" align =LEFT width = 15% colspan = "2"  >
												<script language =javascript src='./js/w8113ma1_txtW9_txtW9.js'></script>
											</TD>
											
										</TR>
										<TR>
											 <TD CLASS="TD51" align =LEFT  ROWSPAN = 2 >
												(10)����ȯ�ޱ� ����û 
											</TD>
											 <TD CLASS="TD51" align =LEFT  >
												ȯ�޹��μ� 
											 </TD>
											
										     <TD CLASS="TD61" align =LEFT  width = 10%>
												<script language =javascript src='./js/w8113ma1_txtW10_1_txtW10_1.js'></script>
											</TD>
											
										</TR>
										<TR>
									
										 	<TD CLASS="TD51" align =LEFT  width = 10%>
												����� �����Ư���� 
											 </TD>
											 <TD CLASS="TD61" align =LEFT width = 10%>
												<script language =javascript src='./js/w8113ma1_txtW10_2_txtW10_2.js'></script>
											</TD>
										 </TR>
										
					
											
									</TABLE>
						   </FIELDSET>	
						   			
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
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('VIEW')" Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('PRINT')"   Flag=1>�μ�</BUTTON></TD>
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
<INPUT TYPE=HIDDEN NAME="txtW154" tag="24" ALT="">
<INPUT TYPE=HIDDEN NAME="txtW134" tag="24" ALT="">
<INPUT TYPE=HIDDEN NAME="txtW150" tag="24" ALT="">


</FORM>
<DIV ID="MousePT" NAME="MousePT">

<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

