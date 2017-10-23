
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================

'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ��68ȣ��չ��μұް�����û�� 
'*  3. Program ID           : W8105MA1
'*  4. Program Name         : W8105MA1.asp
'*  5. Program Desc         : ��68ȣ��չ��μұް�����û�� 
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
Const BIZ_MNU_ID = "W8105MA1"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID = "W8105Mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID = "W8111OA1"


Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 

Dim dblOverRate , dblDownRate
Dim dblOverRate_View , dblDownRate_View

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
    Call AppendNumberPlace("6","3","2")
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
    
      Call fu_CompanyYYMMDD()
     
    
    
    call CommonQueryRs("REFERENCE_1,REFERENCE_2"," ufn_TB_Configuration('W2018','" & C_REVISION_YM & "')   "," Minor_cd = '1' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)   '1������ 
         dblDownRate = unicdbl(lgF0)
         dblDownRate_View = replace(lgF1,Chr(11),"")
    call CommonQueryRs("REFERENCE_1,REFERENCE_2"," ufn_TB_Configuration('W2018','" & C_REVISION_YM & "')   "," Minor_cd = '2' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)   '1���ʰ� 
         dblOverRate = unicdbl(lgF0)
         dblOverRate_View = replace(lgF1,Chr(11),"")
	   
	


End Sub

' ----------------------  ���� -------------------------
Function  Verification()
  dim IntRetCD

	Verification = False
    if frm1.txtw6.text < 0 then
       IntRetCD = DisplayMsgBox("WC0006","X",  frm1.txtw6.Alt, "X") 
       Exit function 
    end if   
	
	if unicdbl(frm1.txtw14.text) < unicdbl(frm1.txtw14.text) - unicdbl(frm1.txtw12.text)  then
       IntRetCD = DisplayMsgBox("WC0010","X",  "��������� �������鼼��", "���������� ������ ����") 
       Exit function 
    end if   
    
    if unicdbl(frm1.txtw15.text) > unicdbl(frm1.txtw12.text) then
       IntRetCD = DisplayMsgBox("WC0010","X",  frm1.txtw12.Alt, frm1.txtw15.Alt)    '%1�� '%2���� ���ų� �۾ƾߵ˴ϴ� 
       Exit function 
    end if   
    
	Verification = True	
End Function
'============================================  ���������  ====================================


function fu_CompanyYYMMDD 

  Dim sFiscYear, sRepType, sCoCd, iGap, IntRetCD

    sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value


   '����� 
		IntRetCD =  CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		if IntRetCD = false then
		   IntRetCD = DisplayMsgBox("X", "X", "��ȸ�� ������ ����̷��� �����ϴ�.", "X") 
		
		else
		
		   	frm1.txtw1_s.text = replace(lgF0, Chr(11),"")
			frm1.txtw1_e.text = replace(lgF1, Chr(11),"")
	
        end if
			
		'������������� 
		call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear -1 & "' AND REP_TYPE='1'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
		frm1.txtw2_s.text = replace(lgF0, Chr(11),"")
		frm1.txtw2_e.text = replace(lgF1, Chr(11),"")
End function


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
	
	' �·ε�� ���۷����޽��� �����´�.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf
	Call selectColor(frm1.txtw6)
    Call selectColor(frm1.txtw8)
    Call selectColor(frm1.txtw11)

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"
	Call ggoOper.LockField(Document, "N") 
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If


   '***���� 
   ' W6 �� 3ȣ ������ (107) �� ����⵵�ҵ�ݾ� 
   ' W8 ���� ������� �� 3ȣ ������ (112) ����ǥ�� 
   ' W9 ���� ������� �� 3ȣ ������ (117) ���� 
   ' W11 ���� ������� (121) �������鼼��(��) + (123) �������鼼��(��)


	call CommonQueryRs("w6,W8,  W11","dbo.ufn_TB_68_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	IF lgF0 = "" THEN EXIT Function 
    if unicdbl(lgF0) * (-1) > 0 then
       frm1.txtw6.value = unicdbl(replace(lgF0,chr(11),"")) * (-1) 
    else
       frm1.txtw6.value = 0
    end if
       frm1.txtw8.value = unicdbl(replace(lgF1,chr(11),""))
     
       frm1.txtw11.value = unicdbl(replace(lgF2,chr(11),""))    
    
    
    
      lgBlnFlgChgValue = TRUE
	  Call Fn_CalSum()
   
End Function


Function Fn_CalSum()
 
End Function

'============================================  �̺�Ʈ �Լ�  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub


Function Fn_CalSum1()


    if  unicdbl(frm1.txtw11.text) <= 100000000 * unicdbl(dblDownRate) then														 '(11) * 1�� * ������μ���(1��̸�)
				if ( unicdbl(frm1.txtw8.text)- unicdbl(frm1.txtw11.text)/unicdbl(dblDownRate)) <   unicdbl(frm1.txtw6.text) then      'Min[(8)-(11) / ������μ���(1��̸�),6]
				     frm1.txtw7.text =    unicdbl(frm1.txtw8.text)- unicdbl(frm1.txtw11.text)/unicdbl(dblDownRate) 
				Else
				    frm1.txtw7.text =    unicdbl(frm1.txtw6.text)
				End if

    Else
				if  (unicdbl(frm1.txtw8.text)-(100000000 + (unicdbl(frm1.txtw11.text)-100000000*unicdbl(dblDownRate)) /unicdbl(dblDownRate))) < unicdbl(frm1.txtw6.text)  then
				     frm1.txtw7.text  = (unicdbl(frm1.txtw8.text)-(100000000 + (unicdbl(frm1.txtw11.text)-100000000*unicdbl(dblDownRate)) /unicdbl(dblOverRate) ))
				else
				     frm1.txtw7.text  =  unicdbl(frm1.txtw6.text)
				end if
    
    End if 
  	

 

end function


Function Fn_CalSum()
  
    
    if unicdbl(frm1.txtw8.text) > 100000000 then
       frm1.txtw9.value = dblOverRate_View
       frm1.txtw9_value.value = (dblOverRate)
    else
	   frm1.txtw9.value = dblDownRate_View
       frm1.txtw9_value.value = (dblDownRate)
    
    end if   
         
  	
    if unicdbl(frm1.txtw8.text) <=0 then

       frm1.txtw10.text =0
       
   
    elseif  unicdbl(frm1.txtw8.text) > 0 and unicdbl(frm1.txtw8.text)  <= 100000000 then
        frm1.txtw10.text = unicdbl(frm1.txtw8.text) * unicdbl(dblDownRate)
    
    elseif  unicdbl(frm1.txtw8.text) >  100000000 then
       
       frm1.txtw10.text =( 100000000  * unicdbl(dblDownRate))  + (unicdbl(frm1.txtw8.text) - 100000000) *  unicdbl(dblOverRate)
    end if
     
      frm1.txtw12.text = unicdbl(frm1.txtw10.text)-unicdbl(frm1.txtw11.text)
      frm1.txtw13.text = unicdbl(frm1.txtw10.text)
      
     if	unicdbl(frm1.txtw8.text)-unicdbl(frm1.txtw7.text)  <=0  then
			 frm1.txtw14.text = 0
     elseif (unicdbl(frm1.txtw8.text)-unicdbl(frm1.txtw7.text)) >0 and (unicdbl(frm1.txtw8.text)-unicdbl(frm1.txtw7.text))  <=100000000 then
			 frm1.txtw14.text  = (unicdbl(frm1.txtw8.text)-unicdbl(frm1.txtw7.text)) * dblDownRate			
	 elseif (unicdbl(frm1.txtw8.text)-unicdbl(frm1.txtw7.text)) > 100000000		 then
           frm1.txtw14.text  = (100000000 * dblDownRate) + ((unicdbl(frm1.txtw8.text)-unicdbl(frm1.txtw7.text))-100000000) *  unicdbl(dblOverRate)
           
     end if   
     
     
   
         frm1.txtw15.text = unicdbl(frm1.txtw13.text) - unicdbl(frm1.txtw14.text)
    

end function


Function CheckMessage(ByVal Obj)
dim IntRetCD
    if  UNICDbl(Obj.value) < 0 then
        Obj.value = 0
        Obj.focus

    end if
    
end function



Sub txtW6_CHANGE()
    lgBlnFlgChgValue = TRUE
    Call Fn_CalSum
End Sub


Sub txtW7_CHANGE()
    lgBlnFlgChgValue = TRUE
    Call Fn_CalSum
End Sub


Sub txtW8_CHANGE()
    lgBlnFlgChgValue = TRUE
    Call Fn_CalSum
End Sub


Sub txtW11_CHANGE()
    lgBlnFlgChgValue = TRUE
    Call Fn_CalSum
End Sub

Sub txtW11_CHANGE()
    lgBlnFlgChgValue = TRUE
    Call Fn_CalSum
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
     Call fu_CompanyYYMMDD()
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
    
    
    
    if  frm1.txtW1_S.text ="" or frm1.txtW1_E.text ="" then
       Call DisplayMsgBox("X","x","��� ������ �Ǵ� �������� �����ϴ�.","x")  		
       Exit Function
    end if
    
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

    If Verification = False Then Exit Function 
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


'==========================================================================================
Sub txtW1_S_KeyDown(KeyCode, Shift)
	 
End Sub

'======================================================================================================
'   Event Name : txtW1_S_KeyDown
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtW1_S_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtW1_S.Action = 7
		frm1.txtW1_S.focus
	End If
End Sub


'======================================================================================================
'   Event Name : txtW1_E_KeyDown
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtW1_E_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtW1_E.Action = 7
		frm1.txtW1_E.focus
	End If
End Sub



'======================================================================================================
'   Event Name : txtW2_S_KeyDown
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtW2_S_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtW2_S.Action = 7
		frm1.txtW2_S.focus
	End If
End Sub

'======================================================================================================
'   Event Name : txtW2_E_KeyDown
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtW2_E_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtW2_E.Action = 7
		frm1.txtW2_E.focus
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
									<TD CLASS="TD6"><script language =javascript src='./js/w8105ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
					<TD WIDTH=620 valign=top  >
					   
					    
									<TABLE bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
									  
									
										<TR>
											 <TD CLASS="TD51" align =LEFT width =15% >
												(1)��ջ������ 
											</TD>
											
										    <TD CLASS="TD61" align =LEFT >
													<script language =javascript src='./js/w8105ma1_txtW1_S_txtW1_S.js'></script>~
													<script language =javascript src='./js/w8105ma1_txtW1_E_txtW1_E.js'></script>
											</TD>
											
										    <TD CLASS="TD51" align =LEFT  >
												(2)����������� 
											</TD>
											 <TD CLASS="TD61" align =LEFT     >
													<script language =javascript src='./js/w8105ma1_txtW2_S_txtW2_S.js'></script>~
													<script language =javascript src='./js/w8105ma1_txtW2_E_txtW2_E.js'></script>
											</TD>
											
											
										</TR>
										
										<TR>
											 <TD CLASS="TD51" align =LEFT  rowspan =2>
												(2)��ջ������<br>&nbsp;&nbsp;��ձݾ� 
											</TD>
											<TD CLASS="TD51" align =LEFT  colspan =2>
												(6)��ձݾ� 
											</TD>
											 
										    <TD CLASS="TD61" align =LEFT   >
												<script language =javascript src='./js/w8105ma1_txtW6_txtW6.js'></script>
											</TD>
											
										</TR>
										<TR>
											 <TD CLASS="TD51" align =LEFT  >
												(7)�ұް������� ��ձݾ� 
											</TD>
											 <TD  CLASS="TD51"  align =LEFT   ><BUTTON NAME="btnCb_autoisrt"  ONCLICK="VBScript: Fn_CalSum1()" >�ڵ����</BUTTON></TD>
										     <TD CLASS="TD61" align =LEFT  >
												<script language =javascript src='./js/w8105ma1_txtW7_txtW7.js'></script>
											</TD>
											
										</TR>
										<TR>
										   <TD CLASS="TD51" align =LEFT  rowspan =5>
												(4)�����������<br>&nbsp;&nbsp;���μ��װ�� 
											</TD>
											 <TD CLASS="TD51" align =LEFT  colspan =2>
												(8)����ǥ�� 
											</TD>
											
										    <TD CLASS="TD61" align =LEFT  >
												<script language =javascript src='./js/w8105ma1_txtW8_txtW8.js'></script>
											</TD>
											
										</TR>
										
										<TR>
										   
											 <TD CLASS="TD51" align =LEFT  colspan =2>
												(9)���� 
											</TD>
											
										    <TD CLASS="TD61" align=right nowrap  >
												<INPUT TYPE=TEXT id="txtw9" NAME="txtw9" Size=35 tag="24" style=""></OBJECT>
											</TD>
											
											
										</TR>
										<TR>
										   
											 <TD CLASS="TD51" align =LEFT  colspan =2>
												(10)���⼼�� 
											</TD>
											
										    <TD CLASS="TD61" align =LEFT  >
												<script language =javascript src='./js/w8105ma1_txtW10_txtW10.js'></script>
											</TD>
											
											
										</TR>
										<TR>
										   
											 <TD CLASS="TD51" align =LEFT  colspan =2>
												(11)�������鼼�� 
											</TD>
											
										    <TD CLASS="TD61" align =LEFT  >
												<script language =javascript src='./js/w8105ma1_txtW11_txtW11.js'></script>
											</TD>
											
											
										</TR>
										<TR>
										   
											 <TD CLASS="TD51" align =LEFT  colspan =2>
												(12)��������((10)-(11))
											</TD>
											
										    <TD CLASS="TD61" align =LEFT  >
												<script language =javascript src='./js/w8105ma1_txtW12_txtW12.js'></script>
											</TD>
											
											
										</TR>
					
										<TR>
										   <TD CLASS="TD51" align =LEFT  rowspan =5>
												(5)ȯ�޽�û<br>&nbsp;&nbsp;���� ��� 
											</TD>
											 <TD CLASS="TD51" align =LEFT  colspan =2>
												(13)��������������μ���((13)=(10))
											</TD>
											
										    <TD CLASS="TD61" align =LEFT  >
												<script language =javascript src='./js/w8105ma1_txtW13_txtW13.js'></script>
											</TD>
											
										</TR>
										
										<TR>
										   
											 <TD CLASS="TD51" align =LEFT  colspan =2>
											(14)������ ����[((8)-(7))*����][(14) ��(10)-(12)]
											</TD>
											
										    <TD CLASS="TD61" align =LEFT  >
												<script language =javascript src='./js/w8105ma1_txtW14_txtW14.js'></script>
											</TD>
											
											
										</TR>
										<TR>
										   
											 <TD CLASS="TD51" align =LEFT  colspan =2>
												(15)ȯ�޽�û����((13)-(14))(15��(12))
											</TD>
											
										    <TD CLASS="TD61" align =LEFT  >
												<script language =javascript src='./js/w8105ma1_txtW15_txtW15.js'></script>
											</TD>
											
											
										</TR>
					
					
											
									</TABLE>
						
						   			
					</TD>
				</TR>
		    
				
			</TABLE>
		</TD>
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
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtw9_VALUE" tag="24">

<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24" >
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" >



</FORM>
<DIV ID="MousePT" NAME="MousePT">

<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

