
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================

'*  1. Module Name          : ���μ� 
'*  2. Function Name        : 33ȣ ���������� �������� 
'*  3. Program ID           : W3103MA1
'*  4. Program Name         : W3103MA1.asp
'*  5. Program Desc         : 33ȣ ���������� �������� 
'*  6. Modified date(First) : 2005/03/18
'*  7. Modified date(Last)  : 2006/01/23
'*  8. Modifier (First)     : ȫ���� 
'*  9. Modifier (Last)      : HJO 
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
Const BIZ_MNU_ID = "W3103MA1"
Const BIZ_PGM_ID = "W3103MB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID  = "w3103oa1"

Const C_SHEETMAXROWS = 100




Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 

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

Sub InitSpreadComboBox()
    Dim iCodeArr , IntRetCD1
    Dim iNameArr
    Dim iDx



End Sub


Sub SetSpreadLock()
 
   
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub


Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
         
          
       
	

    End Select    
End Sub

'============================================  ��ȸ���� �Լ�  ====================================

'============================================  �� �Լ�  ====================================




Sub Form_Load()

    Call LoadInfTB19029     
                                             <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet()                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1110100000001111")										<%'��ư ���� ���� %>

	' �����Ѱ� 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
    Call SetDefaultVal
	
  
End Sub

Sub SetDefaultVal()

    frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"     

    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

End Sub
'============================================  �̺�Ʈ �Լ�  ====================================

Function  Verification()

	
	Dim IntRetCD
	dim strWhere
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	    

   
          if  unicdbl(frm1.txtW21.value) <0 then
              IntRetCD = DisplayMsgBox("WC0006", parent.VB_INFORMATION, "�������迹ġ��", "0") 
              Exit Function
          end if
         
  

	Verification = True	
End Function



Function OpenRefMenu()

    Dim arrRet
    Dim arrParam(2)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True
    
    arrRet = window.showModalDialog("../w5/W5105RA1.asp", Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
End Function


Function GetRef()	
    Dim IntRetCD , i
    Dim sMesg
   
   
    'Call ggoOper.ClearField(Document, "2")

	if wgConfirmFlg = "Y" then    'Ȯ���� 
	   Exit function
	end if   
	

	sMesg = wgRefDoc & vbCrLf & vbCrLf
    call SelectColor(frm1.txtW1)  
    call SelectColor(frm1.txtW2)  
    call SelectColor(frm1.txtW3) 
    call SelectColor(frm1.txtW10)  
    call SelectColor(frm1.txtW13)  
    call SelectColor(frm1.txtW14) 
    call SelectColor(frm1.txtW18) 
    call SelectColor(frm1.txtW19) 
    call SelectColor(frm1.txtW20) 
    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"
     Call ggoOper.LockField(Document, "N")
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
     CALL getdata()
    	
End Function





Function GetData()	

	Dim IntRetCD1
	dim strWhere
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	    
    strWhere = FilterVar(Trim(frm1.txtCO_CD.value ),"","S")  
    strWhere = strWhere & " ," & FilterVar(Trim(frm1.txtFISC_YEAR.text ),"","S")
    strWhere = strWhere & " ," & FilterVar(Trim(frm1.cboREP_TYPE.value ),"","S") 
	
	
	call CommonQueryRs("w1,w2,w3,w10,w13,w14"," dbo.ufn_TB_33_GetRef("& strWhere &")" ,,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
          frm1.txtW1.value= unicdbl(replace(lgF0 ,Chr(11),""))
          frm1.txtW2.value= unicdbl(replace(lgF1 ,Chr(11),""))  '32(4) -32(5) - 32(7) - 32(12) - 32(w15)
          frm1.txtW3.value= unicdbl(replace(lgF2 ,Chr(11),""))
          frm1.txtW10.value=unicdbl(replace(lgF3 ,Chr(11),""))
          frm1.txtW13.value=unicdbl(replace(lgF4 ,Chr(11),""))
          frm1.txtW14.value=unicdbl(replace(lgF5 ,Chr(11),""))
   

    call CommonQueryRs("w18,w19,w20"," dbo.ufn_TB_33_GetRef("& strWhere &")", ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
         
          frm1.txtW18.value= unicdbl(replace(lgF0 ,Chr(11),""))
          frm1.txtW19.value= unicdbl(replace(lgF1 ,Chr(11),""))
          frm1.txtW20.value= unicdbl(replace(lgF2 ,Chr(11),""))
          

End Function


Function CheckData()	




End function


Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub

Function Find_Max(Byval a, byval b)
    if a > b then
       Find_Max = unicdbl(a)
    else
       Find_Max = unicdbl(b)
    end if 
    
    
End Function


Sub txt32w1_Change( )    ' 32ȣ ���� 17ȣ �ݾ� 
    lgBlnFlgChgValue  = True 
    Frm1.txtw12.text = unicdbl(Frm1.txtw32.value) -unicdbl(Frm1.txtw19.value)
 
End Sub

Sub txtw1_Change( )      ' ��⸻ ���� �� ������� �����޿� �߰�� 
    lgBlnFlgChgValue  = True 
    Frm1.txtw5.text = Find_Max(unicdbl(Frm1.txtw1.value) - unicdbl(Frm1.txtw4.value),0)
 
End Sub


Sub txtw2_Change( )       ' ��λ����ܾ� 
    lgBlnFlgChgValue  = True 
    Frm1.txtw4.text = Find_Max(unicdbl(Frm1.txtw2.value) - unicdbl(Frm1.txtw3.value) ,0)
End Sub


Sub txtw3_Change( )        ' ���δ���� 
   lgBlnFlgChgValue  = True  
   call txtw2_Change( ) 
End Sub


Sub txtw4_Change( )         ' ������ 
    lgBlnFlgChgValue  = True 
    call txtw1_Change( ) 
End Sub

Sub txtw5_Change( )         ' ���������� �ձݻ��Դ����ѵ��� 
    lgBlnFlgChgValue  = True 
    Frm1.txtw7.text = Find_Max(unicdbl(Frm1.txtw5.value) - unicdbl(Frm1.txtw6.value),0)
End Sub


Sub txtw6_Change( )         ' �̹̼ձݻ����� ������ 
    lgBlnFlgChgValue  = True 
    Call txtw5_Change
End Sub


Sub txtw7_Change( ) 
     lgBlnFlgChgValue  = True       
    if unicdbl(Frm1.txtw7.text ) > unicdbl(Frm1.txtw8.text) then
       Frm1.txtw9.text =  unicdbl(Frm1.txtw8.text)
    else
        Frm1.txtw9.text =  unicdbl(Frm1.txtw7.text)
    end if 
End Sub

Sub txtw8_Change( )
            
    call txtw7_Change
End Sub


Sub txtw9_Change( )
    lgBlnFlgChgValue  = True         
     Frm1.txtw11.text = unicdbl(Frm1.txtw9.value) - unicdbl(Frm1.txtw10.value) 
End Sub

Sub txtw10_Change( )        
    call txtw9_Change
End Sub

Sub txtw12_Change( )        
    lgBlnFlgChgValue  = True
    If  unicdbl(Frm1.txtw16.value)<0 then
		Frm1.txtw17.text = unicdbl(Frm1.txtw12.value) - unicdbl(0)
	Else
		 Frm1.txtw17.text = unicdbl(Frm1.txtw12.value) - unicdbl(Frm1.txtw16.value)
	End IF
End Sub

Sub txtw13_Change( )   
	Dim tmpVal     
    lgBlnFlgChgValue  = True 
	Frm1.txtw16.text =unicdbl(Frm1.txtw13.value) - unicdbl(Frm1.txtw14.value) - unicdbl(Frm1.txtw15.value)
End Sub

Sub txtw14_Change( )        
    call txtw13_Change
End Sub
Sub txtw15_Change( )        
    call txtw13_Change
End Sub
Sub txtw16_Change( )         ' �̹̼ձݻ����� ������ 
    lgBlnFlgChgValue  = True 
    If unicdbl(Frm1.txtw16.value)<0 Then 
  		Frm1.txtw6.text =unicdbl(0)
  	Else
  		Frm1.txtw6.text =unicdbl(Frm1.txtw16.value)  	
  	End IF
    call txtw12_Change    
End Sub
Sub txtw21_Change( )        
    lgBlnFlgChgValue  = True 
    Frm1.txtw12.text = unicdbl(Frm1.txtw21.value)
End Sub
Sub txtw17_Change( )         ' �ձݻ��Դ����� �� 
    lgBlnFlgChgValue  = True 
    Frm1.txtw8.text = unicdbl(Frm1.txtw17.value)
End Sub

Sub txtw18_Change( )         ' �̹̼ձݻ����� ������ 
    lgBlnFlgChgValue  = True 
    Frm1.txtw21.text = unicdbl(Frm1.txtw18.value) -unicdbl(Frm1.txtw19.value)+unicdbl(Frm1.txtw20.value)
End Sub


Sub txtw19_Change( )         ' ������������ ��ġ�ݵ� ���� �� �ؾ�� 
    lgBlnFlgChgValue  = True 

    Frm1.txtw15.text = unicdbl(Frm1.txtw19.value)
    call txtw18_Change
End Sub

Sub txtw20_Change( )         ' �̹̼ձݻ����� ������ 
    lgBlnFlgChgValue  = True 
     call txtw18_Change
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

     Call SetToolbar("1110110000001111")          '��: ��ư ���� ���� 
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
    if Verification = False then Exit Function
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

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	 Call ggoOper.LockField(Document, "N")
	 
	Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
    Call InitData
	'1 ����üũ 
	If wgConfirmFlg = "Y" Then

	    Call SetToolbar("1100000000011111")	
		
	Else
	   '2 ���ȯ�氪 , �ε��ȯ�氪 �� 
		  Call SetToolbar("1111111100111111")									<%'��ư ���� ���� %>
	
	End If
	
	
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE
    Call SetSpreadColor(-1,-1)  
    lgBlnFlgChgValue = false
    Call SetToolbar("1111100000011111")										<%'��ư ���� ���� %>
  
		
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



Function OpenRef()	'�ҵ�ݾ��հ�ǥ 

    Dim arrRet
    Dim arrParam(4)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
	Dim arrRowVal
    Dim arrColVal, lLngMaxRow
    Dim iDx
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("WB001RA1")
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "WB001RA1", "x")
        IsOpenPop = False
        Exit Function
    End If
    
	arrParam(0) = frm1.txtCO_CD.Value
	arrParam(1) = frm1.txtCO_NM.Value		
	arrParam(2) = frm1.txtFISC_YEAR.Text		
	arrParam(3) = frm1.cboREP_TYPE.Value		

    arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	
    
    IsOpenPop = False
    
    
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
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;
								<a href="vbscript:GetRef()">�ݾ׺ҷ�����</A>|<A href="vbscript:OpenRefMenu">�ҵ�ݾ��հ�ǥ��ȸ</A></TD>
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
					<TD WIDTH=1024 valign=top HEIGHT="100" >
					   
					      <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>1.�������� ���� ����� ���� </LEGEND>
									<TABLE width = 100% bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
									   
										
										<TR>
											<TD CLASS="TD51" align =center width = 10% ROWSPAN=2>
												(1)�⸻���� ����� ��<BR>�ӿ� ������ ������<br>�����޿��߰� 
											</TD>
											
										    <TD CLASS="TD51" align =center width = 15%  COLSPAN=3>
												��⸻ ���� ������������ 
											</TD>
											<TD CLASS="TD51" align =center width = 15%   ROWSPAN=2>
												(5)����������<BR>�ձݻ��Դ���<BR> �ѵ���((1)-(4))
											</TD>
											
												<TD CLASS="TD61" align =center width = 15%>
											</TD>
										</TR>
										<TR>
											<TD CLASS="TD51" align =center width = 10% >
												(2)��λ�⸻�ܾ� 
											</TD>
											
										    <TD CLASS="TD51" align =center width = 15%  >
												(3)���δ���� 
											</TD>
											<TD CLASS="TD51" align =center width = 15%  >
												(4)�� �� ��((2)-(3))
											</TD>
											<TD CLASS="TD61" align =center width = 15%>
											</TD>
										</TR>
										
										<TR>
											
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW1" name=txtW1 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS="TD61" align =center width = 15%>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW2" name=txtW2 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW3" name=txtW3 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2" width = 100%></OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS="TD61" align =center width = 15%>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW4" name=txtW4 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2" width = 100%></OBJECT>');</SCRIPT>
											</TD>
												<TD CLASS="TD61" align =center width = 15%>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW5" name=txtW5 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2" width = 100%></OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS="TD61" align =center width = 15%>
											</TD>
										</TR>
										
										<TR>
											<TD CLASS="TD51" align =center width = 10% >
												(6)�̹̼ձݻ�����<BR>������(16)
											</TD>
											
										    <TD CLASS="TD51" align =center width = 15% x>
												(7)�ձݻ���<BR>�ѵ���((5)-(6))
											</TD>
											<TD CLASS="TD51" align =center width = 15%   >
												(8)�ձݻ��Դ��<BR>������(17)
											</TD>
												<TD CLASS="TD51" align =center width = 10% >
												(9)�ձݻ��Թ�����<BR>((7)��(8)�� �����ݾ�)
											</TD>
											
										    <TD CLASS="TD51" align =center width = 15% x>
												(10)ȸ��ձݰ��� 
											</TD>
											<TD CLASS="TD51" align =center width = 15%   >
												(11)�����ݾ�<BR>((9)-(10))
											</TD>
											
											
										</TR>
										
										
										<TR>
											
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW6" name=txtW6 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS="TD61" align =center width = 15%>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW7" name=txtW7 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW8" name=txtW8 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2" width = 100%></OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS="TD61" align =center width = 15%>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW9" name=txtW9 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2" width = 100%></OBJECT>');</SCRIPT>
											</TD>
												<TD CLASS="TD61" align =center width = 15%>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW10" name=txtW10 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2" width = 100%></OBJECT>');</SCRIPT>
											</TD>
											</TD>
											<TD CLASS="TD61" align =center width = 11%>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW11" name=txtW11 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2" width = 100%></OBJECT>');</SCRIPT>
											</TD>
											
											
										</TR>
										
										
						
									</TABLE>
						   </FIELDSET>				
						   			
					</TD>
				</TR>
				
					<TR>
					<TD WIDTH=1024 valign=top >
					   
					      <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>2.�̹� �ձݻ����� ����� ���� ��� </LEGEND>
									<TABLE width = 100% bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">	
										<TR>
											<TD CLASS="TD51" align =center width = 10% >
												(12)�������� <BR>��ġ�ݵ��(21)
											</TD>											
										    <TD CLASS="TD51" align =center width = 15% >
												(13)������������ <BR>���ݵ�����⸻<BR>�Ű����������Ѽձݻ��Ծ� 
											</TD>
											<TD CLASS="TD51" align =center width = 15%   >
												(14)�����������ݵ�ձݺ��δ���� 
											</TD>
												<TD CLASS="TD51" align =center width = 10% >
												(15)�������� ����ݵ�<BR> ���� �� �ؾ�� 
											</TD>
										    <TD CLASS="TD51" align =center width = 15% >
												(16)�̹� �ձݻ����� ������<BR>((13)��(14)��(15))
											</TD>
											<TD CLASS="TD51" align =center width = 15%   >
												(17)�ձݻ��Դ������<BR>((12)��(16))
											</TD>											
										</TR>
										<TR>											
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW12" name=txtW12 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS="TD61" align =center width = 15%>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW13" name=txtW13 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW14" name=txtW14 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2" width = 100%></OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS="TD61" align =center width = 15%>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW15" name=txtW15 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2" width = 100%></OBJECT>');</SCRIPT>
											</TD>
												<TD CLASS="TD61" align =center width = 15%>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW16" name=txtW16 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2" width = 100%></OBJECT>');</SCRIPT>
											</TD>
											</TD>
											<TD CLASS="TD61" align =center width = 15%>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW17" name=txtW17 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2" width = 100%></OBJECT>');</SCRIPT>
											</TD>
										</TR>										
										<TR>
											<TD CLASS="TD51" align =center width = 10% >
												(18)�����������迹ġ�ݵ� 
											</TD>
										    <TD CLASS="TD51" align =center width = 15% COLSPAN=2 >
												(19)�����������迹ġ�ݵ���ɹ��ؾ�� 
											</TD>
											<TD CLASS="TD51" align =center width = 15%  >
												(20)����������迹ġ�ݵ��� ���Ծ� 
											</TD>
										    <TD CLASS="TD51" align =center width = 15% COLSPAN=2 >
												(21)�������迹ġ�ݵ� ��(18-19+20)
											</TD>							
										</TR>
										<TR>											
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW18" name=txtW18 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="(18)�����������迹ġ�ݵ�" tag="21X2Z" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS="TD61" align =center width = 15% COLSPAN=2>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW19" name=txtW19 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="(19)�����������迹ġ�ݵ���ɹ��ؾ��" tag="21X2Z" width = 100% ></OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS="TD61" align =center width = 15% >
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW20" name=txtW20 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="(20)����������迹ġ�ݵ��� ���Ծ�" tag="21X2Z" width = 100%></OBJECT>');</SCRIPT>
											</TD>
											<TD CLASS="TD61" align =center width = 15% COLSPAN=2>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW21" name=txtW21 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="(21)�������迹ġ�ݵ� ��(18-19+20)" tag="24X2" width = 100%></OBJECT>');</SCRIPT>
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
<INPUT TYPE=HIDDEN NAME="txt32w1" tag="24">
<INPUT TYPE=HIDDEN NAME="txt32w2" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">

<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

