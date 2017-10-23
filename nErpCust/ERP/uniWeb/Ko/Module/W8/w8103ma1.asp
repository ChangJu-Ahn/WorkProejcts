<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ���μ� ���� 
'*  3. Program ID           : W8103MA1
'*  4. Program Name         : W8103MA1.asp
'*  5. Program Desc         : ��58ȣ ���μ��߰������Ű��ΰ�꼭 
'*  6. Modified date(First) : 2005/01/28
'*  7. Modified date(Last)  : 2006/01/27
'*  8. Modifier (First)     : �ֿ��� 
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
<STYLE>
	.CHECKBOX {
		BORDER: 0;
	}

</STYLE>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  ���/���� ����  ====================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID		= "W8103MA1"
Const BIZ_PGM_ID		= "W8103mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_REF_PGM_ID	= "W8103mb2.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID		= "W8109OA1"

' -- �׸��� �÷� ���� 

Dim C_W1	' �׸��� ��(����� ��)
Dim C_W2
Dim C_W3
Dim C_W4_1
Dim C_W4_2
Dim C_W5_1
Dim C_W5_2
Dim C_W6_1
Dim C_W6_2
Dim C_W7
Dim C_W8
Dim C_W9
Dim C_W10_1

Dim C_W01	
Dim C_W02	
Dim C_W03	
Dim C_W04	
Dim C_W05	
Dim C_W06	
Dim C_W07	
Dim C_W09	
Dim C_W10	
Dim C_W11	
Dim C_W12	
Dim C_W13	
Dim C_W14	
Dim C_W15	

Dim C_W31	
Dim C_W32	
Dim C_W33	
Dim C_W34	
Dim C_W35	
Dim C_W36	
Dim C_W37	
Dim C_W38	
Dim C_W39	
Dim C_W40	
Dim C_W41	

Dim C_W51	

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 
Dim lgCurrGrid, lgvspdData(2)
Dim	lgFISC_START_DT, lgFISC_END_DT, lgMonGap, lgW2018

Dim IsRunEvents	' �Ф� �����̺�Ʈ�ݺ��� ���� 

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	C_W1	= 0	
	C_W2	= 1
	C_W3	= 2
	C_W4_1	= 3
	C_W4_2	= 4
	C_W5_1	= 5
	C_W5_2	= 6
	C_W6_1	= 7
	C_W6_2	= 8
	C_W7	= 9
	C_W8	= 10
	C_W9	= 11
	C_W10_1	= 12
	
	C_W01	= 13
	C_W02	= 14
	C_W03	= 15
	C_W04	= 16
	C_W05	= 17
	C_W06	= 18
	C_W07	= 19
	C_W09	= 20
	C_W10	= 21
	C_W11	= 22
	C_W12	= 23
	C_W13	= 24
	C_W14	= 25
	C_W15	= 26

	C_W31	= 27
	C_W32	= 28
	C_W33	= 29
	C_W34	= 30
	C_W35	= 31
	C_W36	= 32
	C_W37	= 33
	C_W38	= 34
	C_W39	= 35
	C_W40	= 36
	C_W41	= 37

	C_W51	= 38
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

    IsRunEvents = False
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub



'============================================  �Ű��� �޺� �ڽ� ä���  ====================================

Sub InitComboBox()
	' ��ȸ����(����)
	Dim IntRetCD1
	Call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))

	Call InitSpreadPosVariables

 	call CommonQueryRs("REFERENCE_1"," ufn_TB_Configuration('W2018','" & C_REVISION_YM & "') ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
           
    lgW2018 = Split(lgF0 , chr(11))
    	
	IsRunEvents = True	' OBJECT �� �������� �̺�Ʈ�� �߻��ϴ°��� ���� 
	
	Call CommonQueryRs("MINOR_CD,MINOR_NM"," dbo.ufn_TB_MINOR('W1065', '" & C_REVISION_YM & "') ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtData(C_W1) ,lgF0  ,lgF1  ,Chr(11))
    
 	Call CommonQueryRs("MINOR_CD,MINOR_NM"," dbo.ufn_TB_MINOR('W1064', '" & C_REVISION_YM & "') ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtData(C_W2) ,lgF0  ,lgF1  ,Chr(11))
 
    IsRunEvents = False
    
    
End Sub

Sub InitSpreadSheet()

	Call AppendNumberPlace("6","3","2")
	Call AppendNumberPlace("7","2","0")
	
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitData()

	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
    
	Call GetFISC_DATE

End Sub

Sub InitSpreadComboBox()

End Sub

'============================== ���۷��� �Լ�  ========================================

Function GetRef()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg, arrW1, arrW2, iRow, iMaxRows, sTmp

	' �·ε�� ���۷����޽��� �����´�.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
			
	IntRetCD = CommonQueryRs("W1, W2"," dbo.ufn_TB_58_GetRef('" & sCoCd & "','" & sFiscYear & "','" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD = True Then
		arrW1		= Split(lgF0, chr(11))
		arrW2		= Split(lgF1, chr(11))
		iMaxRows	= UBound(arrW1)

		For iRow = 0 To iMaxRows -1
			sTmp = "frm1.txtData(C_W" & arrW1(iRow) & ").Text = """ & CStr(arrW2(iRow)) & """"	
			Execute sTmp	' -- ������ ��� �ִ� ����� �����Ѵ�.  ** ufn_TB_3_GetRef�� W1�ʵ��� ���� �ڵ�� ���� �ʰ�, �ڵ��� �迭�ε��������� �ϸ� �̷��� ���ص� ��.
		Next

	End If
	
	Call SetHeadReCalc()
End Function

' ��� ���� 
Sub SetHeadReCalc()	
	Dim dblSum, dblW4(100)	

	If IsRunEvents Then Exit Sub	' �Ʒ� .vlaue = ���� �̺�Ʈ�� �߻��� ����Լ��� ���°� ���´�.
	
	IsRunEvents = True
	
	With frm1
		dblW4(C_W01) = UNICDbl(.txtData(C_W01).value)
		dblW4(C_W02) = UNICDbl(.txtData(C_W02).value)
		dblW4(C_W03) = UNICDbl(.txtData(C_W03).value)
		
		dblW4(C_W04) = dblW4(C_W01) - dblW4(C_W02) + dblW4(C_W03)
		.txtData(C_W04).value = dblW4(C_W04)	' (104) Ȯ������ 
		
		dblW4(C_W05) = UNICDbl(.txtData(C_W05).value)
		dblW4(C_W06) = UNICDbl(.txtData(C_W06).value)
		
		dblW4(C_W07) = dblW4(C_W04) - dblW4(C_W05) - dblW4(C_W06)
		.txtData(C_W07).value = dblW4(C_W07)	' (107) �������� 
		
		If dblW4(C_W07) < 0 Then
			dblW4(C_W09) = 0
		Else
			dblW4(C_W09) = dblW4(C_W07) * 6 / lgMonGap 
		End If
		.txtData(C_W09).value = dblW4(C_W09)	' (108) �߰��������� 
			
		dblW4(C_W10) = UNICDbl(.txtData(C_W10).value)
		
		dblW4(C_W11) = dblW4(C_W09) - dblW4(C_W10) 
		.txtData(C_W11).value = dblW4(C_W11)	' (110) �����߰��������� 

		dblW4(C_W12) = UNICDbl(.txtData(C_W12).value)
		
		dblW4(C_W13) = dblW4(C_W11) + dblW4(C_W12) 
		.txtData(C_W13).value = dblW4(C_W13)	' (112) ���꼼�� 
		
		If dblW4(C_W11) <= 10000000 Then
			dblW4(C_W14) = 0
		ElseIf dblW4(C_W11) > 10000000 And dblW4(C_W11) <= 20000000 Then
			dblW4(C_W14) = dblW4(C_W11) - 10000000
		ElseIf dblW4(C_W11) > 20000000 Then
			dblW4(C_W14) = dblW4(C_W11) * 0.5
		End If
		.txtData(C_W14).value = dblW4(C_W14)	' (113) �г�����		
		
		dblW4(C_W15) = dblW4(C_W13) - dblW4(C_W14) 
		.txtData(C_W15).value = dblW4(C_W15)	' (114) ���μ��� 
		
		dblW4(C_W31) = UNICDbl(.txtData(C_W31).value)
		If dblW4(C_W31) < 0 Then
			dblW4(C_W32) = 0
		ElseIf (dblW4(C_W31) * 12 / 6) <= 100000000 Then
			dblW4(C_W32) = lgW2018(0) ' �̻��� 
		Else
			dblW4(C_W32) = lgW2018(1) ' �ʰ����� 
		End If
		.txtData(C_W32).Text = (dblW4(C_W32) * 100) & "%"	' (116) ���� 
		
		If  (dblW4(C_W31) * 12 / 6) > 100000000 Then
			dblW4(C_W33) = ((dblW4(C_W31) * 12 / 6) - 100000000) * (dblW4(C_W32) * 6/12) + ((100000000 * lgW2018(0)) * 6/12)
		ElseIf (dblW4(C_W31) * 12 / 6) <= 100000000 Then
			dblW4(C_W33) = ((dblW4(C_W31) * 12 / 6) * lgW2018(0)) * 12 / 6
		End If
		.txtData(C_W33).Text = dblW4(C_W33)	' (116) ���� 
		
		' 33�ڵ� ����߰� 
		'dblW4(C_W33) = UNICDbl(.txtData(C_W33).value)
		dblW4(C_W34) = UNICDbl(.txtData(C_W34).value)
		dblW4(C_W35) = UNICDbl(.txtData(C_W35).value)
		dblW4(C_W36) = UNICDbl(.txtData(C_W36).value)
		
		dblW4(C_W37) = dblW4(C_W33) - dblW4(C_W34) - dblW4(C_W35) - dblW4(C_W36) 
		.txtData(C_W37).value = dblW4(C_W37)	' (121) �߰��������� 
		
		dblW4(C_W38) = UNICDbl(.txtData(C_W38).value)
		
		dblW4(C_W39) = dblW4(C_W37) + dblW4(C_W38) 
		.txtData(C_W39).value = dblW4(C_W39)	' (123) �����Ҽ��װ� 
		
		If dblW4(C_W39) <= 10000000 Then
			dblW4(C_W40) = 0
		ElseIf dblW4(C_W39) > 10000000 And dblW4(C_W39) <= 20000000 Then
			dblW4(C_W40) = dblW4(C_W39) - 10000000
		ElseIf dblW4(C_W39) > 20000000 Then
			dblW4(C_W40) = dblW4(C_W39) * 0.5
		End If
		.txtData(C_W40).value = dblW4(C_W40)	' (125) ���μ��� 

		dblW4(C_W41) = dblW4(C_W39) - dblW4(C_W40) 
		.txtData(C_W41).value = dblW4(C_W41)	' (125) �����Ҽ��� 
					
	End With

	lgBlnFlgChgValue= True ' ���濩�� 
	IsRunEvents = False	' �̺�Ʈ �߻������� ������ 
End Sub

Sub GetFISC_DATE()	' ������ ��ȸ���ǿ� �����ϴ� ������,�������� �����´�.
	Dim sFiscYear, sRepType, sCoCd, ret, datFISC_START_DT, datFISC_END_DT, iRet, sRepTypeNm
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	sRepTypeNm	= frm1.cboREP_TYPE.options(frm1.cboREP_TYPE.selectedIndex).text
	
	iRet = CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If iRet = False Then
		Call DisplayMsgBox("WC0037", parent.VB_INFORMATION, sFiscYear & "��", sRepTypeNm)	
		frm1.cboREP_TYPE.value = "1"
		Exit Sub
	End If
	
	' ���� �Ⱓ�� �ʼ��Է� 
	lgFISC_START_DT = CDate(lgF0)
	lgFISC_END_DT = CDate(lgF1)

	With frm1
		
	IsRunEvents = True

		.txtData(C_W5_1).Text	= lgFISC_START_DT
		.txtData(C_W5_2).Text	= lgFISC_END_DT
		'.txtData(C_W4).Text		= "6"
	
		' ��������������� ���� 
		sFiscYear	= UNIFormatDate(CDate(frm1.txtFISC_YEAR.text)-1)
	
		ret = CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		If ret Then
			datFISC_START_DT = CDate(lgF0)
			datFISC_END_DT = CDate(lgF1)
			lgMonGap = DateDiff("m", datFISC_START_DT, datFISC_END_DT)+1
		Else
			lgMonGap = 12
		End If

		.txtData(C_W7).Text	= lgMonGap
		
	IsRunEvents = False
	
	End With
	
End Sub

'====================================== �� �Լ� =========================================

'============================================  ��ȸ���� �Լ�  ====================================


'============================================  �� �Լ�  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         

    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>
	  
	' �����Ѱ� 
	Call InitComboBox
	
	IsRunEvents = True
	
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
	'Call ggoOper.FormatDate(frm1.txtW2 , parent.gDateFormat,3)
	
	Call InitData 
	'
     
    If frm1.cboREP_TYPE.value <> "2" Then
		Call DisplayMsgBox("W80004", parent.VB_INFORMATION, "", "X")  
		Call SetToolbar("1100000000000111")	
    End If
    
    IsRunEvents = False
	
	Call FncQuery
	   
    'Call ChangeCombo(2,False)
End Sub


'============================================  �̺�Ʈ �Լ�  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub

Sub cboREP_TYPE_onChange()	' �Ű������ �ٲٸ�..
	Call GetFISC_DATE
End Sub

Sub ChkW3(Byval Index)	
	Dim i
	With frm1
		IsRunEvents = True

		For i = 0 To 2
			.txtW3(i).checked = False
		Next
		.txtW3(Index).checked = True
		.txtData(C_W3).value  = Index+1
		
		IsRunEvents = False
		lgBlnFlgChgValue = True
	End With
End Sub

Sub ChkW4_1(Byval Index)
	Dim i
	With frm1
		IsRunEvents = True
		For i = 0 To 4
			.txtW4_1(i).checked = False
		Next
		.txtW4_1(Index).checked = True
		.txtData(C_W4_1).value  = Index+1	
		
		IsRunEvents = False
		lgBlnFlgChgValue = True
	End With
End Sub

Sub ChkW10_1(Byval Index)
	Dim i
	With frm1
		IsRunEvents = True

		For i = 0 To 1
			.txtW10_1(i).checked = False
		Next
		.txtW10_1(Index).checked = True
		.txtData(C_W10_1).value  = Index+1
						
		IsRunEvents = False
		lgBlnFlgChgValue = True
	End With
End Sub

Sub ChangeCombo(strWhere,strTag)
	lgBlnFlgChgValue = strTag
	Dim i
	
	with frm1
		select case strWhere
			Case "2"
		
				If .txtData(1).value=1 then 
				
					For i = 13 To 26
						Select Case i
							Case  13,14,15,17,18,20,21,23,25
								Call ggoOper.setreqAttr(.txtData(i),"D") 'N,R								

						End Select
						
					Next
					For i=27 to 37
						Select Case i
							Case 27,30,31,32,34,36
								Call ggoOper.setreqAttr(.txtData(i),"Q")
						End Select
						.txtData(i).Text=""
					Next
				ElseIf .txtData(1).value=2 Then
				
					For i = 13 To 26
						Select Case i
							Case  13,14,15,17,18,20,21,23,25
								Call ggoOper.setreqAttr(.txtData(i),"Q") 'N,R
						End Select		
						.txtData(i).Text=""		
					Next
					For i=27 to 37
						Select Case i
							Case 27,30,31,32,34,36
								Call ggoOper.setreqAttr(.txtData(i),"D") 'N,R

						End Select 
					Next			
				End If
		End Select
	End with
	
End Sub

'============================================  �׸��� �̺�Ʈ   ====================================
'============================================  �������� �Լ�  ====================================

Function FncQuery() 
    Dim IntRetCD , i, blnChange
    
    FncQuery = False                                                        
    blnChange = False
    
    Err.Clear                                                               <%'Protect system from crashing%>

	
<%  '-----------------------
    'Check previous data area
    '----------------------- %>  
    If lgBlnFlgChgValue Or blnChange Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call InitVariables													<%'Initializes local global variables%>
    Call InitData                              
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
	
     
    CALL DBQuery()
    
End Function

Function FncSave() 
    Dim blnChange, i, sMsg
    
    blnChange = False
    
    FncSave = False                                                         
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
    'On Error Resume Next                                                   <%'��: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    If lgBlnFlgChgValue = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If
	    
	If Not chkField(Document, "2") Then                             '��: Check contents area
	   Exit Function
	End If

	With frm1
		If Not .txtW3(0).checked And Not .txtW3(1).checked And Not .txtW3(2).checked Then
			Call DisplayMsgBox("WC0030", "X", "(3) ���α���", "X")                          <%'No data changed!!%>
			Exit Function
		End If
		If Not .txtW4_1(0).checked And Not .txtW4_1(1).checked And Not .txtW4_1(2).checked And Not .txtW4_1(3).checked And Not .txtW4_1(4).checked Then
			Call DisplayMsgBox("WC0030", "X", "(4) ������ ����", "X")                          <%'No data changed!!%>
			Exit Function
		End If
		If Not .txtW10_1(0).checked And Not .txtW10_1(1).checked Then
			Call DisplayMsgBox("WC0030", "X", "(7) �Ű��� ���", "X")                          <%'No data changed!!%>
			Exit Function
		End If
	End With
	
    If Verification = False Then Exit Function
    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '��: Save db data
    
    FncSave = True                                                          
    
End Function

' ----------------------  ���� -------------------------
Function  Verification()

	Verification = False

	
	Verification = True	
End Function

'========================================================================================
Function FncNew() 
    Dim IntRetCD 

    FncNew = False

  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    IsRunEvents = True
    
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")
    Call InitVariables
    Call InitData

    Call SetToolbar("1100100000000111")

	frm1.txtCO_CD.focus

	IsRunEvents = False
    FncNew = True

End Function


Function FncCopy() 
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCopy = False                                                               '��: Processing is NG

 	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 

End Function

Function FncInsertRow(ByVal pvRowCnt) 

End Function

Function FncDeleteRow() 

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
	
    If lgBlnFlgChgValue Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
Function FncDelete() 
    Dim IntRetCD

    FncDelete = False
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF

    Call DbDelete

    FncDelete = True
End Function

'============================================  DB �＼�� �Լ�  ====================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key   
        'strVal = strVal     & "&txtMaxRows="         & lgvspdData(lgCurrGrid).MaxRows            
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '��:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	lgBlnFlgChgValue = False
	
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = parent.OPMD_UMODE
	
	With frm1
		.txtW3(UNICDbl(.txtData(C_W3).value)-1).checked = True
		.txtW4_1(UNICDbl(.txtData(C_W4_1).value)-1).checked = True
		.txtW10_1(UNICDbl(.txtData(C_W10_1).value)-1).checked = True
	End With
	
	' �������� ���� : ���ߵǸ� ���ȴ�.
	Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	'1 ����üũ : �׸��� �� 
	If wgConfirmFlg = "N" Then
		Call SetToolbar("1101100000000111")			
	Else
		
		'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
		Call SetToolbar("1100000000000111")										<%'��ư ���� ���� %>
	End If
	Call ChangeCombo(2,False)
	'lgvspdData(lgCurrGrid).focus			
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow, lCol, lMaxRows, lMaxCols , i    
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
	With frm1
	
		For i = C_W1 To C_W51	
			Select Case i
				Case C_W4_1
					strVal = strVal & .txtData(i).Value & Parent.gColSep& .txtData(i).Value  & Parent.gColSep	
				Case C_W4_2						
				Case C_W5_1, C_W5_2, C_W6_1, C_W6_2, C_W8, C_W32
					strVal = strVal & .txtData(i).Text & Parent.gColSep				
				Case Else
					strVal = strVal & .txtData(i).Value & Parent.gColSep
			End Select
		Next 

	End With

	Frm1.txtSpread.value      =  strVal
	Frm1.txtMode.value        =  Parent.UID_M0002
	frm1.txtHeadMode.value	  =  lgIntFlgMode
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		
	Dim iRow											        <%' ���� ������ ���� ���� %>
	Call InitVariables
	
    Call MainQuery()
End Function

'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
    strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key            
	
	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function


'========================================================================================
Function DbDeleteOk()
	Call FncNew()
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
<SCRIPT LANGUAGE=javascript FOR=txtData EVENT=Change>
<!--
    SetHeadReCalc();
//-->
</SCRIPT>
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
		<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:GetRef()">�ݾ׺ҷ�����</A></TD>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w8103ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
									<TD CLASS="TD5">���θ�</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
									</TD>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">�Ű���</TD>
									<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE" ALT="�Ű���" STYLE="WIDTH: 50%" tag="14X"></SELECT>
									</TD>
									<TD CLASS="TD5"></TD>
									<TD CLASS="TD6"></TD>
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
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto"><% ' -- overflow=auto : ������ ������ ������ ũ�⿡ ���� ��ũ�ѹٰ� �����ǰ� �Ѵ� %>
						<TABLE <%=LR_SPACE_TYPE_60%> border="0" height=100% width="100%">
						   <TR>
								<TD>1. �߰� ���� �Ⱓ</TD>
						   </TR>
						   <TR>
								<TD width="100%">
								<TABLE <%=LR_SPACE_TYPE_60%> border="1" height=100% width="100%">
									 <TR>
										   <TD CLASS="TD51" width=18%>(1) ���Ằ ����</TD>
										   <TD CLASS="TD61" COLSPAN=2 width=20%><SELECT id="txtData" name=txtData STYLE="WIDTH: 200"  tag="23" onChange="vbscript:ChangeCombo 1,True"></SELECT></TD>
										   <TD CLASS="TD51" COLSPAN=2 width=20%>(2) ���װ�����</TD>
										   <TD CLASS="TD61" COLSPAN=2 width=42%><SELECT id="txtData" name=txtData STYLE="WIDTH: 200"  tag="23" onChange="vbscript:ChangeCombo 2,True"></SELECT></TD>
									 </TR>
									 <TR>
										   <TD CLASS="TD51" ROWSPAN=2>(3) ���α���</TD>
										   <TD CLASS="TD61" ROWSPAN=2 COLSPAN=2><INPUT TYPE=HIDDEN ID="txtData" name="txtData">
										   <TABLE>
											<TR>
												<TD ALIGN=CENTER><INPUT TYPE=CHECKBOX CLASS="CHECKBOX" NAME=txtW3 VALUE="1" ID=txtW3_1 BORDER=0 onclick="vbscript:ChkW3(0)"><LABEL FOR="txtW3_1">1. ����</LABEL></TD>
												<TD ALIGN=CENTER><INPUT TYPE=CHECKBOX CLASS="CHECKBOX" NAME=txtW3 VALUE="2" ID=txtW3_2 BORDER=0 onclick="vbscript:ChkW3(1)"><LABEL FOR="txtW3_2">2. �ܱ�</LABEL></TD>
												<TD ALIGN=CENTER><INPUT TYPE=CHECKBOX CLASS="CHECKBOX" NAME=txtW3 VALUE="3" ID=txtW3_3 BORDER=0 onclick="vbscript:ChkW3(2)"><LABEL FOR="txtW3_3">3. ����</LABEL></TD>
											</TR>
										   </TABLE>
										   </TD>
										   <TD CLASS="TD51" ROWSPAN=2 COLSPAN=2>(4) ������ ����</TD>
										   <TD CLASS="TD61" ALIGN=CENTER width=15%>��������</TD>
										   <TD CLASS="TD61" ALIGN=CENTER width=27%>�񿵸�����</TD>
									</TR>
									<TR>
										   <TD CLASS="TD61"><INPUT TYPE=HIDDEN ID="txtData" name="txtData">
										   <TABLE <%=LR_SPACE_TYPE_20%>>
											<TR>
												<TD ALIGN=CENTER><INPUT TYPE=CHECKBOX CLASS="CHECKBOX" NAME=txtW4_1 VALUE="1" ID=txtW4_11 BORDER=0 onclick="vbscript:ChkW4_1(0)"><LABEL FOR="txtW4_11">1. �߼�</LABEL></TD>
												<TD ALIGN=CENTER><INPUT TYPE=CHECKBOX CLASS="CHECKBOX" NAME=txtW4_1 VALUE="2" ID=txtW4_12 BORDER=0 onclick="vbscript:ChkW4_1(1)"><LABEL FOR="txtW4_12">2. �Ϲ�</LABEL></TD>
											</TR>
										   </TABLE>
										   </TD>
										   <TD CLASS="TD61"><INPUT TYPE=HIDDEN ID="txtData" name="txtData">
										   <TABLE <%=LR_SPACE_TYPE_20%>>
											<TR>
												<TD ALIGN=CENTER><INPUT TYPE=CHECKBOX CLASS="CHECKBOX" NAME=txtW4_1 VALUE="3" ID=txtW4_13 BORDER=0 onclick="vbscript:ChkW4_1(2)"><LABEL FOR="txtW4_13">3. ��������</LABEL></TD>
												<TD ALIGN=CENTER><INPUT TYPE=CHECKBOX CLASS="CHECKBOX" NAME=txtW4_1 VALUE="4" ID=txtW4_14 BORDER=0 onclick="vbscript:ChkW4_1(3)"><LABEL FOR="txtW4_14">4. �߼�</LABEL></TD>
												<TD ALIGN=CENTER><INPUT TYPE=CHECKBOX CLASS="CHECKBOX" NAME=txtW4_1 VALUE="5" ID=txtW4_15 BORDER=0 onclick="vbscript:ChkW4_1(4)"><LABEL FOR="txtW4_15">5. �Ϲ�</LABEL></TD>
											</TR>
										   </TABLE>
										   </TD>
									 </TR>
									 <TR>
										   <TD CLASS="TD51">(5) �������</TD>
										   <TD CLASS="TD61" COLSPAN=2><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script> ~ <script language =javascript src='./js/w8103ma1_txtData_N515311688.js'></script></TD>
										   <TD CLASS="TD51" COLSPAN=2>(6) �����Ⱓ</TD>
										   <TD CLASS="TD61" COLSPAN=2><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script> ~ <script language =javascript src='./js/w8103ma1_txtData_N287269365.js'></script></TD>
									 </TR>
									 <TR>
										   <TD CLASS="TD51" width=18%>(7) ���������������</TD>
										   <TD CLASS="TD61" width=10%><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script> ����</TD>
										   <TD CLASS="TD51" COLSPAN=2 width=15%>(8) �Ű���</TD>
										   <TD CLASS="TD61" width=15%><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
										   <TD CLASS="TD51" width=20%>(9) ���Աݾ�</TD>
										   <TD CLASS="TD61" width=42%><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									 </TR>
									 <TR>
										   <TD CLASS="TD51">(10) �Ű��� ���</TD>
										   <TD CLASS="TD61" COLSPAN=6><INPUT TYPE=HIDDEN ID="txtData" name="txtData">
										   <TABLE <%=LR_SPACE_TYPE_20%> >
											<TR>
												<TD ALIGN=CENTER><INPUT TYPE=CHECKBOX CLASS="CHECKBOX" NAME=txtW10_1 VALUE="1" ID=txtW10_11 BORDER=0 onclick="vbscript:ChkW10_1(0)"><LABEL FOR="txtW10_11">1. �� �� �� ��</LABEL></TD>
												<TD ALIGN=CENTER><INPUT TYPE=CHECKBOX CLASS="CHECKBOX" NAME=txtW10_1 VALUE="2" ID=txtW10_12 BORDER=0 onclick="vbscript:ChkW10_1(1)"><LABEL FOR="txtW10_12">2. �� �� �� �� ��</LABEL></TD>
											</TR>
										   </TABLE>
										   </TD>
									 </TR>
								</TABLE>
								</TD>
						   </TR>
						   <TR>
								<TD>2. �Ű� �� ���� ���� ���</TD>
						   </TR>
						   <TR>
								<TD width="100%">
									<TABLE <%=LR_SPACE_TYPE_60%> border="1" height=100% width="100%">
									 <TR>
										   <TD CLASS="TD61" COLSPAN=3 ALIGN=CENTER>�� ��</TD>
										   <TD CLASS="TD61" COLSPAN=2 ALIGN=CENTER>�� �� ��</TD>
									 </TR>
									 <TR>
										   <TD CLASS="TD51" width="5%" ROWSPAN=14 ALIGN=CENTER>
										   <TABLE <%=LR_SPACE_TYPE_20%>>
											<TR>
												<TD ALIGN=CENTER>(1)<br><br>��<br>��<br>��<br>��<br>��<br>��<br>��<br>��<br>��<br>��<br>��</TD>
												<TD ALIGN=CENTER>��<br><br>��<br>6<br>3<br>ȣ<br><br>��<br>1<br>��</TD>
											</TR>
										   </TABLE></TD>
										   <TD CLASS="TD51" width="5%" ROWSPAN=7 ALIGN=CENTER>
										   <TABLE <%=LR_SPACE_TYPE_20%>>
											<TR>
												<TD ALIGN=CENTER>��<br>��<br>��<br>��<br>��<br>��</TD>
												<TD ALIGN=CENTER>��<br><br><br>��<br><br><br><br>��</TD>
											</TR>
										   </TABLE></TD>
										   <TD CLASS="TD51" width="45%">(101) �� �� �� ��</TD>
										   <TD CLASS="TD61" width="5%" ALIGN=CENTER>01</TD>
										   <TD CLASS="TD61" width="30%"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
										   <TD CLASS="TD51">(102) �� �� �� �� �� ��</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>02</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>										   
										   <TD CLASS="TD51">(103) �� �� �� ��</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>03</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51">(104) Ȯ �� �� �� [(101) - (102) + (103)]</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>04</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51">(105) �� �� �� �� �� ��</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>05</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>									  
									<TR>
									       <TD CLASS="TD51">(106) �� õ �� �� �� ��</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>06</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>	
									<TR>
									       <TD CLASS="TD51">(107) �� �� �� �� [(104) - (105) - (106)]</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>07</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>
									       <TABLE <%=LR_SPACE_TYPE_20%> border="0">
												<TR>
													<TD ALIGN=CENTER WIDTH=30%>(108) �߰���������</TD>
													<TD ALIGN=CETER WIDTH=50%>
													<TABLE <%=LR_SPACE_TYPE_20%> border="0">
														<TR>
															<TD ROWSPAN=3 WIDTH=30% ALIGN=RIGHT>[(107) X&nbsp;&nbsp;</TD>
															<TD ALIGN=CENTER>6</TD>
															<TD ROWSPAN=3 WIDTH=10%>]</TD>
														</TR>
														<TR>
															<TD HEIGHT=1 BGCOLOR=BLACK></TD>
														</TR>
														<TR>
															<TD ALIGN=CENTER>���������������</TD>
														</TR>
													</TABLE>	
													</TD>
													<TD WIDTH=20%>&nbsp;</TD>
												</TR>
											</TABLE></TD>															
										   <TD CLASS="TD61" ALIGN=CENTER>09</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(109) �� �� �� �� �� �� �� ��</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>10</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(110) �� �� �� �� �� �� �� �� [(108) - (109)]</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>11</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>											
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(111) �� �� �� �� </TD>
										   <TD CLASS="TD61" ALIGN=CENTER>12</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>											
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(112) �� �� �� �� �� �� [(110) + (111)]</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>13</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>											
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(113) �� �� �� ��</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>14</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>											
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(114) �� �� �� �� [(112) - (113)]</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>15</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>											
									<TR>
									     <TD HEIGHT=5></TD>
									</TR>
									 <TR>
										   <TD CLASS="TD61" COLSPAN=3 ALIGN=CENTER>�� ��</TD>
										   <TD CLASS="TD61" COLSPAN=2 ALIGN=CENTER>�� �� ��</TD>
									 </TR>
									<TR>
										   <TD CLASS="TD51" width="5%" ROWSPAN=11 ALIGN=CENTER>
										   <TABLE <%=LR_SPACE_TYPE_60%>>
											<TR>
												<TD ALIGN=CENTER>(2)<br><br>��<br>��<br>��<br>��<br>��<br>��</TD>
												<TD ALIGN=CENTER>��<br><br>��<br>6<br>3<br>ȣ<br><br>��<br>4<br>��</TD>
											</TR>
										   </TABLE></TD>
									       <TD CLASS="TD51" COLSPAN=2>(115) �� �� ǥ ��</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>31</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>												 											 
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(116)�� ��</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>32</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>											 
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(117) �� �� �� ��</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>33</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(118) �� �� �� �� �� ��</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>34</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(119) �� �� �� �� �� ��</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>35</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(120) �� õ �� �� �� ��</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>36</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(121) �� �� �� �� �� �� [(117) - (118) - (119) - (120)]</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>37</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(122) �� �� �� ��</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>38</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(123) �� �� �� �� �� �� [(121) + (122)]</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>39</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(124)�� �� �� ��</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>40</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=2>(125) �� �� �� �� [(123) - (124)]</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>41</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
									<TR>
									     <TD HEIGHT=5></TD>
									</TR>
									<TR>
									       <TD CLASS="TD51" COLSPAN=3>(126) �� �� �� �� �� [(114) �Ǵ� (125)�� ������ ����]</TD>
										   <TD CLASS="TD61" ALIGN=CENTER>51</TD>
										   <TD CLASS="TD61"><script language =javascript src='./js/w8103ma1_txtData_txtData.js'></script></TD>
									</TR>
											 
									</TABLE>
								</TD>
							</TR>
							
						  </TABLE>
						</DIV>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" style="display:'none'"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHeadMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtW14" tag="24"><INPUT TYPE=HIDDEN NAME="txtW35" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

