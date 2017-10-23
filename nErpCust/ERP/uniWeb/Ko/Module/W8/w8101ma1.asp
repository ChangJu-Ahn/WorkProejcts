<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ���μ� ���� 
'*  3. Program ID           : W8101MA1
'*  4. Program Name         : W8101MA1.asp
'*  5. Program Desc         : ��3ȣ ���μ�����ǥ�� �� ����������꼭 
'*  6. Modified date(First) : 2005/01/27
'*  7. Modified date(Last)  : 2005/01/27
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
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
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  ���/���� ����  ====================================
<!-- #Include file="../../inc/lgvariables.inc" -->



Const BIZ_MNU_ID		= "W8101MA1"
Const BIZ_PGM_ID		= "W8101mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_REF_PGM_ID	= "W8101mb2.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_POP_ID		= "W8101MA2.asp"
Const EBR_RPT_ID		 = "W8107OA1"

Const TYPE_1	= 0		' �׸��� �迭��ȣ �� ����� W_TYPE �÷��� ��. 
Const TYPE_2	= 1		' �� ��Ƽ �׸��� PG������ ���� ���̺��� �ڵ�� �����ȴ�.

' -- �׸��� �÷� ���� 
Dim C_W1	' �׸��� �� 
Dim C_W2
Dim C_W2_1
Dim C_W2_2
Dim C_W3
Dim C_W4

Dim C_W01	' �׸��� ��(����� ��)
Dim C_W02	
Dim C_W03	
Dim C_W04	
Dim C_W05	
Dim C_W54	
Dim C_W06	
Dim C_W06_1
Dim C_W07	
Dim C_W08	
Dim C_W09	
Dim C_W10	
Dim C_W10_1
Dim C_W11	
Dim C_W12	
Dim C_W13	
Dim C_W14	
Dim C_W15	
Dim C_W16	
Dim C_W16_1
Dim C_W17	
Dim C_W18	
Dim C_W19	
Dim C_W20	
Dim C_W21	
Dim C_W22	
Dim C_W23	
Dim C_W24	
Dim C_W25	
Dim C_W26	
Dim C_W27	
Dim C_W28	
Dim C_W29	
Dim C_W30	
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
Dim C_W42	
Dim C_W43	
Dim C_W44	
Dim C_W45	
Dim C_W46	
Dim C_W55	
Dim C_W47	
Dim C_W48	
Dim C_W49	
Dim C_W50	
Dim C_W51	
Dim C_W52	
Dim C_W53	

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 
Dim lgCurrGrid, lgvspdData(2)
Dim	lgFISC_START_DT, lgFISC_END_DT, lgW2018

Dim IsRunEvents	' �Ф� �����̺�Ʈ�ݺ��� ���� 

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	' -- �׸��� 1
	C_W1	= 1
	C_W2	= 2
	C_W2_1	= 3
	C_W2_2	= 4
	C_W3	= 5
	C_W4	= 6
	
	C_W01	= 0
	C_W02	= 1
	C_W03	= 2
	C_W04	= 3
	C_W05	= 4
	C_W54	= 5
	C_W06	= 6
	
	C_W06_1 = 7
	C_W07	= 8
	C_W08	= 9
	C_W09	= 10
	C_W10	= 11
	
	C_W10_1	= 12
	C_W11	= 13
	C_W12	= 14
	C_W13	= 15
	C_W14	= 16
	C_W15	= 17
	C_W16	= 18
	
	C_W16_1	= 19
	C_W17	= 20
	C_W18	= 21
	C_W19	= 22
	C_W20	= 23
	C_W21	= 24
	C_W22	= 25
	C_W23	= 26
	C_W24	= 27
	C_W25	= 28
	C_W26	= 29
	C_W27	= 30
	C_W28	= 31
	
	' -- �׸���2
	C_W29	= 32
	C_W30	= 33
	C_W31	= 34
	C_W32	= 35
	C_W33	= 36
	C_W34	= 37
	C_W35	= 38
	C_W36	= 39
	C_W37	= 40
	C_W38	= 41
	C_W39	= 42
	C_W40	= 43
	C_W41	= 44
	C_W42	= 45
	C_W43	= 46
	C_W44	= 47
	C_W45	= 48
	C_W46	= 49
	C_W55	= 50 ' <-- 2003.03.07 �����߰� 
	C_W47	= 51
	C_W48	= 52
	C_W49	= 53
	C_W50	= 54
	C_W51	= 55
	C_W52	= 56
	C_W53	= 57
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

    lgCurrGrid = TYPE_1
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
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))

	Call InitSpreadPosVariables

End Sub

Sub InitSpreadSheet()

	Call AppendNumberPlace("6","3","2")
	Call AppendNumberPlace("8","15","0")	' -- �ݾ� 15�ڸ� ���� : ���ϰ˻���ġ 
	
End Sub

Sub InitComboBox2()
	
	call CommonQueryRs("MINOR_CD, MINOR_NM + ' ('+ REFERENCE_2 + ')', REFERENCE_1, REFERENCE_2"," ufn_TB_Configuration('W2014','" & C_REVISION_YM & "') ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetComboX(frm1.txtData(C_W14) , lgF0, lgF1, lgF2, lgF3, Chr(11))
 
 	call CommonQueryRs("MINOR_CD, MINOR_NM+ ' ('+ REFERENCE_2 + ')', REFERENCE_1, REFERENCE_2"," ufn_TB_Configuration('W2013','" & C_REVISION_YM & "') ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetComboX(frm1.txtData(C_W35) , lgF0, lgF1, lgF2, lgF3, Chr(11))
    
 	call CommonQueryRs("REFERENCE_1"," ufn_TB_Configuration('W2018','" & C_REVISION_YM & "') ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
           
    lgW2018 = Split(lgF0 , chr(11))
    
End Sub

'============================================  �׸��� �Լ�  ====================================

Sub InitData()

	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
       
	Call GetFISC_DATE

	Call InitComboBox2()
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
			
	IntRetCD = CommonQueryRs("W1, W2"," dbo.ufn_TB_3_GetRef_" & C_REVISION_YM & "('" & sCoCd & "','" & sFiscYear & "','" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD = True Then
		arrW1		= Split(lgF0, chr(11))
		arrW2		= Split(lgF1, chr(11))
		iMaxRows	= UBound(arrW1)

		For iRow = 0 To iMaxRows -1
			sTmp = "frm1.txtData(C_W" & arrW1(iRow) & ").Text = """ & CStr(arrW2(iRow)) & """"	
			Execute sTmp	' -- ������ ��� �ִ� ����� �����Ѵ�.  ** ufn_TB_3_GetRef�� W1�ʵ��� ���� �ڵ�� ���� �ʰ�, �ڵ��� �迭�ε��������� �ϸ� �̷��� ���ص� ��.
		Next
	Else
		Call DisplayMsgBox("900014", parent.VB_INFORMATION, "", "")             '�� : No data is found.
	End If
	
	Call SetHeadReCalc()
End Function

' ��� ���� 
Sub SetHeadReCalc()	
	Dim dblSum, dblW4(100), dblMonGap
	
	If IsRunEvents Then Exit Sub	' �Ʒ� .vlaue = ���� �̺�Ʈ�� �߻��� ����Լ��� ���°� ���´�.
	
	IsRunEvents = True
	
	With frm1
		dblW4(C_W01) = UNICDbl(.txtData(C_W01).value)
		dblW4(C_W02) = UNICDbl(.txtData(C_W02).value)
		dblW4(C_W03) = UNICDbl(.txtData(C_W03).value)
		
		dblW4(C_W04) = dblW4(C_W01) + dblW4(C_W02) - dblW4(C_W03)
		.txtData(C_W04).value = dblW4(C_W04)	' (104) �������ҵ�ݾ� 
		
		dblW4(C_W05) = UNICDbl(.txtData(C_W05).value)
		dblW4(C_W54) = UNICDbl(.txtData(C_W54).value)
		
		dblW4(C_W06) = dblW4(C_W04) + dblW4(C_W05) - dblW4(C_W54)
		.txtData(C_W06).value = dblW4(C_W06)	' (107) ����������ҵ�ݾ� 
		.txtData(C_W06_1).value = dblW4(C_W06)
			
		dblW4(C_W07) = UNICDbl(.txtData(C_W07).value)
		dblW4(C_W08) = UNICDbl(.txtData(C_W08).value)
		dblW4(C_W09) = UNICDbl(.txtData(C_W09).value)		
		
		dblW4(C_W10) = dblW4(C_W06) - dblW4(C_W07) - dblW4(C_W08) - dblW4(C_W09)
		.txtData(C_W10).value = dblW4(C_W10)	' (112) ����ǥ�� 
		
		' -- 2006-01-04: 200603 ������ 
		dblW4(C_W10_1) = dblW4(C_W10) + UNICDbl(.txtW55_1.value)
		.txtData(C_W10_1).value = dblW4(C_W10_1)

		' -- 114-115�� �����: �������ظ��� 
		If frm1.cboREP_TYPE.value = "2" Then
			dblMonGap = 6
		Else
			dblMonGap = DateDiff("m", lgFISC_START_DT, lgFISC_END_DT)+1
		End If
		
		' -- 2006-01-04: 200603 ������ 
		'dblSum = dblW4(C_W10) * 12 / dblMonGap
		dblSum = dblW4(C_W10_1) * 12 / dblMonGap

		If dblSum <= 100000000 Then
			.txtData(C_W11).value = lgW2018(0) ' 1������ 
		Else
			.txtData(C_W11).value = lgW2018(1)	'1���ʰ� 
		End If

		If dblSum <= 0 Then
			.txtData(C_W12).value = 0 ' 1������ 
		Else
			If dblSum <= 100000000 Then
				.txtData(C_W12).value = (dblSum * lgW2018(0) * dblMonGap) / 12
			Else
				.txtData(C_W12).value = ((dblSum * lgW2018(1) * dblMonGap) / 12) - ( ( 100000000 * (lgW2018(1) - lgW2018(0)) * dblMonGap) / 12)
			End If
		End If
		'zzzzzzzzzzzzzz
			.txtData(C_W12).value = "782839149"	
		dblW4(C_W11) = UNICDbl(.txtData(C_W11).value)
		dblW4(C_W12) = UNICDbl(.txtData(C_W12).value)
		
		dblW4(C_W13) = UNICDbl(.txtData(C_W13).value)	
		If 	.txtData(C_W14).value <> "" Then
			dblW4(C_W14) = UNICDbl(.txtData(C_W14).options(.txtData(C_W14).selectedIndex).VAL)		' -- ���̳��� �޺���: B_Configuration.SeqNo=1
			.txtW14.value = .txtData(C_W14).options(.txtData(C_W14).selectedIndex).VIEW				' -- ���̳��� �޺���: B_Configuration.SeqNo=2
		Else
			dblW4(C_W14) = 0
		End If
		
		dblW4(C_W15) = dblW4(C_W13) * dblW4(C_W14)
		.txtData(C_W15).value = dblW4(C_W15)	' (118) ���⼼�� 

		dblW4(C_W16) = dblW4(C_W12) + dblW4(C_W15)	
		.txtData(C_W16).value = dblW4(C_W16)	' (119) �հ� 
		.txtData(C_W16_1).value = dblW4(C_W16)
		
		dblW4(C_W17) = UNICDbl(.txtData(C_W17).value)
		dblW4(C_W18) = dblW4(C_W16) - dblW4(C_W17)	
		.txtData(C_W18).value = dblW4(C_W18)	' (122) �������� 
		
		dblW4(C_W19) = UNICDbl(.txtData(C_W19).value)
		dblW4(C_W20) = UNICDbl(.txtData(C_W20).value)

		dblW4(C_W21) = dblW4(C_W18) - dblW4(C_W19) + dblW4(C_W20)
		.txtData(C_W21).value = dblW4(C_W21)	' (125) ������ 
		
		dblW4(C_W22) = UNICDbl(.txtData(C_W22).value)
		dblW4(C_W23) = UNICDbl(.txtData(C_W23).value)
		dblW4(C_W24) = UNICDbl(.txtData(C_W24).value)		
		dblW4(C_W25) = UNICDbl(.txtData(C_W25).value)		
		
		dblW4(C_W26) = dblW4(C_W22) + dblW4(C_W23) + dblW4(C_W24) + dblW4(C_W25)
		.txtData(C_W26).value = dblW4(C_W26)	' (130) �Ұ� 
		
		dblW4(C_W27) = UNICDbl(.txtData(C_W27).value)	
		dblW4(C_W28) = dblW4(C_W26) + dblW4(C_W27)
		.txtData(C_W28).value = dblW4(C_W28)	' (132) �հ� 
		
		dblW4(C_W29) = UNICDbl(.txtData(C_W29).value)	
		dblW4(C_W30) = dblW4(C_W21) - dblW4(C_W28) + dblW4(C_W29)
		.txtData(C_W30).value = dblW4(C_W30)	' (134) ���������Ҽ���	
		
		dblW4(C_W31) = UNICDbl(.txtData(C_W31).value)
		dblW4(C_W32) = UNICDbl(.txtData(C_W32).value)
		dblW4(C_W33) = UNICDbl(.txtData(C_W33).value)	
		dblW4(C_W34) = dblW4(C_W31) + dblW4(C_W32) - dblW4(C_W33)
		.txtData(C_W34).value = dblW4(C_W34)	' (138) ����ǥ�� 
		
		If .txtData(C_W35).value <> "" Then	
			dblW4(C_W35) = UNICDbl(.txtData(C_W35).options(.txtData(C_W35).selectedIndex).VAL)		' -- ���̳��� �޺���: B_Configuration.SeqNo=1	
			'dblW4(C_W36) = dblW4(C_W34) * dblW4(C_W35)
			.txtW35.value = .txtData(C_W35).options(.txtData(C_W35).selectedIndex).VIEW				' -- ���̳��� �޺���: B_Configuration.SeqNo=2
		Else
			'dblW4(C_W36) = 0
		End If
		
		'If pMode <> 1 Then	' -- �ڵ����İ�꿡�� ���� 
		'	.txtData(C_W36).value = dblW4(C_W36)	' (140) ���⼼�� 
		'End If
		dblW4(C_W36) = UNICDbl(.txtData(C_W36).value)
		
		dblW4(C_W37) = UNICDbl(.txtData(C_W37).value)	
		dblW4(C_W38) = dblW4(C_W36) - dblW4(C_W37)
		.txtData(C_W38).value = dblW4(C_W38)	' (142) �������� 
		
		dblW4(C_W39) = UNICDbl(.txtData(C_W39).value)
		dblW4(C_W40) = UNICDbl(.txtData(C_W40).value)	
		dblW4(C_W41) = dblW4(C_W38) - dblW4(C_W39) + dblW4(C_W40)
		.txtData(C_W41).value = dblW4(C_W41)	' (145) ������ 
			
		dblW4(C_W42) = UNICDbl(.txtData(C_W42).value)
		dblW4(C_W43) = UNICDbl(.txtData(C_W43).value)	
		dblW4(C_W44) = dblW4(C_W42) + dblW4(C_W43)
		.txtData(C_W44).value = dblW4(C_W44)	' (148) �� 
		
		dblW4(C_W45) = dblW4(C_W41) - dblW4(C_W44)
		.txtData(C_W45).value = dblW4(C_W45)	' (149) ���������Ҽ��� 
		
		dblW4(C_W46) = dblW4(C_W30) + dblW4(C_W45)
		.txtData(C_W46).value = dblW4(C_W46)	' (150) ���������Ҽ��װ� 
		
		dblW4(C_W55) = UNICDbl(.txtData(C_W55).value)
		
		dblW4(C_W47) = dblW4(C_W46) - dblW4(C_W20) - dblW4(C_W29) - dblW4(C_W40) + dblW4(C_W27) - dblW4(C_W55)	' -- 200603 ���׼��� 
		If dblW4(C_W47) < 0 Then
			.txtData(C_W47).value = 0
		Else
			.txtData(C_W47).value = dblW4(C_W47)	' (151) �г����װ������� 
		End If
		
		If dblW4(C_W47) <= 10000000 Then
			dblW4(C_W50) = 0
		ElseIf dblW4(C_W47) > 10000000 AND dblW4(C_W47) <= 20000000 Then
			dblW4(C_W50) = dblW4(C_W47) - 10000000
		ElseIf dblW4(C_W47) > 20000000 Then
			dblW4(C_W50) = Fix(dblW4(C_W47) * 0.5)
		End If
		.txtData(C_W50).value = dblW4(C_W50)	' (154) �� 

		dblW4(C_W49) = UNICDbl(.txtData(C_W49).value)
		'dblW4(C_W50) = UNICDbl(.txtData(C_W50).value)	
		dblW4(C_W48) = dblW4(C_W50) - dblW4(C_W49) 
		.txtData(C_W48).value = dblW4(C_W48)	' (152) ���ݳ��� 

		dblW4(C_W55) = UNICDbl(.txtData(C_W55).value)
		dblW4(C_W53) = dblW4(C_W46) - dblW4(C_W55) - dblW4(C_W50)
		.txtData(C_W53).value = dblW4(C_W53)	' (157) �� 
						
		dblW4(C_W52) = UNICDbl(.txtData(C_W52).value)
		'dblW4(C_W53) = UNICDbl(.txtData(C_W53).value)
		dblW4(C_W51) = dblW4(C_W53) - dblW4(C_W52)
		.txtData(C_W51).value = dblW4(C_W51)	' (155) ���ݳ��� 

	End With

	lgBlnFlgChgValue= True ' ���濩�� 
	IsRunEvents = False	' �̺�Ʈ �߻������� ������ 
End Sub

Sub GetFISC_DATE()	' ������ ��ȸ���ǿ� �����ϴ� ������,�������� �����´�.
	Dim sFiscYear, sRepType, sCoCd, iGap
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	' ���� �Ⱓ�� �ʼ��Է� 
	lgFISC_START_DT = CDate(lgF0)
	lgFISC_END_DT = CDate(lgF1)

End Sub

Function OpenW07()	'�̿���ձ� �˾� 

    Dim arrRet
    Dim arrParam(4)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
	Dim arrRowVal
    Dim arrColVal, lLngMaxRow
    Dim iDx
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True
    
	arrParam(0) = frm1.txtCO_CD.Value
	arrParam(1) = frm1.txtFISC_YEAR.Text		
	arrParam(2) = frm1.cboREP_TYPE.Value		
	arrParam(3) = UNICDbl(frm1.txtData(C_W06_1).value)		

    arrRet = window.showModalDialog(BIZ_POP_ID, Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0) <> "" Then
		frm1.txtData(C_W07).value = arrRet(0)
		
		Call SetHeadReCalc
	End IF
    
    IsOpenPop = False
    
    
End Function


Function txtData_onchange()
	Call SetHeadReCalc
End Function

'====================================== �� �Լ� =========================================

'============================================  ��ȸ���� �Լ�  ====================================


'============================================  �� �Լ�  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         

    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>
	 
	Call InitComboBox
	
	' �����Ѱ� 
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
	
	Call InitData 
	
    'Call ggoOper.FormatDate(frm1.txtW2 , parent.gDateFormat,3)
	Call FncQuery
	'
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
    'Call InitData                              
    															
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

	With frm1
		If UNICDbl(.txtData(C_W06_1).value) > 0 Then
			If UNICDbl(.txtData(C_W06_1).value) < 0 And  UNICDbl(.txtData(C_W07).value) > 0 Then
				Call DisplayMsgBox("W80001", "X", "(108)�� ������� �ҵ�ݾ�", "(109)�̿���ձ�")                          <%'No data changed!!%>
				Exit Function
			End If

			If UNICDbl(.txtData(C_W06_1).value) < UNICDbl(.txtData(C_W07).value) Then
				Call DisplayMsgBox("WC0010", "X", "(109)�̿����", "(108)�� ������� �ҵ�ݾ�")                          <%'No data changed!!%>
				Exit Function
			End If
		End If
		
		If UNICDbl(.txtData(C_W18).value) < UNICDbl(.txtData(C_W19).value) Then
			Call DisplayMsgBox("WC0010", "X", "(123)�������鼼��(��)", "(122)��������")                          <%'No data changed!!%>
			Exit Function
		End If
						
	End With
	
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
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")
    Call InitVariables
    Call InitData

    Call SetToolbar("1100100000000111")

	frm1.txtCO_CD.focus

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
		    
	' �������� ���� : ���ߵǸ� ���ȴ�.
	Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	'1 ����üũ : �׸��� �� 
	If wgConfirmFlg = "N" Then
		' ���� �ڵ� ȯ�氪�� ����� ���� ���� 
		With frm1
			If .txtData(C_W14).value <>"" Then
				If .txtData(C_W14).options(.txtData(C_W14).selectedIndex).VIEW <> .txtW14.value & "%" Then
					Call DisplayMsgBox("WC0029", "X", "(117) �� ��", "W2014")                          <%'No data changed!!%>
					Exit Function
				End If
			ElseIf .txtData(C_W35).value <>"" Then
				If .txtData(C_W35).options(.txtData(C_W35).selectedIndex).VIEW <> .txtW35.value & "%" Then
					Call DisplayMsgBox("WC0029", "X", "(139) �� ��", "W2013")                          <%'No data changed!!%>
					Exit Function
				End If
			End If
		End With
		Call SetToolbar("11011000000000111")
	Else
		
		'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
		Call SetToolbar("11001000000000111")										<%'��ư ���� ���� %>
	End If
	
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
	
		For i = C_W01 To C_W53	
			strVal = strVal & .txtData(i).Value & Parent.gColSep
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
	try {
		if (this.noevent == null)
			SetHeadReCalc();
    } catch(e) {
    }
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
									<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFISC_YEAR CLASS=FPDTYYYY title=FPDATETIME ALT="�������" tag="14X1" id=txtFISC_YEAR></OBJECT>');</SCRIPT>
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
						<TABLE <%=LR_SPACE_TYPE_60%> BORDER=0>
                            <TR>
                                <TD WIDTH="50%" VALIGN=TOP HEIGHT=100%>
								<TABLE <%=LR_SPACE_TYPE_20%> border="0" height=100% width="100%">
								   <TR>
										<TD>
											<TABLE <%=LR_SPACE_TYPE_20%> border="1" height=100% width="100%">
											 <TR>
												   <TD CLASS="TD51" width="5%" ROWSPAN=7 ALIGN=CENTER>(1)<br>��<br>��<br>��<br>��<br>��<br>��<br>��<br>��<br>��</TD>
												   <TD CLASS="TD51" width="60%"COLSPAN=2>(101) �� �� �� �� �� �� �� �� ��</TD>
												   <TD CLASS="TD61" width="5%" ALIGN=CENTER>01</TD>
												   <TD CLASS="TD61" width="30%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
											       <TD CLASS="TD51"  width="10%" ROWSPAN=2 ALIGN=CENTER>�� �� �� �� �� ��</TD>
												   <TD CLASS="TD51">(102) �� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>02</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>										   
												   <TD CLASS="TD51">(103) �� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>03</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
											     <TD CLASS="TD51" COLSPAN=2>(104) �� �� �� �� �� �� ��<br>[(101) + (102) - (103)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>04</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
											     <TD CLASS="TD51" COLSPAN=2>(105) �� �� �� �� �� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>05</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>									  
											<TR>
											     <TD CLASS="TD51" COLSPAN=2>(106) �� �� �� �� �� �� �� ��<br>�� �� �� �� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>54</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>	
											<TR>
											     <TD CLASS="TD51" COLSPAN=2>(107)�� �� �� �� �� �� �� �� ��<br>[(104) + (105) - (106)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>06</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											</TABLE>
										</TD>
									</TR>
									<TR>
									     <TD HEIGHT=2></TD>
									</TR>
									<TR>
										<TD>
											<TABLE <%=LR_SPACE_TYPE_20%> border="1" height=100% width="100%">
											 <TR>
												   <TD CLASS="TD51" width="5%" ROWSPAN=5 ALIGN=CENTER>(2)<br>��<br>��<br>ǥ<br>��<br>��<br>��</TD>
												   <TD CLASS="TD51" width="60%">(108)�� ������� �ҵ�ݾ�<br>[(108)=(107)]</TD>
												   <TD CLASS="TD61" width="5%" ALIGN=CENTER>&nbsp;</TD>
												   <TD CLASS="TD61" width="30%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											 <TR>
												   <TD CLASS="TD51">(109)�� �� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>07</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>											
											 <TR>
												   <TD CLASS="TD51">(110)�� �� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>08</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>	
											 <TR>
												   <TD CLASS="TD51">(111)�� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>09</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>	
											 <TR>
												   <TD CLASS="TD51">(112)�� �� ǥ ��<br>[(108) - (109) - (110) - (111)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>10</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>	
											</TABLE>
										</TD>
									</TR>
									<TR>
									     <TD HEIGHT=2></TD>
									</TR>
									<TR>
										<TD>
											<TABLE <%=LR_SPACE_TYPE_20%> border="1" height=100% width="100%">
											 <TR>
												   <TD CLASS="TD51" width="5%" ALIGN=CENTER>&nbsp;&nbsp;</TD>
												   <TD CLASS="TD51" width="60%">(159) �� �� ǥ �� �� ��</TD>
												   <TD CLASS="TD61" width="5%" ALIGN=CENTER>55</TD>
												   <TD CLASS="TD61" width="30%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW55_1" name=txtW55_1 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											</TABLE>
										</TD>
									</TR>
									<TR>
									     <TD HEIGHT=2></TD>
									</TR>
									<TR>
										<TD>
											<TABLE <%=LR_SPACE_TYPE_20%> border="1" height=100% width="100%">
											 <TR>
												   <TD CLASS="TD51" width="5%" ROWSPAN=7 ALIGN=CENTER>(3)<br>��<br>��<br>��<br>��<br>��<br>��</TD>
												   <TD CLASS="TD51" width="60%">(113) ���� ǥ�� �ݾ�2 [(112) + (159)]</TD>
												   <TD CLASS="TD61" width="5%" ALIGN=CENTER>56</TD>
												   <TD CLASS="TD61" width="30%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											 <TR>
												   <TD CLASS="TD51">(114) �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>11</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X6Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>											
											 <TR>
												   <TD CLASS="TD51">(115) �� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>12</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>	
											 <TR>
												   <TD CLASS="TD51">(116) �� �� �� �� �� ��<br>(�� �� 96��)</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>13</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>	
											 <TR>
												   <TD CLASS="TD51">(117) �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>14</TD>
												   <TD CLASS="TD61"><SELECT NAME=txtData STYLE="Width: 100%" tag="25X8Z" onChange="vbscript:SetHeadReCalc()"><OPTION VALUE="" VAL="0" VIEW=""></OPTION></SELECT></TD>
											</TR>	
											 <TR>
												   <TD CLASS="TD51">(118) �� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>15</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											 <TR>
												   <TD CLASS="TD51">(119) �� �� {(115)+(118)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>16</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											</TABLE>
										</TD>
									</TR>
									<TR>
									     <TD HEIGHT=2></TD>
									</TR>
									<TR>
										<TD>
											<TABLE <%=LR_SPACE_TYPE_20%> border="1" height=100% width="100%">
											 <TR>
												   <TD CLASS="TD51" width="5%" ROWSPAN=13 ALIGN=CENTER>(4)<br>��<br>��<br>��<br><br>��<br>��<br>��<br>��</TD>
												   <TD CLASS="TD51" width="60%" COLSPAN=3>(120)�� �� �� �� [(120) = (119)]</TD>
												   <TD CLASS="TD61" width="5%" ALIGN=CENTER>&nbsp;</TD>
												   <TD CLASS="TD61" width="30%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											 <TR>
												   <TD CLASS="TD51" width="60%" COLSPAN=3>(121)�� �� �� �� �� ��(��)</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>17</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>											
											 <TR>
												   <TD CLASS="TD51" COLSPAN=3>(122)�� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>18</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>	
											 <TR>
												   <TD CLASS="TD51" COLSPAN=3>(123)�� �� �� �� �� �� (��)</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>19</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>	
											 <TR>
												   <TD CLASS="TD51" COLSPAN=3>(124)�� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>20</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>	
											 <TR>
												   <TD CLASS="TD51" COLSPAN=3>(125)������ [(122)-(123)+(124)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>21</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											 <TR>
												   <TD CLASS="TD51" width="5%" ROWSPAN=7 ALIGN=CENTER>��γ��μ���</TD>
												   <TD CLASS="TD51" width="5%" ROWSPAN=5 ALIGN=CENTER>���ѳ� ���μ���</TD>
												   <TD CLASS="TD51" width="50%">(126)�߰� ���� ����</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>22</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											 <TR>
												   <TD CLASS="TD51">(127)���� �ΰ� ����</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>23</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											 <TR>
												   <TD CLASS="TD51">(128)��õ ���� ����</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>24</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											 <TR>
												   <TD CLASS="TD51">(129)��������ȸ����� �ܱ����μ���<INPUT type=hidden name=txtW25_NM STYLE="WIDTH: 50%" tag="25" maxlength=20></TD>
												   <TD CLASS="TD61" ALIGN=CENTER>25</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											 <TR>
												   <TD CLASS="TD51">(130)�� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>26</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											 <TR>
												   <TD CLASS="TD51" COLSPAN=2>(131)�Ű� ������ ���꼼��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>27</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>											
											 <TR>
												   <TD CLASS="TD51" COLSPAN=2>(132)�� �� [(130)+(131)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>28</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>												
											
											</TABLE>
										</TD>
									</TR>										
								  </TABLE>
								</TD>
                                <TD WIDTH="50%" VALIGN=TOP HEIGHT=100%>
								<TABLE <%=LR_SPACE_TYPE_20%> border="1" width="100%">
								   <TR>
										<TD>
											<TABLE <%=LR_SPACE_TYPE_20%> border="1" height=100% width="100%">
											 <TR>
												   <TD CLASS="TD51" width="5%" ROWSPAN=2></TD>
												   <TD CLASS="TD51" width="60%">(133) �� �� �� �� �� �� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>29</TD>
												   <TD CLASS="TD61" width="30%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51">(134) �� �� �� �� �� �� �� ��<br>[(125) - (132) + (133)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>30</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											</TABLE>
										</TD>
									</TR>
									<TR>
									     <TD HEIGHT=2></TD>
									</TR>
								   <TR>
										<TD>
											<TABLE <%=LR_SPACE_TYPE_20%> border="1" width="100%">
											 <TR>
												   <TD CLASS="TD51" width="5%" ROWSPAN=15 ALIGN=CENTER>(5)<br>��<br>��<br>��<br><br>��<br>��<br>��<br>��<br>��<br>��<br>��<br><br>��<br>��<br>��<br><br>��<br>��</TD>
												   <TD CLASS="TD51" width="10%" ROWSPAN=2 ALIGN=CENTER VALIGN=CENTER>�� ��<br>�� ��</TD>
												   <TD CLASS="TD51" width="50%">(135) �� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>31</TD>
												   <TD CLASS="TD61" width="30%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51">(136) �� �� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>32</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(137) �� �� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>33</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(138) ���� ǥ�� [(135) + (136) - (137)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>34</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(139) �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>35</TD>
												   <TD CLASS="TD61"><SELECT NAME=txtData STYLE="Width: 100%" tag="25X8Z" onChange="vbscript:SetHeadReCalc()"><OPTION VALUE="" VAL="0" VIEW=""></OPTION></SELECT</TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(140) �� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>36</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100% AutoCalc="No"></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(141) �� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>37</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(142) �� �� �� �� [(140) - (141)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>38</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(143) �� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>39</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(144) �� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>40</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(145) ������ [(142) - (143) + (144)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>41</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" ROWSPAN=3 ALIGN=CENTER>�ⳳ��<br>����</TD>
												   <TD CLASS="TD51">(146) �� �� �� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>42</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51">(147) (<INPUT name=txtW43_NM STYLE="WIDTH: 50%" tag="25" maxlength=20>) �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>43</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51">(148) �� [(143)+(147)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>44</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(149) �� �� �� �� �� �� ��<br>[(145) - (148)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>45</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											</TABLE>
										</TD>
									</TR>
									<TR>
									     <TD HEIGHT=2></TD>
									</TR>
								   <TR>
										<TD>
											<TABLE <%=LR_SPACE_TYPE_20%> border="1" width="100%">
											 <TR>
												   <TD CLASS="TD51" width="5%" ROWSPAN=9 ALIGN=CENTER>(6)<br>��<br>��<br>��</TD>
												   <TD CLASS="TD51" width="60%" COLSPAN=2>(150) �� �� �� �� �� �� �� ��<br>[(134) + (149)]</TD>
												   <TD CLASS="TD61" width="5%" ALIGN=CENTER>46</TD>
												   <TD CLASS="TD61" width="30%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(151) ��ǰ� �ٸ� ȸ��ó��<br>�� �� �� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>57</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" COLSPAN=2>(152) �� �� �� �� �� �� �� �� ��<br>[(150) - (124) - (133) - (144) + (131) - (151)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>47</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" width="10%" ROWSPAN=3 ALIGN=CENTER>�г���<br>����</TD>
												   <TD CLASS="TD51">(153) �� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>48</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X22" width = 100% noevent></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51">(154) �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>49</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51">(155) �� [(153) + (154)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>50</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51" ROWSPAN=3 ALIGN=CENTER>����<br>����<br>����</TD>
												   <TD CLASS="TD51">(156) �� �� �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>51</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X22" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51">(157) �� ��</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>52</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X8Z" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												   <TD CLASS="TD51">(158) �� [(156) + (157)]<br>[(158) = (150) - (151) - (155)]</TD>
												   <TD CLASS="TD61" ALIGN=CENTER>53</TD>
												   <TD CLASS="TD61"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X82" width = 100%></OBJECT>');</SCRIPT></TD>
											</TR>

											</TABLE>
										</TD>
									</TR>
									<TR>
									     <TD HEIGHT=2></TD>
									</TR>
								   <TR>
										<TD height=100>
											<TABLE <%=LR_SPACE_TYPE_20%> border="1" width="100%">
											 <TR>
												   <TD width="100%" HEIGHT=100%>&nbsp;</TD>
											</TR>
											</TABLE>
										</TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex=-1></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" style="display:'none'" tabindex=-1></TEXTAREA>
<INPUT TYPE=HIDDEN TABINDEX=-1 NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN TABINDEX=-1 NAME="txtHeadMode" tag="24">
<INPUT TYPE=HIDDEN TABINDEX=-1 NAME="txtW14" tag="24"><INPUT TYPE=HIDDEN TABINDEX=-1 NAME="txtW35" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<input type="hidden" name="uname" TABINDEX="-1">
	<input type="hidden" name="dbname" TABINDEX="-1">
	<input type="hidden" name="filename" TABINDEX="-1">
	<input type="hidden" name="strUrl" TABINDEX="-1">
	<input type="hidden" name="date" TABINDEX="-1">
</FORM>
</BODY>
</HTML>

