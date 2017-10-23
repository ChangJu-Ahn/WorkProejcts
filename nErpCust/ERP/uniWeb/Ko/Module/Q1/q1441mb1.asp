<%@LANGUAGE = VBScript%> 
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->	
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1441MB1
'*  4. Program Name         : ������ �ý����� �׷����� �׸���.
'*  5. Program Desc         : 
'*  6. Component List       : PQBG120
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- ChartFX�� ����� ����ϱ� ���� Include ���� -->
<!-- #include file="../../inc/CfxIE.inc" -->
<%													
On Error Resume Next

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("I", "*", "NOCOOKIE", "QB")

Dim dLotSize
Dim dSamplesize
Dim dAcceptCount
Dim dProcessDefectRatio

'Dim dASNSamplesize1
'Dim dASNSamplesize2
'Dim dASNAcceptanceCnt1
'Dim dASNAcceptanceCnt2
Dim dASNProcessDefRatio

Dim strReplace

On Error Resume Next

dLotSize = UNICDbl(UNIConvNum(Request("txtLotsize"), 0), 0)
dSamplesize = UNICDbl(UNIConvNum(Request("txtSamplesize"), 0), 0)
dProcessDefectRatio = UNICDbl(UNIConvNum(Request("txtProcessDefectRatio"), 0), 0)
dAcceptCount = UNICDbl(UNIConvNum(Request("txtAcceptCount"), 0), 0)

'dASNSamplesize1 = UNICDbl(UNIConvNum(Request("txtSamplesize1"), 0), 0)
'dASNSamplesize2 = UNICDbl(UNIConvNum(Request("txtSamplesize2"), 0), 0)
'dASNAcceptanceCnt1 = UNICDbl(UNIConvNum(Request("txtAccept1"), 0), 0)
'dASNAcceptanceCnt2 = UNICDbl(UNIConvNum(Request("txtAccept2"), 0), 0)
'dASNProcessDefRatio = UNICDbl(UNIConvNum(Request("txtDefectRatio"), 0), 0)

dASNProcessDefRatio = dProcessDefectRatio
'	DisVA = ReadCookie("txtInsVA")
strReplace = Request("txtReplaceMode")

'++++++++++++++++++++++++++++++++++++++++++  2.5.1 ATI � ����Լ� +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
	Dim R
	
    Dim LotAcc 	                		'LOT�� �հ�Ȯ���� ���Ѵ�.
    Dim RetATI(250) 
    Dim ATI
    Dim Temp 
    
    Dim i
    Dim j
    Dim k
    Dim a
    
    Dim ScaleDefRatio
    Dim DefRatioDiv
    Dim ATIMax
    Dim ASNMax
    Dim OCMax
    Dim AOQMax
    
    DIm Biono_Samplesize				'���׺����� ���� ����ϱ� ���� n�� �޾� ���̴� �κ� 
    Dim Biono_DefectRatio				'���׺����� ���� ����ϱ� ���� p�� �޾� ���̴� �κ� 
    Dim bino_val
    
    Biono_Samplesize = dSamplesize
    Biono_DefectRatio = dProcessDefectRatio
    
    ScaleDefRatio = 0 
    DefRatioDiv = 0

    If 0.1 > dProcessDefectRatio then
		ScaleDefRatio = 0.15
    Elseif dProcessDefectRatio > 0.11 and dProcessDefectRatio =< 0.20 then
    	ScaleDefRatio = 0.25
    Elseif dProcessDefectRatio > 0.21 and dProcessDefectRatio =< 0.30 then
    	ScaleDefRatio = 0.35
    Elseif dProcessDefectRatio > 0.31 and dProcessDefectRatio =< 0.40 then
    	ScaleDefRatio = 0.45    
    Elseif dProcessDefectRatio > 0.41 and dProcessDefectRatio =< 0.50 then
    	ScaleDefRatio = 0.55
    Elseif dProcessDefectRatio > 0.51 and dProcessDefectRatio =< 0.60 then
    	ScaleDefRatio = 0.65    
    Elseif dProcessDefectRatio > 0.61 and dProcessDefectRatio =< 0.70 then
    	ScaleDefRatio = 0.75
    Elseif dProcessDefectRatio > 0.71 and dProcessDefectRatio =< 0.80 then
    	ScaleDefRatio = 0.85    
    Elseif dProcessDefectRatio > 0.81 and dProcessDefectRatio =< 0.90 then
    	ScaleDefRatio = 0.95    
    Elseif dProcessDefectRatio > 0.96 then
    	ScaleDefRatio = 1.0
    End if
    
    ' ScaleDefRatio = dProcessDefectRatio * 13.3			'�Է¹��� �ҷ����� �������� ����Ͽ� �����ϰ� �����Ѵ�.
    DefRatioDiv = ScaleDefRatio / 250				'ȭ�鿡 �°� x���� ������ �ݴϴ�.
    
    
	RetATI(0) = dSamplesize   
	ATIMax = dSamplesize
    For R=1 to 250
	    Biono_DefectRatio = DefRatioDiv * R				'�ҷ����Է°��� ���ϰ� �մϴ�.
	    LotAcc = 0
		
		For k = 0 to dAcceptCount
		    a = k								'�Ѱ��ִ� �μ��� for���� ������ ������ Error�߻�.
	        bino_val = Bino(dSamplesize, a, Biono_DefectRatio)        		'���׺��� �Լ��� ȣ���Ѵ�.
	        
	        LotAcc = LotAcc + bino_val  					'x=0 ~ x=c ������ �������� ���Ѵ�.
	    Next 

		Select Case strReplace  			'�ҷ�ǰ ��ü ���ο� ���� ������ ����˴ϴ�.
			Case 0					'�ҷ�ǰ�� ��ǰ���� ��ü���� �ʴ� ��� 
		    	ATI = dSamplesize * LotAcc + dLotSize * (1 - LotAcc) 			'ATI���� ���Ѵ�.
			Case 1					'�ҷ�ǰ�� ��ǰ���� ��ü ��� 
		    	ATI = (dSamplesize+(dLotSize-dSamplesize*(1-LotAcc)))/(1-LotAcc)	'ATI���� ���Ѵ�.	    	
		End Select	   
		
		If ATIMax < ATI Then
			ATIMax = ATI
		End If
		RetATI(R) = ATI           			 			'��ȯ ���� ���� �迭�� �����ݴϴ�. 
	Next

ATIMax = ATIMax * 1.1					'�׷��� �׸��� Y�ప�� �ִ밪���� ���� ũ�� �Ѵ�.

'++++++++++++++++++++++++++++++++++++++++++  2.5.2 ASN � ����Լ� +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

'	Dim p_n1                 'ASN���� ���� ���� 
'	Dim p_n2 
'	Dim p_c1 
'	Dim p_c2 

 '  	p_n1 = dASNSamplesize1			'������ ���ؼ� �Է¹޾ƾ� ��.
  ' 	p_n2 = dASNSamplesize2
   '	p_c1 = dASNAcceptanceCnt1
   	'p_c2 = dASNAcceptanceCnt2
   '	p = dASNProcessDefRatio

   	Dim LotAcc_ASN				'ASN�� ���Ǵ� �հ�Ȯ�� 
   	Dim LOTRej 		                   	'LOT�� ���հ�Ȯ���� ���մϴ�.
   	Dim ASN 
   	
	Dim Prio_Prob 
	Dim RetASN(250)
	Dim ScaleDefRatio2
	Dim DefRatioDiv2
	Dim S   
	
'   	Biono_Samplesize = 0			'ASN�� ���ϱ� ���� ���׺����� n���� �ʱ�ȭ �����ݴϴ�.
'  	Biono_Samplesize = p_n1
' 	Biono_DefectRatio = 0
   	
	S = 0
'	bino_val = 0

    If 0.1 > dASNProcessDefRatio then
		ScaleDefRatio2 = 0.15
	Elseif dASNProcessDefRatio > 0.11 and dASNProcessDefRatio =< 0.20 then
		ScaleDefRatio2 = 0.25
	Elseif dASNProcessDefRatio > 0.21 and dASNProcessDefRatio =< 0.30 then
		ScaleDefRatio2 = 0.35
	Elseif dASNProcessDefRatio > 0.31 and dASNProcessDefRatio =< 0.40 then
		ScaleDefRatio2 = 0.45    
	Elseif dASNProcessDefRatio > 0.41 and dASNProcessDefRatio =< 0.50 then
		ScaleDefRatio2 = 0.55
	Elseif dASNProcessDefRatio > 0.51 and dASNProcessDefRatio =< 0.60 then
		ScaleDefRatio2 = 0.65    
	Elseif dASNProcessDefRatio > 0.61 and dASNProcessDefRatio =< 0.70 then
		ScaleDefRatio2 = 0.75
	Elseif dASNProcessDefRatio > 0.71 and dASNProcessDefRatio =< 0.80 then
		ScaleDefRatio2 = 0.85    
	Elseif dASNProcessDefRatio > 0.81 and dASNProcessDefRatio =< 0.90 then
		ScaleDefRatio2 = 0.95    
	Elseif dASNProcessDefRatio > 0.96 then
		ScaleDefRatio2 = 1.0
	End if

  	DefRatioDiv2 = ScaleDefRatio2 / 250			'ȭ�鿡 �°� x���� ������ �ݴϴ�.
    	
'	IF dASNSamplesize1 > 0 then			'2 ȸ �˻��� ��쿡 ���� 
'
'	   	For S=1 to 250
'		
'			bino_val = 0                            		'�ʱ�ȭ 
'			LOTRej = 0                              		'�ʱ�ȭ 
'		
'		    	Biono_DefectRatio = DefRatioDiv2 * S				'�ҷ����Է°��� ���ϰ� �մϴ�.
'			
'	 		LotAcc_ASN = Bino(p_n1, p_c1, Biono_DefectRatio)          		'LOT�� �հ�Ȯ���� ���մϴ�.
'	  
'		              'i�� �ٸ� ������ �����ؾ� �����ذ�, ��ġ�鼭 ���� �߻�   		'���ϰ� = �Լ��̸� �� 
'			For w = 0 To p_c2
'				a = w						'������ �����ϱ� ���ؼ� 
'				bino_val = Bino(p_n1, a, Biono_DefectRatio)
'			            	LOTRej = LOTRej + bino_val      			'���׺����� ���ؼ� �������� ���Ѵ�.
'			Next 
'				LOTRej = 1 - LOTRej             			'��꿡 �ʿ��� �������� ���Ѵ�.
'			   
'			Prio_Prob = LotAcc_ASN + LOTRej             			'ASN�� LOT�� �հ�Ȯ���� ���հ�Ȯ���� ���մϴ�.
'			
'			ASN = p_n1 + p_n2 * (1 - Prio_Prob)       			'���� ASN����� ���մϴ�.
'			
'			RetASN(S) = ASN						'������� ��̿� �Ҵ��մϴ�.
'	
'			if ASN > ASNMax then					'����߿��� �ִ밪 ã�� 
'				ASNMax = ASN	
'			End if
'		Next	
'	End IF
	
	ASNMax = dSamplesize + dSamplesize * 0.1			'����߿��� �ִ밪 ã��			
	ASN = dSamplesize					'������ strASNSamplesize���� �о�;� �� 

	For S=0 to 250	
		RetASN(S) = ASN						'������� ��̿� �Ҵ��մϴ�.
	Next	

'++++++++++++++++++++++++++++++++++++++++++  2.5.3 OC � ����Լ� +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
    Dim LotAcc_OC 	                		'LOT�� �հ�Ȯ���� ���Ѵ�.
    Dim RetOC(250) 
    Dim OC
       
    Dim ScaleDefRatio3
    Dim DefRatioDiv3
    
    ScaleDefRatio3 = 0 
    DefRatioDiv3 = 0

	If 0.05 => dProcessDefectRatio then
		ScaleDefRatio3 = 0.075
	Elseif dProcessDefectRatio > 0.51 and dProcessDefectRatio =< 0.10 then
		ScaleDefRatio3 = 0.15
	Elseif dProcessDefectRatio > 0.11 and dProcessDefectRatio =< 0.20 then
		ScaleDefRatio3 = 0.25
	Elseif dProcessDefectRatio > 0.21 and dProcessDefectRatio =< 0.30 then
		ScaleDefRatio3 = 0.35
	Elseif dProcessDefectRatio > 0.31 and dProcessDefectRatio =< 0.40 then
		ScaleDefRatio3 = 0.45    
	Elseif dProcessDefectRatio > 0.41 and dProcessDefectRatio =< 0.50 then
		ScaleDefRatio3 = 0.55
	Elseif dProcessDefectRatio > 0.51 and dProcessDefectRatio =< 0.60 then
		ScaleDefRatio3 = 0.65    
	Elseif dProcessDefectRatio > 0.61 and dProcessDefectRatio =< 0.70 then
		ScaleDefRatio3 = 0.75
	Elseif dProcessDefectRatio > 0.71 and dProcessDefectRatio =< 0.80 then
		ScaleDefRatio3 = 0.85    
	Elseif dProcessDefectRatio > 0.81 and dProcessDefectRatio =< 0.90 then
		ScaleDefRatio3 = 0.95    
	Elseif dProcessDefectRatio > 0.96 then
		ScaleDefRatio3 = 1.0
	End if

    DefRatioDiv3 = ScaleDefRatio3 / 250				'ȭ�鿡 �°� x���� ������ �ݴϴ�.

    For R=0 to 250
	    Biono_DefectRatio= DefRatioDiv3 * R						'�ҷ����Է°��� ���ϰ� �մϴ�.
	                 								'���ϰ� = �Լ��̸� �� 
	    LotAcc_OC = 0
	    
		For k= 0 to dAcceptCount
			a=k								'�Ѱ��ִ� �μ��� for���� ������ ������ Error�߻�.
			bino_val = Bino(dSamplesize, a, Biono_DefectRatio)        		'���׺��� �Լ��� ȣ���Ѵ�.
			LotAcc_OC = LotAcc_OC + bino_val 				'x=0 ~ x=c ������ �������� ���Ѵ�.
		Next 
	    	
	    RetOC(R) = LotAcc_OC           			 		'��ȯ ���� ���� �迭�� �����ݴϴ�. 
	Next

	OCMax = RetOC(1) 							'�׷��� �׸��� Y�ప�� �ִ밪���� ���� ũ�� �Ѵ�.

'++++++++++++++++++++++++++++++++++++++++++  2.5.4 AOQ � ����Լ� +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
    Dim LotAcc_AOQ 	                		'LOT�� �հ�Ȯ���� ���Ѵ�.
    Dim RetAOQ(250) 
    Dim AOQ
     
    Dim ScaleDefRatio4
    Dim DefRatioDiv4
      
  '  Biono_Samplesize = dSamplesize
  '  Biono_DefectRatio = dProcessDefectRatio
    
    ScaleDefRatio4 = 0 
    DefRatioDiv4 = 0
    
	If 0.05 => dProcessDefectRatio then
		ScaleDefRatio4 = 0.075
	Elseif dProcessDefectRatio > 0.51 and dProcessDefectRatio =< 0.10 then
	    ScaleDefRatio4 = 0.15
	Elseif dProcessDefectRatio > 0.11 and dProcessDefectRatio =< 0.20 then
	    ScaleDefRatio4 = 0.25
	Elseif dProcessDefectRatio > 0.21 and dProcessDefectRatio =< 0.30 then
	    ScaleDefRatio4 = 0.35
	Elseif dProcessDefectRatio > 0.31 and dProcessDefectRatio =< 0.40 then
	    ScaleDefRatio4 = 0.45    
	Elseif dProcessDefectRatio > 0.41 and dProcessDefectRatio =< 0.50 then
	    ScaleDefRatio4 = 0.55
	Elseif dProcessDefectRatio > 0.51 and dProcessDefectRatio =< 0.60 then
	    ScaleDefRatio4 = 0.65    
	Elseif dProcessDefectRatio > 0.61 and dProcessDefectRatio =< 0.70 then
	    ScaleDefRatio4 = 0.75
	Elseif dProcessDefectRatio > 0.71 and dProcessDefectRatio =< 0.80 then
	    ScaleDefRatio4 = 0.85    
	Elseif dProcessDefectRatio > 0.81 and dProcessDefectRatio =< 0.90 then
	    ScaleDefRatio4 = 0.95    
	Elseif dProcessDefectRatio > 0.96 then
	    ScaleDefRatio4 = 1.0
	End if    

    DefRatioDiv4 = ScaleDefRatio4 / 250				'ȭ�鿡 �°� x���� ������ �ݴϴ�.
    
    AOQMax=0

    For R=0 to 250
	    Biono_DefectRatio= DefRatioDiv4 * R				'�ҷ����Է°��� ���ϰ� �մϴ�.
	                 						'���ϰ� = �Լ��̸� �� 
	    LotAcc_AOQ = 0
	    
		For k= 0 to dAcceptCount
			a=k								'�Ѱ��ִ� �μ��� for���� ������ ������ Error�߻�.
			bino_val = Bino(dSamplesize, a, Biono_DefectRatio)        		'���׺��� �Լ��� ȣ���Ѵ�.
			LotAcc_AOQ = LotAcc_AOQ + bino_val  					'x=0 ~ x=c ������ �������� ���Ѵ�.
		Next 
	    	
	Select Case strReplace  				'�ҷ�ǰ ��ü ���ο� ���� ������ ����˴ϴ�.
		Case 0					'�ҷ�ǰ�� ��ǰ���� ��ü���� �ʴ� ��� 
	    	AOQ = (LotAcc_AOQ*Biono_DefectRatio*(dLotSize-dSamplesize))/((dLotSize-Biono_DefectRatio*dSamplesize)-(1-LotAcc_AOQ)*Biono_DefectRatio*(dLotSize-dSamplesize))
	    	'AOQ���� ���Ѵ�.
		Case 1					'�ҷ�ǰ�� ��ǰ���� ��ü�ϴ� ��� 
		AOQ = Biono_DefectRatio*LotAcc_AOQ*(1 - dSamplesize/dLotSize)		'AOQ���� ���Ѵ�.
	End Select
	
	RetAOQ(R) = AOQ           			 			'��ȯ ���� ���� �迭�� �����ݴϴ�. 
	 
	if AOQ > AOQMax then					'����߿��� �ִ밪 ã�� 
		AOQMax = AOQ
	End if
		
Next

'++++++++++++++++++++++++++++++++++++++++++  2.5.1 ���׺��� ����Լ� +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function Bino(n,a,p)    '���׺������� ����մϴ�.

		Dim temp1 
		Dim temp2 
		Dim temp3
		Dim Multemp1 
		Dim Multemp2 
		Dim Comb_val 
    
	    temp1 =  Biono_Samplesize
	    temp2 = a
   	    temp3=  Biono_DefectRatio

	    Multemp1 = 1
	    Multemp2 = 1
	  
	    For i = (Biono_Samplesize - temp2 + 1) To temp1			'���� �κ��� ���� 
	        	Multemp1 = Multemp1 * i
	    Next 
	    
	    If temp2 = 0 Then      						'�и� 0�� ���, ����� 1�� ǥ���մϴ�.
	        	Comb_val = 1
	    Else	        
	        For j = 1 To temp2
	            	Multemp2 = Multemp2 * j
	        Next 
	        	Comb_val = Multemp1 / Multemp2
	    End If
	    
	    Bino = Comb_val * (temp3 ^ temp2) * ((1 - temp3) ^ (temp1- temp2)) 	'�Լ��̸� = ���� �� 
    	    
End Function

%>

<Script Language=vbscript>
	
		Dim ATIMeas_Val
		Redim ATIMeas_Val(250)
		
		Dim ASNMeas_Val
		Redim ASNMeas_Val(250)
		
		Dim OCMeas_Val
		Redim OCMeas_Val(250)
		
		Dim AOQMeas_Val
		Redim AOQMeas_Val(250)
<%	
		Dim Cnt 		    
		
		For Cnt = 0 to 250
%>
		ATIMeas_Val(<%=Cnt%>) = "<%=RetATI(Cnt)%>"
		ASNMeas_Val(<%=Cnt%>) = "<%=RetASN(Cnt)%>"
		OCMeas_Val(<%=Cnt%>) = "<%=RetOC(Cnt)%>"
		AOQMeas_Val(<%=Cnt%>) = "<%=RetAOQ(Cnt)%>"
<%
 		Next
 %>

 	Dim i
		
	'��ƮFX1 - ATI� �׸��� 
	With Parent.frm1.ChartFX1
		.Title_(2) = "ATI�"
		.Gallery = <% = LINES%>
		.Axis(<%=AXIS_Y%>).Max = <% = ATIMax %>
		.Axis(<%=AXIS_Y%>).Decimals = 0
		
		.MarkerShape = <%=MK_NONE%>
		.AXIS(<%=AXIS_X%>).PixPerUnit = 1						'X���� ������ 0���� ����� �� ����.
		
		.OpenDataEx <%=COD_VALUES%>, 1, 251				'��Ʈ FX���� ������ ä�� �����ֱ� 
			For i = 0 to 250	
				.ValueEx(0, i) = ATIMeas_Val(i)
			Next
		.CloseData <%=COD_VALUES%>
		
		.Series(0).MarkerShape = <%=MK_NONE%>
		.Series(0).Visible = True
		
		.Axis(<%=AXIS_Y%>).Max = .Axis(<%=AXIS_Y%>).Max * 1.1
	
		.Axis(<%=AXIS_X%>).TickMark = <%=TS_NONE%>		
		'.Axis(<%=AXIS_X%>).Max = <% = DefRatioDiv %>
		.Axis(<%=AXIS_X%>).Max = 100
		.Axis(<%=AXIS_X%>).Step = 50
		
		'���� % --> * 100
		.Axis(<%=AXIS_X%>).Label(0)  = "<%=UniNumClientFormat(0, 2, 0)%>"						
		.Axis(<%=AXIS_X%>).Label(50) = "<%=UniNumClientFormat(DefRatioDiv * 50 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(100) = "<%=UniNumClientFormat(DefRatioDiv * 100 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(150) = "<%=UniNumClientFormat(DefRatioDiv * 150 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(200) = "<%=UniNumClientFormat(DefRatioDiv * 200 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(250) = "<%=UniNumClientFormat(DefRatioDiv * 250 * 100, 2, 0)%>"
		
	End With
	
	
	'��ƮFX2 - ASN� �׸��� 
	With Parent.frm1.ChartFX2
		.Title_(2) = "ASN �"
		.Gallery = <% = LINES%>
		.Axis(<%=AXIS_Y%>).Max = <% = ASNMax %>
		.Axis(<%=AXIS_Y%>).Step = <% = ASNMax %> / 6
		.Axis(<%=AXIS_Y%>).Decimals = 0
		
		.MarkerShape = <%=MK_NONE%>
		.AXIS(<%=AXIS_X%>).PixPerUnit = 1						'X���� ������ 0���� ����� �� ����.
		
		.OpenDataEx <%=COD_VALUES%>, 1, 251					'��Ʈ FX���� ������ ä�� �����ֱ� 
			For i = 0 to 250	
				.ValueEx(0, i) = ASNMeas_Val(i)
			Next			
		.CloseData <%=COD_VALUES%>
		
		.Series(0).MarkerShape = <%=MK_NONE%>
		.Series(0).Visible = True
		
		.Axis(<%=AXIS_Y%>).Max = .Axis(<%=AXIS_Y%>).Max * 1.1	
		
		.Axis(<%=AXIS_X%>).TickMark = <%=TS_NONE%>		
		'.Axis(<%=AXIS_X%>).Max = <% = DefRatioDiv2 %>
		.Axis(<%=AXIS_X%>).Max = 100
		.Axis(<%=AXIS_X%>).Step = 50
		
		'���� % --> * 100
		.Axis(<%=AXIS_X%>).Label(0)  = "<%=UniNumClientFormat(0, 2, 0)%>"						
		.Axis(<%=AXIS_X%>).Label(50) = "<%=UniNumClientFormat(DefRatioDiv2 * 50 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(100) = "<%=UniNumClientFormat(DefRatioDiv2 * 100 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(150) = "<%=UniNumClientFormat(DefRatioDiv2 * 150 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(200) = "<%=UniNumClientFormat(DefRatioDiv2 * 200 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(250) = "<%=UniNumClientFormat(DefRatioDiv2 * 250 * 100, 2, 0)%>"
		
	End With
	
	'��ƮFX3 - OC� �׸��� 
	With Parent.frm1.ChartFX3
		.Title_(2) = "OC �"
		.Gallery = <% = LINES%>
		 .Axis(<%=AXIS_Y%>).Max = <% = OCMax %>
		
		.MarkerShape = <%=MK_NONE%>
		.AXIS(<%=AXIS_X%>).PixPerUnit = 1						'X���� ������ 0���� ����� �� ����.
		
		.OpenDataEx <%=COD_VALUES%>, 1, 251				'��Ʈ FX���� ������ ä�� �����ֱ� 
			For i = 0 to 250	
				.ValueEx(0, i) = OCMeas_Val(i)
			Next
		.CloseData <%=COD_VALUES%>
		
		.Series(0).MarkerShape = <%=MK_NONE%>
		.Series(0).Visible = True
		
		.Axis(<%=AXIS_Y%>).Max = .Axis(<%=AXIS_Y%>).Max * 1.1
	
		.Axis(<%=AXIS_X%>).TickMark = <%=TS_NONE%>		
		'.Axis(<%=AXIS_X%>).Max = <% = DefRatioDiv3 %>
		.Axis(<%=AXIS_X%>).Max = 100
		.Axis(<%=AXIS_X%>).Step = 50
		'���� % --> * 100
		.Axis(<%=AXIS_X%>).TickMark = <%=TS_NONE%>
		.Axis(<%=AXIS_X%>).Label(0)  = "<%=UniNumClientFormat(0, 2, 0)%>"						
		.Axis(<%=AXIS_X%>).Label(50) = "<%=UniNumClientFormat(DefRatioDiv3 * 50 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(100) = "<%=UniNumClientFormat(DefRatioDiv3 * 100 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(150) = "<%=UniNumClientFormat(DefRatioDiv3 * 150 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(200) = "<%=UniNumClientFormat(DefRatioDiv3 * 200 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(250) = "<%=UniNumClientFormat(DefRatioDiv3 * 250 * 100, 2, 0)%>"
		
	End With

	'��ƮFX4 - AOQ� �׸��� 
	With Parent.frm1.ChartFX4
		.Title_(2) = "AOQ �"
		.Gallery = <% = LINES%>
		.Axis(<%=AXIS_Y%>).Max = <% = AOQMax %>
		.Axis(<%=AXIS_Y%>).Step = <% = AOQMax %> / 6
		.Axis(<%=AXIS_Y%>).Decimals = 4							'AOQ�� ��� �ִ밪�� �ʹ� �۾Ƽ� �Ҽ��� 4�ڸ����� ǥ���Ѵ�.
		 
		.MarkerShape = <%=MK_NONE%>
		.AXIS(<%=AXIS_X%>).PixPerUnit = 0						'X���� ������ 0���� ����� �� ����.
		
		.OpenDataEx <%=COD_VALUES%>, 1, 251				'��Ʈ FX���� ������ ä�� �����ֱ� 
			For i = 0 to 250	
				.ValueEx(0, i) = AOQMeas_Val(i)
			Next
		.CloseData <%=COD_VALUES%>
		
		.Series(0).MarkerShape = <%=MK_NONE%>
		.Series(0).Visible = True
		
		.Axis(<%=AXIS_Y%>).Max = .Axis(<%=AXIS_Y%>).Max * 1.1
		
		.Axis(<%=AXIS_X%>).TickMark = <%=TS_NONE%>		
		'.Axis(<%=AXIS_X%>).Max = <% = DefRatioDiv4 %>
		.Axis(<%=AXIS_X%>).Max = 100
		.Axis(<%=AXIS_X%>).Step = 50
		'���� % --> * 100
		.Axis(<%=AXIS_X%>).Label(0)  = "<%=UniNumClientFormat(0, 2, 0)%>"						
		.Axis(<%=AXIS_X%>).Label(50) = "<%=UniNumClientFormat(DefRatioDiv4 * 50 * 100, 2, 0)%>"	
		.Axis(<%=AXIS_X%>).Label(100) = "<%=UniNumClientFormat(DefRatioDiv4 * 100 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(150) = "<%=UniNumClientFormat(DefRatioDiv4 * 150 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(200) = "<%=UniNumClientFormat(DefRatioDiv4 * 200 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(250) = "<%=UniNumClientFormat(DefRatioDiv4 * 250 * 100, 2, 0)%>"
		
	End With
		
</Script>	
