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
'*  4. Program Name         : 전문가 시스템의 그래픽을 그린다.
'*  5. Program Desc         : 
'*  6. Component List       : PQBG120
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- ChartFX용 상수를 사용하기 위한 Include 지정 -->
<!-- #include file="../../inc/CfxIE.inc" -->
<%													
On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
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

'++++++++++++++++++++++++++++++++++++++++++  2.5.1 ATI 곡선 계산함수 +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
	Dim R
	
    Dim LotAcc 	                		'LOT의 합격확률을 구한다.
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
    
    DIm Biono_Samplesize				'이항분포의 값을 계산하기 위해 n을 받아 들이는 부분 
    Dim Biono_DefectRatio				'이항분포의 값을 계산하기 위해 p를 받아 들이는 부분 
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
    
    ' ScaleDefRatio = dProcessDefectRatio * 13.3			'입력받은 불량률을 스케일을 고려하여 적절하게 설정한다.
    DefRatioDiv = ScaleDefRatio / 250				'화면에 맞게 x축을 나누어 줍니다.
    
    
	RetATI(0) = dSamplesize   
	ATIMax = dSamplesize
    For R=1 to 250
	    Biono_DefectRatio = DefRatioDiv * R				'불량률입력값을 변하게 합니다.
	    LotAcc = 0
		
		For k = 0 to dAcceptCount
		    a = k								'넘겨주는 인수가 for문의 변수와 같으면 Error발생.
	        bino_val = Bino(dSamplesize, a, Biono_DefectRatio)        		'이항분포 함수를 호출한다.
	        
	        LotAcc = LotAcc + bino_val  					'x=0 ~ x=c 까지의 누적값을 구한다.
	    Next 

		Select Case strReplace  			'불량품 대체 여부에 따라 계산식이 변경됩니다.
			Case 0					'불량품을 양품으로 대체하지 않는 경우 
		    	ATI = dSamplesize * LotAcc + dLotSize * (1 - LotAcc) 			'ATI값을 구한다.
			Case 1					'불량품을 양품으로 대체 경우 
		    	ATI = (dSamplesize+(dLotSize-dSamplesize*(1-LotAcc)))/(1-LotAcc)	'ATI값을 구한다.	    	
		End Select	   
		
		If ATIMax < ATI Then
			ATIMax = ATI
		End If
		RetATI(R) = ATI           			 			'반환 받은 값을 배열에 보여줍니다. 
	Next

ATIMax = ATIMax * 1.1					'그래프 그릴때 Y축값을 최대값보다 조금 크게 한다.

'++++++++++++++++++++++++++++++++++++++++++  2.5.2 ASN 곡선 계산함수 +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

'	Dim p_n1                 'ASN관련 변수 선언 
'	Dim p_n2 
'	Dim p_c1 
'	Dim p_c2 

 '  	p_n1 = dASNSamplesize1			'쿼리에 의해서 입력받아야 함.
  ' 	p_n2 = dASNSamplesize2
   '	p_c1 = dASNAcceptanceCnt1
   	'p_c2 = dASNAcceptanceCnt2
   '	p = dASNProcessDefRatio

   	Dim LotAcc_ASN				'ASN에 사용되는 합격확률 
   	Dim LOTRej 		                   	'LOT의 불합격확률을 구합니다.
   	Dim ASN 
   	
	Dim Prio_Prob 
	Dim RetASN(250)
	Dim ScaleDefRatio2
	Dim DefRatioDiv2
	Dim S   
	
'   	Biono_Samplesize = 0			'ASN을 구하기 위해 이항분포의 n값을 초기화 시켜줍니다.
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

  	DefRatioDiv2 = ScaleDefRatio2 / 250			'화면에 맞게 x축을 나누어 줍니다.
    	
'	IF dASNSamplesize1 > 0 then			'2 회 검사의 경우에 적용 
'
'	   	For S=1 to 250
'		
'			bino_val = 0                            		'초기화 
'			LOTRej = 0                              		'초기화 
'		
'		    	Biono_DefectRatio = DefRatioDiv2 * S				'불량률입력값을 변하게 합니다.
'			
'	 		LotAcc_ASN = Bino(p_n1, p_c1, Biono_DefectRatio)          		'LOT의 합격확률을 구합니다.
'	  
'		              'i를 다른 것으로 변경해야 문제해결, 곁치면서 문제 발생   		'리턴값 = 함수이름 꼭 
'			For w = 0 To p_c2
'				a = w						'에러를 방지하기 위해서 
'				bino_val = Bino(p_n1, a, Biono_DefectRatio)
'			            	LOTRej = LOTRej + bino_val      			'이항분포에 의해서 사전값을 구한다.
'			Next 
'				LOTRej = 1 - LOTRej             			'계산에 필요한 실제값을 구한다.
'			   
'			Prio_Prob = LotAcc_ASN + LOTRej             			'ASN을 LOT의 합격확률과 불합격확률을 더합니다.
'			
'			ASN = p_n1 + p_n2 * (1 - Prio_Prob)       			'실제 ASN결과를 구합니다.
'			
'			RetASN(S) = ASN						'계산결과를 어레이에 할당합니다.
'	
'			if ASN > ASNMax then					'어레이중에서 최대값 찾기 
'				ASNMax = ASN	
'			End if
'		Next	
'	End IF
	
	ASNMax = dSamplesize + dSamplesize * 0.1			'어레이중에서 최대값 찾기			
	ASN = dSamplesize					'원래는 strASNSamplesize에서 읽어와야 함 

	For S=0 to 250	
		RetASN(S) = ASN						'계산결과를 어레이에 할당합니다.
	Next	

'++++++++++++++++++++++++++++++++++++++++++  2.5.3 OC 곡선 계산함수 +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
    Dim LotAcc_OC 	                		'LOT의 합격확률을 구한다.
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

    DefRatioDiv3 = ScaleDefRatio3 / 250				'화면에 맞게 x축을 나누어 줍니다.

    For R=0 to 250
	    Biono_DefectRatio= DefRatioDiv3 * R						'불량률입력값을 변하게 합니다.
	                 								'리턴값 = 함수이름 꼭 
	    LotAcc_OC = 0
	    
		For k= 0 to dAcceptCount
			a=k								'넘겨주는 인수가 for문의 변수와 같으면 Error발생.
			bino_val = Bino(dSamplesize, a, Biono_DefectRatio)        		'이항분포 함수를 호출한다.
			LotAcc_OC = LotAcc_OC + bino_val 				'x=0 ~ x=c 까지의 누적값을 구한다.
		Next 
	    	
	    RetOC(R) = LotAcc_OC           			 		'반환 받은 값을 배열에 보여줍니다. 
	Next

	OCMax = RetOC(1) 							'그래프 그릴때 Y축값을 최대값보다 조금 크게 한다.

'++++++++++++++++++++++++++++++++++++++++++  2.5.4 AOQ 곡선 계산함수 +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
    Dim LotAcc_AOQ 	                		'LOT의 합격확률을 구한다.
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

    DefRatioDiv4 = ScaleDefRatio4 / 250				'화면에 맞게 x축을 나누어 줍니다.
    
    AOQMax=0

    For R=0 to 250
	    Biono_DefectRatio= DefRatioDiv4 * R				'불량률입력값을 변하게 합니다.
	                 						'리턴값 = 함수이름 꼭 
	    LotAcc_AOQ = 0
	    
		For k= 0 to dAcceptCount
			a=k								'넘겨주는 인수가 for문의 변수와 같으면 Error발생.
			bino_val = Bino(dSamplesize, a, Biono_DefectRatio)        		'이항분포 함수를 호출한다.
			LotAcc_AOQ = LotAcc_AOQ + bino_val  					'x=0 ~ x=c 까지의 누적값을 구한다.
		Next 
	    	
	Select Case strReplace  				'불량품 대체 여부에 따라 계산식이 변경됩니다.
		Case 0					'불량품을 양품으로 대체하지 않는 경우 
	    	AOQ = (LotAcc_AOQ*Biono_DefectRatio*(dLotSize-dSamplesize))/((dLotSize-Biono_DefectRatio*dSamplesize)-(1-LotAcc_AOQ)*Biono_DefectRatio*(dLotSize-dSamplesize))
	    	'AOQ값을 구한다.
		Case 1					'불량품을 양품으로 대체하는 경우 
		AOQ = Biono_DefectRatio*LotAcc_AOQ*(1 - dSamplesize/dLotSize)		'AOQ값을 구한다.
	End Select
	
	RetAOQ(R) = AOQ           			 			'반환 받은 값을 배열에 보여줍니다. 
	 
	if AOQ > AOQMax then					'어레이중에서 최대값 찾기 
		AOQMax = AOQ
	End if
		
Next

'++++++++++++++++++++++++++++++++++++++++++  2.5.1 이항분포 계산함수 +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function Bino(n,a,p)    '이항분포값을 계산합니다.

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
	  
	    For i = (Biono_Samplesize - temp2 + 1) To temp1			'분자 부분의 곱셈 
	        	Multemp1 = Multemp1 * i
	    Next 
	    
	    If temp2 = 0 Then      						'분모가 0인 경우, 결과를 1로 표시합니다.
	        	Comb_val = 1
	    Else	        
	        For j = 1 To temp2
	            	Multemp2 = Multemp2 * j
	        Next 
	        	Comb_val = Multemp1 / Multemp2
	    End If
	    
	    Bino = Comb_val * (temp3 ^ temp2) * ((1 - temp3) ^ (temp1- temp2)) 	'함수이름 = 수식 꼭 
    	    
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
		
	'차트FX1 - ATI곡선 그리기 
	With Parent.frm1.ChartFX1
		.Title_(2) = "ATI곡선"
		.Gallery = <% = LINES%>
		.Axis(<%=AXIS_Y%>).Max = <% = ATIMax %>
		.Axis(<%=AXIS_Y%>).Decimals = 0
		
		.MarkerShape = <%=MK_NONE%>
		.AXIS(<%=AXIS_X%>).PixPerUnit = 1						'X축의 간격을 0으로 만드는 것 같음.
		
		.OpenDataEx <%=COD_VALUES%>, 1, 251				'차트 FX와의 데이터 채널 열어주기 
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
		
		'단위 % --> * 100
		.Axis(<%=AXIS_X%>).Label(0)  = "<%=UniNumClientFormat(0, 2, 0)%>"						
		.Axis(<%=AXIS_X%>).Label(50) = "<%=UniNumClientFormat(DefRatioDiv * 50 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(100) = "<%=UniNumClientFormat(DefRatioDiv * 100 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(150) = "<%=UniNumClientFormat(DefRatioDiv * 150 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(200) = "<%=UniNumClientFormat(DefRatioDiv * 200 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(250) = "<%=UniNumClientFormat(DefRatioDiv * 250 * 100, 2, 0)%>"
		
	End With
	
	
	'차트FX2 - ASN곡선 그리기 
	With Parent.frm1.ChartFX2
		.Title_(2) = "ASN 곡선"
		.Gallery = <% = LINES%>
		.Axis(<%=AXIS_Y%>).Max = <% = ASNMax %>
		.Axis(<%=AXIS_Y%>).Step = <% = ASNMax %> / 6
		.Axis(<%=AXIS_Y%>).Decimals = 0
		
		.MarkerShape = <%=MK_NONE%>
		.AXIS(<%=AXIS_X%>).PixPerUnit = 1						'X축의 간격을 0으로 만드는 것 같음.
		
		.OpenDataEx <%=COD_VALUES%>, 1, 251					'차트 FX와의 데이터 채널 열어주기 
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
		
		'단위 % --> * 100
		.Axis(<%=AXIS_X%>).Label(0)  = "<%=UniNumClientFormat(0, 2, 0)%>"						
		.Axis(<%=AXIS_X%>).Label(50) = "<%=UniNumClientFormat(DefRatioDiv2 * 50 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(100) = "<%=UniNumClientFormat(DefRatioDiv2 * 100 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(150) = "<%=UniNumClientFormat(DefRatioDiv2 * 150 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(200) = "<%=UniNumClientFormat(DefRatioDiv2 * 200 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(250) = "<%=UniNumClientFormat(DefRatioDiv2 * 250 * 100, 2, 0)%>"
		
	End With
	
	'차트FX3 - OC곡선 그리기 
	With Parent.frm1.ChartFX3
		.Title_(2) = "OC 곡선"
		.Gallery = <% = LINES%>
		 .Axis(<%=AXIS_Y%>).Max = <% = OCMax %>
		
		.MarkerShape = <%=MK_NONE%>
		.AXIS(<%=AXIS_X%>).PixPerUnit = 1						'X축의 간격을 0으로 만드는 것 같음.
		
		.OpenDataEx <%=COD_VALUES%>, 1, 251				'차트 FX와의 데이터 채널 열어주기 
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
		'단위 % --> * 100
		.Axis(<%=AXIS_X%>).TickMark = <%=TS_NONE%>
		.Axis(<%=AXIS_X%>).Label(0)  = "<%=UniNumClientFormat(0, 2, 0)%>"						
		.Axis(<%=AXIS_X%>).Label(50) = "<%=UniNumClientFormat(DefRatioDiv3 * 50 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(100) = "<%=UniNumClientFormat(DefRatioDiv3 * 100 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(150) = "<%=UniNumClientFormat(DefRatioDiv3 * 150 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(200) = "<%=UniNumClientFormat(DefRatioDiv3 * 200 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(250) = "<%=UniNumClientFormat(DefRatioDiv3 * 250 * 100, 2, 0)%>"
		
	End With

	'차트FX4 - AOQ곡선 그리기 
	With Parent.frm1.ChartFX4
		.Title_(2) = "AOQ 곡선"
		.Gallery = <% = LINES%>
		.Axis(<%=AXIS_Y%>).Max = <% = AOQMax %>
		.Axis(<%=AXIS_Y%>).Step = <% = AOQMax %> / 6
		.Axis(<%=AXIS_Y%>).Decimals = 4							'AOQ의 경우 최대값이 너무 작아서 소수는 4자리까지 표현한다.
		 
		.MarkerShape = <%=MK_NONE%>
		.AXIS(<%=AXIS_X%>).PixPerUnit = 0						'X축의 간격을 0으로 만드는 것 같음.
		
		.OpenDataEx <%=COD_VALUES%>, 1, 251				'차트 FX와의 데이터 채널 열어주기 
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
		'단위 % --> * 100
		.Axis(<%=AXIS_X%>).Label(0)  = "<%=UniNumClientFormat(0, 2, 0)%>"						
		.Axis(<%=AXIS_X%>).Label(50) = "<%=UniNumClientFormat(DefRatioDiv4 * 50 * 100, 2, 0)%>"	
		.Axis(<%=AXIS_X%>).Label(100) = "<%=UniNumClientFormat(DefRatioDiv4 * 100 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(150) = "<%=UniNumClientFormat(DefRatioDiv4 * 150 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(200) = "<%=UniNumClientFormat(DefRatioDiv4 * 200 * 100, 2, 0)%>"
		.Axis(<%=AXIS_X%>).Label(250) = "<%=UniNumClientFormat(DefRatioDiv4 * 250 * 100, 2, 0)%>"
		
	End With
		
</Script>	
