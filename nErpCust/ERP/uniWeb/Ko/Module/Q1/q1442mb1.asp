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
'*  3. Program ID           : Q1442MB1
'*  4. Program Name         : 전문가 시스템의 그래픽을 그린다.
'*  5. Program Desc         : 
'*  6. Component List       : PQBG120
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/08/01
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Koh Jae Woo
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

Call HideStatusWnd		

Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("I", "*", "NOCOOKIE", "QB")

Dim strUpperAcceptCount
Dim strLowerAcceptCount
Dim strSD
Dim strSampleSize
Dim strInsCri

Dim dUpperAcceptCount
Dim dLowerAcceptCount
Dim dSD
Dim dSampleSize

strInsCri = Request("txtInsCri") 	

strUpperAcceptCount = Request("txtUpperAcceptCount")
strLowerAcceptCount = Request("txtLowerAcceptCount")
strSD = Request("txtSD")
strSampleSize = Request("txtSampleSize")

If strUpperAcceptCount <> "" then
	dUpperAcceptCount = UNICDbl(UNIConvNum(strUpperAcceptCount, 0), 0)
End If

If strLowerAcceptCount <> "" then
	dLowerAcceptCount = UNICDbl(UNIConvNum(strLowerAcceptCount, 0), 0)
End If

If strSD <> "" then
	dSD = UNICDbl(UNIConvNum(strSD, 0), 0)
End If

dSampleSize = UNICDbl(UNIConvNum(strSampleSize, 0), 0)

'++++++++++++++++++++++++++++++++++++++++++  2.5.3 OC 곡선 계산함수 +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
	Dim OCMax
    Dim LotAcc_OC 	                		'LOT의 합격확률을 구한다.
    Dim RetOC(600) 
    Dim OC
    Dim Zp_value
    DIm Z
    Dim Lp
    Dim m
       
    Dim ScaleDefRatio3
    Dim DefRatioDiv
    Dim UM
    Dim LM
    Dim XLabel
    Dim R
    Dim dInterval
	Dim dStart
	Dim dEnd
	Dim Xi    
	
    ScaleDefRatio3 = 0 
    DefRatioDiv = 0
	
'상한합격판정치가 주어졌을때 
IF strLowerAcceptCount = "" then
	OCMax = 0	    	    	
	DefRatioDiv = dUpperAcceptCount	
	
	dStart = DefRatioDiv - 3.2 * dSD
	
	dInterval = (DefRatioDiv - dStart) / 300 
	dStart= DefRatioDiv - dInterval * 300
	Xi = dStart
	For R = 0 To 600
	   Zp_value=((Xi-dUpperAcceptCount)*sqr(dSampleSize))/dSD
	   Lp=1-p_normal(Zp_value)
	   RetOC(R) = Lp       			 			'반환 받은 값을 배열에 보여줍니다. 
		    		    	
	   Xi = dStart + (R * dInterval)
	 		
	   If OCMax < RetOC(R) then
	   		OCMax=RetOC(R)
	   End if
	Next
	
	dEnd = dStart + (600 * dInterval)
End if	 		

'하한합격판정치가 주어졌을때 
IF strUpperAcceptCount = "" then	
	OCMax = 0	    	    	
	DefRatioDiv = dLowerAcceptCount	
	
	dStart = DefRatioDiv - 3.2 * dSD

	dInterval = (DefRatioDiv - dStart) / 300 
	dStart= DefRatioDiv - dInterval * 300
	Xi = dStart
	For R = 0 To 600
	   Zp_value=((dLowerAcceptCount-Xi)*sqr(dSampleSize))/dSD
	   Lp=1-p_normal(Zp_value)
	   RetOC(R) = Lp       			 			'반환 받은 값을 배열에 보여줍니다. 
		    		    	
	   Xi = dStart + (R * dInterval)
	 		
	   If OCMax < RetOC(R) then
			OCMax=RetOC(R)
	   End if
	Next
	
	dEnd = dStart + (600 * dInterval)
End if

IF strUpperAcceptCount <> "" and strLowerAcceptCount <> "" then
	DefRatioDiv=(dUpperAcceptCount + dLowerAcceptCount)/2
	    	    
	OCMax = 0	    
	dStart = DefRatioDiv - 3.2 * dSD

	If dLowerAcceptCount < dStart Then
		dStart = dLowerAcceptCount - 0.1 * (dUpperAcceptCount - dLowerAcceptCount)
	End If
	
	dInterval = (DefRatioDiv - dStart) / 300 
	dStart= DefRatioDiv - dInterval * 300
	Xi = dStart
	For R = 0 To 600
		If Xi <= DefRatioDiv then 
			Zp_value=((dLowerAcceptCount-Xi)*sqr(dSampleSize))/dSD
			Lp=1-p_normal(Zp_value)
			RetOC(R) = Lp       					'반환 받은 값을 배열에 보여줍니다. 
		Else
			Zp_value=((Xi-dUpperAcceptCount)*sqr(dSampleSize))/dSD
			Lp=1-p_normal(Zp_value)
			RetOC(R) = Lp       			 		'반환 받은 값을 배열에 보여줍니다. 	    
		End If
			 	
		Xi = dStart + (R * dInterval)
		
		If OCMax < RetOC(R) then
			OCMax=RetOC(R)
		End if				
	Next
	
	dEnd = dStart + (600 * dInterval)
End IF
				
'++++++++++++++++++++++++++++++++++++++++++  2.5.1 정규분포 계산함수 +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function Zp(p)

	Dim i
	Dim startval
	Dim endval
	
	startval = 0
	endval = 3
	
	If p < 0.5 Then
		startval = -3
	    	endval = 0
	End If
	
		For i = startval To endval Step 0.01
		    
		    	If p < p_normal(i) Then
		        		Zp = i
				Exit For
			End If
	
		Next 
		
End Function


Public Function p_normal(x)
	
	Dim LOG_PI
	Dim term
	Dim result
	Dim k
	   
	LOG_PI = 1.1447298858494  
	    
    		If x >= 0 Then
        			p_normal = 0.5 * (1 + p_gamma(0.5, 0.5 * x * x, LOG_PI / 2))
    		Else
        			p_normal = 0.5 * q_gamma(0.5, 0.5 * x * x, LOG_PI / 2)
    		End If
End Function


Public Function p_gamma(a, x, loggamma_a)

	Dim k 
	Dim result, term, previous 


	If x >= (1 + a) Then
    		p_gamma = 1 - q_gamma(a, x, loggamma_a)
    		Exit Function
	ElseIf x = 0 Then
    		p_gamma = 0
    		Exit Function
	End If

	term = Exp(a * Log(x) - x - loggamma_a) / a
	result = term

	For k = 1 To 1000

	    	term = term * x / (a + k)
	    	previous = result
	    	result = result + term
	    	
	    	If result = previous Then
	        		p_gamma = result
	        		Exit Function
	    	End If
	Next 

	p_gamma = result
	
End Function


Public Function q_gamma(a, x, loggamma_a)
	
	Dim k 
	Dim result, w, temp, previous
	Dim la
	Dim lb
	
	la = 1
	lb = 1 + x - a
	
	If x < (1 + a) Then
		q_gamma = 1 - p_gamma(a, x, loggamma_a)
	    	Exit Function
	End If
	
	w = Exp(a * Log(x) - x - loggamma_a)
	result = w / lb
	
	For k = 2 To 1000
	    temp = ((k - 1 - a) * (lb - la) + (k + x) * lb) / k
	    la = lb
	    lb = temp
	    w = w * (k - 1 - a) / k
	    temp = w / (la * lb)
	    previous = result
	    result = result + temp
	    
	    If result = previous Then
	    	q_gamma = result
	        	Exit Function
	    End If
	    
	Next 
	
	q_gamma = result
	
End Function

%>

<Script Language=vbscript>

		Dim OCMeas_Val
		Redim OCMeas_Val(600)
		
		
<%	
		Dim Cnt 		    
		
		For Cnt = 0 to 600
%>
		OCMeas_Val(<%=Cnt%>) = "<%=RetOC(Cnt)%>"
		
<%
 		Next
 %>

 	Dim i
		
			
	'차트FX1 - OC곡선 그리기 
		
		With Parent.frm1.ChartFX1
			
			.Title_(2) = "검사특성 그래프"
			.Gallery = <% = LINES%>
			.Axis(<%=AXIS_Y%>).Max = <% = OCMax %>
			
			.MarkerShape = <%=MK_NONE%>
			
			.AXIS(<%=AXIS_X%>).PixPerUnit = 1						'X축의 간격을 0으로 만드는 것 같음.
			
			.OpenDataEx <%=COD_VALUES%>, 1, 601				'차트 FX와의 데이터 채널 열어주기 
				For i = 0 to 600	
					.ValueEx(0, i) = OCMeas_Val(i)
				Next
			.CloseData <%=COD_VALUES%>
			
			.Series(0).MarkerShape = <%=MK_NONE%>
			.Series(0).Visible = True
			
			.Axis(<%=AXIS_Y%>).Max = .Axis(<%=AXIS_Y%>).Max * 1.1
		
			.Axis(<%=AXIS_X%>).TickMark = <%=TS_NONE%>
			.Axis(<%=AXIS_X%>).Max = <% = dEnd %>
			.Axis(<%=AXIS_X%>).Step = 50
			
			.Axis(<%=AXIS_X%>).Label(0)  = "<%=UniNumClientFormat(dStart, 4, 0)%>"
			.Axis(<%=AXIS_X%>).Label(100) = "<%=UniNumClientFormat(dStart + 100 * dInterval, 4, 0)%>"
			.Axis(<%=AXIS_X%>).Label(200) = "<%=UniNumClientFormat(dStart + 200 * dInterval, 4, 0)%>"
			.Axis(<%=AXIS_X%>).Label(300) = "<%=UniNumClientFormat(dStart + 300 * dInterval, 4, 0)%>"
			.Axis(<%=AXIS_X%>).Label(400) = "<%=UniNumClientFormat(dStart + 400 * dInterval, 4, 0)%>"	
			.Axis(<%=AXIS_X%>).Label(500) = "<%=UniNumClientFormat(dStart + 500 * dInterval, 4, 0)%>"	
			.Axis(<%=AXIS_X%>).Label(600) = "<%=UniNumClientFormat(dEnd, 4, 0)%>"						
			
		End With
	
</Script>	