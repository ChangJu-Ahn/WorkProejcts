<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Reference등록)
'*  3. Program ID           : B1a05mb1.asp
'*  4. Program Name         : B1a05mb1.asp
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             +B25011ManagePlant
'                             +B25011ManagePlant
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 2000/09/20
'*  8. Modified date(Last)  : 2002/12/16
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd

On Error Resume Next									'☜: 
'Err.Clear												'☜: Protect system from crashing
Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strSpread
Dim strMajor, strMinor

Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount   

Call LoadBasisGlobalInf()

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strSpread = Trim(Request("txtSpread"))

Select Case strMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
    Const B398_EG1_E1_seq_no = 0  
    Const B398_EG1_E1_reference = 1
    Const B398_EG1_E1_ref_type = 2

	Dim ObjPB2G121
    Dim I1_b_major
    Dim I2_b_minor
	Dim Export_Array1

	strMajor = Request("txtMajor")
	strMinor = Request("txtMinor")

    I1_b_major =  Trim(strMajor)
    I2_b_minor = Trim(strMinor)    
    
    Set ObjPB2G121 = server.CreateObject("PB2G121.cBListConfiguration")    
    on error resume next
    Err.Clear 
    Export_Array1 = ObjPB2G121.B_LIST_CONFIGURATION(gStrGlobalCollection,I1_b_major,I2_b_minor)
    Set ObjPB2G121 = nothing

    If CheckSYSTEMError(Err,True) = True Then              
        Response.End  
    End If
    on error goto 0
    
%>
<Script Language=vbscript>
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp
    Dim strData
	
	With parent		
		 LngMaxRow = 0
<%      
        If isEmpty(Export_Array1) Then
            GroupCount = -1
        Else
            GroupCount = Ubound(Export_Array1,1)
        End If
        
	    For LngRow = 0 To GroupCount
%>        
            strData = strData & Chr(11) & "<%=strMinor%>"
            strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array1(LngRow,B398_EG1_E1_seq_no)))%>"
            strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array1(LngRow,B398_EG1_E1_reference)))%>"
                        
            If "<%=ConvSPChars(Trim(Export_Array1(LngRow,B398_EG1_E1_ref_type)))%>" = "S" Then
                strData = strData & Chr(11) & "시스템 정의"
            Else
                strData = strData & Chr(11) & "사용자 정의"
            End If
            strData = strData & Chr(11) & LngMaxRow + <%=LngRow%> + 1 
            strData = strData & Chr(11) & Chr(12)
<%      		
       Next
%>    		

	    .ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowData strData

		.DbQueryOkFinal()
			
	End With
</Script>	
<%    

Case CStr(UID_M0002)																'☜: 다음Data조회요청을 받음 

    If Request("txtMaxRows") = "" Then
		Call DisplayMsgBox("700117", vbInformation, "", "", I_MKSCRIPT)
		Response.End 
	End If
	

    Dim Obj2PB2G121
    Dim iErrorPosition

    Set Obj2PB2G121 = server.CreateObject ("PB2G121.cBCtlConfiguration")    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear        
    Call Obj2PB2G121.B_CONTROL_CONFIGURATION(gStrGlobalCollection,strSpread,iErrorPosition)
    Set Obj2PB2G121 = nothing

    If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
		Response.End 
    End If
    on error goto 0                                                             
%>
<Script Language=vbscript>
	With parent																		'☜: 화면 처리 ASP 를 지칭함 
		.DbSaveOk
	End With
</Script>
<%
End Select
%>
