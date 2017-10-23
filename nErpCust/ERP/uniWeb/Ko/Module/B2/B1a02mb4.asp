<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect 
'*  2. Function Name        : Master Data(User Defined Minor Code등록)
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             +B25011ManagePlant
'                             +B25011ManagePlant
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2002/12/10
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************

%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd

On Error Resume Next									'☜: 
Err.Clear												'☜: Protect system from crashing

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strSpread

Dim strMajor, strPrevKey

Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount          

Call LoadBasisGlobalInf()

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strSpread = Request("txtSpread")

Select Case strMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 

	strMajor = Request("txtMajor")
	Call LookUpMajor(strMajor)
	
    Const B388_EG1_E1_b_minor_minor_cd = 0
    Const B388_EG1_E1_b_minor_minor_nm = 1
    Const B388_EG1_E1_b_minor_minor_type = 2


	Dim ObjPB2G031
    Dim I1_b_major_major_cd
	
	Dim Export_Array

    I1_b_major_major_cd = Trim(strMajor) 
    
    Set ObjPB2G031 = server.CreateObject ("PB2G031.cBListMinorCode")    
    on error resume next
    Err.Clear 
    Export_Array = ObjPB2G031.B_LIST_MINOR_CODE(gStrGlobalCollection,I1_b_major_major_cd)
    Set ObjPB2G031 = nothing

    If CheckSYSTEMError(Err,True) = True Then                                   'Minor코드정보가 없습니다.
        Response.End
    End If
    on error goto 0

	Call LookupMajor(strMajor)
    
%>
<Script Language=vbscript>
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp
    Dim strData1
    Dim strData2
	Dim iRow1
	Dim iRow2
	
	With parent																	'☜: 화면 처리 ASP 를 지칭함 
		LngMaxRow = 0										'Save previous Maxrow
		iRow1 = 0
		iRow2 = 0
<%      
        GroupCount = Ubound(Export_Array,1)
	    For LngRow = 0 To GroupCount
%>      
		If "<%=ConvSPChars(Trim(Export_Array(LngRow,2)))%>" = "S" Then
            iRow1 = iRow1 + 1
            strData1 = strData1 & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B388_EG1_E1_b_minor_minor_cd)))%>"'MINOR_CD
            strData1 = strData1 & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B388_EG1_E1_b_minor_minor_nm)))%>"'MINOR_NM
			strData1 = strData1 & Chr(11) & "S"	'3			
			strData1 = strData1 & Chr(11) & LngMaxRow + iRow1				'5
			strData1 = strData1 & Chr(11) & Chr(12)											
		Else
		    iRow2 = iRow2 + 1
            strData2 = strData2 & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B388_EG1_E1_b_minor_minor_cd)))%>"'MINOR_CD
            strData2 = strData2 & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B388_EG1_E1_b_minor_minor_nm)))%>"'MINOR_NM
			strData2 = strData2 & Chr(11) & "U"	'3			
			strData2 = strData2 & Chr(11) & LngMaxRow + iRow2				'5
			strData2 = strData2 & Chr(11) & Chr(12)											
		End if

<%      
    Next
%>    
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowData strData1
		
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowData strData2
		
		.frm1.hMajor.value = "<%=ConvSPChars(Request("txtMajor"))%>"
		.DbQueryOk(LngMaxRow + 1)
				
	End With
</Script>	
<%    

Case CStr(UID_M0002)																'☜: 다음Data조회요청을 받음 
									
    Err.Clear																		'☜: Protect system from crashing
	If Request("txtMaxRows") = "" Then
		Call DisplayMsgBox("700117", vbInformation, "", "", I_MKSCRIPT)
		Response.End 
	End If

    Dim Obj2PB2G031
    Dim iErrorPosition

    Set Obj2PB2G031 = server.CreateObject ("PB2G031.cBControlMinorCode")    

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear        
    Call Obj2PB2G031.B_CONTROL_MINOR_CODE(gStrGlobalCollection,strSpread,iErrorPosition)
    Set Obj2PB2G031 = nothing

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

<%
'==============================================================================
' Function : LookUp...
' Description : 저장시 Lookup
'==============================================================================
Function LookUpMajor(Byval strCode)
    Const B386_I1_major_cd = 0
    Const B386_I1_major_nm = 1
    
    Const B386_E1_major_cd = 0
    Const B386_E1_major_nm = 1
    Const B386_E1_minor_len = 2
    Const B386_E1_type = 3

	Dim ObjPB2S012		
	Dim I1_b_major
	Dim E1_b_major
	
    Dim lg_major_major_cd
    Dim lg_major_major_nm
    Dim lg_C_major_minor_len
    Dim lg_major_type

    ReDim I1_b_major(B386_I1_major_nm)
    ReDim E1_b_major(B386_E1_type)
    
    I1_b_major(B386_I1_major_cd) = strCode

    Set ObjPB2S012 = server.CreateObject ("PB2S012.cBLookMajorCode")    
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status
    E1_b_major = ObjPB2S012.B_LOOKUP_MAJOR_CODE(gStrGlobalCollection,I1_b_major)
    Set ObjPB2S012 = nothing    

    If CheckSYSTEMError(Err,True) = True Then                                              
%>
<Script Language=vbscript>
    
    With parent.frm1
    
	.txtMajorNm.value = ""
	.txtMajor.focus
	.txtMajor2.value = ""
	.txtMajorNm2.value = ""
	.txtMinLen.value = ""
	End With
</Script>
<%  
        Response.End 
	End If
    on error goto 0

    lg_major_major_cd = E1_b_major(B386_E1_major_cd)
    lg_major_major_nm = E1_b_major(B386_E1_major_nm)
    lg_C_major_minor_len = E1_b_major(B386_E1_minor_len)
    lg_major_type = E1_b_major(B386_E1_type)

	If lg_major_type = "S" Then 
%>
<Script Language=vbscript>
    
    Call parent.DisplayMsgBox("122400", "X", "X", "X")    

    With parent.frm1
    
	.txtMajorNm.value = ""
	.txtMajor.focus
	.txtMajor2.value = ""
	.txtMajorNm2.value = ""
	.txtMinLen.value = ""
	End With
</Script>
<%  
        Response.End 
    End If
%>

<Script Language=vbscript>
    With parent.frm1
    
	.txtMajor.value = "<%=ConvSPChars(lg_major_major_cd)%>"
	.txtMajorNm.value = "<%=ConvSPChars(lg_major_major_nm)%>"
	.txtMajor2.value = "<%=ConvSPChars(lg_major_major_cd)%>"
	.txtMajorNm2.value = "<%=ConvSPChars(lg_major_major_nm)%>"
	.txtMinLen.value = "<%=lg_C_major_minor_len%>" 
	End With
	
	If "<%=lg_major_type%>" = "S" Then 
		parent.frm1.rdoChargeCd2.Checked = True
	Else 
		parent.frm1.rdoChargeCd1.Checked = True
	End If 
	
</Script>	
<%
End Function
%>
