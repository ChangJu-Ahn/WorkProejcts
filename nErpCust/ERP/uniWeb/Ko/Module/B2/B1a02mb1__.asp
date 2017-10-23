<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect 
'*  2. Function Name        : Master Data(Minor Code등록)
'*  3. Program ID           : B1a02mb1.asp
'*  4. Program Name         : B1a02mb1.asp
'*  5. Program Desc         :
'*  6. Comproxy List        : +B1a012LookupPreNextMajor
'                             +B1a021ControlMinorCode
'                             +B1a028ListMinorCode
'                             +B1a019LookupMajorCode
'*  7. Modified date(First) : 2000/09/19
'*  8. Modified date(Last)  : 2002/12/16
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
Call HideStatusWnd                                      '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다    

On Error Resume Next									'☜: 
'Err.Clear												'☜: Protect system from crashing

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strSpread

Dim strMajor, Major, sMajor
Dim strPrevKey

Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount

Call LoadBasisGlobalInf()

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strSpread = Request("txtSpread")

Major = Request("txtMajor")
strPrevKey = Request("lgStrPrevKey")

If Major = "" Then
	Response.End
End If

If strMode <> CStr(UID_M0002) Then
	If strMode = "Q"  Or strMode = "Q2" Then		
		Call ListMinor(Major)
	Else
        Dim ObjPB2G032
        Dim E1_b_major_array  
        
        Dim I1_select_char
        Dim I2_major_cd
        
        Const B382_E1_major_cd = 0
        Const B382_E1_major_nm = 1
        Const B382_E1_minor_len = 2
        Const B382_E1_type = 3

        ReDim E1_b_major_array(B382_E1_type)

	    I1_select_char = strMode
	    I2_major_cd = Major
	    
        Set ObjPB2G032 = server.CreateObject ("PB2G032.cBLookPreNextMajor")    
        on error resume next
        Err.Clear 
        E1_b_major_array = ObjPB2G032.B_LOOKUP_PRE_NEXT_MAJOR(gStrGlobalCollection,I1_select_char,I2_major_cd)
        Set ObjPB2G032 = nothing
        
        If (Err.number <> 0 and inStr(Err.Description ,"900011")) > 0 _
           or (Err.number <> 0 and inStr(Err.Description ,"900012")) > 0 Then
            If CheckSYSTEMError(Err,True) = True Then                                   'Major코드정보가 없습니다.            
            End If   
            
            Call LookupMajor(Major)
            
        Else
            If CheckSYSTEMError(Err,True) = True Then  
                Response.End                                  'Major코드정보가 없습니다.            
            End If   

            sMajor = E1_b_major_array(B382_E1_major_cd)
        
            Call LookupMajor(sMajor)
        End If
%>
<Script Language=vbscript>
	With parent.frm1
         parent.dbPrevNextOk
    End With
</Script>	

<%    

	End If
Else
	If Request("txtMaxRows") = "" Then
		Call DisplayMsgBox("700117", vbInformation, "", "", I_MKSCRIPT)
		Response.End 
	End If

    Dim Obj2PB2G031
    Dim iErrorPosition


    Set Obj2PB2G031 = server.CreateObject ("PB2G031.cBControlMinorCode")    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear        
    Call Obj2PB2G031.B_CONTROL_MINOR_CODE (gStrGlobalCollection,strSpread,iErrorPosition)
    Set Obj2PB2G031 = nothing

    If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
		Response.End 
    End If
    on error goto 0                                                             
%>
<Script Language=vbscript>
	With parent																		'☜: 화면 처리 ASP 를 지칭함 
		'window.status = "저장 성공"		
		.DbSaveOk
	End With
</Script>
<%					
End If
%>

<%
'==============================================================================
' Function : LookUp...
' Description : 저장시 Lookup
'==============================================================================
Function ListMinor(ByVal strMajor)
    Const B388_EG1_E1_b_minor_minor_cd = 0
    Const B388_EG1_E1_b_minor_minor_nm = 1
    Const B388_EG1_E1_b_minor_minor_type = 2


	Dim ObjPB2G031
    Dim I1_b_major_major_cd
	
	Dim Export_Array

    I1_b_major_major_cd = strMajor  
    
	Call LookupMajor(strMajor)
    
    Set ObjPB2G031 = server.CreateObject ("PB2G031.cBListMinorCode")    
    on error resume next
    Err.Clear 
    Export_Array = ObjPB2G031.B_LIST_MINOR_CODE(gStrGlobalCollection,I1_b_major_major_cd)
    Set ObjPB2G031 = nothing
    
    If CheckSYSTEMError(Err,True) = True Then                                   'Minor코드정보가 없습니다.
        Exit Function
    End If
    on error goto 0
    
%>
<Script Language=vbscript>
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp
    Dim strData
	
	With parent																	'☜: 화면 처리 ASP 를 지칭함 
		 LngMaxRow = 0
		 
<%      
        GroupCount = Ubound(Export_Array,1)
	    For LngRow = 0 To GroupCount
%>        
            strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B388_EG1_E1_b_minor_minor_cd)))%>"'MINOR_CD
            strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B388_EG1_E1_b_minor_minor_nm)))%>"'MINOR_NM
            If "<%=ConvSPChars(Trim(Export_Array(LngRow,B388_EG1_E1_b_minor_minor_type)))%>" = "S" Then
                strData = strData & Chr(11) & "시스템 정의"'MINOR_TYPE
            Else
                strData = strData & Chr(11) & "사용자 정의"'MINOR_TYPE
            End If
            strData = strData & Chr(11) & LngMaxRow + <%=LngRow%> + 1
            strData = strData & Chr(11) & Chr(12)
            
<%      		
        Next
%>    		
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData strData
		.frm1.vspdData.ReDraw = True		

		.frm1.hMajor.value = "<%=ConvSPChars(Request("txtMajor"))%>"
		.DbQueryOk
	End With
</Script>	
<%
End Function
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
	parent.frm1.txtMajorNm.value = ""
	parent.frm1.txtMinorLen.value = ""
	parent.frm1.txtMajor.focus
</Script>
<%  
    Response.End 
	End If
    on error goto 0
     
    lg_major_major_cd = E1_b_major(B386_E1_major_cd)
    lg_major_major_nm = E1_b_major(B386_E1_major_nm)
    lg_C_major_minor_len = E1_b_major(B386_E1_minor_len)
    lg_major_type = E1_b_major(B386_E1_type)
    
%>
    
<Script Language=vbscript>
	parent.frm1.txtMajor.value = "<%=ConvSPChars(lg_major_major_cd)%>"
	parent.frm1.txtMajorNm.value = "<%=ConvSPChars(lg_major_major_nm)%>"
	parent.frm1.txtMinorLen.value = "<%=lg_C_major_minor_len%>" 
	
</Script>
<% 
	If strMode = "Q" Then
%>
<Script Language=vbscript>
		parent.ggoSpread.Source = parent.frm1.vspdData
		parent.ggoSpread.SSSetEdit 1, "Minor코드", 26,,,CInt(parent.frm1.txtMinorLen.value), 2
</Script>
<%
	End If
%>
<Script Language=vbscript>
	If "<%=lg_major_type%>" = "S" Then 
		parent.frm1.rdoChargeCd2.Checked = True
	Else 
		parent.frm1.rdoChargeCd1.Checked = True
	End If 
</Script>	
<%
End Function
%>
