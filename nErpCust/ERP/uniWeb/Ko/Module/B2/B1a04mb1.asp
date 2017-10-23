<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect 
'*  2. Function Name        : Master Data(사용자Minor Code등록)
'*  3. Program ID           : B1a04mb1.asp
'*  4. Program Name         : B1a04mb1.asp
'*  5. Program Desc         :
'*  6. Comproxy List        : +B1a032LookupPreNextUdMajor
'                             +B1a041ControlUserMinorCode
'                             +B1a048ListUserMinorCode
'                             +B1a039LookupMajorCode
'*  7. Modified date(First) : 2001/07/09
'*  8. Modified date(Last)  : 2002/12/13
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

Dim strMajor, Major, sMajor
Dim strPrevKey

Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount

Call LoadBasisGlobalInf()

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strSpread = Trim(Request("txtSpread"))

Major = Request("txtMajor")
strPrevKey = Request("lgStrPrevKey")

If Major = "" Then
	Response.End
End If

If strMode <> CStr(UID_M0002) Then
	If strMode = "Q"  Or strMode = "Q2" Then		
		Call ListMinor(Major)
	Else
        Dim ObjPB2G202
        Dim E1_b_major_array  
        
        Dim I1_select_char
        Dim I2_ud_major_cd
        
        Const B391_E1_ud_major_cd = 0
        Const B391_E1_ud_major_nm = 1
        Const B391_E1_ud_minor_len = 2

        ReDim E1_b_major_array(B391_E1_ud_minor_len)

	    I1_select_char = strMode
	    I2_ud_major_cd = Major
	    
        Set ObjPB2G202 = server.CreateObject ("PB2G202.cBLookPreNextUdMajor")    
        on error resume next
        Err.Clear 
        E1_b_major_array = ObjPB2G202.B_LOOKUP_PRE_NEXT_UD_MAJOR(gStrGlobalCollection,I1_select_char,I2_ud_major_cd)
        Set ObjPB2G202 = nothing

        If CheckSYSTEMError(Err,True) = True Then                                   'Major코드정보가 없습니다.            
	    	Response.End                                                            '처음 데이터입니다.
        End If   
        on error goto 0

        sMajor = E1_b_major_array(B391_E1_ud_major_cd)
        
		Call LookupMajor(sMajor)
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

    Dim Obj2PB2G201
    Dim iErrorPosition

    Set Obj2PB2G201 = server.CreateObject ("PB2G201.cBCtlUserMinorCode")    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear        
    Call Obj2PB2G201.B_CONTROL_USER_MINOR_CODE(gStrGlobalCollection,strSpread,iErrorPosition)
    Set Obj2PB2G201 = nothing

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
    Dim I1_b_user_defined_major_cd
    Dim I2_b_user_defined_minor
    
    Const B395_I2_ud_minor_cd = 0    '[CONVERSION INFORMATION]  View Name : import b_user_defined_minor
    Const B395_I2_ud_minor_nm = 1

    Const B395_EG1_E1_ud_minor_cd = 0    '[CONVERSION INFORMATION]  View Name : export_item b_user_defined_minor
    Const B395_EG1_E1_ud_minor_nm = 1
    Const B395_EG1_E1_ud_reference = 2


	Dim ObjPB2G201
	Dim Export_Array

    ReDim I2_b_user_defined_minor(B395_I2_ud_minor_nm)
    
	I1_b_user_defined_major_cd =  strMajor   

	Call LookupMajor(strMajor)

    Set ObjPB2G201 = server.CreateObject ("PB2G201.cBListUserMinorCode")    
    on error resume next
    Err.Clear 
    Export_Array = ObjPB2G201.B_LIST_USER_MINOR_CODE (gStrGlobalCollection,I1_b_user_defined_major_cd,I2_b_user_defined_minor)
    Set ObjPB2G201 = nothing

    If CheckSYSTEMError(Err,True) = True Then                               
	   Exit Function													'☜: 비지니스 로직 처리를 종료함 
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
																'☜: 화면 처리 ASP 를 지칭함 
		 LngMaxRow = 0
		 '.frm1.vspdData.MaxRows= 0
<%      
        GroupCount = Ubound(Export_Array,1)
	    For LngRow = 0 To GroupCount
%>        
            strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B395_EG1_E1_ud_minor_cd)))%>"'
            strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B395_EG1_E1_ud_minor_nm)))%>"'
            strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B395_EG1_E1_ud_reference)))%>"'
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
    Const B393_I1_ud_major_cd = 0
    Const B393_I1_ud_major_nm = 1

    Const B393_E1_ud_major_cd = 0
    Const B393_E1_ud_major_nm = 1
    Const B393_E1_ud_minor_len = 2
    
	Dim ObjPB2G192	
	Dim I1_b_user_defined_major
	Dim E1_b_major
	
    Dim lg_major_major_cd
    Dim lg_major_major_nm
    Dim lg_C_major_minor_len
	
    ReDim I1_b_user_defined_major(B393_I1_ud_major_nm)
    ReDim E1_b_major(B393_E1_ud_minor_len)
    

    I1_b_user_defined_major(B393_I1_ud_major_cd) = strCode
	
    Set ObjPB2G192 = server.CreateObject ("PB2G192.cBLookUserMajorCode")    
    On Error Resume Next                                                                 
    Err.Clear                                                                            
    E1_b_major = ObjPB2G192.B_LOOKUP_USER_MAJOR_CODE(gStrGlobalCollection,I1_b_user_defined_major)
    Set ObjPB2G192 = nothing    
    
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

    lg_major_major_cd = E1_b_major(B393_E1_ud_major_cd)
    lg_major_major_nm = E1_b_major(B393_E1_ud_major_nm)
    lg_C_major_minor_len = E1_b_major(B393_E1_ud_minor_len)
%>
   
<Script Language=vbscript>
	parent.frm1.txtMajor.value = "<%=ConvSPChars(lg_major_major_cd)%>"
	parent.frm1.txtMajorNm.value = "<%=ConvSPChars(lg_major_major_nm)%>"
	parent.frm1.txtMinorLen.value = "<%=lg_C_major_minor_len%>"
</Script>
<%
End Function
%>
