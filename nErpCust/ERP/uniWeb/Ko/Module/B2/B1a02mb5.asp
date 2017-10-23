<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect 
'*  2. Function Name        : Master Data(User Defined Minor Code등록)
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +B1a013LookupPNMajorForUser
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2002/12/10
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%													                    '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd

On Error Resume Next									                '☜: 
Err.Clear												                '☜: Protect system from crashing

Dim strMode																'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strMajor

Dim ObjPB2G041
Dim E1_b_major_array  
        
Dim I1_ief_supplied_select_char
Dim I2_b_major_major_cd
        
Const B383_E1_major_cd = 0
Const B383_E1_major_nm = 1
Const B383_E1_minor_len = 2
Const B383_E1_type = 3

ReDim E1_b_major_array(B383_E1_type)

Call LoadBasisGlobalInf()

strMode = Request("txtMode")											'☜ : 현재 상태를 받음 
strMajor = Request("txtMajor")

Select Case strMode
    Case CStr(UID_M0003)	
	    I1_ief_supplied_select_char = "P"
	    I2_b_major_major_cd = strMajor
	    
        Set ObjPB2G041 = server.CreateObject ("PB2G041.cBLookPNMajorForUser")     
        on error resume next
        Err.Clear 
        E1_b_major_array = ObjPB2G041.B_LOOKUP_P_N_MAJOR_FOR_USER(gStrGlobalCollection,I1_ief_supplied_select_char,I2_b_major_major_cd)
        Set ObjPB2G041 = nothing

        If CheckSYSTEMError(Err,True) = True Then                                   'Major코드정보가 없습니다.            
	    	on error goto 0                                                         '마지막 데이터입니다.    
	    	Response.End                                                            '처음 데이터입니다.
        End If   
        on error goto 0

        strMajor = E1_b_major_array(B383_E1_major_cd)
        
        Call LookupMajor(strMajor)
%>
<Script Language=vbscript>
	With parent.frm1
         parent.dbPrevNextOk
    End With
</Script>	

<%    
    Case CStr(UID_M0004)	
        I1_ief_supplied_select_char = "N"
	    I2_b_major_major_cd = strMajor
	    
        Set ObjPB2G041 = server.CreateObject ("PB2G041.cBLookPNMajorForUser")    
        on error resume next
        Err.Clear 
        E1_b_major_array = ObjPB2G041.B_LOOKUP_P_N_MAJOR_FOR_USER(gStrGlobalCollection,I1_ief_supplied_select_char,I2_b_major_major_cd)
        Set ObjPB2G041 = nothing

        If CheckSYSTEMError(Err,True) = True Then                                   'Major코드정보가 없습니다.            
	    	on error goto 0                                                         '마지막 데이터입니다.    
	    	Response.End                                                            '처음 데이터입니다.
        End If   
        on error goto 0

        strMajor = E1_b_major_array(B382_E1_major_cd)
        
        Call LookupMajor(strMajor)
%>
<Script Language=vbscript>
	With parent.frm1
         parent.dbPrevNextOk
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
	
	.txtMajorNm2.value = ""
	.txtMinLen.value = ""
    End With	
</Script>
<%  
        Exit Function
	End If
    on error goto 0

    lg_major_major_cd = E1_b_major(B386_E1_major_cd)
    lg_major_major_nm = E1_b_major(B386_E1_major_nm)
    lg_C_major_minor_len = E1_b_major(B386_E1_minor_len)
    lg_major_type = E1_b_major(B386_E1_type)
%>
    
<Script Language=vbscript>
	With parent.frm1
    .txtMajor.value = "<%=ConvSPChars(lg_major_major_cd)%>"
	.txtMajorNm.value = "<%=ConvSPChars(lg_major_major_nm)%>"
	
	.txtMajor2.value = "<%=ConvSPChars(lg_major_major_cd)%>"
	.txtMajorNm2.value = "<%=ConvSPChars(lg_major_major_nm)%>"
	.txtMinLen.value = "<%=lg_C_major_minor_len%>"
    End With	
	
</Script>
<%
End Function
%>
