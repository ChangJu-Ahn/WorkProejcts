<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect 
'*  2. Function Name        : Master Data(Minor Code���)
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             +B25011ManagePlant
'                             +B25011ManagePlant
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2002/07/10
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%													                    '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd

On Error Resume Next									                '��: 
Err.Clear												                '��: Protect system from crashing

Dim strMode																'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strMajor

Dim E1_b_major_array  
Dim E1_b_major_array1  
        
Dim I1_select_char
Dim I2_major_cd
        
Const B382_E1_major_cd = 0
Const B382_E1_major_nm = 1
Const B382_E1_minor_len = 2
Const B382_E1_type = 3

Dim lg_major_major_cd
Dim lg_major_major_nm
Dim lg_C_major_minor_len
Dim lg_major_type

Call LoadBasisGlobalInf()

strMode = Request("txtMode")											'�� : ���� ���¸� ���� 

Select Case strMode

    Case CStr(UID_M0003)													'��: ���� ��ȸ/Prev/Next ��û�� ���� 
    
    	strMajor = Request("txtMajor")

        Dim ObjPB2G032
        ReDim E1_b_major_array(B382_E1_type)

	    I1_select_char = "P"
	    I2_major_cd = strMajor
	    
        Set ObjPB2G032 = server.CreateObject ("PB2G032.cBLookPreNextMajor")    
        on error resume next
        Err.Clear 
        E1_b_major_array = ObjPB2G032.B_LOOKUP_PRE_NEXT_MAJOR(gStrGlobalCollection,I1_select_char,I2_major_cd)
        Set ObjPB2G032 = nothing

        If CheckSYSTEMError(Err,True) = True Then                                   'Major�ڵ������� �����ϴ�.            
%>
<Script Language=vbscript>
	parent.frm1.txtMajorNm.value = ""
	parent.frm1.txtMinorLen.value = ""
	parent.frm1.txtMajor.focus
</Script>
<%  
        
	    	Response.End                                                            'ó�� �������Դϴ�.
        End If   
        on error goto 0
        
    lg_major_major_cd = E1_b_major_array(B386_E1_major_cd)
    lg_major_major_nm = E1_b_major_array(B386_E1_major_nm)
    lg_C_major_minor_len = E1_b_major_array(B386_E1_minor_len)
    lg_major_type = E1_b_major_array(B386_E1_type)

%>
<Script Language=vbscript>
	
	With parent.frm1
	    																	'��: ȭ�� ó�� ASP �� ��Ī�� 
    .txtMajor.value = "<%=ConvSPChars(lg_major_major_cd)%>"
	'.txtMajorNm.value = "<%=ConvSPChars(lg_major_major_nm)%>"
	'.txtMinorLen.value = "<%=lg_C_major_minor_len%>"
 
     parent.dbPrevNextOk
    
    End With
</Script>	

<%    
    Response.End
    
    Case CStr(UID_M0004)													'��: ���� ��ȸ/Prev/Next ��û�� ���� 
    
    	strMajor = Request("txtMajor")

        Dim Obj2PB2G032
        ReDim E1_b_major_array1(B382_E1_type)
        
	    I1_select_char = "N"
	    I2_major_cd = strMajor
	    
        Set Obj2PB2G032 = server.CreateObject ("PB2G032.cBLookPreNextMajor")    
        on error resume next
        Err.Clear 
        E1_b_major_array1 = Obj2PB2G032.B_LOOKUP_PRE_NEXT_MAJOR(gStrGlobalCollection,I1_select_char,I2_major_cd)
        Set Obj2PB2G032 = nothing

        If CheckSYSTEMError(Err,True) = True Then                                   'Major�ڵ������� �����ϴ�.            
%>
<Script Language=vbscript>
	parent.frm1.txtMajorNm.value = ""
	parent.frm1.txtMinorLen.value = ""
	parent.frm1.txtMajor.focus
</Script>
<%  
	    	
	    	Response.End                                                            'ó�� �������Դϴ�.
        End If   
        on error goto 0

    lg_major_major_cd = E1_b_major_array1(B386_E1_major_cd)
    lg_major_major_nm = E1_b_major_array1(B386_E1_major_nm)
    lg_C_major_minor_len = E1_b_major_array1(B386_E1_minor_len)
    lg_major_type = E1_b_major_array1(B386_E1_type)

%>
<Script Language=vbscript>
	
	With parent.frm1
	    																	'��: ȭ�� ó�� ASP �� ��Ī�� 
    .txtMajor.value = "<%=ConvSPChars(lg_major_major_cd)%>"
    
    parent.dbPrevNextOk
    
    End With
</Script>	

<%    
    End Select
   
%>
