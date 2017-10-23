<%@ LANGUAGE="VBSCRIPT" %>
<% Option explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M1212MB1
'*  4. Program Name         : 공급처칼렌다생성 
'*  5. Program Desc         : 공급처칼렌다생성 
'*  6. Component List       : PM1G222.cMbatchSpplCalSvr
'*  7. Modified date(First) : 2001/01/16
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Min, Hak-jun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->

<%
	Call LoadBasisGlobalInf()
	
   	Dim PM1G222																'☆ : 입력/수정용 ComProxy Dll 사용 변수 
	DIM imp_group
	dim imp_m_user_id
	dim imp_year_m_str_wks
	dim imp_b_biz_partner	
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
    
	imp_m_user_id = Request("txtInsrtUserId")
	imp_year_m_str_wks =  Request("cboYear")
	imp_b_biz_partner =   Trim(Request("txtBpCd"))
	
	ReDim imp_group(7)
	imp_group(1) =  Request("chkSun")
    imp_group(2) =  Request("chkMon")
    imp_group(3) =  Request("chkTue")
    imp_group(4) =  Request("chkWed")
    imp_group(5) =  Request("chkThu")
    imp_group(6) =  Request("chkFri")
    imp_group(7) =  Request("chkSat") 
    
    Set PM1G222 =  CreateObject("PM1G222.cMbatchSpplCalSvr")
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true then 					
		Set PM1G222 = Nothing
		Response.END
	End If
	
    CALL PM1G222.BATCH_SPPL_CAL_SVR(gStrGlobalCollection, cstr(imp_year_m_str_wks), cstr(imp_m_user_id), imp_group, cstr(imp_b_biz_partner))
   
   If CheckSYSTEMError2(Err,True,"","","","","") = true then 	
		Set PM1G222 = Nothing
		Response.END
	End If
	
	Set PM1G222 = Nothing
  
    Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent "               & vbCr
    Response.Write " .GenOk "		    	    & vbCr 
	Response.Write "End With" & vbCr
	Response.Write "</Script>" & vbCr

%>
