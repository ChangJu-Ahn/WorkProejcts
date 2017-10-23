<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%
	call LoadBasisGlobalInf()

'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m5211mb5.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 수입 B/L 회계처리 확정 Transaction 처리용 ASP								*
'*  7. Modified date(First) : 2000/04/21																*
'*  8. Modified date(Last)  : 2003/05/26																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/21 : Coding Start												*
'********************************************************************************************************
 Dim lgOpModeCRUD
    
 Call HideStatusWnd

 lgOpModeCRUD  = Request("txtMode") 
                             
 Select Case lgOpModeCRUD
  Case CStr(5211)                '☜: 현재 Save 요청을 받음 
     Call SubRelease()
 End Select 
'============================================================================================================
' Name : SubRelease
' Desc : 확정,확정취소 요청을 받음 
'============================================================================================================
Sub SubRelease()
 
 On Error Resume next
 Err.Clear 
 
 Dim iPM8G211
 Dim L_SelectChar
 Dim I3_m_batch_ap_post_wks
 Dim pvCB

 Dim IG1_imp_dtl_group    '☜: Protect system from crashing
	
	Const M557_I3_ap_dt_type = 0
    Const M557_I3_ap_dt = 1
    Const M557_I3_import_flg = 2
    
    Const M557_IG1_I1_count = 0
    Const M557_IG1_I2_iv_no = 1
    Const M557_IG1_I3_ap_dt = 2
    
	'--------------------
	'수정(2003.06.09)	    
	ReDim IG1_imp_dtl_group(0, 2)
	    
	IG1_imp_dtl_group(0, M557_IG1_I2_iv_no)  = Trim(Request("txtIvNo"))
	IG1_imp_dtl_group(0, M557_IG1_I1_count) = 1
	IG1_imp_dtl_group(0, M557_IG1_I3_ap_dt) = UNIConvDate(Trim(Request("txtBlIssueDt")))

	ReDim I3_m_batch_ap_post_wks(2)
	I3_m_batch_ap_post_wks(M557_I3_ap_dt_type) = ""
	'I3_m_batch_ap_post_wks(M557_I3_ap_dt) = UNIConvDate(Trim(Request("txtBlIssueDt")))
	I3_m_batch_ap_post_wks(M557_I3_import_flg) = "Y"
    
    
 IF Request("txtPost") = "D" then
	L_SelectChar  = "N"
 Else
	L_SelectChar  = "Y"
 End if
 
 pvCB = "F"
 	
 Set iPM8G211 = server.CreateObject("PM8G211.cMPostApS")  
 
 If CheckSYSTEMError(Err,True) = true Then   
	Set iPM8G211 = Nothing
	Exit Sub              
 End if
   
 Call iPM8G211.M_POST_AP_SVR(pvCB, gStrGlobalCollection, L_SelectChar, IG1_imp_dtl_group, I3_m_batch_ap_post_wks)

 If CheckSYSTEMError2(Err,True, "","","","","") = True Then
    Set iPM8G211 = Nothing
     Response.Write "<Script Language=VBScript>" & vbCr
     Response.Write "parent.frm1.btnPosting.disabled = False" & vbCr
     Response.Write "</Script>" & vbCr
    Exit Sub
 End If

  Response.Write "<Script language=vbs> " & vbCr     
  Response.Write "		Parent.InitVariables() "      & vbCr  
  Response.Write "		Parent.MainQuery() "      & vbCr       '☜: 화면 처리 ASP 를 지칭함 
  Response.Write "</Script> " & vbCr

 Set iPM8G211 = Nothing

End sub   
%>
