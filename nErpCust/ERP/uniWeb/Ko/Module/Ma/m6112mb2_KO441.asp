<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%				
	Dim lgOpModeCRUD				'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
	On Error Resume Next			'☜: Protect system from crashing
	Err.Clear 
	
	Call LoadBasisGlobalInf()
	Call HideStatusWnd									'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
	lgOpModeCRUD  = Request("txtMode") 

	Select Case lgOpModeCRUD
	        Case CStr(UID_M0001)
	        Case CStr(UID_M0002)
	             Call SubBizSave()
		    Case CStr(UID_M0003)
	End Select
	
	Response.End 

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data into Db
'============================================================================================================

Sub SubBizSave()
	On Error Resume Next			'☜: Protect system from crashing
	Err.Clear 

	Dim iPMAG182 
	
	Dim I2_b_plant_plant_cd
    Dim I3_processs_step
    Dim I4_DisbQryDt				' -- 경비 배부대상 추출일(To)
    Dim I5_Disb_Batch_Job_Dt		' -- 경비 배부일(재고Posting일자)
    Dim I11_DisbFrQryDt				' -- 경비 배부대상 추출일(From)
    Dim E1_dist_ref_no 
            
    Dim strYear, strMonth, strDay
    Dim strFirstDay, strTempDay, strLastDay
       	
	I2_b_plant_plant_cd								= Trim(Request("txtPlantCd"))    
	I3_processs_step									= Trim(Request("txtStep"))   
	'I4_DisbQryDt									= Trim(Request("txtToDisbQryDt")) 
	I4_DisbQryDt									= UNIConvDate(Trim(Request("txtToDisbQryDt")))  'KSJ 수정 
	Call ExtractDateFrom(Request("txtDisbDt"), gDateFormatYYYYMM, gComDateType, strYear, strMonth, strDay)
	strFirstDay = UNIConvDate(UniConvYYYYMMDDToDate(gDateFormat,strYear,strMonth,"01"))
	strTempDay	=  UNIDateAdd("M",1,strFirstDay,gServerDateFormat)
	strLastDay	=  UNIDateAdd("D",-1,strTempDay,gServerDateFormat)
	I5_Disb_Batch_Job_Dt =  strLastDay
	'I11_DisbFrQryDt									= Trim(Request("txtFrDisbQryDt")) 
	I11_DisbFrQryDt									= UNIConvDate(Trim(Request("txtFrDisbQryDt")))  'KSJ 수정 
'Call ServerMesgBox(I3_processs_step, vbCritical, I_MKSCRIPT)  

    Set iPMAG182 = Server.CreateObject("PMAG182.cMMaintDistSvr")    

    If CheckSYSTEMError(Err,True) = true Then
		Exit Sub
	End If
   Call iPMAG182.M_MAINT_DISTRIBUT_SVR(gStrGlobalCollection, "D", I2_b_plant_plant_cd, I3_processs_step, I4_DisbQryDt, I5_Disb_Batch_Job_Dt, "", "", I11_DisbFrQryDt, E1_dist_ref_no)

	If CheckSYSTEMError(Err,True) = true Then
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtDistRefNo.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		Set iPMAG182 = Nothing 		
		Exit Sub
	End If
	
	Set iPMAG182 = Nothing									'☜: ComProxy Unload
'Call ServerMesgBox("sucess", vbCritical, I_MKSCRIPT)  	
	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "	parent.frm1.txtDistRefNo.value = """ & E1_dist_ref_no & """" & vbCrLf
	Response.Write "	With parent" & vbCr															
	Response.Write "		.DbSaveOk" & vbCr
	Response.Write "	End With" & vbCr
	Response.Write "</Script>" & vbCr

End Sub
%>
