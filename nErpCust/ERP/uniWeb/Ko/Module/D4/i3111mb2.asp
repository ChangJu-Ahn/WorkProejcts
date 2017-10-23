<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!--
'**********************************************************************************************
'*  1. Module Name          : Long-term Inv Analysis
'*  2. Function Name        : 
'*  3. Program ID           : I3111MB2
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Component List       : PI3G110
'*  7. Modified date(First) : 2006/05/25
'*  8. Modified date(Last)  : 2006/05/25
'*  9. Modifier (First)     : KiHong Han
'* 10. Modifier (Last)      : KiHong Han
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************-->
<%
	
'import I_LONGTERM_INV_ANAL_CONFG
Const I1_plant_cd = 0
Const I1_longterm_stock_cal_period = 1
Const I1_pernicious_stock_cal_period = 2
Const I1_plan_flag = 3
Const I1_plan_stock_cal_period = 4
    
Call LoadBasisGlobalInf
                                         
On Error Resume Next
Call HideStatusWnd																'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim objPI3G110																	
Dim lgIntFlgMode	
Dim iCommandSent
Dim iArrImport
		
lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
		
Set objPI3G110 = Server.CreateObject("PI3G110.cIMtLongtermInvAnalConfg")

If CheckSYSTEMError(Err,True) = True Then
   Response.End
End If
	    
If lgIntFlgMode = OPMD_CMODE Then
	iCommandSent = "CREATE"
ElseIf lgIntFlgMode = OPMD_UMODE Then
	iCommandSent = "UPDATE"
End If

Redim iArrImport(4)
	
iArrImport(I1_plant_cd) = UCASE(Request("txtPlantCd2"))
iArrImport(I1_longterm_stock_cal_period) = Request("txtLongtermStockCalPeriod")
iArrImport(I1_pernicious_stock_cal_period) = Request("txtPerniciousStockCalPeriod")
iArrImport(I1_plan_flag) = ""
iArrImport(I1_plan_stock_cal_period) = ""
'iArrImport(I1_plan_flag) = Request("cboplanflag") 
'iArrImport(I1_plan_stock_cal_period) = Request("txtplanStockCalPeriod") 
		
Call objPI3G110.I_MT_LONGTERM_INV_ANAL_CONFG_SVR(gStrGlobalCollection, iCommandSent, iArrImport)

If CheckSYSTEMError(Err,True) = true Then
   Set objPI3G110 = Nothing
   Response.End
End If

Set objPI3G110 = Nothing
%>
<Script Language=vbscript>
With parent																			
	.DbSaveOk
End With
</Script>
