<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 판매계획관리 
'*  3. Program ID           : S2211MB1
'*  4. Program Name         : 판매계획환경설정 
'*  5. Program Desc         :
'*  6. Comproxy List        : PS2G211.dll
'*  7. Modified date(First) : 2003/01/07
'*  8. Modified date(Last)  : 2003/02/12
'*  9. Modifier (First)     : Park yongsik
'* 10. Modifier (Last)      : Hwang Seong Bae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
	Call LoadBasisGlobalInf()
	Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
%>

<%													
	Dim iStrMode

	On Error Resume Next														
	Err.Clear

	Call HideStatusWnd
	 
	iStrMode = Request("txtMode")												'☜ : 현재 상태를 받음 

	Select Case iStrMode
	Case CStr(UID_M0001)
		Call SubBizQuery()
	Case CStr(UID_M0002)		
		Call SubBizSave()
	Case CStr(UID_M0003)		
		Call SubBizDelete()
	End Select
	
'===============================================================
' Name	: SubBizQuery
' Desc	: Query Data from DB
'===============================================================
Sub SubBizQuery()
	
	Dim iObjPS2G211
	Dim iArrRsOut
	
    On Error Resume Next

    Err.Clear 
    
    Set iObjPS2G211 = Server.CreateObject("PS2G211.cGetSSpConfig")
    
    If CheckSYSTEMError(Err,True) = True Then
        Exit Sub
    End If

    Call iObjPS2G211.ListRow(gStrGlobalCollection, Request("txtSpType"), iArrRsOut)
        
    If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set iObjPS2G211 = Nothing												'☜: ComProxy Unload
        Exit Sub
	End if
	
	Set iObjPS2G211 = Nothing

	If UBound(iArrRsOut) < 0 Then ' 판매계획 환경설정정보가 없을 경우 
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "Call Parent.DisplayMsgBox(""800076"", ""X"", ""X"", ""X"")" & vbCr   
		Response.Write "</Script>" & vbCr
		Exit Sub
	end if
	
	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent.frm1" & vbCr
	Response.Write ".cboSpType.Value = """ & iArrRsOut(0,0) & """" & vbCr			' 판매계획구분 
	Response.Write ".txtFixedInterval.text = """ & iArrRsOut(1,0) & """" & vbCr		' 확정구간 
	Response.Write ".txtFcInterval.text = """ & iArrRsOut(2,0) & """" & vbCr		' 예시구간 
	Response.Write ".cboDistrMethodCfm.Value = """ & iArrRsOut(3,0) & """" & vbCr	' 배분방법(확정)
	Response.Write ".cboDistrMethodFc.Value	= """ & iArrRsOut(4,0) & """" & vbCr	' 배분방법(예시)
	Response.Write ".cboPmRmnQty.Value = """ & iArrRsOut(5,0) & """" & vbCr			' 잔량처리방법 
	Response.Write ".cboPriceRule.Value	= """ & iArrRsOut(6,0) & """" & vbCr		' 단가 적용 규칙 
	Response.Write ".cboXchgRateFg.Value = """ & iArrRsOut(7,0) & """" & vbCr		' 환율적용규칙 
	Response.Write ".cboPmNonXchgRate.Value	= """ & iArrRsOut(8,0) & """" & vbCr	' 환율처리방법 
	
	' 프로그램 사용여부 
	if Cint(iArrRsOut(9,0)) AND 2^9 then
		Response.Write ".chkUseStep1.checked = true" & vbCr
	End if
	
	if Cint(ConvSPChars(iArrRsOut(9,0))) AND 2^12 then
		Response.Write ".chkUseStep2.checked = true" & vbCr
	End if
	
	' 전단계 범위내 수정 
	if Cint(ConvSPChars(iArrRsOut(10,0))) AND 2^9 then
		Response.Write ".chkSameQtyFlag1.checked = true" & vbCr
	End if
	
	'if Cint(ConvSPChars(iArrRsOut(10,0))) AND 2^12 then
	'	Response.Write ".chkSameQtyFlag2.checked = true" & vbCr
	'End if
	
	' 영업그룹별 진행 허용여부 
	if Cint(ConvSPChars(iArrRsOut(11,0))) AND 2^8 then
		Response.Write ".chkProcessBySg1.checked = true" & vbCr
	Else
		Response.Write ".chkProcessBySg1.checked = False" & vbCr
	End if
	
	'if Cint(ConvSPChars(iArrRsOut(11,0))) AND 2^11 then
	'	Response.Write ".chkProcessBySg2.checked = true" & vbCr
	'End if

	' 공장별 진행 여부 
	if Cint(ConvSPChars(iArrRsOut(12,0))) AND 2^10 then
		Response.Write ".chkProcessByPlant1.checked = true" & vbCr
	End if
	
	if Cint(ConvSPChars(iArrRsOut(12,0))) AND 2^13 then
		Response.Write ".chkProcessByPlant2.checked = true" & vbCr
	End if
	
	Response.Write "End With" & vbCr
	Response.Write "parent.DbQueryOk" & vbCr																'☜: 조화가 성공 
	Response.Write "</Script>" & vbCr

End Sub

'===============================================================-
' Name	: SubBizSave
' Desc	: Save Data into DB
'===============================================================
Sub SubBizSave()
	On Error Resume Next

	Dim iObjPS2G211

	Set iObjPS2G211 = Server.CreateObject("PS2G211.cMaintSSpConfig")
	
    If CheckSYSTEMError(Err,True) = True Then
        Exit Sub
    End If

	Call iObjPS2G211.Maintain(gStrGlobalCollection, Trim(request("txtSpreadIns")), Trim(request("txtSpreadUpd")), "")	

	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set iObjPS2G211 = Nothing												'☜: ComProxy Unload
        Exit Sub
	End if
	
	Set iObjPS2G211 = Nothing
	
	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "parent.DbSaveOk" & vbCr
	Response.Write "</Script>" & vbCr
End Sub

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
	
	On Error Resume Next

	Dim iObjPS2G211

	Set iObjPS2G211 = Server.CreateObject("PS2G211.cMaintSSpConfig")
	
    If CheckSYSTEMError(Err,True) = True Then
        Exit Sub
    End If

	Call iObjPS2G211.Maintain(gStrGlobalCollection, "", "", Trim(Request("txtSpreadDel")))	

	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set iObjPS2G211 = Nothing												'☜: ComProxy Unload
        Exit Sub
	End if

	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "parent.DbDeleteOk" & vbCr
	Response.Write "</Script>" & vbCr
End Sub
%>
