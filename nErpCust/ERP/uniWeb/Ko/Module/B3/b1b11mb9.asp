<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b11mb9.asp
'*  4. Program Name         : Update Item by Plant (Detail)
'*  5. Program Desc         :
'*  6. Component List       : PB3S111.cBMngItemByPltDtl 
'*  7. Modified date(First) : 2001/03/13
'*  8. Modified date(Last)  : 2002/11/14
'*  9. Modifier (First)     : Jung Yu Kyung
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next														'☜: 
Err.Clear

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim pPB3S111																	'☆ : 저장용 Component Dll 사용 변수 

Dim strPlantCd
Dim iErrorPosition

Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount

Dim ii

strPlantCd = Request("hPlantCd")
	
If strPlantCd = "" Then
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)				'⊙:
	Response.End
End If

itxtSpread = ""
             
iCUCount = Request.Form("txtCUSpread").Count
             
itxtSpreadArrCount = -1
             
ReDim itxtSpreadArr(iCUCount)
    
For ii = 1 To iCUCount
    itxtSpreadArrCount = itxtSpreadArrCount + 1
    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
Next

itxtSpread = Join(itxtSpreadArr,"")

'-----------------------
'Com action result check area(OS,internal)
'-----------------------
Set pPB3S111 = Server.CreateObject("PB3S111.cBMngItemByPltDtl")    
If CheckSYSTEMError(Err,True) = True Then
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "	Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

Call pPB3S111.B_UPDATE_ITEM_BY_PLANT_DETAIL(gStrGlobalCollection, _
											itxtSpread, _
											strPlantCd, _
											iErrorPosition)

If CheckSYSTEMError2(Err, True, iErrorPosition & "행", "", "", "", "") = True Then
	Set pPB3S111 = Nothing															'☜: Unload Component
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "	Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

Set pPB3S111 = Nothing															'☜: Unload Component

Response.Write "<Script Language=vbscript>" & vbCrLf
Response.Write "	parent.DbSaveOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End
%>
