<%Option Explicit
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S2211MB2
'*  4. Program Name         : 판매계획단위별올림정보 
'*  5. Program Desc         : 판매계획단위별올림정보 
'*  6. Comproxy List        : PS2G212.dll
'*  7. Modified date(First) : 2003/1/7
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : Heeyoung Lee
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%Call LoadBasisGlobalInf()
Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB") 	

On Error Resume Next														

Call HideStatusWnd

Const	C_LocExpFlag	= 1
Const	C_FrSpPeriod	= 3
Const	C_ToSpPeriod	= 4
Const	C_SalesGrp		= 6
Const	C_SoldToParty	= 7
Const	C_ItemCd		= 8

Dim iStrMode
Dim iStrData
Dim iObjPS2G212
Dim iArrListOut			' Result of recordset.getrow(), it means iArrListOut is two dimension array (column, row)
Dim iLngRow
Dim iLngLastRow			' The last row number in the spread
Dim iLngSheetMaxRows	' Row numbers to be displayed in the spread.
Dim iLngPageNo
Dim iLngErrorPosition
Dim iStrPrevKey

iStrMode = Request("txtMode")												'☜ : 현재 상태를 받음 

Select Case iStrMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
	err.Clear
	iLngSheetMaxRows = CLng(100)
	iLngLastRow = CLng(Request("txtSheetLastRow"))
	
	iLngPageNo = Request("lgPageNo")
	
    Set iObjPS2G212 = Server.CreateObject("PS2G212.cListSSpUnitRoundingPolicy")    
    
    If CheckSYSTEMError(Err,True) = True Then
		Response.End
    End If
    
    Call iObjPS2G212.ListRows(gStrGlobalCollection, iLngSheetMaxRows, iLngPageNo, Request("txtWhere"),  iArrListOut)
	
    Set iObjPS2G212 = Nothing		                                                 '☜: Unload Comproxy DLL

	If CheckSYSTEMError(Err,True) = True Then
       Response.End 
    End If   

	' Set Next key
	If iLngSheetMaxRows = Ubound(iArrListOut,2) Then
		iStrPrevKey = iArrListOut(0, iLngSheetMaxRows)
		iLngSheetMaxRows  = iLngSheetMaxRows - 1
	    iLngPageNo  = iLngPageNo + 1   
	Else	
		iStrPrevKey = ""
		iLngPageNo = ""
		iLngSheetMaxRows = Ubound(iArrListOut,2)
	End If

	For iLngRow = 0 To iLngSheetMaxRows
   	 	iStrData = iStrData & gColSep & ConvSPChars(iArrListOut(0, iLngRow))			' 단위 
   	 	iStrData = iStrData & gColSep	
   	 	iStrData = iStrData & gColSep & ConvSPChars(iArrListOut(1, iLngRow))			' 소수점자릿수 
   	 	iStrData = iStrData & gColSep & ConvSPChars(iArrListOut(2, iLngRow))			' 올림처리단위 
   	 	iStrData = iStrData & gColSep & ConvSPChars(iArrListOut(3, iLngRow))			' 올림구분 
	 	iStrData = iStrData & gColSep & Cstr(iLngRow + 1 + iLngLastRow) & gColSep & gRowSep
	Next

	Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> "		& vbCr           
    Response.Write " Parent.ggoSpread.Source = Parent.frm1.vspdData	 "	& vbCr
    Response.Write " Parent.ggoSpread.SSShowDataByClip  """	& iStrData		& """" & vbCr    
    Response.Write " parent.frm1.hConUnit.value = """	& iStrPrevKey   & """" & vbCr
'    Response.Write " Parent.frm1.hConUnit.value  = """ & Request("txtConUnit")	  	 & """" & vbCr
    Response.Write " Parent.lgPageNo = " & UNIConvNum(iLngPageNo,0)		& vbCr  
    Response.Write " Parent.DbQueryOk" & vbCr   
	Response.Write "</SCRIPT> "		
	Response.End 

Case CStr(UID_M0002)																'☜: 저장 요청을 받음 

    Set iObjPS2G212 = Server.CreateObject("PS2G212.cMaintSSpUnitRoundingPolicy")  

    If CheckSYSTEMError(Err,True) = True Then
		Response.End		
    End If

	Call iObjPS2G212.Maintain(gStrGlobalCollection, Trim(Request("txtSpreadIns")), Trim(Request("txtSpreadUpd")), Trim(Request("txtSpreadDel")), iLngErrorPosition )

	Set iObjPS2G212 = Nothing
	If CheckSYSTEMError2(Err, True, iLngErrorPosition & "행","","","","") = True Then
		Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> " & vbCr   
		Response.Write " Call Parent.SubSetErrPos(" & iLngErrorPosition & ")" & vbCr
		Response.Write "</SCRIPT> "		
		Response.End 
	End If

    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> " 													'☜: Row 의 상태 
    
End Select
%>
  