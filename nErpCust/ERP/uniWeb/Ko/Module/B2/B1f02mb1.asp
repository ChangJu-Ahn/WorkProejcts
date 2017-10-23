<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Translation of Unit for Item)
'*  3. Program ID           : B1f02mb1.asp
'*  4. Program Name         : B1f02mb1.asp
'*  5. Program Desc         :
'*  6. Comproxy List        :
'                             +B1f021CtrlTranslaOfItemUnit
'                             +B1f028ListTranslaOfItemUnit
'							  +B1f039LookupUnitOfMeasure
'							  +B1f024LookupPreNextUnit
'*  7. Modified date(First) : 2000/09/15
'*  8. Modified date(Last)  : 2002/12/12
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd

Dim pB1f021												'☆ : 입력/수정용 ComProxy Dll 사용 변수 
Dim pB1f028												'☆ : 조회용 ComProxy Dll 사용 변수 

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strSpread
Dim StrNextKey		' 다음 값 
Dim StrNextUnit		' 다음 값 
Dim StrNextToUnit	' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim lgStrPrevUnit	' 이전 값 
Dim lgStrPrevToUnit	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount          

Dim strData
Dim arrVal, arrTemp																'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim strStatus																	'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
Dim	lGrpCnt																		'☜: Group Count
Dim strUsrId

On Error Resume Next														'☜: 
Err.Clear

Call LoadBasisGlobalInf()

Call loadInfTB19029B("I", "B","NOCOOKIE","MB")
strMode   = Request("txtMode")												'☜ : 현재 상태를 받음 
strSpread = Trim(Request("txtSpread"))

LngMaxRow = CInt(Request("txtMaxRows"))	
	
Select Case strMode
    Case CStr(UID_M0001), "P", "N",	"R"													'☜: 현재 조회/Prev/Next 요청을 받음 
         Call SubBizQuery()
    Case CStr(UID_M0002)																'☜: 저장 요청을 받음 
         Call SubBizSave()
End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim I1_b_unit_of_measure
    Dim I2_b_unit_of_measure_unit
    Dim I3_b_item_item_cd
    
    Const B438_I1_unit = 0
    Const B438_I1_dimension = 1
 
    Const B438_EG1_E4_item_cd = 0 
    Const B438_EG1_E4_item_nm = 1
    Const B438_EG1_E1_unit = 2    
    Const B438_EG1_E2_to_unit = 3 
    Const B438_EG1_E3_from_factor = 4 
    Const B438_EG1_E3_to_factor = 5

	Dim ObjPB2G111
    ReDim I1_b_unit_of_measure(B438_I1_dimension)	
	Dim Export_Array
	
    I1_b_unit_of_measure(B438_I1_dimension) = Request("txtDim")

	If strMode <> "P" And strMode <> "N" And strMode <> "R" Then
	    Call LookUpUnit(Request("txtDim"), Request("txtUnit"))

        I1_b_unit_of_measure(B438_I1_unit) = Request("txtUnit")
        I2_b_unit_of_measure_unit = LookUpPrevNextUnit("Q",Request("txtUnit"), "")
        I3_b_item_item_cd = ""
    ElseIf strMode = "R" Then
	    I1_b_unit_of_measure(B438_I1_unit) = Request("txtFUnit")
	    I2_b_unit_of_measure_unit = Request("txtToUnit")
	    I3_b_item_item_cd = ""
%>
<Script Language=vbscript>
	parent.frm1.vspdData.MaxRows = 0
	parent.frm1.txtFUnit.value   = "<%=Request("txtFUnit")%>"
	parent.frm1.txtFUnitNm.value = "<%=Request("txtFUnitNm")%>"
	parent.frm1.txtTUnit.value   = "<%=Request("txtToUnit")%>"
	parent.frm1.txtTUnitNm.value = "<%=Request("txtTUnitNm")%>"
</Script>
<%
	Else		
	    I1_b_unit_of_measure(B438_I1_unit) = Request("txtUnit")
	    I2_b_unit_of_measure_unit = LookUpPrevNextUnit(strMode, Request("txtUnit"), Request("txtToUnit"))		
	    I3_b_item_item_cd = ""
	End If

    Set ObjPB2G111 = server.CreateObject("PB2G111.cBListTranOfItemUnit")    

    on error resume next
    Err.Clear 
    Export_Array = ObjPB2G111.B_LIST_TRANSLA_OF_ITEM_UNIT (gStrGlobalCollection,I1_b_unit_of_measure,I2_b_unit_of_measure_unit,I3_b_item_item_cd)
    Set ObjPB2G111 = nothing

    If CheckSYSTEMError(Err,True) = True Then                               
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
    End If






    ArrCount = Ubound(Export_Array,1)
    
    ReDim ArrData(ArrCount)
    
    For LngRow = 0 To ArrCount
		strData =           Chr(11) & ConvSPChars(Trim(Export_Array(LngRow,B438_EG1_E4_item_cd )))
		strData = strData & Chr(11) & " "                                                           
		strData = strData & Chr(11) & ConvSPChars(Trim(Export_Array(LngRow,B438_EG1_E4_item_nm )))
		strData = strData & Chr(11) & ConvSPChars(Trim(Export_Array(LngRow,B438_EG1_E1_unit )))
		strData = strData & Chr(11) & ":"                                                           
		strData = strData & Chr(11) & ConvSPChars(Trim(Export_Array(LngRow,B438_EG1_E2_to_unit )))
		strData = strData & Chr(11) & "="                                                           
		strData = strData & Chr(11) & UNINumClientFormat(Trim(Export_Array(LngRow,B438_EG1_E3_from_factor )), ggQty.DecPoint,0)
		strData = strData & Chr(11) & ":"                                                           
		strData = strData & Chr(11) & UNINumClientFormat(Trim(Export_Array(LngRow,B438_EG1_E3_to_factor )), ggQty.DecPoint,0)
		strData = strData & Chr(11) & LngMaxRow + LngRow + 1
		strData = strData & Chr(11) & Chr(12)
		ArrData(LngRow) = strData
    Next
    
    strData = Join(ArrData,"")








%>
<Script Language=vbscript>
    Dim LngLastRow
    Dim LngMaxRow
    Dim LngRow
    Dim strTemp
    Dim strData

	With parent		
																'☜: 화면 처리 ASP 를 지칭함 
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData "<%=strData%>"

	    .frm1.hDimension.value = "<%=Request("cboDimension")%>"
	    .DbQueryOk
	End With
</Script>	
<%    

End Sub

'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    
    If Request("txtMaxRows") = "" Then
		Call DisplayMsgBox("700117", vbInformation, "", "", I_MKSCRIPT)
		Exit Sub
	End If
    
    Dim Obj2PB2G111
    Dim iErrorPosition

    Set Obj2PB2G111 = server.CreateObject("PB2G111.cBCtrlTranOfItemUnit")    

    On Error Resume Next                                               
    Err.Clear       
    Call Obj2PB2G111.B_CTRL_TRANSLA_OF_ITEM_UNIT (gStrGlobalCollection,strSpread,iErrorPosition)
    Set Obj2PB2G111 = nothing
    
    If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
		Exit Sub
    End If
    on error goto 0                                                             

%>
<Script Language=vbscript>
    Parent.DbSaveOk
</Script>
<%					

End Sub	

Function LookUpUnit(Byval strDim, Byval strCode)
    Const B267_I1_unit = 0
    Const B267_I1_unit_nm = 1
    Const B267_I1_dimension = 2

    Const B267_E1_unit = 0
    Const B267_E1_unit_nm = 1
    Const B267_E1_dimension = 2

	Dim ObjPB0C006		
	Dim I1_b_unit_of_measure
	Dim E1_b_unit_of_measure
	
    ReDim I1_b_unit_of_measure(B267_I1_dimension)
    ReDim E1_b_unit_of_measure(B267_E1_dimension)
    
    I1_b_unit_of_measure(B267_I1_unit) = strCode
    I1_b_unit_of_measure(B267_I1_dimension) = strDim
    
    Set ObjPB0C006 = server.CreateObject ("PB0C006.CB0C006")    
    
    On Error Resume Next
    Err.Clear                                                                            '☜: Clear Error status
    E1_b_unit_of_measure = ObjPB0C006.B_SELECT_UNIT_OF_MEASURE (gStrGlobalCollection,I1_b_unit_of_measure)
    Set ObjPB0C006 = nothing    
    
%>
<Script Language=vbscript>
	parent.frm1.txtUnitNm.value   = "<%=ConvSPChars(E1_b_unit_of_measure(B267_E1_unit_nm))%>"
</Script>
<%
    If CheckSYSTEMError(Err,True) = True Then                                              
        on error goto 0

%>
<Script Language=vbscript>
	parent.frm1.vspdData.MaxRows = 0
	
	parent.frm1.txtFUnit.value   = ""
	parent.frm1.txtFUnitNm.value = ""
	parent.frm1.txtTUnit.value   = ""
	parent.frm1.txtTUnitNm.value = ""
	parent.frm1.txtUnit.Focus
	
</Script>
<%
        Response.End 
    End If
    
End Function


Function LookUpPrevNextUnit(Byval strFlag, Byval strFrom, Byval strCode)
    Dim E1_from_unit
    Dim E2_to_unit

    Const B437_E1_unit = 0 
    Const B437_E1_unit_nm = 1

    Const B437_E2_unit = 0 
    Const B437_E2_unit_nm = 1

    Dim I1_select_char
    Dim I2_from_unit
    Dim I3_to_unit
    
    ReDim E1_from_unit(B437_E1_unit_nm)
    ReDim E2_to_unit(B437_E2_unit_nm)
	Dim ObjPB2G112
    
	LookUpPrevNextUnit = ""
		
  	I1_select_char = strFlag
	I2_from_unit   = strFrom
	I3_to_unit     = strCode

    Set ObjPB2G112 = server.CreateObject("PB2G112.cBLookPreNextUnit")    

    on error resume next
    Err.Clear 
    Call ObjPB2G112.B_LOOKUP_PRE_NEXT_UNIT (gStrGlobalCollection, _
         I1_select_char,I2_from_unit,I3_to_unit, _
         E1_from_unit,E2_to_unit)
         
    Set ObjPB2G112 = nothing

    If Err.number <> 0 And (inStr(Err.Description ,"900011") > 0 Or inStr(Err.Description ,"900012") > 0) then
        If CheckSYSTEMError(Err,True) = True Then                               
            Response.End 
        End If
    End If

			    
%>
<Script Language=vbscript>
	parent.frm1.vspdData.MaxRows = 0
	
	parent.frm1.txtFUnit.value   = "<%=ConvSPChars(E1_from_unit(B437_E1_unit))%>"
	parent.frm1.txtFUnitNm.value = "<%=ConvSPChars(E1_from_unit(B437_E1_unit_nm))%>"
	parent.frm1.txtTUnit.value   = "<%=ConvSPChars(E2_to_unit(B437_E2_unit))%>"
	parent.frm1.txtTUnitNm.value = "<%=ConvSPChars(E2_to_unit(B437_E2_unit_nm))%>"
	
</Script>
<%

	LookUpPrevNextUnit = ConvSPChars(E2_to_unit(B437_E2_unit))
End Function
%>





