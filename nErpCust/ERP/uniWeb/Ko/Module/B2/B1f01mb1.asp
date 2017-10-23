<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Translation of Unit)
'*  3. Program ID           : B1f01mb1.asp
'*  4. Program Name         : B1f01mb1.asp
'*  5. Program Desc         :
'*  6. Comproxy List        : +B1f039LookupUnitOfMeasure
'                             +B1f011CtrlTranslationOfUnit
'                             +B1f018ListTranslationOfUnit
'*  7. Modified date(First) : 2000/09/08
'*  8. Modified date(Last)  : 2002/12/11
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
On Error Resume Next														'☜:
Err.Clear

Dim comB1F011																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
Dim comB1F018																'☆ : 조회용 ComProxy Dll 사용 변수 

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strSpread

Dim StrNextKey		' 다음 값(Unit)
Dim StrNextKey2		' 다음 값(ToUnit)
Dim StrNextKey3		' 다음 값(Dimension)
Dim lgStrPrevKey	' 이전 값(Unit)
Dim lgStrPrevKey2	' 이전 값(ToUnit)
Dim lgStrPrevKey3	' 이전 값(Dimension)
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount

Call LoadBasisGlobalInf()

Call loadInfTB19029B("I", "B","NOCOOKIE","MB")

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strSpread = Trim(Request("txtSpread"))

Select Case strMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 

    Dim I1_b_unit_of_measure
    
    Const B434_I1_unit = 0
    Const B434_I1_dimension = 1
 
    Const B434_EG1_E1_minor_nm = 0
    Const B434_EG1_E2_dimension = 1
    Const B434_EG1_E2_unit = 2    
    Const B434_EG1_E3_to_unit = 3 
    Const B434_EG1_E4_from_factor = 4
    Const B434_EG1_E4_to_factor = 5

	Dim ObjPB2G101
    ReDim I1_b_unit_of_measure(B434_I1_dimension)	
	Dim Export_Array
	
%>
<Script Language=vbscript>
	parent.frm1.txtUnitNm.value = "<%=ConvSPChars(LookUpUnit(Request("txtUnit")))%>"
</Script>
<%

    I1_b_unit_of_measure(B434_I1_unit) = Request("txtUnit")
    I1_b_unit_of_measure(B434_I1_dimension) = Request("txtDimensionCd")

    Set ObjPB2G101 = server.CreateObject("PB2G101.cBListTransOfUnit")    

    on error resume next
    Err.Clear 
    Export_Array = ObjPB2G101.B_LIST_TRANSLATION_OF_UNIT(gStrGlobalCollection,I1_b_unit_of_measure)
    Set ObjPB2G101 = nothing

    If CheckSYSTEMError(Err,True) = True Then                               
		Response.End														'☜: 비지니스 로직 처리를 종료함 
    End If
    on error goto 0

%>
<Script Language=vbscript>
    Dim LngLastRow
    Dim LngMaxRow
    Dim LngRow
    Dim strTemp
    Dim strData

	With parent		
																'☜: 화면 처리 ASP 를 지칭함 
		LngMaxRow = 0
<%      
        GroupCount = Ubound(Export_Array,1)
	    For LngRow = 0 To GroupCount
%>
		strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B434_EG1_E1_minor_nm )))%>"     
		strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B434_EG1_E2_dimension )))%>"  
		strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B434_EG1_E2_unit )))%>" 
        strData = strData & Chr(11) & " "
        strData = strData & Chr(11) & ":"
		strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B434_EG1_E3_to_unit)))%> "      
        strData = strData & Chr(11) & " "
        strData = strData & Chr(11) & "="
		strData = strData & Chr(11) & "<%=UNINumClientFormat(Trim(Export_Array(LngRow,B434_EG1_E4_from_factor)), ggQty.DecPoint,0)%>"  
		strData = strData & Chr(11) & ":"
		strData = strData & Chr(11) & "<%=UNINumClientFormat(Trim(Export_Array(LngRow,B434_EG1_E4_to_factor)), ggQty.DecPoint,0)%>"  
		strData = strData & Chr(11) & LngMaxRow + <%=LngRow%> + 1
		strData = strData & Chr(11) & Chr(12)
<%
    Next
%>    

		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowData strData

		.frm1.hDimension.value = "<%=Request("cboDimension")%>"
		.frm1.hUnit.value = "<%=ConvSPChars(Request("txtUnit"))%>"
		.DbQueryOk

	End With
</Script>
<%

Case CStr(UID_M0002)																'☜: 저장 요청을 받음 
    Dim Obj2PB2G101
    Dim iErrorPosition

    Set Obj2PB2G101 = server.CreateObject ("PB2G101.cBCtrlTransOfUnit")    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear        
    Call Obj2PB2G101.B_CTRL_TRANSLATION_OF_UNIT(gStrGlobalCollection,strSpread,iErrorPosition)
    Set Obj2PB2G101 = nothing

    If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
		Response.End 
    End If
    on error goto 0                                                             
%>
<Script Language=vbscript>
	With parent																		'☜: 화면 처리 ASP 를 지칭함 
		'window.status = "저장 성공"
		.DbSaveOk
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
Function LookUpUnit(Byval strCode)
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
    
    Set ObjPB0C006 = server.CreateObject ("PB0C006.CB0C006")    
    
    On Error Resume Next
    Err.Clear                                                                            '☜: Clear Error status
    E1_b_unit_of_measure = ObjPB0C006.B_SELECT_UNIT_OF_MEASURE (gStrGlobalCollection,I1_b_unit_of_measure)
    Set ObjPB0C006 = nothing    
    
    If Err.number <> 0 and inStr(Err.Description ,"124000") > 0 then
  	LookUpUnit = ""
    Else
        If CheckSYSTEMError(Err,True) = True Then                                              
        	Exit Function
	    End If
        on error goto 0
        LookUpUnit = E1_b_unit_of_measure(B267_E1_unit_nm)
    End If
End Function
%>
