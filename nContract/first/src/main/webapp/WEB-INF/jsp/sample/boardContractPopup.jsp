<%@ page language="java" contentType="text/html; charset=utf-8" 
	pageEncoding="utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<%@ include file="/WEB-INF/include/include-header.jspf" %>
<title>Insert title here</title>
</head>
<body>
	<table>
	 <tr>
	 	<td><label style="font-weight: bold;">Filter(고객사_약칭) : </label></td>
	 	<td><input type='text' id='txtFilter_BP' onkeyup='{filter_BP();return false}' onkeypress='javascript:if(event.keyCode==13){ filter_BP(); return false;}'></td>
	 </tr>
	 <tr>
	 	<td><label style="font-weight: bold;">Filter(계약 번호) : </label></td>
	 	<td><input type='text' id='txtFilter_No' onkeyup='{filter_No();return false}' onkeypress='javascript:if(event.keyCode==13){ filter_No(); return false;}'></td>
	 </tr>
	</table>
	<br/>
	
	<table class="board_list">
        <colgroup>
		<col width="15%"/>
		<col width="15%"/>
		<col width="20%"/>
	</colgroup>
	<thead>
		<tr>
			<th scope="col">계약번호</th>
			<th scope="col">계약서명</th>
			<th scope="col">고객사</th>
		</tr>
	</thead>
	<tbody id="TBody">
		<c:choose>
			<c:when test="${fn:length(list) > 0}">
				<c:forEach items="${list }" var="row">
					<tr id = '${fn:toUpperCase(row.BP_NM1) }'>
						<td ><a href = "javascript:returnParent('${row }')" style="color: blue;" id="${fn:toUpperCase(row.CONTRACT_NO) }">
							${row.CONTRACT_NO }</a>
						</td>
						<td>${row.CONTRACT_NM }</td>
						<td>${row.BP_NM1 }</td>
					</tr>
				</c:forEach>
			</c:when>
			<c:otherwise>
				<tr>
					<td colspan="2">조회된 결과가 없습니다.</td>
				</tr>
			</c:otherwise>
		</c:choose>
        </tbody>
    </table>    
</body>
<script type="text/javascript">

	function returnParent(raw){
		var returnValue = new Array();
		
		raw = raw.replace("{","");
		raw = raw.replace("}","");
		
		returnValue = raw.split(',');
				
		window.opener.getReturnValue(returnValue);
		window.close();
		
	}
	
	function filter_BP(){
		var parentId = "";
		
		if($('#txtFilter_BP').val()=="")
			$("#TBody tr").css('display','');
		else{
			parentId = $("#TBody tr[id*='"+ $('#txtFilter_BP').val().toUpperCase() + "']");
			
			$("#TBody tr").css('display','none');
			parentId.css('display','');
		}
		return false;
	}
	
	function filter_No(){
		var parentId = "";
		
		if($('#txtFilter_No').val()=="")
			$("#TBody tr").css('display','');
		else{
			parentId = $("#TBody tr td a[id*='" + $('#txtFilter_No').val().toUpperCase() + "']").parent().parent(); //a태그 id를 기준으로 부모-부모의 오브젝트 참고
			
			$("#TBody tr").css('display','none');
			parentId.css('display',''); // a태그의 부모-> 부모태그인 tr을 보이도록 설정
		}
		return false;
	}
</script>
</html>