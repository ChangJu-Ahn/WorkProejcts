<%@ page language="java" contentType="text/html; Charset=EUC-KR" pageEncoding="utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<%@ include file="/WEB-INF/include/include-header.jspf" %>
<title>Insert title here</title>
</head>
<body>
	<table>
	 <tr>
	 	<td><label style="font-weight: bold;">Filter : </label></td>
	 	<td><input type='text' id='txtFilter' onkeyup='{filter();return false}' onkeypress='javascript:if(event.keyCode==13){ filter(); return false;}'></td>
	 </tr>
	 <tr>
	 	<td colspan='2'><b>[기준은 거래처  전체이름 입니다.(Full Name)]</b></td>
	 </tr>
	</table>
	<table class="board_list">
        <colgroup>
		<col width="15%"/>
		<col width="10%"/>
	</colgroup>
	<thead>
		<tr>
			<th scope="col">CODE</th>
			<th scope="col">CODE_NM</th>
		</tr>
	</thead>
	<tbody id="TBody">
		<c:choose>
			<c:when test="${fn:length(list) > 0}">
				<c:forEach items="${list }" var="row">
					<tr id = '${fn:toUpperCase(row.CODE_NM) }'>
						<td><a href = "javascript:returnParent('${row.CODE }','${row.CODE_NM }')" style="color: blue;">${row.CODE }</a></td>
						<td>
							${row.CODE_NM }
						</td>
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
	function returnParent(CODE,CODE_NM){
		var returnValue = new Array();
		
		returnValue[0] = CODE;
		returnValue[1] = CODE_NM;
		
		window.opener.getReturnValue(returnValue);
		window.close();
	}
	
	function filter(){
		if($('#txtFilter').val()=="")
			$("#TBody tr").css('display','');
		else{
			$("#TBody tr").css('display','none');
			
			$("#TBody tr[id*='"+$('#txtFilter').val().toUpperCase()+"']").css('display','');
		}
		return false;
	}
</script>
</html>