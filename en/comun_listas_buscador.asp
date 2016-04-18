<form action="<%=form_action%>" method="post" name="myform" onSubmit="primerapag();">
 <input type="hidden" name='busc' value='<%=busc%>' />
 <input type="hidden" name='pag' value='<%=pag%>' />
 <input type="hidden" name='hr' value='<%=hr%>' />
 <input type="hidden" name='arr' value='<%=arr%>' />
 <input type="hidden" name='ordenacion' value='<%=ordenacion%>' />
 <input type="hidden" name='nregs' value='<%=nregs%>' />
<table class="tabla3" width="90%" align="center">
	<tr>
  		<td colspan="3" class="subtitulo3">Substance search</td>
  	</tr>
	<tr>
		<td align="right"><strong>Name</strong></td>
		<td><input type="text" name="nombre" value="<%=nombre%>" />
			<select name="tipobus">
				<option value="exacto" <%if tipobus="exacto" then response.write "selected"%>>exact name</option>
				<option value="parte" <%if tipobus="parte" then response.write "selected"%>>part of the name</option>
			</select>
		</td>
	</tr>
	<tr>
		<td align="right"><strong>CAS/EC/Index No</strong></td>
		<td><input type="text" name="numero" value="<%=numero%>" /></td>
	</tr>
	<tr>
		<td colspan="2" align="center"><input type="submit" value="Search"/> <input type="reset" value="Erase" /></td>
	</tr>
</table>




<%
if busc<>"" then
	if hr=0  then
%>
		<fieldset id="flashmsg"><legend class="advertencia"><strong>Warning</strong></legend>No substances found.</fieldset>
<%
	else
		response.Write("<p class='neg' style='margin:15px 0; padding:10px;'>" &hr& " substances found. Showing from  " &registroini+1& " to " &registrofin+1& ":</p>")
%>
		<%=tablares%>
<%
if hr>nregs then
%>
		<div align='center' style="margin:20px 10px; background-color: #3399CC; padding:3px;"><%paginacion%></div>
<%
end if
%>
<%
	end if
end if
%>
</form>