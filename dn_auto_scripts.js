// #####################################################################function dame_ids_checks(f, nombre_check){	// Devuelve los ids de los checkboxes chequeados con el nombre indicado	idchecks = "";	for (i=0; i<f.length; i++)	{		if ((f.elements[i].type == "checkbox") && (f.elements[i].name == nombre_check) && (f.elements[i].checked))		{			if (idchecks == "")			{				idchecks = f.elements[i].value;			}			else			{				idchecks += ", " + f.elements[i].value;			}		}	}	return(idchecks);	}// #####################################################################function eliminar(){	if (confirm("¿Seguro que desea eliminar los productos seleccionados?"))	{		var f=document.getElementById("cesta_productos");		// Cogemos los ids seleccionados y construimos cadena con todos ellos		f.ids_chequeados.value = dame_ids_checks(f,"idcheck");		// Si ha seleccionado algo, actuamos. Si no, avisamos		if (f.ids_chequeados.value != "")		{			// Establecemos el action del formulario			f.action="dn_auto_eliminar.asp";			// Enviamos el formulario			f.submit();		}		else		{			alert("Seleccione los productos que desea eliminar");		}	}}// #####################################################################function autoevaluar(){	var f=document.getElementById("cesta_productos");
	// Mostramos mensaje "Buscando..."	document.getElementById("mensaje_cesta").innerHTML="<div class='mensaje_ajax'><img src='imagenes/progress.gif' hspace='5' align='absmiddle'/><strong>Autoevaluando...</strong> Por favor, espere.<br/><br/></div>";
	change('mensaje_cesta', 'muestra');
	// Cogemos los ids seleccionados y construimos cadena con todos ellos	idchecks = dame_ids_checks(f,"idcheck");		// Enviamos consulta AJAX	new ajax('dn_auto_evaluar.asp', {postBody: 'idcheck='+idchecks, update: $('dn_auto_evaluador_cuerpo'), onComplete: autoevaluar_completed});}// #####################################################################function autoevaluar_completed(){	// Borra mensaje "Buscando..."
	document.getElementById("mensaje_cesta").innerHTML="";
	change('mensaje_cesta', 'oculta');}// #####################################################################function busca_sustancia(origen){	// Cogemos los campos para la búsqueda	numero_tipo = document.getElementById(origen+"_numero_tipo").value;	numero = document.getElementById(origen+"_numero").value;
	nombre = document.getElementById(origen+"_nombre").value;
	nombre = quitar_tildes(nombre);
	// Mostramos mensaje "Buscando..."	//document.getElementById("busqueda_"+origen).innerHTML="<center><img src='imagenes/progress.gif' hspace='5' align='absmiddle'/><strong>Buscando sustancia...</strong> Por favor, espere.</center>";	// Enviamos consulta AJAX	new ajax('dn_auto_busca.asp', {postBody: 'origen='+origen+"&numero_tipo="+numero_tipo+"&numero="+encodeURI(numero)+"&nombre="+encodeURI(nombre), update: $('busqueda_'+origen), onComplete: busca_sustancia_completed});}// #####################################################################function busca_sustancia_completed(){}
origen_live="";
nombre_live="";
numero_live="";

// #####################################################################

function busca_sustancia_live(origen)
{
	// Live search
	origen_live=origen;

	// Cogemos los campos para la búsqueda	numero = document.getElementById(origen+"_numero").value;
	nombre = document.getElementById(origen+"_nombre").value;
	nombre = quitar_tildes(nombre);

	// Si ha cambiado alguno de los campos, buscamos
	// Sólo si la longitud es mayor que 3
	if ((nombre_live != nombre) || (numero_live != numero))
	{
		busca_sustancia(origen_live);
		nombre_live = nombre;
		numero_live = numero;
	}

	// Volvemos a mirar dentro de un rato
	temporizador_live = setTimeout ('busca_sustancia_live(origen_live)', 1000); 
}
// ###################################################################

function busca_sustancia_live_stop()
{
	clearTimeout(temporizador_live);}


// ###################################################################function selecciona_sustancia(nombre, numero, frases, origen){	// Pone nombre y frase en campos del componente indicado por "origen"	// Y borra los resultados de búsqueda	document.getElementById(origen+"_nombre").value = nombre;
	document.getElementById(origen+"_numero").value = numero;	document.getElementById(origen+"_frases_r").value = frases;	document.getElementById("busqueda_"+origen).innerHTML = "";	}// ###################################################################function anadir_componente(maxComponentes){	// Cogemos numero de componentes
	campo = document.getElementById("num_componentes");	campo.value = parseInt(campo.value) + 1;		if (parseInt(campo.value) <= maxComponentes)	{
	// Mostramos las tablas del componente
	change('tabla_comp_'+campo.value, 'componente muestra');

	// Actualizamos nombre de botón eliminar último componente
	document.getElementById('boton_eliminar_componente').value="Eliminar componente "+campo.value;	}	else	{		alert ("No se permiten más de "+maxComponentes+" componentes por producto.");	}}function anadir_componente_completed(){
}// ###################################################################

function eliminar_componente()
{
	// Borra el último componente y reduce el contador

	// Cogemos numero de componentes
	campo = document.getElementById("num_componentes");

	if (parseInt(campo.value) > 1)
	{
		change('tabla_comp_'+campo.value, 'componente oculta');
		campo.value = parseInt(campo.value) - 1;
		document.getElementById('boton_eliminar_componente').value="Eliminar componente "+campo.value;
	}
	else
	{
		alert("No se permite eliminar el primer componente, cada producto debe tener al menos un componente.");
	}
}

// ###################################################################function frasesr(idcampo){	abreVentanaCentrada('dn_auto_frases_r.asp?idcampo='+idcampo, 'frasesr', '640', '480', 'yes', 'yes');}// ###################################################################function validarProducto(){	// Comprueba que se ha introducido un nombre para el producto y para cada componente	var errores="";	var f=document.forms["form_prod"];	var num_componentes = f.num_componentes.value;	if (f.elements["prod_nombre"].value == "")	{		errores += "\n\n* El nombre del producto está en blanco";	}	for (i=1; i<=num_componentes; i++)	{		if (f.elements["comp"+i+"_nombre"].value== "")		{			errores += "\n\n* El nombre del componente "+i+" está en blanco.";		}	}	if (errores != "")	{		alert("ATENCIÓN, corrija los siguientes errores:"+errores);	}	else	{		f.submit();	}}// #####################################################################function mostrar_producto(id_producto){	// Abre ficha de producto no editable en ventana emergente	abreVentanaCentrada('dn_auto_producto_mostrar.asp?id_producto='+id_producto, 'verproducto'+id_producto, '780', '550', 'yes', 'yes');}

// #####################################################################function mostrar_ayuda(razon){	// Abre ficha de ayuda correspondiente en ventana emergente	abreVentanaCentrada('dn_auto_ayuda.asp?razon='+razon, 'verayuda', '780', '300', 'auto', 'yes');}

// #####################################################################

function quitar_tildes(cadena)
{
	// Devuelve la cadena sin tildes
	cadena=replace_all(cadena, "á", "a");
	cadena=replace_all(cadena, "é", "e");
	cadena=replace_all(cadena, "í", "i");
	cadena=replace_all(cadena, "ó", "o");
	cadena=replace_all(cadena, "ú", "u");
	cadena=replace_all(cadena, "Á", "A");
	cadena=replace_all(cadena, "É", "E");
	cadena=replace_all(cadena, "Í", "I");
	cadena=replace_all(cadena, "Ó", "O");
	cadena=replace_all(cadena, "Ú", "U");
	return cadena;
}

// #####################################################################
function replace_all(cadena, busca, reemplaza)
{
	// Reemplaza todas las ocurrencias por la nueva
	while(cadena.indexOf(busca)!=-1)
	{
		cadena=cadena.replace(busca, reemplaza);
	}
	return cadena;
}

// #####################################################################

