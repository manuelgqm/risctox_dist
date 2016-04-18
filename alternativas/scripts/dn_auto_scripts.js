// #####################################################################

function dame_ids_checks(f, nombre_check)
{
	// Devuelve los ids de los checkboxes chequeados con el nombre indicado
	idchecks = "";
	for (i=0; i<f.length; i++)
	{
		if ((f.elements[i].type == "checkbox") && (f.elements[i].name == nombre_check) && (f.elements[i].checked))
		{
			if (idchecks == "")
			{
				idchecks = f.elements[i].value;
			}
			else
			{
				idchecks += ", " + f.elements[i].value;
			}
		}
	}
	return(idchecks);	
}

// #####################################################################

function autoevaluar()
{
	var f=document.getElementById("cesta_productos");

	// Cogemos los ids seleccionados y construimos cadena con todos ellos
	idchecks = dame_ids_checks(f,"idcheck");
	
	// Enviamos consulta AJAX
	new ajax('dn_auto_evaluar.asp', {postBody: 'idcheck='+idchecks, update: $('dn_auto_evaluador_cuerpo'), onComplete: actualizaAutoevaluar});
}

// #####################################################################

function actualizaAutoevaluar()
{
	// Despliega el acordeón
	myAccordion.showThisHideOpen(myDivs[2]);
}

// #####################################################################

function eliminar()
{
	if (confirm("¿Seguro que desea eliminar los productos seleccionados?"))
	{
		var f=document.getElementById("cesta_productos");

		// Cogemos los ids seleccionados y construimos cadena con todos ellos
		idchecks = dame_ids_checks(f,"idcheck");
	
		// Enviamos consulta AJAX
		new ajax('dn_auto_eliminar.asp', {postBody: 'idcheck='+idchecks, update: $('dn_auto_cesta_cuerpo'), onComplete: actualizaEliminar});
	}
}

// #####################################################################

function actualizaEliminar()
{

}

// #####################################################################

function nuevo()
{
	var f=document.getElementById("nuevo_producto");

	// Enviamos consulta AJAX
	new ajax('dn_auto_nuevo.asp', {postBody: 'nombre_producto='+escape(f.nombre_producto.value), update: $('dn_auto_cesta_cuerpo'), onComplete: actualizaNuevo});
}

// #####################################################################

// #####################################################################

function actualizaNuevo()
{
	// Despliega el acordeón
	myAccordion.showThisHideOpen(myDivs[0]);
}

// #####################################################################

function seleccionaProducto(id)
{
	// Enviamos consulta AJAX
	new ajax('dn_auto_modificar.asp', {postBody: 'id='+id, update: $('dn_auto_producto_cuerpo'), onComplete: actualizaSeleccionaProducto});
}

// #####################################################################

function actualizaSeleccionaProducto()
{
	// Despliega el acordeón
	myAccordion.showThisHideOpen(myDivs[1]);
}

// #####################################################################

function modificarProducto()
{
	var f=document.getElementById("form_modificar_producto");

	// Enviamos consulta AJAX
	new ajax('dn_auto_modificar_2.asp', {postBody: 'id='+f.id_producto.value+'&nombre_producto='+escape(f.nombre_producto.value), update: $('dn_auto_producto_cuerpo'), onComplete: seleccionaProducto(f.id_producto.value)});
}

// #####################################################################
