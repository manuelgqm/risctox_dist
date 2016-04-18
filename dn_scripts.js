// ###################################################################

function abreVentanaCentrada(url, nombre, ancho, alto, scrollbars, resizable)
{    
		anchoPantalla = screen.availWidth;
		altoPantalla = screen.availHeight;
	
		posX = Math.round((anchoPantalla - ancho)/2);
		posY = Math.round((altoPantalla - alto)/2);		
		opciones = "width=" + ancho + ",height=" + alto + ",top=" + posY + ",left=" + posX + ",scrollbars=" + scrollbars + ",resizable=" + resizable;		
    
    nueva=open(url,nombre,opciones);
    nueva.focus();
}

// ###################################################################
