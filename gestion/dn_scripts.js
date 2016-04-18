function abreVentanaCentrada(url, ancho, alto)
{    
		anchoPantalla = screen.availWidth;
		altoPantalla = screen.availHeight;
	
		posX = Math.round((anchoPantalla - ancho)/2);
		posY = Math.round((altoPantalla - alto)/2);		
		opciones = "width=" + ancho + ",height=" + alto + ",top=" + posY + ",left=" + posX + ",scrollbars=yes,resizable=yes";		
    
    nueva=open(url,'_blank',opciones);
    nueva.focus();
}
