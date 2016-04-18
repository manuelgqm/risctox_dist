function toggle(id_objeto, id_imagen) {
    if (Element.visible(id_objeto)){
      $(id_imagen).src="../imagenes/desplegar.gif";
    } else {
      $(id_imagen).src="../imagenes/plegar.gif";
    };
    new Effect.toggle(id_objeto,"appear");
};

function toggle_texto(id_objeto, texto) {
    new Effect.toggle(id_objeto,"appear");
};