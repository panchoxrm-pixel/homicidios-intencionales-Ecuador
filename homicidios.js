function ocultarSecciones(){
    if(document.getElementById("inicio")) { document.getElementById("inicio").classList.remove("activa"); }
    if(document.getElementById("area")) { document.getElementById("area").classList.remove("activa"); }
    if(document.getElementById("sexo")) { document.getElementById("sexo").classList.remove("activa"); }
    if(document.getElementById("edad")) { document.getElementById("edad").classList.remove("activa"); }
    if(document.getElementById("comparacionEntreAnios")) { document.getElementById("comparacionEntreAnios").classList.remove("activa"); }    
}

function mostrarSeccion(id){
    ocultarSecciones();
    
    // 1. Activa la sección correspondiente si existe
    const seccion = document.getElementById(id);
    if(seccion) {
        seccion.classList.add("activa");
    }
    
    // 2. Remueve el resaltado previo de todos los enlaces del sidebar
    const enlaces = document.querySelectorAll('.sidebar-menu label');
    enlaces.forEach(enlace => {
        enlace.classList.remove('active');
    });
    
    // 3. CORRECCIÓN: Búsqueda exacta y estricta del ID entre comillas simples o dobles
    const enlaceActivo = Array.from(enlaces).find(enlace => {
        const accionClick = enlace.getAttribute('onclick');
        if (!accionClick) return false;
        
        // Verifica si incluye exactamente la secuencia con comillas: 'id' o "id"
        return accionClick.includes(`'${id}'`) || accionClick.includes(`"${id}"`);
    });
    
    if(enlaceActivo) {
        enlaceActivo.classList.add('active');
    }

    // 4. Cierra el menú de hamburguesa automáticamente
    const menuToggle = document.getElementById("menu-toggle");
    if(menuToggle) { 
        menuToggle.checked = false; 
    }
}

// Inicialización de la app
mostrarSeccion("inicio");