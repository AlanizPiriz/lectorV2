    let data = [];
    let typingTimer;
    let ticketsAcumulados = localStorage.getItem("tickets") || "";

        // ---------- TUS CÓDIGOS ALTERNATIVOS CARGADOS A MANO (CORREGIDO) ----------
    // Forzamos a que posCode sea un texto puro envolviendo los números entre comillas ""
    let alternativosData = [
      { "posCode": "02006764", "Cod. Barras": "7730205108531" },
      { "posCode": "02006765", "Cod. Barras": "7730205043177" },
      { "posCode": "02006736", "Cod. Barras": "8410791501501" },
      { "posCode": "02006737", "Cod. Barras": "8410791501518" },
      { "posCode": "02006759", "Cod. Barras": "7730205043160" },
      { "posCode": "02006738", "Cod. Barras": "8410791501525" },
      { "posCode": "02006739", "Cod. Barras": "8410791501532" },
      { "posCode": "02015458", "Cod. Barras": "7730205065940" }
    ];


    // Función que buscará el archivo grande en GitHub en el futuro
    async function cargarAlternativosDesdeGitHub() {
      try {
        const response = await fetch('./alternativos.json');
        if (response.ok) {
          const datosArchivo = await response.json();
          if (datosArchivo && datosArchivo.length > 0) {
            alternativosData = datosArchivo; // Reemplaza los de prueba por los miles del archivo
            console.log(`🌐 ¡Éxito! Se cargaron ${alternativosData.length} códigos alternativos desde el archivo de GitHub.`);
          }
        } else {
          console.log("ℹ️ No se detectó un archivo 'alternativos.json' en GitHub. Usando los códigos cargados a mano.");
        }
      } catch (error) {
        console.log("ℹ️ Modo manual activo (No se encontró o no se pudo leer alternativos.json aún).");
      }
    }

    // ---------- IndexedDB Helpers ----------
    const dbName = "ExcelDB";
    const storeName = "excelData";

    function openDB() {
      return new Promise((resolve, reject) => {
        const request = indexedDB.open(dbName, 1);
        request.onupgradeneeded = function (event) {
          const db = event.target.result;
          if (!db.objectStoreNames.contains(storeName)) {
            db.createObjectStore(storeName, { keyPath: "id" });
          }
        };
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
    }

    async function saveDataToIndexedDB(data) {
      const db = await openDB();
      const tx = db.transaction(storeName, "readwrite");
      const store = tx.objectStore(storeName);
      store.put({ id: 1, data });
      return tx.complete;
    }

    async function loadDataFromIndexedDB() {
      const db = await openDB();
      const tx = db.transaction(storeName, "readonly");
      const store = tx.objectStore(storeName);
      return new Promise((resolve) => {
        const request = store.get(1);
        request.onsuccess = () => resolve(request.result ? request.result.data : null);
        request.onerror = () => resolve(null);
      });
    }

    // ---------- Mostrar información del archivo cargado ----------
    function showFileInfo() {
      const info = JSON.parse(localStorage.getItem("lastFileInfo") || "null");
      if (info) {
        const div = document.getElementById("fileInfo");
        div.innerHTML = `📂 <b>${info.name}</b> (cargado: ${info.date})`;
      }
    }

    // ---------- Cargar datos guardados al iniciar la app ----------
    window.addEventListener('load', async () => {
      const storedData = await loadDataFromIndexedDB();
      if (storedData) {
        data = storedData;
        document.getElementById('searchInput').disabled = false;
        document.getElementById('searchBtn').disabled = false;
        console.log("Datos restaurados desde IndexedDB ✅");
      }
      showFileInfo();

      // Intenta buscar el archivo grande en GitHub de forma automática
      await cargarAlternativosDesdeGitHub();
    });

    // ---------- Cargar archivo Excel ----------
    document.getElementById('excelFile').addEventListener('change', function(e) {
      const file = e.target.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = async function(e) {
        const workbook = XLSX.read(e.target.result, { type: 'binary' });
        const firstSheet = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheet];
        data = XLSX.utils.sheet_to_json(sheet);

        // Guardar en IndexedDB
       await saveDataToIndexedDB(data);

        // Guardar nombre y fecha en localStorage
        localStorage.setItem("lastFileInfo", JSON.stringify({
          name: file.name,
          date: new Date().toLocaleString()
        }));

        alert(`Archivo "${file.name}" cargado correctamente ✅`);
        document.getElementById('searchInput').disabled = false;
        document.getElementById('searchBtn').disabled = false;
        document.getElementById('searchInput').focus();
        showFileInfo();
      };
      reader.readAsBinaryString(file);
    });

    // ---------- Buscar automáticamente ----------
    const input = document.getElementById('searchInput');
    input.addEventListener('input', function() {
      clearTimeout(typingTimer);
      typingTimer = setTimeout(() => {
        document.getElementById('searchBtn').click();
      }, 300);
    });

    input.addEventListener('keydown', function(e) {
      if (e.key === 'Enter') {
        e.preventDefault();
        document.getElementById('searchBtn').click();
      }
    });

    // ---------- Buscar producto ----------
        // ---------- Buscar producto (CON PUENTE CORREGIDO A COLUMNA 'CÓDIGO') ----------
    document.getElementById('searchBtn').addEventListener('click', function() {
      const term = document.getElementById('searchInput').value.trim().toLowerCase();
      if (!term) return;

      const resultDiv = document.getElementById('result');
      const infoDiv = document.getElementById('productInfo');
      
      const normalize = str => str.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

      // 1. PASO A: Búsqueda directa en el Excel Principal cargado
      let found = data.find(row =>
        Object.values(row).some(val => String(val).toLowerCase().includes(term))
      );

      let esAlternativo = false;

      // 2. PASO B: Si NO se encuentra, buscamos en los alternativos descargados de GitHub
      if (!found && typeof alternativosData !== 'undefined' && alternativosData.length > 0) {
        console.log("No encontrado directo. Buscando en códigos alternativos...");
        
        // Buscamos si coincide con la columna "Cod. Barras" del JSON alternativo
        const altMatch = alternativosData.find(row => {
          const barra = row["Cod. Barras"] ? String(row["Cod. Barras"]).trim().toLowerCase() : "";
          return barra === term || barra.includes(term);
        });

        // 3. PASO C: Si hallamos la barra alternativa, extraemos su "posCode"
        if (altMatch) {
          const codigoPosOriginal = altMatch["posCode"] ? String(altMatch["posCode"]).trim().toLowerCase() : "";
          
          if (codigoPosOriginal) {
            console.log(`¡Código alternativo asociado a posCode: ${codigoPosOriginal}! Buscando en principal...`);
            
            // 4. PASO D: Volvemos al Excel Principal a buscar el producto usando la columna exacta 'Código'
            found = data.find(row => {
              // Obtenemos el valor de la columna 'Código' del Excel Principal (manejando posibles diferencias de tildes o mayúsculas)
              const valorCodigoPrincipal = row["Código"] || row["Codigo"] || row["CÓDIGO"];
              if (!valorCodigoPrincipal) return false;

              const cleanPrincipal = String(valorCodigoPrincipal).trim().toLowerCase().replace(/^0+/, '');
              const cleanTarget = codigoPosOriginal.replace(/^0+/, '');
              
              // Compara los códigos de forma estricta o limpiando los ceros iniciales
              return String(valorCodigoPrincipal).trim().toLowerCase() === codigoPosOriginal || cleanPrincipal === cleanTarget;
            });
            
            if (found) {
              esAlternativo = true; // Marcamos que hizo puente con éxito
              console.log("puente ok")
            }
          }
        }
      }

      // 5. Renderizado final del resultado (Tu lógica de negocio y UI original)
      if (found) {
        window.ultimoProducto = found;
        const entries = Object.entries(found);

        // Separamos el PVP del resto
        const pvpEntry = entries.find(([key]) => normalize(key) === 'pvp');
        // Buscamos la columna 'Artículo' según tu nueva captura de pantalla
        const articuloEntry = entries.find(([key]) => normalize(key) === 'articulo' || normalize(key) === 'articulo');
        
        let html = '';

        // Aviso visual en pantalla si el producto se detectó mediante el puente
        if (esAlternativo) {
          html += `<div style="background:#FFF3CD; color:#856404; padding:6px; border-radius:4px; margin-bottom:8px; font-size:0.85rem; font-weight:bold; border:1px solid #FFEEBA;">🔄 Producto por Código Alternativo</div>`;
        }

        // Mostrar Artículo primero (usando la clave real encontrada en el Excel)
        if (articuloEntry) {
          const [key, value] = articuloEntry;
          html += `<b>${key}:</b> ${value}<br>`;
        } else {
          // En caso de que no encuentre la key normalizada, buscamos la propiedad exacta 'Artículo' de tu Excel
          const nombreArticulo = found["Artículo"] || found["Articulo"];
          if (nombreArticulo) {
            html += `<b>Artículo:</b> ${nombreArticulo}<br>`;
          }
        }

        // Primero el PVP en rojo (si existe)
        if (pvpEntry) {
          const [key, value] = pvpEntry;
          html += `<b style="color:red;">${key}:</b> <span style="color:red;">${value}</span><br>`;
        }

        // Asignamos al contenedor y mostramos
        infoDiv.innerHTML = html;
        resultDiv.style.display = 'block';

        // ✅ Limpiar input y volver a enfocar automáticamente
        document.getElementById('searchInput').value = "";
        setTimeout(() => document.getElementById('searchInput').focus(), 300);

      } else {
        resultDiv.style.display = 'none';
        //alert("Producto no encontrado ❌");
        document.getElementById('searchInput').value = "";
        setTimeout(() => document.getElementById('searchInput').focus(), 300);
      }

      // Animación de scroll hacia abajo original
      setTimeout(() => {
        window.scrollTo({
            top: document.body.scrollHeight,
            behavior: 'smooth'
        });
      }, 100);
    });

    // ---------- Botones ----------
    // document.getElementById('noBtn').addEventListener('click', function() {
    //   document.getElementById('searchInput').value = "";
    //   document.getElementById('result').style.display = 'none';
    //   document.getElementById('searchInput').focus();
    // });

    let cbosTab; // Variable global

//document.getElementById('yesBtn').addEventListener('click', function() {
//  const url = "https://cbos.arcadiasuite.com/cbos/storeLabelGenerateFind.html";
//  
//  if (!cbosTab || cbosTab.closed) {
//    cbosTab = window.open(url, "cbosTab");
//  } else {
//    cbosTab.focus();
//  }
//});

function limpiarTexto(texto) {
  return String(texto || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function formatearNombre(nombre) {
  nombre = limpiarTexto(nombre);

  if (nombre.length > 30) {
    nombre = nombre.substring(0, 30);
    nombre = nombre.substring(0, nombre.lastIndexOf(" "));
  }

  return nombre;
}

function generarEtiqueta(producto) {
  const nombre = formatearNombre(producto["Artículo"]);
  const precio = producto["PVP"] || "";
  const codigo = String(producto["Cód. barras ppal."] || "").replace(/\D/g, "");

  return `
n
L
H30
PE
z
D11
191100300850095${nombre}
191100700500150$${precio}
1E1202000040112B${codigo}
Q1
E
`;
}


// ---------- Agregar ticket ----------
function agregarTicket(producto) {
    const ticket = generarEtiqueta(producto);
    
    // Actualizamos la variable global y el storage
    ticketsAcumulados += ticket;
    localStorage.setItem("tickets", ticketsAcumulados);

    // Notificación visual rápida (Toast)
    mostrarToast("Ticket agregado ✅");

    // En PDA, limpiar el valor antes del focus asegura que el siguiente escaneo sea limpio
    const input = document.getElementById('searchInput');
    input.value = ''; 
    input.focus();
}

function mostrarToast(mensaje) {
    const toast = document.getElementById("toast");
    toast.innerText = mensaje;
    toast.style.visibility = "visible";
    
    setTimeout(() => { 
        toast.style.visibility = "hidden"; 
    }, 1500); // 1.5 segundos es ideal para no saturar al operario
}


document.getElementById('yesBtn').addEventListener('click', function() {
  if (window.ultimoProducto) {
    agregarTicket(window.ultimoProducto);
  } else {
    alert("No hay producto ❌");
  }
});

// ---------- Descargar ----------
document.getElementById('downloadTickets').addEventListener('click', function() {
  if (!ticketsAcumulados.trim()) {
    alert("No hay tickets ❌");
    return;
  }

  const blob = new Blob([ticketsAcumulados], { type: "text/plain" });
  const url = URL.createObjectURL(blob);

  const fecha = new Date();
  const fechaFormateada = `${fecha.getDate()}-${(fecha.getMonth() + 1)}-${fecha.getFullYear()}`;

  const a = document.createElement("a");
  a.href = url;
  a.download = `tickets_${fechaFormateada}.txt`;
  localStorage.removeItem("tickets");
  ticketsAcumulados = "";
  a.click();

  URL.revokeObjectURL(url);
});

// ---------- Borrar ----------
document.getElementById('clearTickets').addEventListener('click', function() {
  if (confirm("¿Borrar todos los tickets?")) {
    ticketsAcumulados = "";
    localStorage.removeItem("tickets");
    alert("Borrado ✅");
  }
});
