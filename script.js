
let data = [];
let typingTimer;
let ticketsAcumulados = localStorage.getItem("tickets") || "";


// ---------- DATOS ALTERNATIVOS ----------
let alternativosData = [
  { "posCode": "006764", "Cod. Barras": "7730205108531" },
  { "posCode": "006765", "Cod. Barras": "7730205043177" },
  { "posCode": "006736", "Cod. Barras": "8410791501501" },
  { "posCode": "006737", "Cod. Barras": "8410791501518" },
  { "posCode": "006759", "Cod. Barras": "7730205043160" },
  { "posCode": "006738", "Cod. Barras": "8410791501525" },
  { "posCode": "006739", "Cod. Barras": "8410791501532" },
  { "posCode": "015458", "Cod. Barras": "7730205065940" }
];

function requireLogin() {
  const storeId = localStorage.getItem("storeId");

  if (!storeId) {
    document.getElementById("loginBox").style.display = "block";
    document.getElementById("app").style.display = "none";
    return null;
  }

  return storeId;
}

const API_URL = "http://localhost:3000";

// ---------- SOCKET (UNO SOLO) ----------
const socket = io(API_URL);

// ---------- CONECTAR SOCKET ----------
function conectarSocket(storeId) {

  if (!storeId) return;

  socket.off("print");

  socket.emit("register", {
    storeId
  });

  socket.on("print", (data) => {
    console.log("Confirmación de impresión recibida");
    mostrarToast("Imprimiendo... 🖨️");
});
}

// ---------- LOGIN ----------
document.getElementById("loginBtn").addEventListener("click", async () => {

  const user = document.getElementById("user").value;
  const password = document.getElementById("pass").value;

  const res = await fetch(`${API_URL}/login`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ user, password })
  });

  const data = await res.json();

  if (!data.ok) {
    document.getElementById("error").innerText = "Login incorrecto";
    return;
  }

  // guardar tienda
  localStorage.setItem("storeId", data.storeId);

  // UI
  document.getElementById("loginBox").style.display = "none";
  document.getElementById("app").style.display = "block";

  // socket
  conectarSocket(data.storeId);
});

// ---------- AUTO LOGIN ----------
window.addEventListener("load", () => {

  const storeId = requireLogin();

  if (storeId) {

    document.getElementById("loginBox").style.display = "none";
    document.getElementById("app").style.display = "block";

    conectarSocket(storeId);
  }
});

    // Función que buscará el archivo grande en GitHub en el futuro
    async function cargarAlternativosDesdeGitHub() {
      try {
        const response = await fetch('Alternativos.json');
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

  // 1. PASO A: Búsqueda directa en el Excel Principal (Coincidencia exacta corregida)
  let found = data.find(row =>
    Object.values(row).some(val => String(val).trim().toLowerCase() === term)
  );

  let esAlternativo = false;

  // 2. PASO B: Si NO se encuentra, buscamos en los alternativos
  if (!found && typeof alternativosData !== 'undefined' && alternativosData.length > 0) {
    console.log("No encontrado directo. Buscando en códigos alternativos...");
    
    // CORREGIDO: Eliminamos .includes() para evitar falsos positivos con números cortos
    const altMatch = alternativosData.find(row => {
      const barra = row["Cod. Barras"] ? String(row["Cod. Barras"]).trim().toLowerCase() : "";
      return barra === term; 
    });

    // 3. PASO C: Si hallamos la barra alternativa, extraemos su "posCode"
    if (altMatch) {
      const codigoPosOriginal = altMatch["posCode"] ? String(altMatch["posCode"]).trim().toLowerCase() : "";
      
      if (codigoPosOriginal) {
        const codigoPosConPrefijo = "02" + codigoPosOriginal;
        console.log(`¡Código alternativo con prefijo: ${codigoPosConPrefijo}! Buscando en principal...`);

        found = data.find(row => {
          const valorCodigoPrincipal = row["Código"] || row["Codigo"] || row["CÓDIGO"];
          if (!valorCodigoPrincipal) return false;
          
          const cleanPrincipal = String(valorCodigoPrincipal).trim().toLowerCase().replace(/^0+/, '');
          const cleanTarget = codigoPosConPrefijo.replace(/^0+/, '');

          return String(valorCodigoPrincipal).trim().toLowerCase() === codigoPosConPrefijo || cleanPrincipal === cleanTarget;
        });
        
        if (found) {
          esAlternativo = true; // Marcamos que hizo puente con éxito
          console.log("puente ok");
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
        mostrarToast("Producto no encontrado ❌");
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
  const nombre =
    formatearNombre(
      producto["Artículo"]
    );

  const precio =
    producto["PVP"] || "";

  const codigo = String(
    producto[
      "Cód. barras ppal."
    ] || ""
  ).replace(/\D/g, "");

  return `n
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

  ticketsAcumulados += ticket;

  localStorage.setItem(
    "tickets",
    ticketsAcumulados
  );
  actualizarContador();
  mostrarToast("Ticket agregado ✅");

  const input =
    document.getElementById(
      "searchInput"
    );

  input.value = "";
  input.focus();
}

function mostrarToast(mensaje) {

  const toast =
    document.getElementById(
      "toast"
    );

  toast.innerText = mensaje;

  toast.style.visibility =
    "visible";

  setTimeout(() => {

    toast.style.visibility =
      "hidden";

  }, 1500);
}

document
  .getElementById("yesBtn")
  .addEventListener(
    "click",
    function () {

      if (
        window.ultimoProducto
      ) {

        agregarTicket(
          window.ultimoProducto
        );

      } else {

        mostrarToast(
          "No hay producto ❌"
        );
      }
    }
  );

// ---------- IMPRIMIR ----------
document
  .getElementById(
    "downloadTickets"
  )
  .addEventListener(
    "click",
    imprimirTickets
  );

  document
  .getElementById(
    "reprintBtn"
  )
  .addEventListener(
    "click",
    reimprimirUltima
  );

async function imprimirTickets() {

  if (
    !ticketsAcumulados.trim()
  ) {

    mostrarToast(
      "No hay tickets ❌"
    );

    return;
  }

  const storeId = localStorage.getItem("storeId");
  const payload = {
    storeId: storeId,  // ← usar ese, no "tienda-1"
    tickets: ticketsAcumulados
  };

  try {

    const response =
      await fetch(
        `${API_URL}/print`,
        {
          method: "POST",

          headers: {
            "Content-Type":
              "application/json"
          },

          body:
            JSON.stringify(
              payload
            )
        }
      );

    const data =
      await response.json();

    if (data.ok) {

      // guardar última impresión
      localStorage.setItem(
        "ultimaImpresion",
        JSON.stringify(
          payload
        )
      );

      mostrarToast(
        "Impresión enviada ✅"
      );

      // limpiar cola actual
      ticketsAcumulados =
        "";

      localStorage.removeItem(
        "tickets"
      );
    }

  } catch (error) {

    console.error(error);

    mostrarToast(
      "Error enviando impresión ❌"
    );
  }
  actualizarContador();
}

// ---------- REIMPRIMIR ----------
async function reimprimirUltima() {

  const ultima =
    localStorage.getItem(
      "ultimaImpresion"
    );

  if (!ultima) {

    mostrarToast(
      "No hay impresión guardada ❌"
    );

    return;
  }

  try {

    const response =
      await fetch(
        `${API_URL}/print`,
        {
          method: "POST",

          headers: {
            "Content-Type":
              "application/json"
          },

          body: ultima
        }
      );

    const data =
      await response.json();

    if (data.ok) {

      mostrarToast(
        "Reimpresión enviada ✅"
      );
    }

  } catch (error) {

    console.error(error);

    mostrarToast(
      "No se pudo reimprimir ❌"
    );
  }
}


// ---------- Borrar ----------
document.getElementById('clearTickets').addEventListener('click', function() {
  if (confirm("¿Borrar todos los tickets?")) {
    ticketsAcumulados = "";
    localStorage.removeItem("tickets");
    actualizarContador();
    mostrarToast("Borrado ✅");
  }
});

function actualizarContador() {

  const cantidad =
    ticketsAcumulados
      .split("Q1")
      .length - 1;

  document.getElementById(
    "ticketCount"
  ).innerText =
    `🖨️ Etiquetas para imprimir: ${cantidad}`;
}


