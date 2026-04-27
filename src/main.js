// Importar Tauri
const { invoke } = window.__TAURI__.core;
const { open } = window.__TAURI__.dialog;

// Elementos del DOM
const docxInput = document.getElementById('docx-input');
const xlsxInput = document.getElementById('xlsx-input');
const docxFilename = document.getElementById('docx-filename');
const xlsxFilename = document.getElementById('xlsx-filename');
const extractBtn = document.getElementById('extract-btn');
const backBtn = document.getElementById('back-btn');
const insertBtn = document.getElementById('insert-btn');
const mainWindow = document.getElementById('main-window');
const reviewWindow = document.getElementById('review-window');
const flashMessage = document.getElementById('flash-message');
const sheetSelect = document.getElementById('sheet-select');

// Campos del formulario de revisión
const fields = {
  numero: document.getElementById('numero'),
  nombre: document.getElementById('nombre'),
  empresa: document.getElementById('empresa'),
  telefono: document.getElementById('telefono'),
  correo: document.getElementById('correo'),
  servicio: document.getElementById('servicio'),
  valor_total: document.getElementById('valor_total'),
  medio: document.getElementById('medio'),
  estado: document.getElementById('estado'),
  trabajo_realizado_en: document.getElementById('trabajo_realizado_en'),
  orden_servicio: document.getElementById('orden_servicio'),
  fecha: document.getElementById('fecha'),
  observacion: document.getElementById('observacion')
};

let xlsxPath = '';
let docxPath = '';

function basename(path) {
  return path.split(/[/\\]/).pop();
}

// Click handlers para abrir diálogos de archivo
docxFilename.addEventListener('click', async () => {
  const selected = await open({
    title: 'Seleccionar archivo .docx',
    filters: [{ name: 'Documentos Word', extensions: ['docx'] }]
  });
  
  if (selected) {
    docxPath = selected;
    docxFilename.textContent = basename(selected);
    updateExtractButton();
  }
});

xlsxFilename.addEventListener('click', async () => {
  const selected = await open({
    title: 'Seleccionar archivo Excel',
    filters: [{ name: 'Archivos Excel', extensions: ['xlsx'] }]
  });
  
  if (selected) {
    xlsxPath = selected;
    xlsxFilename.textContent = basename(selected);
    updateExtractButton();
  }
});

// Habilitar botón de extracción solo cuando ambos archivos están seleccionados
function updateExtractButton() {
  extractBtn.disabled = !(docxPath && xlsxPath);
}

// Variable para acción a ejecutar cuando se presiona OK en el flash
let flashOkAction = null;

// Mostrar flash message (requiere clic en OK para cerrar)
function showFlash(message, type = 'success', onOk = null) {
  const flashText = document.getElementById('flash-text');
  flashText.textContent = message;
  flashMessage.className = `flash ${type}`;
  flashMessage.classList.remove('hidden');
  flashOkAction = onOk;
}

// Botón OK del flash
document.getElementById('flash-ok-btn').addEventListener('click', () => {
  flashMessage.classList.add('hidden');
  if (flashOkAction) {
    flashOkAction();
    flashOkAction = null;
  }
});

// Extraer datos del .docx
extractBtn.addEventListener('click', async () => {
  try {
    extractBtn.disabled = true;
    extractBtn.textContent = '⏳ Extrayendo...';

    const datos = await invoke('extract_cotizacion', { docxPath });

    // Llenar el formulario con los datos extraídos
    fields.numero.value = datos.numero || '';
    fields.nombre.value = datos.nombre || '';
    fields.empresa.value = datos.empresa || '';
    fields.telefono.value = datos.telefono || '';
    fields.correo.value = datos.correo || '';
    fields.servicio.value = datos.servicio || '';
    fields.valor_total.value = datos.valor_total || '';
    fields.medio.value = datos.medio || '';
    fields.estado.value = datos.estado || 'RECIBIDA';
    fields.trabajo_realizado_en.value = datos.trabajo_realizado_en || '';
    fields.orden_servicio.value = datos.orden_servicio || '';
    fields.fecha.value = datos.fecha || '';
    fields.observacion.value = datos.observacion || '';

    // Obtener hojas del Excel
    await loadSheets();

    // Cambiar a ventana de revisión
    mainWindow.classList.remove('active');
    reviewWindow.classList.add('active');

    showFlash('✅ Datos extraídos correctamente', 'success');

  } catch (error) {
    showFlash(`❌ Error: ${error}`, 'error');
    console.error('Error al extraer:', error);
  } finally {
    extractBtn.disabled = false;
    extractBtn.textContent = '🔍 Extraer datos';
  }
});

// Cargar hojas del Excel
async function loadSheets() {
  try {
    const sheets = await invoke('get_excel_sheets', { xlsxPath });
    
    // Limpiar y llenar el select
    sheetSelect.innerHTML = '<option value="">Seleccionar hoja...</option>';
    
    sheets.forEach(sheet => {
      const option = document.createElement('option');
      option.value = sheet;
      option.textContent = sheet;
      sheetSelect.appendChild(option);
    });
    
    // Seleccionar la hoja del mes actual si existe
    const currentMonth = new Date().toLocaleString('es-ES', { month: 'long' }).toLowerCase();
    const monthSheet = sheets.find(s => s.toLowerCase().includes(currentMonth));
    if (monthSheet) {
      sheetSelect.value = monthSheet;
    }
    
  } catch (error) {
    console.error('Error al cargar hojas:', error);
    sheetSelect.innerHTML = '<option value="">Error al cargar hojas</option>';
  }
}

// Botón de volver
backBtn.addEventListener('click', () => {
  reviewWindow.classList.remove('active');
  mainWindow.classList.add('active');
});

// Insertar datos en Excel
insertBtn.addEventListener('click', async () => {
  try {
    insertBtn.disabled = true;
    insertBtn.textContent = '⏳ Insertando...';

    // Construir objeto de datos
    const datos = {
      numero: fields.numero.value || null,
      nombre: fields.nombre.value || null,
      empresa: fields.empresa.value || null,
      telefono: fields.telefono.value || null,
      correo: fields.correo.value || null,
      servicio: fields.servicio.value || null,
      valor_total: fields.valor_total.value || null,
      medio: fields.medio.value || null,
      estado: fields.estado.value || null,
      trabajo_realizado_en: fields.trabajo_realizado_en.value || null,
      orden_servicio: fields.orden_servicio.value || null,
      fecha: fields.fecha.value || null,
      observacion: fields.observacion.value || null
    };

    const sheetName = sheetSelect.value || null;

    const resultado = await invoke('insert_cotizacion', { datos, xlsxPath, sheetName });

    if (resultado) {
      // Limpiar formulario
      Object.entries(fields).forEach(([name, field]) => {
        field.value = name === 'estado' ? 'RECIBIDA' : '';
      });

      // Volver al menú principal y mostrar mensaje
      reviewWindow.classList.remove('active');
      mainWindow.classList.add('active');
      showFlash(`✅ Cotización ${datos.numero} insertada correctamente`, 'success');
    } else {
      // Volver al menú principal y mostrar mensaje
      reviewWindow.classList.remove('active');
      mainWindow.classList.add('active');
      showFlash(`⚠️ La cotización ${datos.numero} ya existe en la hoja`, 'duplicate');
    }

  } catch (error) {
    // Volver al menú principal y mostrar error
    reviewWindow.classList.remove('active');
    mainWindow.classList.add('active');
    showFlash(`❌ Error: ${error}`, 'error');
    console.error('Error al insertar:', error);
  } finally {
    insertBtn.disabled = false;
    insertBtn.textContent = '✓ Insertar en Excel';
  }
});

// Prevenir comportamiento por defecto de arrastrar y soltar
document.addEventListener('dragover', (e) => e.preventDefault());
document.addEventListener('drop', (e) => e.preventDefault());
