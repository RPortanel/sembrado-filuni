import React, { useState, useEffect, useRef } from 'react';
import { DragDropContext, Droppable, Draggable } from '@hello-pangea/dnd';
import * as XLSX from 'xlsx-js-style'; 
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';
import { db } from './firebase'; 
import { doc, onSnapshot, setDoc } from 'firebase/firestore';

// --- 1. GENERADORES DE ESTRUCTURAS ---
const initAuditorio = () => {
  const layout = { banca: [] };
  for (let i = 1; i <= 8; i++) layout[`estrado_silla_${i}`] = [];
  
  for (let f = 1; f <= 16; f++) {
    for (let s = 1; s <= 13; s++) {
      // TV UNAM ahora ocupa Fila 7 y 8, asientos 6, 7 y 8
      if ((f === 7 || f === 8) && (s >= 6 && s <= 8)) continue;
      // Pasillo central ahora va de la Fila 9 a la 16, asientos 6, 7 y 8
      if (f >= 9 && f <= 16 && (s >= 6 && s <= 8)) continue; 
      layout[`fila_${f}_silla_${s}`] = [];
    }
  }
  return layout;
};

const initComida = () => {
  const layout = { banca: [] };
  for (let m = 1; m <= 12; m++) { 
    for (let s = 1; s <= 10; s++) layout[`mesa_${m}_silla_${s}`] = [];
  }
  return layout;
};

const getColorDependencia = (dependencia) => {
  if (!dependencia) return '#cbd5e1'; 
  const depNormalizada = String(dependencia).toLowerCase().trim()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  
  if (depNormalizada === 'unam') return '#1e3a8a'; 
  if (depNormalizada === 'unam cdc') return '#3b82f6'; 
  if (depNormalizada === 'uv') return '#f97316'; 
  if (depNormalizada === 'externo' || depNormalizada === 'ext') return '#22c55e'; 
  if (depNormalizada === 'invitado especial') return '#84cc16'; 
  if (depNormalizada === 'dgpyfe') return '#facc15'; 
  if (depNormalizada.includes('rector')) return '#9333ea';
  
  return '#cbd5e1'; 
};

const getExcelStyle = (dependencia) => {
  const baseStyle = { alignment: { wrapText: true, vertical: 'top', horizontal: 'center' } };
  if (!dependencia) return { ...baseStyle, fill: { patternType: 'solid', fgColor: { rgb: 'CBD5E1' } }, font: { color: { rgb: '000000' } } }; 
  
  const depNormalizada = String(dependencia).toLowerCase().trim()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    
  let bg = 'CBD5E1';
  let text = '000000'; 

  if (depNormalizada === 'unam') { bg = '1E3A8A'; text = 'FFFFFF'; } 
  else if (depNormalizada === 'unam cdc') { bg = '3B82F6'; text = 'FFFFFF'; }
  else if (depNormalizada === 'uv') { bg = 'F97316'; text = 'FFFFFF'; }
  else if (depNormalizada === 'externo' || depNormalizada === 'ext') { bg = '22C55E'; text = 'FFFFFF'; }
  else if (depNormalizada === 'invitado especial') { bg = '84CC16'; text = '000000'; }
  else if (depNormalizada === 'dgpyfe') { bg = 'FACC15'; text = '000000'; }
  else if (depNormalizada.includes('rector')) { bg = '9333EA'; text = 'FFFFFF'; }

  return { fill: { patternType: 'solid', fgColor: { rgb: bg } }, font: { color: { rgb: text }, bold: true }, alignment: { wrapText: true, vertical: 'top', horizontal: 'center' } };
};

const paletaMesas = [
  { bg: '#fee2e2', border: '#ef4444' }, { bg: '#fef3c7', border: '#f59e0b' }, { bg: '#dcfce7', border: '#22c55e' }, 
  { bg: '#e0f2fe', border: '#0ea5e9' }, { bg: '#f3e8ff', border: '#a855f7' }, { bg: '#ffedd5', border: '#f97316' },
  { bg: '#ecfccb', border: '#84cc16' }, { bg: '#ccfbf1', border: '#14b8a6' }, { bg: '#ede9fe', border: '#6366f1' }, 
  { bg: '#ffe4e6', border: '#f43f5e' }, { bg: '#fae8ff', border: '#d946ef' }, { bg: '#f1f5f9', border: '#64748b' } 
];

export default function App() {
  const [vistaActual, setVistaActual] = useState('auditorio'); 
  const [auditorio, setAuditorio] = useState(initAuditorio());
  const [comida, setComida] = useState(initComida());
  const [nombresMesas, setNombresMesas] = useState({
    mesa_1: 'Mesa 1', mesa_2: 'Mesa 2', mesa_3: 'Mesa 3', mesa_4: 'Mesa 4', mesa_5: 'Mesa 5',
    mesa_6: 'Mesa 6', mesa_7: 'Mesa 7', mesa_8: 'Mesa 8', mesa_9: 'Mesa 9', mesa_10: 'Mesa 10', 
    mesa_11: 'Mesa 11', mesa_12: 'Mesa 12' 
  });
  
  const [ordenMesas, setOrdenMesas] = useState([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]);
  const [historial, setHistorial] = useState([]);
  const [modoEdicion, setModoEdicion] = useState(false); 
  const [modoPresentacion, setModoPresentacion] = useState(false);
  const [mostrarMenuDescarga, setMostrarMenuDescarga] = useState(false); 
  const [zoom, setZoom] = useState(1); 
  const [busquedaBanca, setBusquedaBanca] = useState('');
  const [busquedaLienzo, setBusquedaLienzo] = useState('');
  const [invitadoEditando, setInvitadoEditando] = useState(null);
  const pdfRef = useRef(null); 

  // --- 2. CONEXIÓN EN TIEMPO REAL A FIREBASE ---
  useEffect(() => {
    const docRef = doc(db, 'eventos', 'filuni2026');
    const unsub = onSnapshot(docRef, (docSnap) => {
      if (docSnap.exists()) {
        const data = docSnap.data();
        if(data.auditorio) setAuditorio(data.auditorio);
        if(data.comida) setComida(data.comida);
        if(data.nombresMesas) setNombresMesas(data.nombresMesas);
        if(data.ordenMesas) setOrdenMesas(data.ordenMesas);
      } else {
        setDoc(docRef, { auditorio: initAuditorio(), comida: initComida(), nombresMesas, ordenMesas });
      }
    });
    return () => unsub(); 
  }, []);

  const syncToCloud = async (nuevoAuditorio, nuevaComida, nuevosNombres, nuevoOrden) => {
    try {
      await setDoc(doc(db, 'eventos', 'filuni2026'), {
        auditorio: nuevoAuditorio || auditorio,
        comida: nuevaComida || comida,
        nombresMesas: nuevosNombres || nombresMesas,
        ordenMesas: nuevoOrden || ordenMesas
      });
    } catch(error) {
      console.error("Error sincronizando a Firebase:", error);
    }
  };

  const guardarEnHistorial = () => {
    setHistorial(prev => {
      const fotoActual = {
        auditorio: JSON.parse(JSON.stringify(auditorio)),
        comida: JSON.parse(JSON.stringify(comida)),
        nombresMesas: JSON.parse(JSON.stringify(nombresMesas)),
        ordenMesas: [...ordenMesas]
      };
      const nuevoHistorial = [...prev, fotoActual];
      if (nuevoHistorial.length > 20) nuevoHistorial.shift();
      return nuevoHistorial;
    });
  };

  const deshacerUltimaAccion = () => {
    if (historial.length === 0) return;
    const estadoAnterior = historial[historial.length - 1];
    const nuevoHistorial = historial.slice(0, -1);
    
    setHistorial(nuevoHistorial);
    setAuditorio(estadoAnterior.auditorio);
    setComida(estadoAnterior.comida);
    setNombresMesas(estadoAnterior.nombresMesas);
    setOrdenMesas(estadoAnterior.ordenMesas);
    
    syncToCloud(estadoAnterior.auditorio, estadoAnterior.comida, estadoAnterior.nombresMesas, estadoAnterior.ordenMesas);
  };

  const obtenerTodosLosInvitados = () => {
    const mapa = new Map();
    const extraer = (layout) => {
      Object.values(layout).forEach(lista => {
        lista.forEach(inv => { mapa.set(inv.id, inv); });
      });
    };
    extraer(auditorio);
    extraer(comida);
    return Array.from(mapa.values());
  };

  // --- 3. LÓGICAS DE BLOQUEO, ELIMINACIÓN Y EDICIÓN ---
  const toggleBloqueoBanca = (id) => {
    if(!modoEdicion) return;
    guardarEnHistorial(); 
    const propBloqueo = vistaActual === 'auditorio' ? 'bloqueado_auditorio' : 'bloqueado_comida';
    
    const nuevoAuditorio = { ...auditorio };
    Object.keys(nuevoAuditorio).forEach(key => {
        nuevoAuditorio[key] = nuevoAuditorio[key].map(inv => inv.id === id ? { ...inv, [propBloqueo]: !inv[propBloqueo] } : inv);
    });
    if (nuevoAuditorio.banca) nuevoAuditorio.banca.sort((a, b) => (a.bloqueado_auditorio === b.bloqueado_auditorio ? 0 : a.bloqueado_auditorio ? 1 : -1));

    const nuevaComida = { ...comida };
    Object.keys(nuevaComida).forEach(key => {
        nuevaComida[key] = nuevaComida[key].map(inv => inv.id === id ? { ...inv, [propBloqueo]: !inv[propBloqueo] } : inv);
    });
    if (nuevaComida.banca) nuevaComida.banca.sort((a, b) => (a.bloqueado_comida === b.bloqueado_comida ? 0 : a.bloqueado_comida ? 1 : -1));

    setAuditorio(nuevoAuditorio); setComida(nuevaComida);
    syncToCloud(nuevoAuditorio, nuevaComida, null, null);
  };

  const eliminarInvitadoDeBanca = (id) => {
    if(!modoEdicion) return;
    if (!window.confirm("¿Seguro que deseas eliminar a este invitado de todo el evento?")) return;
    guardarEnHistorial(); 

    const purgarLayout = (layout) => {
      const nuevoLayout = { ...layout };
      Object.keys(nuevoLayout).forEach(key => {
        nuevoLayout[key] = nuevoLayout[key].filter(inv => inv.id !== id);
      });
      return nuevoLayout;
    };

    const nuevoAuditorio = purgarLayout(auditorio);
    const nuevaComida = purgarLayout(comida);
    
    setAuditorio(nuevoAuditorio);
    setComida(nuevaComida);
    syncToCloud(nuevoAuditorio, nuevaComida, null, null);
  };

  const guardarEdicion = (invitadoActualizado) => {
    guardarEnHistorial(); 
    const actualizarLayout = (layout) => {
      const nuevoLayout = { ...layout };
      Object.keys(nuevoLayout).forEach(key => {
        nuevoLayout[key] = nuevoLayout[key].map(inv => inv.id === invitadoActualizado.id ? invitadoActualizado : inv);
      });
      return nuevoLayout;
    };
    
    const nuevoAuditorio = actualizarLayout(auditorio);
    const nuevaComida = actualizarLayout(comida);
    setAuditorio(nuevoAuditorio); setComida(nuevaComida);
    syncToCloud(nuevoAuditorio, nuevaComida, null, null);
    setInvitadoEditando(null); 
  };

  const manejarEdicionNombreMesa = (m, valor) => {
    if(!modoEdicion) return;
    setNombresMesas(prev => ({...prev, [`mesa_${m}`]: valor}));
  }
  const manejarBlurNombreMesa = () => syncToCloud(null, null, nombresMesas, null);

  // --- 4. RESPALDOS LOCALES (JSON) ---
  const descargarRespaldo = () => {
    const datos = { auditorio, comida, nombresMesas, ordenMesas };
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(datos));
    const elementoEnlace = document.createElement('a');
    elementoEnlace.setAttribute("href", dataStr);
    elementoEnlace.setAttribute("download", `Respaldo_Filuni_Nube_${new Date().toISOString().slice(0,10)}.json`);
    document.body.appendChild(elementoEnlace); elementoEnlace.click(); document.body.removeChild(elementoEnlace);
  };

  const cargarRespaldo = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const parseado = JSON.parse(evt.target.result);
        if(parseado.auditorio && parseado.comida) {
          guardarEnHistorial(); 
          const auditorioSeguro = { ...initAuditorio(), ...parseado.auditorio };
          const comidaSegura = { ...initComida(), ...parseado.comida };
          const nombresSeguros = { ...nombresMesas, ...(parseado.nombresMesas || {}) };
          const ordenSeguro = parseado.ordenMesas || [1,2,3,4,5,6,7,8,9,10,11,12];

          setAuditorio(auditorioSeguro); setComida(comidaSegura); setNombresMesas(nombresSeguros); setOrdenMesas(ordenSeguro);
          syncToCloud(auditorioSeguro, comidaSegura, nombresSeguros, ordenSeguro);
          alert('✅ Respaldo cargado y sincronizado en la nube correctamente.');
        } else { alert('❌ El archivo no tiene el formato correcto.'); }
      } catch (error) { alert("❌ Error al leer el archivo de respaldo."); }
      e.target.value = null;
    };
    reader.readAsText(file);
  };

  // --- 5. EXPORTACIÓN A EXCEL MÁSTER ---
  const exportarExcel = () => {
    const wb = XLSX.utils.book_new();
    const celdaVaciaBase = { alignment: { wrapText: true, vertical: 'top', horizontal: 'center' } };

    // HOJA 1: EL DIRECTORIO MAESTRO
    const directorio = obtenerTodosLosInvitados().map(inv => ({
      'ID_NO_TOCAR': inv.id,
      'Dependencia': inv.dependencia,
      'Nombre': inv.nombre,
      'Cargo': inv.cargo
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(directorio), "1. Directorio Editable");

    // HOJA 2: LISTA AUDITORIO
    const datosAuditorio = [];
    for (let i = 1; i <= 8; i++) {
      const ocupante = (auditorio[`estrado_silla_${i}`] || [])[0];
      datosAuditorio.push({ 'Ubicación': `Estrado - Silla ${i}`, 'Dependencia': ocupante?.dependencia || '', 'Nombre': ocupante?.nombre || '[ Vacío ]', 'Cargo': ocupante?.cargo || '' });
    }
    for (let f = 1; f <= 14; f++) {
      for (let s = 1; s <= 13; s++) {
        if ((f === 7 || f === 8) && (s >= 6 && s <= 8)) continue;
        if (f >= 9 && f <= 16 && (s >= 6 && s <= 8)) continue;
        const ocupante = (auditorio[`fila_${f}_silla_${s}`] || [])[0];
        datosAuditorio.push({ 'Ubicación': `Fila ${f} - Silla ${s}`, 'Dependencia': ocupante?.dependencia || '', 'Nombre': ocupante?.nombre || '[ Vacío ]', 'Cargo': ocupante?.cargo || '' });
      }
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(datosAuditorio), "2. Lista Auditorio");

    // HOJA 3: LISTA COMIDA
    const datosComida = [];
    ordenMesas.forEach(m => {
      for(let s=1; s<=10; s++) {
        const ocupante = (comida[`mesa_${m}_silla_${s}`] || [])[0];
        datosComida.push({ 'Mesa': nombresMesas[`mesa_${m}`], 'Asiento': `Silla ${s}`, 'Dependencia': ocupante?.dependencia || '', 'Nombre': ocupante?.nombre || '[ Vacío ]', 'Cargo': ocupante?.cargo || '' });
      }
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(datosComida), "3. Lista Comida");

    // HOJA 4: GRÁFICO AUDITORIO
    const matrizAuditorio = [];
    matrizAuditorio.push([{ v: "ESTRADO (Presidium)", s: { font: { bold: true } } }]);
    const filaEstrado = [];
    for(let i=1; i<=8; i++) {
        const oc = (auditorio[`estrado_silla_${i}`] || [])[0];
        filaEstrado.push(oc ? { v: `${oc.nombre}\n${oc.cargo}`, t: 's', s: getExcelStyle(oc.dependencia) } : { v: "[ Vacío ]", t: 's', s: celdaVaciaBase });
    }
    matrizAuditorio.push(filaEstrado); matrizAuditorio.push([]); 
    for(let f=1; f <= 16; f++) {
        matrizAuditorio.push([{ v: `FILA ${f}`, s: { font: { bold: true } } }]);
        const filaAsientos = [];
        for(let s=1; s<=13; s++) {
             if ((f === 7 || f === 8) && (s >= 6 && s <= 8)) {
                 if (f === 7 && s === 7) filaAsientos.push({ v: "[ TV UNAM ]", t: 's', s: { fill: { patternType: 'solid', fgColor: { rgb: "1E293B" } }, font: { color: { rgb: "FFFFFF" }, bold: true }, alignment: { vertical: 'center', horizontal: 'center' } } });
                 else filaAsientos.push({ v: "", t: 's', s: celdaVaciaBase });
                 continue;
             }
             if (f >= 9 && f <= 16 && (s >= 6 && s <= 8)) { 
                 filaAsientos.push({ v: "", t: 's', s: celdaVaciaBase }); 
                 continue; 
             }
             const oc = (auditorio[`fila_${f}_silla_${s}`] || [])[0];
             filaAsientos.push(oc ? { v: `${oc.nombre}\n${oc.cargo}`, t: 's', s: getExcelStyle(oc.dependencia) } : { v: "[ Vacío ]", t: 's', s: celdaVaciaBase });
        }
        matrizAuditorio.push(filaAsientos);
        if (f === 6) { matrizAuditorio.push([]); matrizAuditorio.push(["", "", "", "", "", { v: "============= P A S I L L O =============", t: 's', s: { font: { bold: true, color: { rgb: "475569" } }, alignment: { horizontal: 'center' } } }]); }
        matrizAuditorio.push([]); 
    }
    const ws3 = XLSX.utils.aoa_to_sheet(matrizAuditorio); ws3['!cols'] = Array(16).fill({ wch: 25 }); XLSX.utils.book_append_sheet(wb, ws3, "4. Gráfico Auditorio");

    // HOJA 5: GRÁFICO COMIDA
    const matrizComida = [];
    ordenMesas.forEach((m) => {
        matrizComida.push([{ v: nombresMesas[`mesa_${m}`], s: { font: { bold: true } } }]);
        for(let fila=0; fila<5; fila++) {
            const o1 = (comida[`mesa_${m}_silla_${(fila * 2) + 1}`] || [])[0]; 
            const o2 = (comida[`mesa_${m}_silla_${(fila * 2) + 2}`] || [])[0];
            matrizComida.push([
              o1 ? { v: `${o1.nombre}\n${o1.cargo}`, t: 's', s: getExcelStyle(o1.dependencia) } : { v: "[ Vacío ]", t: 's', s: celdaVaciaBase }, 
              o2 ? { v: `${o2.nombre}\n${o2.cargo}`, t: 's', s: getExcelStyle(o2.dependencia) } : { v: "[ Vacío ]", t: 's', s: celdaVaciaBase }
            ]);
        }
        matrizComida.push([]);
    });
    const ws4 = XLSX.utils.aoa_to_sheet(matrizComida); ws4['!cols'] = Array(6).fill({ wch: 25 }); XLSX.utils.book_append_sheet(wb, ws4, "5. Gráfico Comida");

    XLSX.writeFile(wb, "Sembrado_Invitados_Completo.xlsx");
  };

  // --- 6. CARGAR EXCEL (NUEVOS INVITADOS) ---
  const cargarExcel = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const filasExcel = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);

        guardarEnHistorial(); 
        
        const nuevosInvitados = filasExcel.map((fila, index) => {
          const filaNormalizada = {};
          Object.keys(fila).forEach(key => { filaNormalizada[key.toLowerCase().trim()] = fila[key]; });
          return { id: `excel-${Date.now()}-${index}`, dependencia: filaNormalizada.dependencia || '', nombre: filaNormalizada.nombre || 'Sin Nombre', cargo: filaNormalizada.cargo || 'Sin Cargo', bloqueado_auditorio: false, bloqueado_comida: false };
        });

        const nuevoAuditorio = { ...auditorio, banca: [...auditorio.banca, ...nuevosInvitados].sort((a, b) => (a.bloqueado_auditorio === b.bloqueado_auditorio ? 0 : a.bloqueado_auditorio ? 1 : -1)) };
        const nuevaComida = { ...comida, banca: [...comida.banca, ...nuevosInvitados].sort((a, b) => (a.bloqueado_comida === b.bloqueado_comida ? 0 : a.bloqueado_comida ? 1 : -1)) };
        
        setAuditorio(nuevoAuditorio); setComida(nuevaComida);
        syncToCloud(nuevoAuditorio, nuevaComida, null, null);
        alert(`✅ Se añadieron ${nuevosInvitados.length} invitados a la banca.`); e.target.value = null; 
      } catch (error) { alert("❌ Error al leer el Excel. Revisa el formato."); }
    };
    reader.readAsArrayBuffer(file);
  };

  // --- 6.5 ACTUALIZADOR MASIVO DE TEXTOS ---
  const actualizarDesdeExcel = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const sheetName = workbook.SheetNames.find(n => n.includes('Directorio')) || workbook.SheetNames[0];
        const filasExcel = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

        const actualizaciones = {};
        filasExcel.forEach(fila => {
          const filaNorm = {};
          Object.keys(fila).forEach(k => { filaNorm[k.toUpperCase().trim()] = fila[k]; });
          
          if (filaNorm['ID_NO_TOCAR']) {
            actualizaciones[filaNorm['ID_NO_TOCAR']] = {
              nombre: filaNorm['NOMBRE'] || '',
              cargo: filaNorm['CARGO'] || '',
              dependencia: filaNorm['DEPENDENCIA'] || ''
            };
          }
        });

        if (Object.keys(actualizaciones).length === 0) {
           alert("❌ No se encontró la columna 'ID_NO_TOCAR'. Asegúrate de usar la hoja '1. Directorio Editable'.");
           return;
        }

        guardarEnHistorial();

        const aplicarActualizaciones = (layout) => {
          const nuevoLayout = { ...layout };
          Object.keys(nuevoLayout).forEach(key => {
            nuevoLayout[key] = nuevoLayout[key].map(inv => {
              if (actualizaciones[inv.id]) {
                return { ...inv, ...actualizaciones[inv.id] };
              }
              return inv;
            });
          });
          return nuevoLayout;
        };

        const nuevoAuditorio = aplicarActualizaciones(auditorio);
        const nuevaComida = aplicarActualizaciones(comida);

        setAuditorio(nuevoAuditorio); setComida(nuevaComida);
        syncToCloud(nuevoAuditorio, nuevaComida, null, null);
        alert(`✅ ¡Textos actualizados! Se revisaron y corrigieron ${Object.keys(actualizaciones).length} registros sin perder sus lugares.`);
      } catch (error) {
        alert("❌ Error procesando las correcciones. Asegúrate de que subiste el Excel correcto.");
      }
      e.target.value = null; 
    };
    reader.readAsArrayBuffer(file);
  };

  // --- 7. EXPORTAR PDF ---
  const exportarPDF = async () => {
    const pdf = new jsPDF('l', 'mm', 'a4'); 
    const pdfWidth = pdf.internal.pageSize.getWidth(); const pdfHeightSheet = pdf.internal.pageSize.getHeight(); 

    const contAuditorio = document.getElementById('contenedor-auditorio'); const contComida = document.getElementById('contenedor-comida');
    let overAuditorio = ''; let overComida = '';
    
    if (contAuditorio) { overAuditorio = contAuditorio.style.overflowX; contAuditorio.style.overflowX = 'visible'; }
    if (contComida) { overComida = contComida.style.overflowX; contComida.style.overflowX = 'visible'; }

    const bloquesACapturar = vistaActual === 'auditorio' ? ['auditorio-parte-1', 'auditorio-parte-2'] : ['comida-parte-1', 'comida-parte-2', 'comida-parte-3', 'comida-parte-4'];
    const fechaHoy = new Date().toLocaleDateString('es-MX', { year: 'numeric', month: '2-digit', day: '2-digit' });

    try {
      for (let i = 0; i < bloquesACapturar.length; i++) {
        const element = document.getElementById(bloquesACapturar[i]);
        if (!element) continue;

        const zoomAnterior = element.style.zoom; const widthAnterior = element.style.width;
        element.style.zoom = 1; element.style.width = `${element.scrollWidth}px`; 

        const canvas = await html2canvas(element, { scale: 2, useCORS: true, scrollX: 0, scrollY: 0, width: element.scrollWidth, height: element.scrollHeight, windowWidth: element.scrollWidth, windowHeight: element.scrollHeight });
        element.style.zoom = zoomAnterior; element.style.width = widthAnterior;

        const imgData = canvas.toDataURL('image/png');
        if (i > 0) pdf.addPage();

        pdf.setFont('helvetica', 'bold'); pdf.setFontSize(14); pdf.setTextColor(15, 23, 42); pdf.text("Inauguración Filuni 2026", 10, 12);

        const margenX = 6; const topReserved = 18; const bottomReserved = 12; const maxImgHeight = pdfHeightSheet - topReserved - bottomReserved; 
        let imgWidth = pdfWidth - (margenX * 2); let imgHeight = (canvas.height * imgWidth) / canvas.width;
        
        if (imgHeight > maxImgHeight) { imgHeight = maxImgHeight; imgWidth = (canvas.width * imgHeight) / canvas.height; }
        const xOffset = (pdfWidth - imgWidth) / 2; const yOffset = topReserved + (maxImgHeight - imgHeight) / 2;

        pdf.addImage(imgData, 'PNG', xOffset, yOffset, imgWidth, imgHeight);
        pdf.setFont('helvetica', 'normal'); pdf.setFontSize(9); pdf.setTextColor(100, 116, 139); pdf.text(`Fecha de generation: ${fechaHoy}`, pdfWidth - 10, pdfHeightSheet - 7, { align: 'right' });
      }
      pdf.save(`Protocolo-${vistaActual}.pdf`);
    } finally {
      if (contAuditorio) contAuditorio.style.overflowX = overAuditorio;
      if (contComida) contComida.style.overflowX = overComida;
    }
  };

  // --- 8. ARRASTRAR Y SOLTAR MÁSTER ---
  const onDragEnd = (result) => {
    if(!modoEdicion) return; 

    const { source, destination, type } = result;
    if (!destination) return;
    
    guardarEnHistorial();

    if (type === 'mesa') {
      const srcRow = parseInt(source.droppableId.split('_')[1]);
      const destRow = parseInt(destination.droppableId.split('_')[1]);
      const srcAbsIndex = (srcRow * 3) + source.index;
      const destAbsIndex = (destRow * 3) + destination.index;

      const nuevoOrden = Array.from(ordenMesas);
      const [mesaMovida] = nuevoOrden.splice(srcAbsIndex, 1);
      nuevoOrden.splice(destAbsIndex, 0, mesaMovida);

      setOrdenMesas(nuevoOrden); syncToCloud(null, null, null, nuevoOrden);
      return;
    }

    const estadoActivo = vistaActual === 'auditorio' ? auditorio : comida;

    if (source.droppableId === destination.droppableId) {
      const nuevaLista = Array.from(estadoActivo[source.droppableId] || []);
      const [movido] = nuevaLista.splice(source.index, 1);
      nuevaLista.splice(destination.index, 0, movido);
      
      const nuevoEstado = { ...estadoActivo, [source.droppableId]: nuevaLista };
      if (vistaActual === 'auditorio') setAuditorio(nuevoEstado); else setComida(nuevoEstado);
      syncToCloud(vistaActual === 'auditorio' ? nuevoEstado : null, vistaActual === 'comida' ? nuevoEstado : null, null, null);
      return;
    }

    const nuevoEstado = { ...estadoActivo };
    const origenLista = Array.from(nuevoEstado[source.droppableId] || []);
    const destinoLista = Array.from(nuevoEstado[destination.droppableId] || []);

    const [invitadoMovido] = origenLista.splice(source.index, 1);

    if (destination.droppableId !== 'banca' && destinoLista.length >= 1) {
      const [invitadoDesplazado] = destinoLista.splice(0, 1);
      destinoLista.push(invitadoMovido);
      origenLista.splice(source.index, 0, invitadoDesplazado);
    } else {
      destinoLista.splice(destination.index, 0, invitadoMovido);
    }

    if (destination.droppableId === 'banca' || source.droppableId === 'banca') {
        const propBloqueo = vistaActual === 'auditorio' ? 'bloqueado_auditorio' : 'bloqueado_comida';
        if (destination.droppableId === 'banca') destinoLista.sort((a, b) => (a[propBloqueo] === b[propBloqueo] ? 0 : a[propBloqueo] ? 1 : -1));
        if (source.droppableId === 'banca') origenLista.sort((a, b) => (a[propBloqueo] === b[propBloqueo] ? 0 : a[propBloqueo] ? 1 : -1));
    }

    nuevoEstado[source.droppableId] = origenLista;
    nuevoEstado[destination.droppableId] = destinoLista;

    if (vistaActual === 'auditorio') setAuditorio(nuevoEstado); else setComida(nuevoEstado);
    syncToCloud(vistaActual === 'auditorio' ? nuevoEstado : null, vistaActual === 'comida' ? nuevoEstado : null, null, null);
  };

  const layoutActivo = vistaActual === 'auditorio' ? auditorio : comida;

  const ocupantesAuditorio = Object.keys(auditorio).reduce((acc, key) => {
    if (key !== 'banca' && !key.includes('estrado')) return acc + (auditorio[key] || []).length;
    return acc;
  }, 0);

  const ocupantesComida = Object.keys(comida).reduce((acc, key) => {
    if (key !== 'banca') return acc + (comida[key] || []).length;
    return acc;
  }, 0);

  const bloquesMesas = [];
  for (let i = 0; i < ordenMesas.length; i += 3) {
    bloquesMesas.push(ordenMesas.slice(i, i + 3));
  }

  const renderMesaLayout = (m, indexWithinRow) => (
    <Draggable key={`mesa_draggable_${m}`} draggableId={`mesa_draggable_${m}`} index={indexWithinRow} isDragDisabled={!modoEdicion}>
      {(provided, snapshot) => (
        <div 
          ref={provided.innerRef} 
          {...provided.draggableProps} 
          style={{ ...provided.draggableProps.style, width: '320px', backgroundColor: paletaMesas[m - 1].bg, border: `3px solid ${paletaMesas[m - 1].border}`, borderRadius: '8px', padding: '15px 20px 20px 20px', display: 'flex', flexDirection: 'column', alignItems: 'center', boxShadow: snapshot.isDragging ? '0 15px 25px rgba(0,0,0,0.3)' : '0 4px 6px -1px rgba(0,0,0,0.1)', flexShrink: 0, opacity: snapshot.isDragging ? 0.9 : 1 }}
        >
          {modoEdicion && (
            <div {...provided.dragHandleProps} title="Arrastrar mesa para reordenarla" style={{ width: '100%', height: '20px', display: 'flex', justifyContent: 'center', alignItems: 'center', cursor: snapshot.isDragging ? 'grabbing' : 'grab', marginBottom: '8px', color: paletaMesas[m - 1].border, opacity: 0.6 }}>
              <span style={{ fontSize: '18px', lineHeight: '0' }}>⣿</span>
            </div>
          )}

          <div style={{ width: '100%', marginBottom: '15px', display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
            <input 
              type="text" 
              readOnly={!modoEdicion}
              value={nombresMesas[`mesa_${m}`] || ''} 
              onChange={(e) => manejarEdicionNombreMesa(m, e.target.value)} 
              onBlur={manejarBlurNombreMesa} 
              style={{ width: '100%', height: '40px', lineHeight: '36px', textAlign: 'center', fontWeight: 'bold', fontSize: '16px', padding: '0 10px', borderRadius: '4px', border: modoEdicion ? '1px solid white' : 'none', backgroundColor: modoEdicion ? 'rgba(255,255,255,0.7)' : 'transparent', outline: 'none', boxSizing: 'border-box', fontFamily: 'sans-serif' }} 
            />
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px', width: '100%' }}>
            {Array.from({ length: 10 }, (_, s) => {
              const ocupante = layoutActivo[`mesa_${m}_silla_${s+1}`] || [];
              return <Silla key={`mesa_${m}_silla_${s+1}`} id={`mesa_${m}_silla_${s+1}`} ocupante={ocupante} vista={vistaActual} busqueda={busquedaLienzo} onEdit={setInvitadoEditando} modoEdicion={modoEdicion} />
            })}
          </div>
        </div>
      )}
    </Draggable>
  );

  return (
    <div style={{ display: 'flex', height: '100vh', width: '100vw', fontFamily: 'sans-serif', backgroundColor: '#f1f5f9', overflow: 'hidden' }}>
      
      {/* MODAL EDICIÓN */}
      {invitadoEditando && (
        <div style={{ position: 'fixed', top: 0, left: 0, width: '100vw', height: '100vh', backgroundColor: 'rgba(0,0,0,0.6)', zIndex: 99999, display: 'flex', justifyContent: 'center', alignItems: 'center' }}>
          <div style={{ backgroundColor: 'white', padding: '25px', borderRadius: '8px', width: '350px', display: 'flex', flexDirection: 'column', gap: '15px', boxShadow: '0 10px 25px rgba(0,0,0,0.2)' }}>
            <h3 style={{ margin: 0, fontSize: '18px', color: '#0f172a' }}>Editar Invitado</h3>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '5px' }}>
              <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#475569' }}>Nombre completo</label>
              <input type="text" value={invitadoEditando.nombre} onChange={(e) => setInvitadoEditando({...invitadoEditando, nombre: e.target.value})} style={{ padding: '8px', borderRadius: '4px', border: '1px solid #cbd5e1', outline: 'none' }} />
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '5px' }}>
              <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#475569' }}>Cargo</label>
              <input type="text" value={invitadoEditando.cargo} onChange={(e) => setInvitadoEditando({...invitadoEditando, cargo: e.target.value})} style={{ padding: '8px', borderRadius: '4px', border: '1px solid #cbd5e1', outline: 'none' }} />
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '5px' }}>
              <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#475569' }}>Dependencia (Color)</label>
              <input type="text" value={invitadoEditando.dependencia} onChange={(e) => setInvitadoEditando({...invitadoEditando, dependencia: e.target.value})} style={{ padding: '8px', borderRadius: '4px', border: '1px solid #cbd5e1', outline: 'none' }} placeholder="Ej: UNAM, UV, Externo, DGPyFE, Rectores" />
            </div>
            <div style={{ display: 'flex', gap: '10px', marginTop: '10px' }}>
              <button onClick={() => guardarEdicion(invitadoEditando)} style={{ flex: 1, backgroundColor: '#3b82f6', color: 'white', padding: '10px', borderRadius: '4px', border: 'none', fontWeight: 'bold', cursor: 'pointer' }}>Guardar Cambios</button>
              <button onClick={() => setInvitadoEditando(null)} style={{ flex: 1, backgroundColor: '#cbd5e1', color: '#475569', padding: '10px', borderRadius: '4px', border: 'none', fontWeight: 'bold', cursor: 'pointer' }}>Cancelar</button>
            </div>
          </div>
        </div>
      )}

      <DragDropContext onDragEnd={onDragEnd}>
        
        {/* PANEL BANCA (SOLO VISIBLE EN MODO EDICIÓN Y SI NO ESTÁ EN PANTALLA COMPLETA) */}
        <div style={{ width: '320px', flexShrink: 0, backgroundColor: 'white', padding: '15px', borderRight: '1px solid #cbd5e1', display: (modoEdicion && !modoPresentacion) ? 'flex' : 'none', flexDirection: 'column', zIndex: 10 }}>
          <h2 style={{ fontSize: '18px', fontWeight: 'bold', marginBottom: '15px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            Panel de Edición
            <span style={{ fontSize: '10px', backgroundColor: '#dcfce7', color: '#16a34a', padding: '4px 8px', borderRadius: '12px' }}>🟢 Abierto</span>
          </h2>
          
          <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', marginBottom: '15px' }}>
            <div style={{ display: 'flex', gap: '8px' }}>
              <label style={{ flex: 1, backgroundColor: '#10b981', color: 'white', padding: '8px', textAlign: 'center', borderRadius: '4px', cursor: 'pointer', fontSize: '11px', fontWeight: 'bold' }}>
                + Subir Nuevos <input type="file" accept=".xlsx, .xls" onChange={cargarExcel} style={{ display: 'none' }} />
              </label>
              <label style={{ flex: 1, backgroundColor: '#0ea5e9', color: 'white', padding: '8px', textAlign: 'center', borderRadius: '4px', cursor: 'pointer', fontSize: '11px', fontWeight: 'bold' }} title="Sube el Excel con los textos corregidos en la hoja 'Directorio'">
                🔄 Corregir Textos <input type="file" accept=".xlsx, .xls" onChange={actualizarDesdeExcel} style={{ display: 'none' }} />
              </label>
            </div>
            
            <div style={{ display: 'flex', gap: '8px', marginTop: '5px' }}>
              <button onClick={descargarRespaldo} style={{ flex: 1, backgroundColor: '#6366f1', color: 'white', padding: '10px', borderRadius: '4px', border: 'none', cursor: 'pointer', fontSize: '11px', fontWeight: 'bold' }}>💾 Descargar (JSON)</button>
              <label style={{ flex: 1, backgroundColor: '#f59e0b', color: 'white', padding: '10px', textAlign: 'center', borderRadius: '4px', cursor: 'pointer', fontSize: '11px', fontWeight: 'bold' }}>
                📂 Restaurar (JSON) <input type="file" accept=".json" onChange={cargarRespaldo} style={{ display: 'none' }} />
              </label>
            </div>

            <button onClick={exportarExcel} style={{ backgroundColor: '#16a34a', color: 'white', padding: '10px', borderRadius: '4px', border: 'none', cursor: 'pointer', fontSize: '14px', fontWeight: 'bold', marginTop: '5px' }}>📊 Exportar a Excel (5 Hojas)</button>
            <button onClick={exportarPDF} style={{ backgroundColor: '#ef4444', color: 'white', padding: '10px', borderRadius: '4px', border: 'none', cursor: 'pointer', fontSize: '14px', fontWeight: 'bold' }}>📄 Exportar Plano PDF</button>
          </div>

          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
             <h3 style={{ fontSize: '14px', color: '#475569', margin: 0 }}>En Banca ({(layoutActivo.banca || []).length})</h3>
             <input type="text" placeholder="🔍 Buscar..." value={busquedaBanca} onChange={(e) => setBusquedaBanca(e.target.value)} style={{ width: '130px', padding: '4px 8px', fontSize: '12px', borderRadius: '4px', border: '1px solid #cbd5e1', outline: 'none' }} />
          </div>
          
          <Droppable droppableId="banca" type="invitado" isDropDisabled={!modoEdicion}>
            {(provided) => (
              <div {...provided.droppableProps} ref={provided.innerRef} style={{ flexGrow: 1, overflowY: 'auto', backgroundColor: '#f8fafc', padding: '10px', borderRadius: '4px', minHeight: '100px' }}>
                {(layoutActivo.banca || []).map((invitado, index) => {
                  const isBloqueado = vistaActual === 'auditorio' ? invitado.bloqueado_auditorio : invitado.bloqueado_comida;
                  return <Tarjeta key={invitado.id} invitado={invitado} index={index} isBanca={true} isBloqueado={isBloqueado} onToggleLock={toggleBloqueoBanca} onDelete={eliminarInvitadoDeBanca} busqueda={busquedaBanca} onEdit={setInvitadoEditando} modoEdicion={modoEdicion} />
                })}
                {provided.placeholder}
              </div>
            )}
          </Droppable>
        </div>

        {/* LIENZO PRINCIPAL */}
        <div style={{ flexGrow: 1, overflow: 'auto', padding: '20px', backgroundColor: '#e2e8f0', position: 'relative' }}>
          
          {/* BARRA SUPERIOR FLOTANTE IZQUIERDA (BOTÓN VISTAS) */}
          <div style={{ position: 'fixed', top: '20px', left: (modoEdicion && !modoPresentacion) ? '350px' : '30px', zIndex: 50, transition: 'left 0.2s' }}>
            <button 
              onClick={() => setVistaActual(vistaActual === 'auditorio' ? 'comida' : 'auditorio')}
              style={{ padding: '8px 15px', fontSize: '14px', borderRadius: '8px', border: 'none', backgroundColor: '#3b82f6', color: 'white', fontWeight: 'bold', cursor: 'pointer', boxShadow: '0 2px 10px rgba(0,0,0,0.1)', transition: 'background-color 0.2s' }}
            >
              👁️ Ver {vistaActual === 'auditorio' ? 'Comida' : 'Auditorio'}
            </button>
          </div>

          {/* BARRA SUPERIOR FLOTANTE DERECHA MAESTRA */}
          <div style={{ position: 'fixed', top: '20px', right: '30px', zIndex: 50, display: 'flex', gap: '15px', alignItems: 'center' }}>
            
            {/* BOTÓN DESPLEGABLE DESCARGAR (SOLO LECTURA) */}
            {!modoEdicion && (
              <div style={{ position: 'relative' }}>
                <button 
                  onClick={() => setMostrarMenuDescarga(!mostrarMenuDescarga)} 
                  style={{ padding: '8px 15px', fontSize: '14px', borderRadius: '8px', border: 'none', backgroundColor: '#6366f1', color: 'white', fontWeight: 'bold', cursor: 'pointer', boxShadow: '0 2px 10px rgba(0,0,0,0.1)' }}
                >
                  📥 Descargar ▼
                </button>
                {mostrarMenuDescarga && (
                  <div style={{ position: 'absolute', top: '100%', right: 0, marginTop: '8px', backgroundColor: 'white', borderRadius: '8px', boxShadow: '0 4px 15px rgba(0,0,0,0.1)', display: 'flex', flexDirection: 'column', overflow: 'hidden', width: '220px', zIndex: 100 }}>
                    <button onClick={() => { exportarExcel(); setMostrarMenuDescarga(false); }} style={{ padding: '12px 15px', border: 'none', backgroundColor: 'transparent', textAlign: 'left', cursor: 'pointer', borderBottom: '1px solid #e2e8f0', fontSize: '13px', fontWeight: 'bold', color: '#16a34a' }}>📊 Exportar a Excel (5 Hojas)</button>
                    <button onClick={() => { exportarPDF(); setMostrarMenuDescarga(false); }} style={{ padding: '12px 15px', border: 'none', backgroundColor: 'transparent', textAlign: 'left', cursor: 'pointer', fontSize: '13px', fontWeight: 'bold', color: '#ef4444' }}>📄 Exportar Plano PDF</button>
                  </div>
                )}
              </div>
            )}

            {/* BOTÓN PANTALLA COMPLETA (SOLO EN MODO EDICIÓN) */}
            {modoEdicion && (
              <button 
                onClick={() => setModoPresentacion(!modoPresentacion)}
                title="Ocultar o Mostrar el Panel de Control Lateral"
                style={{ padding: '8px 15px', fontSize: '14px', borderRadius: '8px', border: 'none', backgroundColor: modoPresentacion ? '#f59e0b' : '#3b82f6', color: 'white', fontWeight: 'bold', cursor: 'pointer', boxShadow: '0 2px 10px rgba(0,0,0,0.1)', transition: 'background-color 0.2s' }}
              >
                {modoPresentacion ? '🗗 Salir Pantalla Completa' : '🖥️ Pantalla Completa'}
              </button>
            )}

            {/* BOTÓN ACTIVAR/DESACTIVAR EDICIÓN */}
            <button 
              onClick={() => { setModoEdicion(!modoEdicion); setMostrarMenuDescarga(false); if(modoEdicion) setModoPresentacion(false); }}
              title={modoEdicion ? "Bloquear sembrado y ocultar panel" : "Habilitar panel de control y arrastre"}
              style={{ padding: '8px 15px', fontSize: '14px', borderRadius: '8px', border: 'none', backgroundColor: modoEdicion ? '#f59e0b' : '#10b981', color: 'white', fontWeight: 'bold', cursor: 'pointer', boxShadow: '0 2px 10px rgba(0,0,0,0.1)', transition: 'background-color 0.2s' }}
            >
              {modoEdicion ? '🔒 Cerrar Edición' : '✏️ Editar'}
            </button>

            {modoEdicion && (
              <button 
                onClick={deshacerUltimaAccion} 
                disabled={historial.length === 0} 
                title={historial.length === 0 ? "Nada que deshacer" : "Deshacer último movimiento"}
                style={{ padding: '8px 15px', fontSize: '14px', borderRadius: '8px', border: 'none', backgroundColor: historial.length === 0 ? '#cbd5e1' : '#f43f5e', color: historial.length === 0 ? '#64748b' : 'white', fontWeight: 'bold', cursor: historial.length === 0 ? 'default' : 'pointer', boxShadow: '0 2px 10px rgba(0,0,0,0.1)', transition: 'background-color 0.2s' }}
              >
                ↩️ Deshacer
              </button>
            )}

            <input type="text" placeholder="🔍 Buscar en sembrado..." value={busquedaLienzo} onChange={(e) => setBusquedaLienzo(e.target.value)} style={{ padding: '8px 15px', fontSize: '14px', borderRadius: '8px', border: '1px solid #cbd5e1', outline: 'none', boxShadow: '0 2px 10px rgba(0,0,0,0.1)', width: '250px' }} />
            
            <div style={{ display: 'flex', gap: '5px', backgroundColor: 'white', padding: '5px', borderRadius: '8px', boxShadow: '0 2px 10px rgba(0,0,0,0.1)' }}>
              <button onClick={() => setZoom(prev => Math.max(0.2, prev - 0.1))} style={{ padding: '5px 12px', fontSize: '18px', fontWeight: 'bold', cursor: 'pointer', border: '1px solid #cbd5e1', borderRadius: '4px', backgroundColor: '#f8fafc' }}>-</button>
              <div style={{ padding: '5px 10px', fontWeight: 'bold', minWidth: '60px', textAlign: 'center' }}>{Math.round(zoom * 100)}%</div>
              <button onClick={() => setZoom(prev => Math.min(1.5, prev + 0.1))} style={{ padding: '5px 12px', fontSize: '18px', fontWeight: 'bold', cursor: 'pointer', border: '1px solid #cbd5e1', borderRadius: '4px', backgroundColor: '#f8fafc' }}>+</button>
            </div>
          </div>

          <div ref={pdfRef} style={{ backgroundColor: 'white', padding: '40px', borderRadius: '8px', minWidth: '1300px', minHeight: '100%', zoom: zoom }}>
            
            <h1 style={{ textAlign: 'center', marginBottom: '10px', fontSize: '28px', fontWeight: 'bold', color: '#0f172a' }}>
              {vistaActual === 'auditorio' ? 'Inauguración Filuni 2026' : 'Comida Inaugural'}
            </h1>
            
            <div style={{ textAlign: 'center', marginBottom: '30px', color: '#475569', fontSize: '14px', fontWeight: 'bold', display: 'flex', justifyContent: 'center', gap: '20px' }}>
               {vistaActual === 'auditorio' && (
                 <span style={{ backgroundColor: '#e0f2fe', color: '#0369a1', padding: '6px 12px', borderRadius: '20px' }}>👥 Asistentes Auditorio: {ocupantesAuditorio}</span>
               )}
               {vistaActual === 'comida' && (
                 <span style={{ backgroundColor: '#fef3c7', color: '#b45309', padding: '6px 12px', borderRadius: '20px' }}>🍽️ Asistentes Comida: {ocupantesComida}</span>
               )}
            </div>

            {/* VISTA AUDITORIO */}
            {vistaActual === 'auditorio' && (
              <div id="contenedor-auditorio" style={{ width: '100%', overflowX: 'auto', paddingBottom: '20px' }}>
                <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', minWidth: 'max-content', padding: '0 20px' }}>
                  
                  {/* HOJA 1 */}
                  <div id="auditorio-parte-1" style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '20px', width: '100%', backgroundColor: 'white' }}>
                    <div style={{ backgroundColor: '#1e293b', padding: '20px 40px', borderRadius: '8px', border: '2px solid #0f172a', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.1)', marginBottom: '20px' }}>
                      <h3 style={{ color: 'white', textAlign: 'center', marginBottom: '20px', letterSpacing: '2px' }}>ESTRADO (Presidium)</h3>
                      <div style={{ display: 'flex', gap: '15px', justifyContent: 'center' }}>
                        {Array.from({ length: 8 }, (_, i) => <Silla key={`estrado_silla_${i+1}`} id={`estrado_silla_${i+1}`} ocupante={layoutActivo[`estrado_silla_${i+1}`] || []} vista="auditorio" busqueda={busquedaLienzo} onEdit={setInvitadoEditando} modoEdicion={modoEdicion} />)}
                      </div>
                    </div>
                    
                    {/* ENCABEZADO ASIENTOS - HOJA 1 */}
                    <div style={{ display: 'flex', alignItems: 'flex-start', gap: '15px', backgroundColor: 'transparent', padding: '0 15px', borderRadius: '6px', border: '1px solid transparent', boxSizing: 'border-box' }}>
                      <div style={{ width: '60px', flexShrink: 0 }} /> 
                      <div style={{ display: 'flex', gap: '10px' }}>
                        {Array.from({ length: 13 }, (_, sIndex) => (
                          <div key={`num_guia_b1_${sIndex+1}`} style={{ width: '120px', minWidth: '120px', flexShrink: 0, textAlign: 'center', color: '#64748b', fontSize: '11px', fontWeight: 'bold', textTransform: 'uppercase', letterSpacing: '0.5px', border: '2px solid transparent', boxSizing: 'border-box' }}>
                            Asiento {sIndex + 1}
                          </div>
                        ))}
                      </div>
                    </div>

                    <div style={{ display: 'flex', flexDirection: 'column', gap: '20px', width: '100%', alignItems: 'center' }}>
                      {Array.from({ length: 6 }, (_, fIndex) => {
                        const f = fIndex + 1;
                        return (
                          <div key={`fila_${f}`} style={{ display: 'flex', alignItems: 'flex-start', gap: '15px', backgroundColor: '#f8fafc', padding: '15px', borderRadius: '6px', border: '1px solid #e2e8f0', height: '125px', boxSizing: 'border-box' }}>
                            <strong style={{ width: '60px', textAlign: 'center', color: '#475569', fontSize: '14px', flexShrink: 0, marginTop: '35px' }}>Fila {f}</strong>
                            <div style={{ display: 'flex', gap: '10px' }}>
                              {Array.from({ length: 13 }, (_, sIndex) => {
                                const s = sIndex + 1;
                                return <Silla key={`fila_${f}_silla_${s}`} id={`fila_${f}_silla_${s}`} ocupante={layoutActivo[`fila_${f}_silla_${s}`] || []} vista="auditorio" busqueda={busquedaLienzo} onEdit={setInvitadoEditando} modoEdicion={modoEdicion} />
                              })}
                            </div>
                          </div>
                        )
                      })}
                    </div>
                  </div>

                  {/* PASILLO CENTRAL */}
                  <div style={{ width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '30px 0', height: '45px', backgroundColor: '#cbd5e1', borderRadius: '6px', border: '2px dashed #94a3b8' }}><span style={{ fontWeight: 'bold', letterSpacing: '12px', color: '#475569', fontSize: '16px' }}>P A S I L L O</span></div>

                  {/* HOJA 2 */}
                  <div id="auditorio-parte-2" style={{ display: 'flex', flexDirection: 'column', gap: '20px', width: '100%', alignItems: 'center', backgroundColor: 'white' }}>
                    
                    {/* ENCABEZADO ASIENTOS - HOJA 2 */}
                    <div style={{ display: 'flex', alignItems: 'flex-start', gap: '15px', backgroundColor: 'transparent', padding: '0 15px', borderRadius: '6px', border: '1px solid transparent', boxSizing: 'border-box' }}>
                      <div style={{ width: '60px', flexShrink: 0 }} /> 
                      <div style={{ display: 'flex', gap: '10px' }}>
                        {Array.from({ length: 13 }, (_, sIndex) => (
                          <div key={`num_guia_b2_${sIndex+1}`} style={{ width: '120px', minWidth: '120px', flexShrink: 0, textAlign: 'center', color: '#64748b', fontSize: '11px', fontWeight: 'bold', textTransform: 'uppercase', letterSpacing: '0.5px', border: '2px solid transparent', boxSizing: 'border-box' }}>
                            Asiento {sIndex + 1}
                          </div>
                        ))}
                      </div>
                    </div>

                    {Array.from({ length: 10 }, (_, fIndex) => {
                      const f = fIndex + 7;
                      return (
                        <div key={`fila_${f}`} style={{ display: 'flex', alignItems: 'flex-start', gap: '15px', backgroundColor: '#f8fafc', padding: '15px', borderRadius: '6px', border: '1px solid #e2e8f0', height: '125px', boxSizing: 'border-box' }}>
                          <strong style={{ width: '60px', textAlign: 'center', color: '#475569', fontSize: '14px', flexShrink: 0, marginTop: '35px' }}>Fila {f}</strong>
                          <div style={{ display: 'flex', gap: '10px' }}>
                            {Array.from({ length: 13 }, (_, sIndex) => {
                              const s = sIndex + 1;
                              if (f === 7 && s === 6) return (<div key="tv_unam" style={{ position: 'relative', width: '380px', height: '95px', flexShrink: 0 }}><div style={{ position: 'absolute', top: 0, left: 0, width: '100%', height: '240px', backgroundColor: '#1e293b', color: '#f8fafc', display: 'flex', alignItems: 'center', justifyContent: 'center', borderRadius: '8px', fontWeight: 'bold', fontSize: '18px', letterSpacing: '2px', zIndex: 9999, boxShadow: '0 4px 6px -1px rgba(0,0,0,0.3)' }}>📺 TV UNAM</div></div>);
                              if (f === 7 && (s === 7 || s === 8)) return null;
                              if (f === 8 && s === 6) return <div key="tv_unam_spacer" style={{ width: '380px', height: '95px', flexShrink: 0 }} />;
                              if (f === 8 && (s === 7 || s === 8)) return null;
                              if (f >= 9 && f <= 16 && (s >= 6 && s <= 8)) return <div key={`pasillo_${f}_${s}`} style={{ width: '120px', height: '95px', flexShrink: 0 }} />;
                              
                              return <Silla key={`fila_${f}_silla_${s}`} id={`fila_${f}_silla_${s}`} ocupante={layoutActivo[`fila_${f}_silla_${s}`] || []} vista="auditorio" busqueda={busquedaLienzo} onEdit={setInvitadoEditando} modoEdicion={modoEdicion} />
                            })}
                          </div>
                        </div>
                      )
                    })}
                  </div>
                </div>
              </div>
            )}

            {/* VISTA COMIDA */}
            {vistaActual === 'comida' && (
              <div id="contenedor-comida" style={{ width: '100%', overflowX: 'auto', paddingBottom: '20px' }}>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '40px', alignItems: 'center', minWidth: 'max-content', padding: '0 20px' }}>
                  
                  {bloquesMesas.map((bloque, indexFila) => (
                    <Droppable key={`mesas_fila_${indexFila}`} droppableId={`mesas_${indexFila}`} direction="horizontal" type="mesa" isDropDisabled={!modoEdicion}>
                      {(provided, snapshot) => (
                        <div 
                           id={`comida-parte-${indexFila + 1}`}
                           ref={provided.innerRef} 
                           {...provided.droppableProps} 
                           style={{ display: 'flex', flexDirection: 'row', flexWrap: 'nowrap', justifyContent: 'center', gap: '40px', padding: '20px', backgroundColor: snapshot.isDraggingOver ? '#f8fafc' : 'white', borderRadius: '8px', minHeight: '350px', minWidth: '1100px', transition: 'background-color 0.2s' }}
                        >
                           {bloque.map((m, index) => renderMesaLayout(m, index))}
                           {provided.placeholder}
                        </div>
                      )}
                    </Droppable>
                  ))}

                </div>
              </div>
            )}
          </div>
        </div>
      </DragDropContext>
    </div>
  );
}

// --- 9. COMPONENTES REFINADOS ---
function Silla({ id, ocupante, vista, busqueda, onEdit, modoEdicion }) {
  const isAuditorio = vista === 'auditorio'; const width = isAuditorio ? '120px' : '100%'; const minWidth = isAuditorio ? '120px' : '85px'; const height = isAuditorio ? '95px' : '75px'; const flexShrink = isAuditorio ? 0 : 1; const arrOcupante = Array.isArray(ocupante) ? ocupante : [];
  const estaOcupada = arrOcupante.length >= 1;

  return (
    <Droppable droppableId={id} type="invitado" isDropDisabled={!modoEdicion}>
      {(provided, snapshot) => {
        let borderColor = estaOcupada ? '2px solid transparent' : '2px dashed #94a3b8';
        let bgColor = estaOcupada ? 'transparent' : 'rgba(255,255,255,0.5)';
        
        if (snapshot.isDraggingOver && modoEdicion) {
           borderColor = estaOcupada ? '2px dashed #f59e0b' : '2px dashed #3b82f6';
           bgColor = estaOcupada ? '#fef3c7' : '#eff6ff';
        }

        return (
          <div ref={provided.innerRef} {...provided.droppableProps} style={{ width, minWidth, height, flexShrink, border: borderColor, backgroundColor: bgColor, borderRadius: '4px', display: 'flex', alignItems: 'center', justifyContent: 'center', boxSizing: 'border-box', transition: 'background-color 0.2s, border 0.2s' }}>
            {!estaOcupada && <span style={{ fontSize: '10px', color: snapshot.isDraggingOver ? '#3b82f6' : '#94a3b8', fontWeight: 'bold' }}>{id.includes('silla') ? id.split('_').pop() : 'Silla'}</span>}
            {arrOcupante.map((invitado, index) => {
              const isBloq = vista === 'auditorio' ? invitado.bloqueado_auditorio : invitado.bloqueado_comida;
              return <Tarjeta key={invitado.id} invitado={invitado} index={index} enSilla={true} isBloqueado={isBloq} busqueda={busqueda} onEdit={onEdit} modoEdicion={modoEdicion} />
            })}
            <div style={{ display: 'none' }}>{provided.placeholder}</div>
          </div>
        )
      }}
    </Droppable>
  );
}

function Tarjeta({ invitado, index, enSilla, isBanca, isBloqueado, onToggleLock, onDelete, busqueda, onEdit, modoEdicion }) {
  const colorBorde = getColorDependencia(invitado.dependencia); const textoBusqueda = busqueda ? busqueda.toLowerCase() : '';
  const coincideBusqueda = textoBusqueda !== '' && (invitado.nombre.toLowerCase().includes(textoBusqueda) || invitado.cargo.toLowerCase().includes(textoBusqueda) || invitado.dependencia.toLowerCase().includes(textoBusqueda));
  
  const opacity = (textoBusqueda !== '' && !coincideBusqueda) ? 0.2 : (isBloqueado ? 0.5 : 1); 
  const bgCard = isBloqueado ? '#e2e8f0' : 'white';
  const glow = coincideBusqueda ? '0 0 15px 4px #ec4899' : '0 2px 4px rgba(0,0,0,0.1)'; 
  const escala = coincideBusqueda ? 'scale(1.02)' : 'scale(1)';

  return (
    <Draggable draggableId={invitado.id} index={index} isDragDisabled={!modoEdicion || (isBanca && isBloqueado)}>
      {(provided) => (
        <div ref={provided.innerRef} {...provided.draggableProps} {...provided.dragHandleProps} onDoubleClick={() => modoEdicion && onEdit(invitado)} title={modoEdicion ? "Doble clic para editar" : ""} style={{ userSelect: 'none', padding: '6px', backgroundColor: bgCard, border: '1px solid #cbd5e1', borderTop: `4px solid ${colorBorde}`, borderRadius: '4px', boxShadow: glow, transform: provided.draggableProps.style?.transform || escala, width: enSilla ? '100%' : '100%', height: enSilla ? '100%' : 'auto', marginBottom: enSilla ? '0' : '10px', display: 'flex', flexDirection: 'column', justifyContent: 'center', alignItems: 'center', textAlign: 'center', boxSizing: 'border-box', position: 'relative', transition: 'box-shadow 0.3s, opacity 0.3s', opacity, ...provided.draggableProps.style }}>
          
          {isBanca && modoEdicion && (
            <>
              <button onPointerDown={(e) => e.stopPropagation()} onClick={() => onToggleLock(invitado.id)} title={isBloqueado ? "Desbloquear" : "Bloquear (mandar al final de la lista)"} style={{ position: 'absolute', top: '-6px', left: '-6px', width: '18px', height: '18px', backgroundColor: isBloqueado ? '#8b5cf6' : '#94a3b8', color: 'white', border: 'none', borderRadius: '50%', fontSize: '10px', fontWeight: 'bold', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', boxShadow: '0 2px 4px rgba(0,0,0,0.2)' }} >{isBloqueado ? '🔓' : '🔒'}</button>
              <button onPointerDown={(e) => e.stopPropagation()} onClick={() => onDelete(invitado.id)} title="Eliminar del evento" style={{ position: 'absolute', top: '-6px', right: '-6px', width: '18px', height: '18px', backgroundColor: '#ef4444', color: 'white', border: 'none', borderRadius: '50%', fontSize: '10px', fontWeight: 'bold', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', boxShadow: '0 2px 4px rgba(0,0,0,0.2)' }} >×</button>
            </>
          )}

          <div style={{ fontSize: '10px', fontWeight: 'bold', lineHeight: '1.2', marginBottom: '3px', overflow: 'hidden', display: '-webkit-box', WebkitLineClamp: '2', WebkitBoxOrient: 'vertical' }}>{invitado.nombre}</div>
          <div style={{ fontSize: '9px', color: '#64748b', lineHeight: '1.15', overflow: 'hidden', display: '-webkit-box', WebkitLineClamp: '3', WebkitBoxOrient: 'vertical', width: '100%' }}>{invitado.cargo}</div>
        </div>
      )}
    </Draggable>
  );
}