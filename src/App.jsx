import React, { useState, useEffect, useRef } from 'react';
import { DragDropContext, Droppable, Draggable } from '@hello-pangea/dnd';
import * as XLSX from 'xlsx-js-style'; 
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';
import { db } from './firebase'; // <-- CONEXIÓN A FIREBASE
import { doc, onSnapshot, setDoc } from 'firebase/firestore';

// --- 1. GENERADORES DE ESTRUCTURAS ---
const initAuditorio = () => {
  const layout = { banca: [] };
  for (let i = 1; i <= 7; i++) layout[`estrado_silla_${i}`] = [];
  
  for (let f = 1; f <= 10; f++) {
    for (let s = 1; s <= 13; s++) {
      if ((f === 5 || f === 6) && (s >= 6 && s <= 8)) continue;
      if (f >= 7 && f <= 10 && s === 7) continue;
      layout[`fila_${f}_silla_${s}`] = [];
    }
  }
  return layout;
};

const initComida = () => {
  const layout = { banca: [] };
  for (let m = 1; m <= 11; m++) { 
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
  { bg: '#ffe4e6', border: '#f43f5e' }, { bg: '#fae8ff', border: '#d946ef' } 
];

export default function App() {
  const [vistaActual, setVistaActual] = useState('auditorio'); 
  const [auditorio, setAuditorio] = useState(initAuditorio());
  const [comida, setComida] = useState(initComida());
  const [nombresMesas, setNombresMesas] = useState({
    mesa_1: 'Mesa 1', mesa_2: 'Mesa 2', mesa_3: 'Mesa 3', mesa_4: 'Mesa 4', mesa_5: 'Mesa 5',
    mesa_6: 'Mesa 6', mesa_7: 'Mesa 7', mesa_8: 'Mesa 8', mesa_9: 'Mesa 9', mesa_10: 'Mesa 10', mesa_11: 'Mesa 11' 
  });
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
      } else {
        // Si el documento no existe en la nube (primer uso), lo creamos
        setDoc(docRef, { auditorio: initAuditorio(), comida: initComida(), nombresMesas });
      }
    });
    return () => unsub(); // Limpiar el listener al cerrar
  }, []);

  // Función maestra para guardar en la nube
  const syncToCloud = async (nuevoAuditorio, nuevaComida, nuevosNombres) => {
    try {
      await setDoc(doc(db, 'eventos', 'filuni2026'), {
        auditorio: nuevoAuditorio || auditorio,
        comida: nuevaComida || comida,
        nombresMesas: nuevosNombres || nombresMesas
      });
    } catch(error) {
      console.error("Error sincronizando a Firebase:", error);
    }
  };

  // --- 3. LÓGICAS DE ELIMINACIÓN Y EDICIÓN ---
  const eliminarInvitadoDeBanca = (id) => {
    if (!window.confirm("¿Seguro que deseas eliminar a este invitado de la banca?")) return;
    const nuevoAuditorio = { ...auditorio, banca: auditorio.banca.filter(inv => inv.id !== id) };
    const nuevaComida = { ...comida, banca: comida.banca.filter(inv => inv.id !== id) };
    syncToCloud(nuevoAuditorio, nuevaComida, null);
  };

  const guardarEdicion = (invitadoActualizado) => {
    const actualizarLayout = (layout) => {
      const nuevoLayout = { ...layout };
      Object.keys(nuevoLayout).forEach(key => {
        nuevoLayout[key] = nuevoLayout[key].map(inv => inv.id === invitadoActualizado.id ? invitadoActualizado : inv);
      });
      return nuevoLayout;
    };
    
    const nuevoAuditorio = actualizarLayout(auditorio);
    const nuevaComida = actualizarLayout(comida);
    syncToCloud(nuevoAuditorio, nuevaComida, null);
    setInvitadoEditando(null); 
  };

  const manejarEdicionNombreMesa = (m, valor) => {
    setNombresMesas(prev => ({...prev, [`mesa_${m}`]: valor}));
  };

  const manejarBlurNombreMesa = () => {
    // Sincroniza al terminar de escribir para no saturar la base de datos
    syncToCloud(null, null, nombresMesas);
  };

  // --- 4. EXPORTACIÓN A EXCEL (Intacta) ---
  const exportarExcel = () => {
    const celdaVaciaBase = { alignment: { wrapText: true, vertical: 'top', horizontal: 'center' } };

    const datosAuditorio = [];
    Object.keys(auditorio).forEach(key => {
      if (key === 'banca') return;
      const ocupantes = auditorio[key] || [];
      const ocupante = ocupantes[0];
      let ubicacionStr = key.includes('estrado') ? key.replace('estrado_silla_', 'Estrado - Silla ') : `Fila ${key.split('_')[1]} - Silla ${key.split('_')[3]}`;
      datosAuditorio.push({ 'Ubicación': ubicacionStr, 'Dependencia': ocupante ? ocupante.dependencia : '', 'Nombre': ocupante ? ocupante.nombre : '', 'Cargo': ocupante ? ocupante.cargo : '' });
    });

    const datosComida = [];
    Object.keys(comida).forEach(key => {
      if (key === 'banca') return;
      const ocupantes = comida[key || []] || [];
      const ocupante = ocupantes[0];
      const p = key.split('_');
      datosComida.push({ 'Mesa': nombresMesas[`mesa_${p[1]}`], 'Asiento': `Silla ${p[3]}`, 'Dependencia': ocupante ? ocupante.dependencia : '', 'Nombre': ocupante ? ocupante.nombre : '', 'Cargo': ocupante ? ocupante.cargo : '' });
    });

    const matrizAuditorio = [];
    matrizAuditorio.push([{ v: "ESTRADO (Presidium)", s: { font: { bold: true } } }]);
    const filaEstrado = [];
    for(let i=1; i<=7; i++) {
        const sillaInfo = auditorio[`estrado_silla_${i}`] || [];
        const oc = sillaInfo[0];
        filaEstrado.push(oc ? { v: `${oc.nombre}\n${oc.cargo}`, t: 's', s: getExcelStyle(oc.dependencia) } : { v: "[ Vacío ]", t: 's', s: celdaVaciaBase });
    }
    matrizAuditorio.push(filaEstrado);
    matrizAuditorio.push([]); 
    
    for(let f=1; f<=10; f++) {
        matrizAuditorio.push([{ v: `FILA ${f}`, s: { font: { bold: true } } }]);
        const filaAsientos = [];
        for(let s=1; s<=13; s++) {
             if ((f === 5 || f === 6) && (s >= 6 && s <= 8)) {
                 if (s === 7) filaAsientos.push({ v: "[ TV UNAM ]", t: 's', s: { fill: { patternType: 'solid', fgColor: { rgb: "1E293B" } }, font: { color: { rgb: "FFFFFF" }, bold: true }, alignment: { vertical: 'center', horizontal: 'center' } } });
                 else filaAsientos.push({ v: "", t: 's', s: celdaVaciaBase });
                 continue;
             }
             if (f >= 7 && f <= 10 && s === 7) {
                 filaAsientos.push({ v: "", t: 's', s: celdaVaciaBase }); 
                 continue;
             }
             const sillaInfo = auditorio[`fila_${f}_silla_${s}`] || [];
             const oc = sillaInfo[0];
             filaAsientos.push(oc ? { v: `${oc.nombre}\n${oc.cargo}`, t: 's', s: getExcelStyle(oc.dependencia) } : { v: "[ Vacío ]", t: 's', s: celdaVaciaBase });
        }
        matrizAuditorio.push(filaAsientos);
        if (f === 6) {
          matrizAuditorio.push([]);
          matrizAuditorio.push(["", "", "", "", "", { v: "============= P A S I L L O =============", t: 's', s: { font: { bold: true, color: { rgb: "475569" } }, alignment: { horizontal: 'center' } } }]);
        }
        matrizAuditorio.push([]); 
    }

    const matrizComida = [];
    for(let m=1; m<=11; m++) { 
        matrizComida.push([{ v: nombresMesas[`mesa_${m}`], s: { font: { bold: true } } }]);
        for(let fila=0; fila<5; fila++) {
            const s1 = (fila * 2) + 1;
            const s2 = (fila * 2) + 2;
            const o1 = (comida[`mesa_${m}_silla_${s1}`] || [])[0];
            const o2 = (comida[`mesa_${m}_silla_${s2}`] || [])[0];
            matrizComida.push([
              o1 ? { v: `${o1.nombre}\n${o1.cargo}`, t: 's', s: getExcelStyle(o1.dependencia) } : { v: "[ Vacío ]", t: 's', s: celdaVaciaBase }, 
              o2 ? { v: `${o2.nombre}\n${o2.cargo}`, t: 's', s: getExcelStyle(o2.dependencia) } : { v: "[ Vacío ]", t: 's', s: celdaVaciaBase }
            ]);
        }
        matrizComida.push([]);
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(datosAuditorio), "Lista Auditorio");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(datosComida), "Lista Comida");
    const ws3 = XLSX.utils.aoa_to_sheet(matrizAuditorio); ws3['!cols'] = Array(16).fill({ wch: 25 }); XLSX.utils.book_append_sheet(wb, ws3, "Gráfico Auditorio");
    const ws4 = XLSX.utils.aoa_to_sheet(matrizComida); ws4['!cols'] = Array(6).fill({ wch: 25 }); XLSX.utils.book_append_sheet(wb, ws4, "Gráfico Comida");
    XLSX.writeFile(wb, "Sembrado_Invitados_Completo.xlsx");
  };

  // --- 5. CARGAR EXCEL (A la nube) ---
  const cargarExcel = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const filasExcel = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);

        const nuevosInvitados = filasExcel.map((fila, index) => {
          const filaNormalizada = {};
          Object.keys(fila).forEach(key => { filaNormalizada[key.toLowerCase().trim()] = fila[key]; });
          return { id: `excel-${Date.now()}-${index}`, dependencia: filaNormalizada.dependencia || '', nombre: filaNormalizada.nombre || 'Sin Nombre', cargo: filaNormalizada.cargo || 'Sin Cargo' };
        });

        const nuevoAuditorio = { ...auditorio, banca: [...auditorio.banca, ...nuevosInvitados] };
        const nuevaComida = { ...comida, banca: [...comida.banca, ...nuevosInvitados] };
        
        syncToCloud(nuevoAuditorio, nuevaComida, null);
        alert(`✅ Se cargaron y sincronizaron ${nuevosInvitados.length} invitados.`);
        e.target.value = null; 
      } catch (error) { alert("❌ Error al leer el Excel. Revisa el formato."); }
    };
    reader.readAsArrayBuffer(file);
  };

  // --- 6. EXPORTAR PDF ---
  const exportarPDF = async () => {
    const pdf = new jsPDF('l', 'mm', 'a4'); 
    const pdfWidth = pdf.internal.pageSize.getWidth(); 
    const pdfHeightSheet = pdf.internal.pageSize.getHeight(); 

    const contAuditorio = document.getElementById('contenedor-auditorio');
    const contComida = document.getElementById('contenedor-comida');
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
        pdf.setFont('helvetica', 'normal'); pdf.setFontSize(9); pdf.setTextColor(100, 116, 139); pdf.text(`Fecha de generación: ${fechaHoy}`, pdfWidth - 10, pdfHeightSheet - 7, { align: 'right' });
      }
      pdf.save(`Protocolo-${vistaActual}.pdf`);
    } finally {
      if (contAuditorio) contAuditorio.style.overflowX = overAuditorio;
      if (contComida) contComida.style.overflowX = overComida;
    }
  };

  // --- 7. ARRASTRAR, SOLTAR Y SINCRONIZAR ---
  const onDragEnd = (result) => {
    const { source, destination } = result;
    if (!destination) return;

    const estadoActivo = vistaActual === 'auditorio' ? auditorio : comida;

    if (source.droppableId === destination.droppableId) {
      const nuevaLista = Array.from(estadoActivo[source.droppableId] || []);
      const [movido] = nuevaLista.splice(source.index, 1);
      nuevaLista.splice(destination.index, 0, movido);
      
      const nuevoEstado = { ...estadoActivo, [source.droppableId]: nuevaLista };
      syncToCloud(vistaActual === 'auditorio' ? nuevoEstado : null, vistaActual === 'comida' ? nuevoEstado : null, null);
      return;
    }

    if (destination.droppableId !== 'banca' && (estadoActivo[destination.droppableId] || []).length >= 1) {
      alert('Esta silla ya está ocupada.');
      return;
    }

    const nuevoEstado = { ...estadoActivo };
    const origenLista = Array.from(nuevoEstado[source.droppableId] || []);
    const destinoLista = Array.from(nuevoEstado[destination.droppableId] || []);

    const [invitadoMovido] = origenLista.splice(source.index, 1);
    destinoLista.splice(destination.index, 0, invitadoMovido);

    nuevoEstado[source.droppableId] = origenLista;
    nuevoEstado[destination.droppableId] = destinoLista;

    syncToCloud(vistaActual === 'auditorio' ? nuevoEstado : null, vistaActual === 'comida' ? nuevoEstado : null, null);
  };

  const layoutActivo = vistaActual === 'auditorio' ? auditorio : comida;

  const renderMesaLayout = (m) => (
    <div key={`mesa_${m}`} style={{ width: '320px', backgroundColor: paletaMesas[m - 1].bg, border: `3px solid ${paletaMesas[m - 1].border}`, borderRadius: '8px', padding: '20px', display: 'flex', flexDirection: 'column', alignItems: 'center', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.1)', flexShrink: 0 }}>
      <div style={{ width: '100%', marginBottom: '15px', display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
        <input 
          type="text" 
          value={nombresMesas[`mesa_${m}`] || ''} 
          onChange={(e) => manejarEdicionNombreMesa(m, e.target.value)}
          onBlur={manejarBlurNombreMesa} 
          style={{ width: '100%', height: '40px', lineHeight: '36px', textAlign: 'center', fontWeight: 'bold', fontSize: '16px', padding: '0 10px', borderRadius: '4px', border: '1px solid white', backgroundColor: 'rgba(255,255,255,0.7)', outline: 'none', boxSizing: 'border-box', fontFamily: 'sans-serif' }} 
        />
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px', width: '100%' }}>
        {Array.from({ length: 10 }, (_, s) => <Silla key={`mesa_${m}_silla_${s+1}`} id={`mesa_${m}_silla_${s+1}`} ocupante={layoutActivo[`mesa_${m}_silla_${s+1}`] || []} vista="comida" busqueda={busquedaLienzo} onEdit={setInvitadoEditando} />)}
      </div>
    </div>
  );

  return (
    <div style={{ display: 'flex', height: '100vh', width: '100vw', fontFamily: 'sans-serif', backgroundColor: '#f1f5f9', overflow: 'hidden' }}>
      
      {/* MODAL DE EDICIÓN FLOTANTE */}
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
        {/* BANCA Y CONTROLES */}
        <div style={{ width: '320px', flexShrink: 0, backgroundColor: 'white', padding: '15px', borderRight: '1px solid #cbd5e1', display: 'flex', flexDirection: 'column', zIndex: 10 }}>
          <h2 style={{ fontSize: '18px', fontWeight: 'bold', marginBottom: '15px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            Panel de Control
            <span style={{ fontSize: '10px', backgroundColor: '#dcfce7', color: '#16a34a', padding: '4px 8px', borderRadius: '12px' }}>🟢 Sincronizado</span>
          </h2>
          
          <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', marginBottom: '15px' }}>
            <div style={{ display: 'flex', gap: '8px' }}>
              <label style={{ flex: 1, backgroundColor: '#10b981', color: 'white', padding: '8px', textAlign: 'center', borderRadius: '4px', cursor: 'pointer', fontSize: '13px', fontWeight: 'bold' }}>
                + Cargar Lista
                <input type="file" accept=".xlsx, .xls" onChange={cargarExcel} style={{ display: 'none' }} />
              </label>
              <button onClick={() => setVistaActual(vistaActual === 'auditorio' ? 'comida' : 'auditorio')} style={{ flex: 1, backgroundColor: '#3b82f6', color: 'white', padding: '8px', borderRadius: '4px', border: 'none', cursor: 'pointer', fontSize: '13px', fontWeight: 'bold' }}>
                Ver {vistaActual === 'auditorio' ? 'Comida' : 'Auditorio'}
              </button>
            </div>
            <hr style={{ borderTop: '1px solid #e2e8f0', margin: '5px 0' }} />
            <button onClick={exportarExcel} style={{ backgroundColor: '#16a34a', color: 'white', padding: '10px', borderRadius: '4px', border: 'none', cursor: 'pointer', fontSize: '14px', fontWeight: 'bold' }}>
              📊 Exportar a Excel (4 Hojas)
            </button>
            <button onClick={exportarPDF} style={{ backgroundColor: '#ef4444', color: 'white', padding: '10px', borderRadius: '4px', border: 'none', cursor: 'pointer', fontSize: '14px', fontWeight: 'bold' }}>
              📄 Exportar Plano PDF
            </button>
          </div>

          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
             <h3 style={{ fontSize: '14px', color: '#475569', margin: 0 }}>En Banca ({(layoutActivo.banca || []).length})</h3>
             <input type="text" placeholder="🔍 Buscar..." value={busquedaBanca} onChange={(e) => setBusquedaBanca(e.target.value)} style={{ width: '130px', padding: '4px 8px', fontSize: '12px', borderRadius: '4px', border: '1px solid #cbd5e1', outline: 'none' }} />
          </div>
          
          <Droppable droppableId="banca">
            {(provided) => (
              <div {...provided.droppableProps} ref={provided.innerRef} style={{ flexGrow: 1, overflowY: 'auto', backgroundColor: '#f8fafc', padding: '10px', borderRadius: '4px', minHeight: '100px' }}>
                {(layoutActivo.banca || []).map((invitado, index) => (
                  <Tarjeta key={invitado.id} invitado={invitado} index={index} isBanca={true} onDelete={eliminarInvitadoDeBanca} busqueda={busquedaBanca} onEdit={setInvitadoEditando} />
                ))}
                {provided.placeholder}
              </div>
            )}
          </Droppable>
        </div>

        {/* LIENZO PRINCIPAL */}
        <div style={{ flexGrow: 1, overflow: 'auto', padding: '20px', backgroundColor: '#e2e8f0', position: 'relative' }}>
          <div style={{ position: 'fixed', top: '20px', right: '30px', zIndex: 50, display: 'flex', gap: '15px', alignItems: 'center' }}>
            <input type="text" placeholder="🔍 Buscar en sembrado..." value={busquedaLienzo} onChange={(e) => setBusquedaLienzo(e.target.value)} style={{ padding: '8px 15px', fontSize: '14px', borderRadius: '8px', border: '1px solid #cbd5e1', outline: 'none', boxShadow: '0 2px 10px rgba(0,0,0,0.1)', width: '250px' }} />
            <div style={{ display: 'flex', gap: '5px', backgroundColor: 'white', padding: '5px', borderRadius: '8px', boxShadow: '0 2px 10px rgba(0,0,0,0.1)' }}>
              <button onClick={() => setZoom(prev => Math.max(0.2, prev - 0.1))} style={{ padding: '5px 12px', fontSize: '18px', fontWeight: 'bold', cursor: 'pointer', border: '1px solid #cbd5e1', borderRadius: '4px', backgroundColor: '#f8fafc' }}>-</button>
              <div style={{ padding: '5px 10px', fontWeight: 'bold', minWidth: '60px', textAlign: 'center' }}>{Math.round(zoom * 100)}%</div>
              <button onClick={() => setZoom(prev => Math.min(1.5, prev + 0.1))} style={{ padding: '5px 12px', fontSize: '18px', fontWeight: 'bold', cursor: 'pointer', border: '1px solid #cbd5e1', borderRadius: '4px', backgroundColor: '#f8fafc' }}>+</button>
            </div>
          </div>

          <div ref={pdfRef} style={{ backgroundColor: 'white', padding: '40px', borderRadius: '8px', minWidth: '1300px', minHeight: '100%', zoom: zoom }}>
            <h1 style={{ textAlign: 'center', marginBottom: '40px', fontSize: '28px', fontWeight: 'bold', color: '#0f172a' }}>
              {vistaActual === 'auditorio' ? 'Sembrado: Inauguración Filuni 2026' : 'Sembrado: Comida Inaugural'}
            </h1>

            {/* VISTA AUDITORIO */}
            {vistaActual === 'auditorio' && (
              <div id="contenedor-auditorio" style={{ width: '100%', overflowX: 'auto', paddingBottom: '20px' }}>
                <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', minWidth: 'max-content', padding: '0 20px' }}>
                  
                  {/* HOJA 1 */}
                  <div id="auditorio-parte-1" style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '20px', width: '100%', backgroundColor: 'white' }}>
                    <div style={{ backgroundColor: '#1e293b', padding: '20px 40px', borderRadius: '8px', border: '2px solid #0f172a', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.1)', marginBottom: '20px' }}>
                      <h3 style={{ color: 'white', textAlign: 'center', marginBottom: '20px', letterSpacing: '2px' }}>ESTRADO (Presidium)</h3>
                      <div style={{ display: 'flex', gap: '15px', justifyContent: 'center' }}>
                        {Array.from({ length: 7 }, (_, i) => <Silla key={`estrado_silla_${i+1}`} id={`estrado_silla_${i+1}`} ocupante={layoutActivo[`estrado_silla_${i+1}`] || []} vista="auditorio" busqueda={busquedaLienzo} onEdit={setInvitadoEditando} />)}
                      </div>
                    </div>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '15px', padding: '0 15px', height: '30px', width: '100%' }}>
                      <div style={{ width: '60px', flexShrink: 0 }} /> 
                      <div style={{ display: 'flex', gap: '10px' }}>
                        {Array.from({ length: 13 }, (_, sIndex) => (<div key={`num_guia_b1_${sIndex+1}`} style={{ width: '120px', minWidth: '120px', textAlign: 'center', color: '#64748b', fontSize: '11px', fontWeight: 'bold', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Asiento {sIndex + 1}</div>))}
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
                                if (f === 5 && s === 6) return (<div key="tv_unam" style={{ position: 'relative', width: '380px', height: '95px', flexShrink: 0 }}><div style={{ position: 'absolute', top: 0, left: 0, width: '100%', height: '240px', backgroundColor: '#1e293b', color: '#f8fafc', display: 'flex', alignItems: 'center', justifyContent: 'center', borderRadius: '8px', fontWeight: 'bold', fontSize: '18px', letterSpacing: '2px', zIndex: 9999, boxShadow: '0 4px 6px -1px rgba(0,0,0,0.3)' }}>📺 TV UNAM</div></div>);
                                if (f === 5 && (s === 7 || s === 8)) return null;
                                if (f === 6 && s === 6) return <div key="tv_unam_spacer" style={{ width: '380px', height: '95px', flexShrink: 0 }} />;
                                if (f === 6 && (s === 7 || s === 8)) return null;
                                return <Silla key={`fila_${f}_silla_${s}`} id={`fila_${f}_silla_${s}`} ocupante={layoutActivo[`fila_${f}_silla_${s}`] || []} vista="auditorio" busqueda={busquedaLienzo} onEdit={setInvitadoEditando} />
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
                    <div style={{ display: 'flex', alignItems: 'center', gap: '15px', padding: '0 15px', height: '30px', width: '100%' }}>
                      <div style={{ width: '60px', flexShrink: 0 }} /> 
                      <div style={{ display: 'flex', gap: '10px' }}>
                        {Array.from({ length: 13 }, (_, sIndex) => (<div key={`num_guia_b2_${sIndex+1}`} style={{ width: '120px', minWidth: '120px', textAlign: 'center', color: '#64748b', fontSize: '11px', fontWeight: 'bold', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Asiento {sIndex + 1}</div>))}
                      </div>
                    </div>
                    {Array.from({ length: 4 }, (_, fIndex) => {
                      const f = fIndex + 7;
                      return (
                        <div key={`fila_${f}`} style={{ display: 'flex', alignItems: 'flex-start', gap: '15px', backgroundColor: '#f8fafc', padding: '15px', borderRadius: '6px', border: '1px solid #e2e8f0', height: '125px', boxSizing: 'border-box' }}>
                          <strong style={{ width: '60px', textAlign: 'center', color: '#475569', fontSize: '14px', flexShrink: 0, marginTop: '35px' }}>Fila {f}</strong>
                          <div style={{ display: 'flex', gap: '10px' }}>
                            {Array.from({ length: 13 }, (_, sIndex) => {
                              const s = sIndex + 1;
                              if (f >= 7 && f <= 10 && s === 7) return <div key={`pasillo_${f}_7`} style={{ width: '120px', height: '95px', flexShrink: 0 }} />;
                              return <Silla key={`fila_${f}_silla_${s}`} id={`fila_${f}_silla_${s}`} ocupante={layoutActivo[`fila_${f}_silla_${s}`] || []} vista="auditorio" busqueda={busquedaLienzo} onEdit={setInvitadoEditando} />
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
                  <div id="comida-parte-1" style={{ display: 'flex', flexDirection: 'row', flexWrap: 'nowrap', justifyContent: 'center', gap: '40px', padding: '20px', backgroundColor: 'white', borderRadius: '8px' }}>{Array.from({ length: 3 }, (_, i) => renderMesaLayout(i + 1))}</div>
                  <div id="comida-parte-2" style={{ display: 'flex', flexDirection: 'row', flexWrap: 'nowrap', justifyContent: 'center', gap: '40px', padding: '20px', backgroundColor: 'white', borderRadius: '8px' }}>{Array.from({ length: 3 }, (_, i) => renderMesaLayout(i + 4))}</div>
                  <div id="comida-parte-3" style={{ display: 'flex', flexDirection: 'row', flexWrap: 'nowrap', justifyContent: 'center', gap: '40px', padding: '20px', backgroundColor: 'white', borderRadius: '8px' }}>{Array.from({ length: 3 }, (_, i) => renderMesaLayout(i + 7))}</div>
                  <div id="comida-parte-4" style={{ display: 'flex', flexDirection: 'row', flexWrap: 'nowrap', justifyContent: 'center', gap: '40px', padding: '20px', backgroundColor: 'white', borderRadius: '8px' }}>{Array.from({ length: 2 }, (_, i) => renderMesaLayout(i + 10))}</div>
                </div>
              </div>
            )}
          </div>
        </div>
      </DragDropContext>
    </div>
  );
}

// --- 8. COMPONENTES REFINADOS ---
function Silla({ id, ocupante, vista, busqueda, onEdit }) {
  const isAuditorio = vista === 'auditorio'; const width = isAuditorio ? '120px' : '100%'; const minWidth = isAuditorio ? '120px' : '85px'; const height = isAuditorio ? '95px' : '75px'; const flexShrink = isAuditorio ? 0 : 1; const arrOcupante = Array.isArray(ocupante) ? ocupante : [];
  return (
    <Droppable droppableId={id}>
      {(provided) => (
        <div ref={provided.innerRef} {...provided.droppableProps} style={{ width, minWidth, height, flexShrink, border: arrOcupante.length === 0 ? '2px dashed #94a3b8' : 'none', backgroundColor: arrOcupante.length === 0 ? 'rgba(255,255,255,0.5)' : 'transparent', borderRadius: '4px', display: 'flex', alignItems: 'center', justifyContent: 'center', boxSizing: 'border-box' }}>
          {arrOcupante.length === 0 && <span style={{ fontSize: '10px', color: '#94a3b8', fontWeight: 'bold' }}>{id.includes('silla') ? id.split('_').pop() : 'Silla'}</span>}
          {arrOcupante.map((invitado, index) => <Tarjeta key={invitado.id} invitado={invitado} index={index} enSilla={true} busqueda={busqueda} onEdit={onEdit} />)}
          <div style={{ display: 'none' }}>{provided.placeholder}</div>
        </div>
      )}
    </Droppable>
  );
}

function Tarjeta({ invitado, index, enSilla, isBanca, onDelete, busqueda, onEdit }) {
  const colorBorde = getColorDependencia(invitado.dependencia); const textoBusqueda = busqueda ? busqueda.toLowerCase() : '';
  const coincideBusqueda = textoBusqueda !== '' && (invitado.nombre.toLowerCase().includes(textoBusqueda) || invitado.cargo.toLowerCase().includes(textoBusqueda) || invitado.dependencia.toLowerCase().includes(textoBusqueda));
  const opacity = (textoBusqueda !== '' && !coincideBusqueda) ? 0.2 : 1; const glow = coincideBusqueda ? '0 0 15px 4px #ec4899' : '0 2px 4px rgba(0,0,0,0.1)'; const escala = coincideBusqueda ? 'scale(1.02)' : 'scale(1)';

  return (
    <Draggable draggableId={invitado.id} index={index}>
      {(provided) => (
        <div ref={provided.innerRef} {...provided.draggableProps} {...provided.dragHandleProps} onDoubleClick={() => onEdit(invitado)} title="Doble clic para editar" style={{ userSelect: 'none', padding: '6px', backgroundColor: 'white', border: '1px solid #cbd5e1', borderTop: `4px solid ${colorBorde}`, borderRadius: '4px', boxShadow: glow, transform: provided.draggableProps.style?.transform || escala, width: enSilla ? '100%' : '100%', height: enSilla ? '100%' : 'auto', marginBottom: enSilla ? '0' : '10px', display: 'flex', flexDirection: 'column', justifyContent: 'center', alignItems: 'center', textAlign: 'center', boxSizing: 'border-box', position: 'relative', transition: 'box-shadow 0.3s, opacity 0.3s', opacity, ...provided.draggableProps.style }}>
          {isBanca && (<button onPointerDown={(e) => e.stopPropagation()} onClick={() => onDelete(invitado.id)} title="Eliminar invitado" style={{ position: 'absolute', top: '-6px', right: '-6px', width: '18px', height: '18px', backgroundColor: '#ef4444', color: 'white', border: 'none', borderRadius: '50%', fontSize: '10px', fontWeight: 'bold', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', boxShadow: '0 2px 4px rgba(0,0,0,0.2)' }} >×</button>)}
          <div style={{ fontSize: '10px', fontWeight: 'bold', lineHeight: '1.2', marginBottom: '3px', overflow: 'hidden', display: '-webkit-box', WebkitLineClamp: '2', WebkitBoxOrient: 'vertical' }}>{invitado.nombre}</div>
          <div style={{ fontSize: '9px', color: '#64748b', lineHeight: '1.15', overflow: 'hidden', display: '-webkit-box', WebkitLineClamp: '3', WebkitBoxOrient: 'vertical', width: '100%' }}>{invitado.cargo}</div>
        </div>
      )}
    </Draggable>
  );
}