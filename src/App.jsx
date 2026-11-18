import React, { useState, useRef } from 'react';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import {
  Image as ImageIcon,
  Table,
  Unlock,
  Heart,
  Upload,
  CheckCircle2,
  Lock,
  ArrowLeft,
  FileText,
  ArrowRight,
  Check,
  ArrowUpRight,
  X,
  FolderArchive,
  FileSpreadsheet,
  Eye,
  CheckSquare,
  Square,
} from 'lucide-react';

export default function App() {
  const [selectedTool, setSelectedTool] = useState(null);
  const [showDonateModal, setShowDonateModal] = useState(false);
  const [donateTab, setDonateTab] = useState('peru'); // 'peru' | 'intl'
  const [processing, setProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState(null);
  const [errorMsg, setErrorMsg] = useState(null);
  const [docFile, setDocFile] = useState(null);
  const fileInputRef = useRef(null);

  // Estados específicos para Herramienta 2: Selector inteligente de tablas
  const [detectedTables, setDetectedTables] = useState([]); // [{ id, rows, previewHeaders, rowCount, colCount }]
  const [selectedTableIds, setSelectedTableIds] = useState([]);
  const [excelLayoutMode, setExcelLayoutMode] = useState('separate'); // 'separate' | 'single'

  const PAYPAL_DONATE_URL = 'https://www.paypal.com/donate/?hosted_button_id=A4YJTUMDD5JN6';
  const DEVSTORE_URL = 'https://devstore.vercel.app';

  const resetStates = (tool) => {
    setSelectedTool(tool);
    setDocFile(null);
    setStatusMsg(null);
    setErrorMsg(null);
    setDetectedTables([]);
    setSelectedTableIds([]);
    setExcelLayoutMode('separate');
  };

  const triggerBlobDownload = (blob, filename) => {
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Analiza las tablas apenas el usuario selecciona un .docx en la Herramienta 2
  const parseWordTables = async (file) => {
    setProcessing(true);
    setErrorMsg(null);
    setStatusMsg(null);
    setDetectedTables([]);
    setSelectedTableIds([]);

    try {
      const arrayBuffer = await file.arrayBuffer();
      const zip = await JSZip.loadAsync(arrayBuffer);
      const docXmlFile = zip.file('word/document.xml');

      if (!docXmlFile) {
        setErrorMsg('No se pudo leer el contenido interno de este archivo Word.');
        setProcessing(false);
        return;
      }

      const xmlText = await docXmlFile.async('text');
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(xmlText, 'application/xml');
      const tables = Array.from(xmlDoc.getElementsByTagName('w:tbl'));

      if (tables.length === 0) {
        setErrorMsg('Este documento Word no contiene ninguna tabla o cuadro.');
        setProcessing(false);
        return;
      }

      const parsedList = tables.map((tbl, idx) => {
        const rows = Array.from(tbl.getElementsByTagName('w:tr'));
        const matrix = rows.map((tr) => {
          const cells = Array.from(tr.getElementsByTagName('w:tc'));
          return cells.map((tc) => {
            const paragraphs = Array.from(tc.getElementsByTagName('w:p'));
            return paragraphs
              .map((p) => {
                const texts = Array.from(p.getElementsByTagName('w:t'));
                return texts.map((t) => t.textContent).join('');
              })
              .join(' ')
              .trim();
          });
        });

        const firstRow = matrix[0] || [];
        const previewHeaders = firstRow.filter(Boolean).slice(0, 4).join(' · ') || 'Tabla sin encabezados';

        return {
          id: idx,
          matrix,
          previewHeaders,
          rowCount: matrix.length,
          colCount: firstRow.length,
        };
      });

      setDetectedTables(parsedList);
      setSelectedTableIds(parsedList.map((t) => t.id)); // Todas marcadas por defecto
    } catch (err) {
      setErrorMsg('Error: Asegúrate de subir un archivo Word válido en formato .docx.');
    }
    setProcessing(false);
  };

  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setDocFile(file);
    setStatusMsg(null);
    setErrorMsg(null);

    if (selectedTool === 'word-to-excel') {
      parseWordTables(file);
    }
  };

  const toggleTableSelection = (id) => {
    setSelectedTableIds((prev) =>
      prev.includes(id) ? prev.filter((item) => item !== id) : [...prev, id]
    );
  };

  // Lógica 1: Extraer todas las imágenes en alta calidad de un .docx
  const handleExtractImages = async (e) => {
    e.preventDefault();
    if (!docFile) return;
    setProcessing(true);
    setStatusMsg(null);
    setErrorMsg(null);

    try {
      const arrayBuffer = await docFile.arrayBuffer();
      const zip = await JSZip.loadAsync(arrayBuffer);

      const mediaFiles = Object.keys(zip.files).filter(
        (path) => path.startsWith('word/media/') && !zip.files[path].dir
      );

      if (mediaFiles.length === 0) {
        setErrorMsg('Este documento Word no contiene ninguna imagen incrustada.');
        setProcessing(false);
        return;
      }

      const baseName = docFile.name.replace(/\.docx$/i, '');

      if (mediaFiles.length === 1) {
        const fileData = await zip.files[mediaFiles[0]].async('blob');
        const ext = mediaFiles[0].split('.').pop();
        triggerBlobDownload(fileData, `${baseName}_imagen_1.${ext}`);
        setStatusMsg('¡Listo! Se extrajo y descargó la imagen en su calidad original.');
      } else {
        const outZip = new JSZip();
        for (let i = 0; i < mediaFiles.length; i++) {
          const path = mediaFiles[i];
          const ext = path.split('.').pop();
          const data = await zip.files[path].async('uint8array');
          outZip.file(`imagen_${String(i + 1).padStart(2, '0')}.${ext}`, data);
        }
        const zipBlob = await outZip.generateAsync({ type: 'blob' });
        triggerBlobDownload(zipBlob, `IMAGENES_${baseName}.zip`);
        setStatusMsg(
          `¡Listo! Se extrajeron ${mediaFiles.length} imágenes en alta calidad dentro de un archivo ZIP.`
        );
      }
    } catch (err) {
      setErrorMsg('Error: Asegúrate de subir un archivo Word válido en formato .docx.');
    }
    setProcessing(false);
  };

  // Lógica 2: Exportar tablas seleccionadas a Excel (en hojas separadas o en una sola hoja)
  const handleExportSelectedTables = (e) => {
    e.preventDefault();
    if (!docFile || selectedTableIds.length === 0) return;

    const chosenTables = detectedTables.filter((t) => selectedTableIds.includes(t.id));
    const workbook = XLSX.utils.book_new();

    if (excelLayoutMode === 'separate' || chosenTables.length === 1) {
      chosenTables.forEach((tbl) => {
        const worksheet = XLSX.utils.aoa_to_sheet(tbl.matrix);
        XLSX.utils.book_append_sheet(workbook, worksheet, `Tabla_${tbl.id + 1}`);
      });
    } else {
      // Unir todas las tablas elegidas en una sola hoja con una fila vacía entre cada tabla
      const combinedRows = [];
      chosenTables.forEach((tbl, index) => {
        if (index > 0) combinedRows.push([]);
        combinedRows.push([`--- Tabla ${tbl.id + 1} ---`]);
        tbl.matrix.forEach((row) => combinedRows.push(row));
      });
      const worksheet = XLSX.utils.aoa_to_sheet(combinedRows);
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Tablas_Unificadas');
    }

    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const excelBlob = new Blob([excelBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });

    const baseName = docFile.name.replace(/\.docx$/i, '');
    triggerBlobDownload(excelBlob, `TABLAS_${baseName}.xlsx`);
    setStatusMsg(
      `¡Listo! Se exportaron ${chosenTables.length} tabla(s) a Excel (.xlsx) según tu configuración.`
    );
  };

  // Lógica 3: Desbloquear Word protegido (eliminar restricción de edición)
  const handleUnlockWord = async (e) => {
    e.preventDefault();
    if (!docFile) return;
    setProcessing(true);
    setStatusMsg(null);
    setErrorMsg(null);

    try {
      const arrayBuffer = await docFile.arrayBuffer();
      const zip = await JSZip.loadAsync(arrayBuffer);
      const settingsFile = zip.file('word/settings.xml');

      if (!settingsFile) {
        setErrorMsg('Este archivo .docx no tiene configuración de bloqueo interna.');
        setProcessing(false);
        return;
      }

      let settingsXml = await settingsFile.async('text');
      const hadProtection = /<w:documentProtection\b[^>]*\/?>/i.test(settingsXml);

      settingsXml = settingsXml.replace(/<w:documentProtection\b[^>]*\/?>/gi, '');
      zip.file('word/settings.xml', settingsXml);

      const newDocxBlob = await zip.generateAsync({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });

      const baseName = docFile.name.replace(/\.docx$/i, '');
      triggerBlobDownload(newDocxBlob, `EDITABLE_${baseName}.docx`);

      if (hadProtection) {
        setStatusMsg('¡Bloqueo eliminado! Descargaste una copia 100% editable sin perder el diseño.');
      } else {
        setStatusMsg(
          'El archivo ya estaba libre de bloqueo de edición; se descargó una copia limpia lista para editar.'
        );
      }
    } catch (err) {
      setErrorMsg('Error: Asegúrate de subir un archivo Word en formato .docx.');
    }
    setProcessing(false);
  };

  const tools = [
    {
      id: 'extract-images',
      badge: 'CALIDAD ORIGINAL',
      title: '1. Extraer Imágenes de un Word',
      desc: 'Saca todas las fotos, logos y gráficos incrustados dentro de un archivo Word (.docx) en su resolución original en un solo clic.',
      bullets: [
        'Evita guardar imagen por imagen con clic derecho y baja calidad.',
        'Extrae automáticamente todas las fotos ordenadas en una carpeta ZIP.',
        'Ideal para informes, guías, catálogos y trabajos recibidos.',
      ],
      icon: <ImageIcon size={34} className="text-[#185ABD]" />,
      btnText: 'Subir archivo Word (.docx)',
    },
    {
      id: 'word-to-excel',
      badge: 'SELECCIÓN INTELIGENTE',
      title: '2. Exportar Tablas de Word a Excel',
      desc: 'Detecta los cuadros dentro de tu Word, te deja elegir cuáles exportar y decidir si van en hojas separadas o juntas en una sola hoja.',
      bullets: [
        'Elige exactamente qué tablas quieres exportar si tu Word tiene varias.',
        'Decide si cada tabla va en una pestaña distinta o todas en la misma hoja.',
        'Evita que se mezclen las celdas o se pierdan números al copiar y pegar.',
      ],
      icon: <Table size={34} className="text-emerald-600" />,
      btnText: 'Subir Word con tablas (.docx)',
    },
    {
      id: 'unlock-word',
      badge: '100% EDITABLE',
      title: '3. Desbloquear Word Protegido',
      desc: 'Quita la restricción de "Solo lectura" o bloqueo de edición en formularios, contratos y plantillas Word sin alterar ni un margen.',
      bullets: [
        'Elimina el candado de edición cuando un formato no te deja escribir.',
        'Mantiene intactos los cuadros, fuentes, encabezados y diseño original.',
        'Funciona al instante sin subir tu documento confidencial a internet.',
      ],
      icon: <Unlock size={34} className="text-slate-800" />,
      btnText: 'Subir Word bloqueado (.docx)',
    },
  ];

  const activeToolObj = tools.find((t) => t.id === selectedTool);

  const renderResultExample = () => {
    if (selectedTool === 'extract-images') {
      return (
        <div className="space-y-4">
          <div className="flex items-center justify-between text-xs bg-slate-100 px-3.5 py-2.5 rounded-xl border border-slate-200">
            <span className="font-semibold text-slate-600 truncate">Informe_Final.docx</span>
            <ArrowRight size={14} className="text-[#185ABD] shrink-0 mx-2" />
            <span className="font-bold text-[#185ABD] flex items-center gap-1 shrink-0">
              <FolderArchive size={14} /> IMAGENES.zip
            </span>
          </div>

          <div className="grid grid-cols-2 gap-2.5">
            {[
              { name: 'imagen_01.png', size: '1920 × 1080 px · Original' },
              { name: 'imagen_02.jpg', size: '1600 × 1200 px · Original' },
              { name: 'imagen_03.png', size: '1280 × 720 px · Original' },
              { name: 'imagen_04.jpg', size: '2048 × 1536 px · Original' },
            ].map((img, i) => (
              <div
                key={i}
                className="p-3 rounded-xl bg-slate-50 border border-slate-200/90 flex items-start gap-2.5"
              >
                <div className="w-8 h-8 rounded-lg bg-blue-100 text-[#185ABD] flex items-center justify-center shrink-0">
                  <ImageIcon size={16} />
                </div>
                <div className="min-w-0">
                  <p className="text-xs font-bold text-slate-800 truncate">{img.name}</p>
                  <p className="text-[10px] text-emerald-600 font-medium">{img.size}</p>
                </div>
              </div>
            ))}
          </div>

          <p className="text-[11px] text-slate-500 leading-relaxed">
            Obtendrás todas las fotos sueltas dentro de una carpeta comprimida sin pérdida de nitidez.
          </p>
        </div>
      );
    }

    if (selectedTool === 'word-to-excel') {
      return (
        <div className="space-y-4">
          <div className="flex items-center justify-between text-xs bg-slate-100 px-3.5 py-2.5 rounded-xl border border-slate-200">
            <span className="font-semibold text-slate-600 truncate">Reporte_Ventas.docx</span>
            <ArrowRight size={14} className="text-emerald-600 shrink-0 mx-2" />
            <span className="font-bold text-emerald-700 flex items-center gap-1 shrink-0">
              <FileSpreadsheet size={14} /> TABLAS.xlsx
            </span>
          </div>

          <div className="border border-slate-200 rounded-xl overflow-hidden text-xs">
            <div className="bg-emerald-700 text-white px-3 py-1.5 font-semibold text-[11px] flex items-center justify-between">
              <span>HOJA: Tabla_1 (o Unificadas)</span>
              <span>Celdas separadas</span>
            </div>
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-slate-100 border-b border-slate-200 text-[11px] text-slate-700">
                  <th className="py-2 px-3 border-r border-slate-200">Producto</th>
                  <th className="py-2 px-3 border-r border-slate-200">Cant.</th>
                  <th className="py-2 px-3">Total (S/)</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 text-[11px] text-slate-600 bg-white">
                <tr>
                  <td className="py-1.5 px-3 border-r border-slate-200">Soporte Técnico</td>
                  <td className="py-1.5 px-3 border-r border-slate-200">12</td>
                  <td className="py-1.5 px-3 font-semibold text-slate-900">960.00</td>
                </tr>
                <tr>
                  <td className="py-1.5 px-3 border-r border-slate-200">Mantenimiento PC</td>
                  <td className="py-1.5 px-3 border-r border-slate-200">8</td>
                  <td className="py-1.5 px-3 font-semibold text-slate-900">960.00</td>
                </tr>
              </tbody>
            </table>
          </div>

          <p className="text-[11px] text-slate-500 leading-relaxed">
            Puedes marcar solo las tablas que te interesan y elegir si las quieres en pestañas separadas o juntas en una sola hoja.
          </p>
        </div>
      );
    }

    return (
      <div className="space-y-4">
        <div className="flex items-center justify-between text-xs bg-slate-100 px-3.5 py-2.5 rounded-xl border border-slate-200">
          <span className="font-semibold text-slate-600 truncate">Contrato_Bloqueado.docx</span>
          <ArrowRight size={14} className="text-slate-900 shrink-0 mx-2" />
          <span className="font-bold text-slate-950 flex items-center gap-1 shrink-0">
            <Unlock size={14} className="text-emerald-600" /> EDITABLE.docx
          </span>
        </div>

        <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-2.5">
          <div className="flex items-center justify-between">
            <span className="text-[11px] font-bold text-red-600 bg-red-50 border border-red-200 px-2 py-0.5 rounded-md line-through">
              Solo lectura / Restringido
            </span>
            <span className="text-[11px] font-bold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-md">
              ✓ Edición libre activa
            </span>
          </div>
          <div className="p-3 bg-white rounded-lg border border-slate-200 text-xs text-slate-700 space-y-1">
            <p className="font-bold text-slate-900">CLÁUSULA PRIMERA: OBJETO DEL CONTRATO</p>
            <p className="text-[11px] text-slate-500">
              Ahora puedes escribir, corregir textos y llenar los recuadros del formulario sin perder el formato original...
            </p>
          </div>
        </div>

        <p className="text-[11px] text-slate-500 leading-relaxed">
          Elimina el bloqueo interno manteniendo intactas las tablas, firmas, márgenes y tipografías.
        </p>
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-[#F5F7FB] text-slate-800 flex flex-col justify-between font-sans">
      {/* Barra Superior Corporativa con Sello DevStore */}
      <header className="bg-white/95 backdrop-blur-md border-b border-slate-200/80 sticky top-0 z-20">
        <div className="max-w-6xl mx-auto px-4 h-16 flex items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <button
              onClick={() => resetStates(null)}
              className="flex items-center gap-2 text-left focus:outline-none cursor-pointer"
            >
              <span className="text-2xl font-black tracking-tight text-slate-900">
                Express<span className="text-[#185ABD]">Word</span>
              </span>
            </button>

            <span className="hidden sm:inline-block h-4 w-px bg-slate-200" />

            <a
              href={DEVSTORE_URL}
              target="_blank"
              rel="noreferrer"
              className="hidden sm:inline-flex items-center gap-1.5 text-xs font-medium text-slate-500 hover:text-slate-950 transition"
            >
              <span>Un producto de</span>
              <span className="inline-flex items-center gap-1 font-bold text-slate-900 bg-slate-100 hover:bg-slate-200 px-2 py-0.5 rounded-md transition">
                <span className="w-3.5 h-3.5 rounded-xs bg-slate-950 text-white text-[9px] font-bold flex items-center justify-center">
                  DS
                </span>
                DevStore
              </span>
            </a>
          </div>

          <div className="flex items-center gap-2.5">
            <a
              href={DEVSTORE_URL}
              target="_blank"
              rel="noreferrer"
              className="hidden md:inline-flex items-center gap-1.5 text-xs font-semibold text-slate-700 hover:text-slate-950 px-3.5 py-2 rounded-lg border border-slate-200 hover:border-slate-400 bg-white transition"
            >
              <span>Más herramientas</span>
              <ArrowUpRight size={14} />
            </a>

            <button
              onClick={() => setShowDonateModal(true)}
              className="bg-slate-950 hover:bg-slate-800 text-white font-semibold text-xs sm:text-sm px-4 py-2 rounded-lg flex items-center gap-2 transition shadow-xs cursor-pointer"
            >
              <Heart size={14} className="text-white fill-white" />
              <span>Respaldar proyecto</span>
            </button>
          </div>
        </div>
      </header>

      {/* Contenido Principal */}
      <main className="flex-1 flex flex-col items-center justify-center px-4 py-10">
        {!selectedTool ? (
          <div className="max-w-6xl w-full text-center">
            <h1 className="text-3xl sm:text-4xl font-extrabold text-slate-900 mb-3">
              Resuelve tareas difíciles de Word en segundos
            </h1>
            <p className="text-slate-600 text-base sm:text-lg max-w-2xl mx-auto mb-6">
              Extrae imágenes originales, pasa tablas a Excel o desbloquea documentos protegidos sin instalar nada.
            </p>

            {/* Guía visual de 3 pasos */}
            <div className="inline-flex flex-wrap items-center justify-center gap-2 sm:gap-4 bg-white px-5 py-2.5 rounded-full border border-slate-200 text-xs sm:text-sm font-medium text-slate-700 mb-10 shadow-xs">
              <span className="flex items-center gap-1.5">
                <span className="w-5 h-5 rounded-full bg-[#185ABD] text-white text-xs flex items-center justify-center font-bold">
                  1
                </span>
                Elige una herramienta
              </span>
              <span className="text-slate-300">•</span>
              <span className="flex items-center gap-1.5">
                <span className="w-5 h-5 rounded-full bg-[#185ABD] text-white text-xs flex items-center justify-center font-bold">
                  2
                </span>
                Sube tu archivo .docx
              </span>
              <span className="text-slate-300">•</span>
              <span className="flex items-center gap-1.5">
                <span className="w-5 h-5 rounded-full bg-[#185ABD] text-white text-xs flex items-center justify-center font-bold">
                  3
                </span>
                Descarga el resultado
              </span>
            </div>

            {/* Tarjetas de Herramientas */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 text-left">
              {tools.map((tool) => (
                <div
                  key={tool.id}
                  onClick={() => resetStates(tool.id)}
                  className="bg-white p-7 rounded-2xl shadow-sm hover:shadow-md border-2 border-slate-200/70 hover:border-[#185ABD] transition cursor-pointer flex flex-col justify-between group"
                >
                  <div>
                    <div className="flex items-center justify-between mb-4">
                      <div className="p-3 rounded-xl bg-slate-50 group-hover:scale-105 transition">
                        {tool.icon}
                      </div>
                      <span className="text-[11px] font-bold tracking-wider uppercase px-2.5 py-1 rounded-full bg-slate-100 text-slate-600">
                        {tool.badge}
                      </span>
                    </div>

                    <h2 className="text-xl font-extrabold text-slate-900 mb-2 group-hover:text-[#185ABD] transition">
                      {tool.title}
                    </h2>

                    <p className="text-sm text-slate-600 leading-relaxed mb-5">{tool.desc}</p>

                    <div className="border-t border-slate-100 pt-4 mb-6 space-y-2.5">
                      <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">
                        ¿Para qué te sirve?
                      </p>
                      {tool.bullets.map((item, idx) => (
                        <div key={idx} className="flex items-start gap-2 text-xs text-slate-600 leading-snug">
                          <Check size={15} className="text-emerald-600 shrink-0 mt-0.5" />
                          <span>{item}</span>
                        </div>
                      ))}
                    </div>
                  </div>

                  <div className="w-full py-3 px-4 rounded-xl bg-blue-50 group-hover:bg-[#185ABD] text-[#185ABD] group-hover:text-white font-bold text-sm flex items-center justify-center gap-2 transition">
                    <span>Entrar a esta opción</span>
                    <ArrowRight size={16} />
                  </div>
                </div>
              ))}
            </div>

            <div className="mt-10 inline-flex items-center gap-2 bg-white px-4 py-2 rounded-full border border-slate-200 text-xs text-slate-600 shadow-xs">
              <Lock size={14} className="text-emerald-600 shrink-0" />
              <span>
                <strong>Privacidad garantizada:</strong> Tus documentos nunca se suben a servidores externos; todo se procesa dentro de tu navegador.
              </span>
            </div>
          </div>
        ) : (
          /* PANTALLA DE ACCIÓN DIRECTA CON EJEMPLO DEL RESULTADO AL LADO */
          <div className="max-w-5xl w-full">
            <div className="mb-6 flex justify-center">
              <button
                type="button"
                onClick={() => resetStates(null)}
                className="inline-flex items-center gap-2 text-xs sm:text-sm font-bold text-slate-800 bg-white hover:bg-slate-950 hover:text-white border-2 border-slate-300 hover:border-slate-950 px-5 py-2.5 rounded-xl shadow-xs transition cursor-pointer"
              >
                <ArrowLeft size={16} />
                <span>Volver y elegir otra opción</span>
              </button>
            </div>

            <div className="text-center mb-8">
              <h1 className="text-3xl font-extrabold text-slate-900 mb-2">{activeToolObj.title}</h1>
              <p className="text-slate-600 text-base max-w-2xl mx-auto leading-relaxed">
                {activeToolObj.desc}
              </p>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-stretch">
              {/* Columna Izquierda: Acción de subida y descarga */}
              <div className="lg:col-span-7 bg-white p-7 sm:p-9 rounded-2xl shadow-sm border border-slate-200 flex flex-col justify-center">
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                  className="hidden"
                  onChange={handleFileChange}
                />

                {!docFile ? (
                  <div className="py-6 text-center">
                    <button
                      type="button"
                      onClick={() => fileInputRef.current.click()}
                      className="w-full sm:w-auto bg-[#185ABD] hover:bg-[#134896] text-white font-bold text-base sm:text-lg px-7 py-5 rounded-xl shadow-lg shadow-blue-600/20 inline-flex items-center justify-center gap-3 transition transform hover:-translate-y-0.5 cursor-pointer"
                    >
                      <Upload size={24} className="shrink-0" />
                      <span>{activeToolObj.btnText}</span>
                    </button>
                    <p className="text-xs text-slate-500 mt-4">
                      Paso 1 de 2: Selecciona un archivo Word en formato .docx desde tu equipo
                    </p>
                  </div>
                ) : (
                  <form
                    onSubmit={
                      selectedTool === 'extract-images'
                        ? handleExtractImages
                        : selectedTool === 'word-to-excel'
                        ? handleExportSelectedTables
                        : handleUnlockWord
                    }
                    className="space-y-5 text-left"
                  >
                    <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 flex items-center justify-between">
                      <div className="flex items-center gap-3 overflow-hidden">
                        <FileText className="text-[#185ABD] shrink-0" size={24} />
                        <div className="truncate">
                          <p className="text-sm font-semibold text-slate-800 truncate">{docFile.name}</p>
                          <p className="text-xs text-emerald-600 font-medium">
                            ✓ Documento cargado listo
                          </p>
                        </div>
                      </div>
                      <button
                        type="button"
                        onClick={() => fileInputRef.current.click()}
                        className="text-xs font-semibold text-[#185ABD] hover:underline shrink-0 ml-2 cursor-pointer"
                      >
                        Cambiar archivo
                      </button>
                    </div>

                    {/* Selector interactivo cuando es Herramienta 2 y detecta tablas */}
                    {selectedTool === 'word-to-excel' && detectedTables.length > 0 && (
                      <div className="space-y-4 bg-slate-50/80 border border-slate-200 rounded-xl p-4">
                        <div className="flex items-center justify-between">
                          <p className="text-xs font-extrabold uppercase tracking-wider text-slate-700">
                            Se encontraron {detectedTables.length} tabla(s) · Elige cuáles exportar:
                          </p>
                          <button
                            type="button"
                            onClick={() =>
                              setSelectedTableIds(
                                selectedTableIds.length === detectedTables.length
                                  ? []
                                  : detectedTables.map((t) => t.id)
                              )
                            }
                            className="text-[11px] font-bold text-[#185ABD] hover:underline cursor-pointer"
                          >
                            {selectedTableIds.length === detectedTables.length
                              ? 'Desmarcar todas'
                              : 'Marcar todas'}
                          </button>
                        </div>

                        <div className="space-y-2 max-h-44 overflow-y-auto pr-1">
                          {detectedTables.map((tbl) => {
                            const isChecked = selectedTableIds.includes(tbl.id);
                            return (
                              <div
                                key={tbl.id}
                                onClick={() => toggleTableSelection(tbl.id)}
                                className={`p-3 rounded-xl border transition cursor-pointer flex items-start gap-3 ${
                                  isChecked
                                    ? 'bg-white border-emerald-600 shadow-2xs'
                                    : 'bg-slate-100/70 border-slate-200 opacity-65'
                                }`}
                              >
                                <div className="mt-0.5 text-emerald-600 shrink-0">
                                  {isChecked ? <CheckSquare size={18} /> : <Square size={18} />}
                                </div>
                                <div className="min-w-0 flex-1">
                                  <div className="flex items-center justify-between gap-2">
                                    <span className="text-xs font-bold text-slate-900">
                                      Tabla {tbl.id + 1}
                                    </span>
                                    <span className="text-[11px] font-medium bg-slate-100 px-2 py-0.5 rounded text-slate-600">
                                      {tbl.rowCount} filas × {tbl.colCount} cols
                                    </span>
                                  </div>
                                  <p className="text-[11px] text-slate-500 truncate mt-0.5">
                                    {tbl.previewHeaders}
                                  </p>
                                </div>
                              </div>
                            );
                          })}
                        </div>

                        {/* Si elige más de 1 tabla, puede decidir si van en pestañas separadas o juntas en 1 hoja */}
                        {selectedTableIds.length > 1 && (
                          <div className="pt-2 border-t border-slate-200/80">
                            <p className="text-[11px] font-bold text-slate-700 mb-2">
                              ¿Cómo quieres organizar las {selectedTableIds.length} tablas en el Excel?
                            </p>
                            <div className="grid grid-cols-2 gap-2">
                              <button
                                type="button"
                                onClick={() => setExcelLayoutMode('separate')}
                                className={`py-2 px-3 rounded-lg border text-xs font-bold transition cursor-pointer ${
                                  excelLayoutMode === 'separate'
                                    ? 'bg-slate-950 text-white border-slate-950'
                                    : 'bg-white text-slate-700 border-slate-300 hover:border-slate-500'
                                }`}
                              >
                                En pestañas separadas
                              </button>
                              <button
                                type="button"
                                onClick={() => setExcelLayoutMode('single')}
                                className={`py-2 px-3 rounded-lg border text-xs font-bold transition cursor-pointer ${
                                  excelLayoutMode === 'single'
                                    ? 'bg-emerald-700 text-white border-emerald-700'
                                    : 'bg-white text-slate-700 border-slate-300 hover:border-slate-500'
                                }`}
                              >
                                Juntas en 1 sola hoja
                              </button>
                            </div>
                          </div>
                        )}
                      </div>
                    )}

                    {/* Oculta el botón de descarga si es Word-to-Excel y no hay tablas o no marcó ninguna */}
                    {(selectedTool !== 'word-to-excel' || detectedTables.length > 0) && (
                      <button
                        type="submit"
                        disabled={
                          processing ||
                          (selectedTool === 'word-to-excel' && selectedTableIds.length === 0)
                        }
                        className="w-full bg-[#185ABD] hover:bg-[#134896] disabled:bg-slate-300 text-white font-bold text-lg py-4 rounded-xl shadow-md transition cursor-pointer"
                      >
                        {processing
                          ? 'Procesando documento...'
                          : selectedTool === 'word-to-excel'
                          ? `Descargar Excel (${selectedTableIds.length} tabla${
                              selectedTableIds.length === 1 ? '' : 's'
                            })`
                          : 'Paso 2 de 2: Procesar y Descargar'}
                      </button>
                    )}
                  </form>
                )}

                {errorMsg && (
                  <div className="mt-6 p-4 rounded-xl bg-red-50 border border-red-200 text-left text-sm text-red-800 font-medium">
                    {errorMsg}
                  </div>
                )}

                {statusMsg && (
                  <div className="mt-6 p-4 rounded-xl bg-emerald-50 border border-emerald-200 flex flex-col gap-3 text-left">
                    <div className="flex items-center gap-2.5 text-sm text-emerald-800 font-medium">
                      <CheckCircle2 size={20} className="text-emerald-600 shrink-0" />
                      <span>{statusMsg}</span>
                    </div>
                    <button
                      onClick={() => setShowDonateModal(true)}
                      className="self-start text-xs font-bold bg-white hover:bg-slate-50 text-slate-900 border border-slate-300 px-3.5 py-2 rounded-lg shadow-xs transition cursor-pointer"
                    >
                      Apóyanos para seguir mejorando la plataforma ❤️
                    </button>
                  </div>
                )}
              </div>

              {/* Columna Derecha: Ejemplo visual de cómo queda el resultado */}
              <div className="lg:col-span-5 bg-white p-6 sm:p-7 rounded-2xl shadow-sm border border-slate-200 flex flex-col justify-between text-left">
                <div>
                  <div className="flex items-center gap-2 mb-4 pb-3 border-b border-slate-100">
                    <div className="w-7 h-7 rounded-lg bg-slate-950 text-white flex items-center justify-center">
                      <Eye size={14} />
                    </div>
                    <div>
                      <p className="text-xs font-extrabold uppercase tracking-wider text-slate-900">
                        Ejemplo del resultado
                      </p>
                      <p className="text-[11px] text-slate-500">Así se entregará tu archivo al descargar</p>
                    </div>
                  </div>

                  {renderResultExample()}
                </div>
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Modal Ejecutivo de Respaldo */}
      {showDonateModal && (
        <div className="fixed inset-0 bg-slate-950/65 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-3xl max-w-md w-full p-7 text-center space-y-5 shadow-2xl border border-slate-200 relative">
            <button
              onClick={() => setShowDonateModal(false)}
              aria-label="Cerrar modal"
              className="absolute top-5 right-5 w-8 h-8 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-600 flex items-center justify-center transition cursor-pointer"
            >
              <X size={16} />
            </button>

            <div className="space-y-1.5 pt-1">
              <div className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-slate-950 text-white text-[11px] font-semibold tracking-wider">
                <Heart size={11} className="fill-white" />
                <span>SOSTENIBILIDAD · DEVSTORE</span>
              </div>
              <h3 className="text-xl font-extrabold text-slate-950 tracking-tight">
                Respalda el desarrollo independiente
              </h3>
              <p className="text-xs sm:text-sm text-slate-600 leading-relaxed">
                Mantenemos ExpressWord 100% gratuito y libre de publicidad invasiva. Tu aporte voluntario impulsa nuevas herramientas digitales.
              </p>
            </div>

            <div className="grid grid-cols-2 gap-1.5 p-1 bg-slate-100 rounded-xl border border-slate-200">
              <button
                type="button"
                onClick={() => setDonateTab('peru')}
                className={`py-2 px-3 rounded-lg text-xs font-bold transition cursor-pointer ${
                  donateTab === 'peru'
                    ? 'bg-slate-950 text-white shadow-2xs'
                    : 'text-slate-600 hover:text-slate-950'
                }`}
              >
                Perú · Yape / Plin
              </button>
              <button
                type="button"
                onClick={() => setDonateTab('intl')}
                className={`py-2 px-3 rounded-lg text-xs font-bold transition cursor-pointer ${
                  donateTab === 'intl'
                    ? 'bg-slate-950 text-white shadow-2xs'
                    : 'text-slate-600 hover:text-slate-950'
                }`}
              >
                Internacional · PayPal
              </button>
            </div>

            {donateTab === 'peru' ? (
              <div className="bg-slate-50 border border-slate-200 rounded-2xl p-5 space-y-3">
                <p className="text-[11px] font-bold text-slate-600 uppercase tracking-wider">
                  Escanea con Yape o Plin
                </p>
                <img
                  src="/qr-yape.jpeg"
                  alt="QR Yape"
                  className="w-44 h-44 mx-auto rounded-xl object-contain bg-white p-2 border border-slate-200 shadow-2xs"
                />
                <p className="text-[11px] text-slate-500">
                  Aporte directo sin comisiones · Desde S/ 1.00
                </p>
              </div>
            ) : (
              <div className="bg-slate-50 border border-slate-200 rounded-2xl p-5 space-y-3">
                <p className="text-[11px] font-bold text-slate-600 uppercase tracking-wider">
                  Escanea o usa enlace directo (USD)
                </p>
                <img
                  src="/paypal-qr.png"
                  alt="QR PayPal"
                  className="w-44 h-44 mx-auto rounded-xl object-contain bg-white p-2 border border-slate-200 shadow-2xs"
                />
                <a
                  href={PAYPAL_DONATE_URL}
                  target="_blank"
                  rel="noreferrer"
                  className="w-full bg-slate-950 hover:bg-slate-800 text-white font-bold py-2.5 px-4 rounded-xl text-xs flex items-center justify-center gap-1.5 transition"
                >
                  <span>Continuar en PayPal (Tarjeta o Saldo)</span>
                  <ArrowUpRight size={14} />
                </a>
              </div>
            )}

            <div className="pt-1 border-t border-slate-100 flex items-center justify-between text-xs">
              <a
                href={DEVSTORE_URL}
                target="_blank"
                rel="noreferrer"
                className="font-semibold text-slate-600 hover:text-slate-950 inline-flex items-center gap-1 transition"
              >
                <span>Conoce más proyectos en DevStore</span>
                <ArrowUpRight size={13} />
              </a>
              <button
                onClick={() => setShowDonateModal(false)}
                className="font-bold text-slate-950 hover:underline cursor-pointer"
              >
                Cerrar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Footer Corporativo */}
      <footer className="bg-white border-t border-slate-200 py-6 px-4">
        <div className="max-w-6xl mx-auto flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-slate-500">
          <p>© 2026 ExpressWord · Procesamiento 100% local en tu navegador.</p>
          <a
            href={DEVSTORE_URL}
            target="_blank"
            rel="noreferrer"
            className="inline-flex items-center gap-1.5 font-semibold text-slate-700 hover:text-slate-950 transition"
          >
            <span>Desarrollado por DevStore</span>
            <ArrowUpRight size={13} />
          </a>
        </div>
      </footer>
    </div>
  );
}