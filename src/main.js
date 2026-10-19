/**
 * main.js - Visor GIS Notificaciones & Saneamiento Subsector 1A-4
 * Con Fondos Cartográficos (Fondo 1, Fondo 2, Fondo 3) y Ajuste Fino en Tiempo Real.
 */
import './style.css';
import './mobile.css';
import 'leaflet/dist/leaflet.css';
import L from 'leaflet';
import Papa from 'papaparse';
import { 
    generateDXF, 
    generateKML, 
    generateExcelWorkbook, 
    downloadFile, 
    latLonToUTM 
} from './ExportEngine.js';
import { 
    LeafletSpatialDrawer, 
    isFeatureInsideSelection 
} from './SpatialDrawingEngine.js';

// ====== CONSTANTES Y URLS ======
const GOOGLE_SHEET_TSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQC4s65ZG9ClF8ERvn-6zczNgcYr2pplaUeMLC41gFyULAA4FBGUv8GUGaatmpcOl0LJsecqaPiGoyT/pub?gid=1786723416&single=true&output=tsv";
const GEOJSON_URL = "./1A-4_NOTIFICAR.geojson";

// Paleta de Colores por Asociación
const ASOCIACION_COLORS = {
    'ESPIRITU SANTO': '#38bdf8',          // Cyan
    'PRIMERO DE JULIO': '#10b981',        // Emerald
    'BARRIOS ALTOS': '#f59e0b',          // Amber
    'LAS LOMAS DE CHEN CHEN': '#ec4899', // Pink
    'BICENTENARIO DE CHEN CHEN': '#8b5cf6', // Violet
    'OTRA': '#64748b'
};

// ====== ESTADO GLOBAL ======
const AppState = {
    map: null,
    baseLayers: {},
    currentBaseLayer: null,
    geojsonLayer: null,
    capaFondoGroup: null,
    labelGroup: null,
    spatialDrawer: null,

    rawGeoJSON: null,
    dbSheet: new Map(),         // Normalized Key -> Row Data
    rawSheetRows: [],
    asociacionesSet: new Set(),
    asociacionesActivas: new Set(),

    // Fondos Cartográficos
    dataFondo1: null,
    dataFondo2: null,
    dataFondo3: null,
    fondoActivo: 'fondo1',      // 'fondo1' | 'fondo2' | 'fondo3' | 'ninguno'
    fondoOpacidad: 0.50,
    fondoColor: '#00e5ff',
    isFondoInteractive: false,

    // Filtros
    modo: 'ESTADO',             // 'ESTADO' | 'ASOCIACION' | 'MASIVO'
    onlyNotificar: true,
    isSingleColor: false,
    showSinCarpeta: true,
    showNoTitulado: true,
    showOtros: false,
    batchMatches: new Set(),

    // Visual
    colorSingle: '#38bdf8',
    colorSinCarpeta: '#f59e0b',
    colorNoTitulado: '#ec4899',
    colorOtros: '#64748b',
    opacidad: 0.60,
    strokeWidth: 2,
    strokeStyle: 'solid',

    // Ajuste Fino Offset Cartográfico (metros)
    offsetE: 0,
    offsetN: 0,
    offsetQueued: false,

    // Etiquetas / Leyendas
    showLabels: true,
    labelField: 'MZ_LOTE',
    labelStyle: 'halo',         // Texto Limpio con Halo fijo

    // Capas visibles activas y límites cacheados
    visibleLayers: [],
    highlightedLayer: null,
    subsectorBounds: null,

    // Recopilación en Campo (Fotos y Notificaciones a Google Drive / Sheets)
    fieldAppsScriptUrl: localStorage.getItem('field_apps_script_url') || '',
    fieldRecordsMap: new Map(), // COD_PROMUVI -> Registro de campo
    currentFieldLotProps: null,
    capturedPhotoBase64: null,
    capturedPhotoDataUrl: null,
    currentFieldGps: null,
    gpsUserMarker: null,
    gpsUserCircle: null
};

// ====== UTILIDADES ======

/**
 * Normaliza claves de lote para asegurar 100% de coincidencias
 * Ej: "1A-4_B_08" <-> "1A-4_B_8"
 */
function normCode(code) {
    if (!code) return '';
    const clean = String(code).trim().toUpperCase();
    const parts = clean.split('_');
    if (parts.length === 3) {
        const [sec, mz, lt] = parts;
        const ltNum = parseInt(lt, 10);
        return isNaN(ltNum) ? clean : `${sec}_${mz}_${ltNum}`;
    }
    return clean;
}

/**
 * Toast Notifications
 */
function showToast(message, type = 'info', duration = 3500) {
    const container = document.getElementById('toast-container');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `toast ${type}`;

    let icon = 'ℹ️';
    if (type === 'success') icon = '✅';
    else if (type === 'warning') icon = '⚠️';
    else if (type === 'error') icon = '❌';

    toast.innerHTML = `<span>${icon}</span> <span>${message}</span>`;
    container.appendChild(toast);

    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transform = 'translateY(10px)';
        toast.style.transition = 'all 0.3s ease';
        setTimeout(() => toast.remove(), 300);
    }, duration);
}

function setLoaderText(txt) {
    const el = document.getElementById('loader-text');
    if (el) el.textContent = txt;
}

function hideLoader() {
    const el = document.getElementById('loader');
    if (el) {
        el.style.opacity = '0';
        setTimeout(() => el.classList.add('hidden'), 400);
    }
}

async function fetchJsonSafe(url) {
    try {
        const res = await fetch(url);
        if (res.ok) return await res.json();
        console.warn(`No se pudo cargar ${url} (HTTP ${res.status})`);
    } catch (e) {
        console.warn(`Error al obtener ${url}:`, e);
    }
    return null;
}

// ====== INICIALIZACIÓN DEL MAPA ======
function initMap() {
    // Coordenadas centrales del Subsector 1A-4 (Chen Chen, Moquegua)
    const initialCenter = [-17.20158, -70.92828];

    AppState.map = L.map('map', {
        center: initialCenter,
        zoom: 17,
        maxZoom: 22,
        zoomControl: false
    });

    // Control de Zoom en posición top-right
    L.control.zoom({ position: 'topright' }).addTo(AppState.map);

    // Crear Panes separados para orden de capas
    // paneCapaFondo a z-index 350 para estar debajo de los polígonos principales (z-index 400)
    AppState.map.createPane('paneCapaFondo');
    AppState.map.getPane('paneCapaFondo').style.zIndex = 350;

    // Definición de Mapas Base optimizados para fluidez móvil
    const tileOptions = {
        keepBuffer: 3,
        updateWhenIdle: true
    };

    AppState.baseLayers = {
        'google-hybrid': L.tileLayer('https://mt1.google.com/vt/lyrs=y&x={x}&y={y}&z={z}', {
            maxZoom: 22,
            maxNativeZoom: 20,
            attribution: '© Google Satellite',
            ...tileOptions
        }),
        'google-sat': L.tileLayer('https://mt1.google.com/vt/lyrs=s&x={x}&y={y}&z={z}', {
            maxZoom: 22,
            maxNativeZoom: 20,
            attribution: '© Google Satellite',
            ...tileOptions
        }),
        'esri-sat': L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
            maxZoom: 22,
            maxNativeZoom: 19,
            attribution: '© Esri World Imagery',
            ...tileOptions
        }),
        'esri-dark': L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Dark_Gray_Base/MapServer/tile/{z}/{y}/{x}', {
            maxZoom: 19,
            attribution: '© Esri Dark Gray',
            ...tileOptions
        }),
        'osm': L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            maxZoom: 19,
            attribution: '© OpenStreetMap',
            ...tileOptions
        }),
        'carto-dark': L.tileLayer('https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png', {
            maxZoom: 20,
            attribution: '© CartoDB Dark Matter',
            ...tileOptions
        })
    };

    AppState.currentBaseLayer = AppState.baseLayers['google-hybrid'];
    AppState.currentBaseLayer.addTo(AppState.map);

    // Grupos de capas
    AppState.capaFondoGroup = L.layerGroup([], { pane: 'paneCapaFondo' }).addTo(AppState.map);
    AppState.labelGroup = L.layerGroup().addTo(AppState.map);

    // Inicializar Motor de Selección Espacial solo si está activo en la interfaz
    if (document.getElementById('btnDrawStart')) {
        AppState.spatialDrawer = new LeafletSpatialDrawer(AppState.map, {
            onPointAdded: (count) => {
                const st = document.getElementById('draw-status');
                if (st) {
                    st.style.display = 'block';
                    st.textContent = `Vértice ${count} añadido. Doble clic para cerrar área.`;
                }
            },
            onComplete: (points) => {
                document.getElementById('btnDrawStart')?.classList.remove('hidden');
                document.getElementById('btnDrawFilter')?.classList.remove('hidden');
                document.getElementById('btnDrawClear')?.classList.remove('hidden');
                const st = document.getElementById('draw-status');
                if (st) st.textContent = `Área cerrada con ${points.length} vértices. Clic en FILTRAR.`;
            },
            onClear: () => {
                document.getElementById('btnDrawFilter')?.classList.add('hidden');
                document.getElementById('btnDrawClear')?.classList.add('hidden');
                const st = document.getElementById('draw-status');
                if (st) st.style.display = 'none';
            }
        });
    }

    // Eventos de movimiento en el mapa para refrescar etiquetas
    AppState.map.on('moveend', () => {
        updateLabels();
    });
}

// ====== CARGA Y SINCRONIZACIÓN DE DATOS ======

async function fetchGoogleSheetTSV() {
    setLoaderText('Descargando padrón de notificaciones desde Google Sheets...');
    const response = await fetch(GOOGLE_SHEET_TSV_URL);
    if (!response.ok) throw new Error(`Error HTTP al descargar Google Sheet: ${response.status}`);
    
    const tsvText = await response.text();
    const parsed = Papa.parse(tsvText, {
        header: true,
        delimiter: '\t',
        skipEmptyLines: true
    });

    AppState.dbSheet.clear();
    AppState.rawSheetRows = [];
    AppState.asociacionesSet.clear();

    parsed.data.forEach(row => {
        let cod = row['COD_PROMUVI'] || row['﻿COD_PROMUVI'] || row['CODIGO'] || '';
        cod = cod.trim();
        if (!cod) return;

        const estado = (row['ESTADO'] || 'SIN ESTADO').trim().toUpperCase();
        const asoc = (row['ASOCIACIÓN'] || row['ASOCIACION'] || 'SIN ASOCIACION').trim().toUpperCase();

        const sheetItem = {
            COD_PROMUVI: cod,
            ASOCIACION: asoc,
            MANZANA: (row['MANZANA'] || '').trim(),
            LOTE: (row['LOTE'] || '').trim(),
            DISTRITO: (row['DISTRITO'] || 'Moquegua').trim(),
            ESTADO: estado,
            NRO_CARPETA: (row['Nº CARPETA'] || row['N° CARPETA'] || row['CARPETA'] || '').trim(),
            DNI_1: (row['DNI 1'] || row['DNI1'] || '').trim(),
            TITULAR_1: (row['TITULAR 1'] || row['TITULAR1'] || '').trim(),
            DNI_2: (row['DNI 2'] || row['DNI2'] || '').trim(),
            TITULAR_2: (row['TITULAR 2'] || row['TITULAR2'] || '').trim()
        };

        AppState.rawSheetRows.push(sheetItem);

        AppState.dbSheet.set(cod.toUpperCase(), sheetItem);
        AppState.dbSheet.set(normCode(cod), sheetItem);

        if (asoc) AppState.asociacionesSet.add(asoc);
    });

    AppState.asociacionesActivas = new Set(AppState.asociacionesSet);

    const badgeTxt = document.getElementById('sync-status-text');
    if (badgeTxt) {
        badgeTxt.textContent = `Google Sheets Activo (${AppState.rawSheetRows.length} en Padrón)`;
    }

    renderAsociacionesFilterList();
    renderDashboard();
}

function alignGeoJsonWithFondo(geojsonData) {
    if (!geojsonData || !geojsonData.features) return;
    const centerLat = -17.20158;
    const latMetersPerDegree = 111320;
    const lngMetersPerDegree = 111320 * Math.cos(centerLat * Math.PI / 180);
    // Vector geodésico exacto para alinear 1A-4 con la Base Gráfica / Fondo (+8.503m Este, -2.008m Norte)
    const dLng = 8.503 / lngMetersPerDegree;
    const dLat = -2.008 / latMetersPerDegree;

    const shiftCoords = (coords) => {
        if (!coords) return;
        if (typeof coords[0] === 'number') {
            coords[0] += dLng;
            coords[1] += dLat;
        } else {
            coords.forEach(shiftCoords);
        }
    };

    geojsonData.features.forEach(f => {
        if (f.geometry && f.geometry.coordinates) {
            shiftCoords(f.geometry.coordinates);
        }
    });
}

async function fetchGeoJSON() {
    setLoaderText('Cargando polígonos cartográficos de Subsector 1A-4...');
    AppState.rawGeoJSON = await fetchJsonSafe(GEOJSON_URL);
    if (AppState.rawGeoJSON) {
        alignGeoJsonWithFondo(AppState.rawGeoJSON);
    }
}

async function loadAllData() {
    try {
        await Promise.all([
            fetchGoogleSheetTSV(),
            fetchGeoJSON()
        ]);

        buildVectorLayer();
        aplicarFiltros();
        hideLoader();
        showToast('Cartografía y Padrón cargados con éxito.', 'success');

        // Cargar Fondo 1 en segundo plano
        loadFondoActivo();
    } catch (err) {
        console.error('Error al iniciar la carga:', err);
        showToast(`Error de carga: ${err.message}`, 'error', 7000);
        hideLoader();
    }
}

// ====== MOTOR DE FONDOS CARTOGRÁFICOS ======

async function loadFondoActivo() {
    if (AppState.fondoActivo === 'ninguno') {
        AppState.capaFondoGroup.clearLayers();
        return;
    }

    let data = null;
    if (AppState.fondoActivo === 'fondo1') {
        if (!AppState.dataFondo1) {
            showToast('Cargando Fondo 1...', 'info');
            AppState.dataFondo1 = await fetchJsonSafe('./Fondo.json');
        }
        data = AppState.dataFondo1;
    } else if (AppState.fondoActivo === 'fondo2') {
        if (!AppState.dataFondo2) {
            showToast('Descargando Fondo 2...', 'info');
            AppState.dataFondo2 = await fetchJsonSafe('./Fondo2.json');
        }
        data = AppState.dataFondo2;
    } else if (AppState.fondoActivo === 'fondo3') {
        if (!AppState.dataFondo3) {
            showToast('Descargando Fondo 3...', 'info');
            AppState.dataFondo3 = await fetchJsonSafe('./Fondo3.json');
        }
        data = AppState.dataFondo3;
    }

    if (!data) return;

    AppState.capaFondoGroup.clearLayers();

    const fondoLayer = L.geoJSON(data, {
        pane: 'paneCapaFondo',
        style: () => ({
            color: AppState.fondoColor,
            weight: 1.2,
            fillColor: AppState.fondoColor,
            fillOpacity: AppState.isFondoInteractive ? (AppState.fondoOpacidad * 0.15) : 0,
            opacity: AppState.fondoOpacidad,
            interactive: AppState.isFondoInteractive
        }),
        onEachFeature: (feature, layer) => {
            if (AppState.isFondoInteractive) {
                layer.on('click', (e) => {
                    if (AppState.spatialDrawer && AppState.spatialDrawer.isDrawing) return;
                    showGeoJsonModal(feature.properties, `Información Cartográfica • ${AppState.fondoActivo.toUpperCase()}`);
                    if (e.originalEvent) L.DomEvent.stop(e.originalEvent);
                });
            }
        }
    });

    fondoLayer.addTo(AppState.capaFondoGroup);
    applyOffsetInRealTime(); // Aplicar el desplazamiento actual si existe
}

// ====== MODAL DE ATRIBUTOS PARA FONDO ======
function showGeoJsonModal(properties, title) {
    const titleEl = document.getElementById('modal-title');
    if (titleEl) titleEl.innerText = title;
    
    const b = document.getElementById('modal-body-content');
    
    let rows = '';
    for (let k in properties) {
        let val = properties[k];
        if (typeof val === 'number') {
            if (k.toLowerCase().includes('area')) {
                val = val.toLocaleString('es-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' m²';
            } else if (k.toLowerCase().includes('length') || k.toLowerCase().includes('perimetro') || k.toLowerCase().includes('longitud')) {
                val = val.toLocaleString('es-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' m';
            }
        } else if (k.toLowerCase().includes('fec') && typeof val === 'number' && val > 100000000000) {
            try {
                val = new Date(val).toLocaleDateString('es-ES');
            } catch(e) {}
        }
        
        rows += `
            <div class="info-item" style="display: flex; justify-content: space-between; align-items: center; padding: 7px 12px; background: rgba(0,0,0,0.25);">
                <div class="info-label" style="margin: 0; font-size: 10px; font-weight: 700; color: var(--accent);">${k}</div>
                <div class="info-value" style="font-size: 12px; font-weight: 600; color: #fff;">${val !== null && val !== undefined && val !== '' ? val : '--'}</div>
            </div>
        `;
    }
    
    b.innerHTML = `
        <div style="display: flex; flex-direction: column; gap: 6px; padding: 8px 0;">
            ${rows || '<p style="padding: 10px; color: var(--text-muted);">Sin propiedades disponibles.</p>'}
        </div>
    `;
    
    document.getElementById('property-modal')?.classList.remove('hidden');
}

// ====== AJUSTE FINO DE POLÍGONOS (REAL-TIME OFFSET ENGINE) ======

function calculateOffsetDegrees() {
    const centerLat = -17.20158;
    const latMetersPerDegree = 111320;
    const lngMetersPerDegree = 111320 * Math.cos(centerLat * Math.PI / 180);
    return {
        dLat: AppState.offsetN / latMetersPerDegree,
        dLng: AppState.offsetE / lngMetersPerDegree
    };
}

/**
 * Desplaza directamente las geometrías de capas (GeoJSON y Fondo) y actualiza sus límites en tiempo real.
 */
function applyOffsetInRealTime() {
    const { dLat, dLng } = calculateOffsetDegrees();

    const shiftLayer = (layer) => {
        if (!layer) return;

        // Si es un grupo contenedor (ej. GeoJSON o LayerGroup o MultiPolygon con subcapas)
        if (layer.eachLayer) {
            layer.eachLayer(shiftLayer);
            return;
        }

        // 1. Mutar las coordenadas de los vértices si tiene getLatLngs
        if (layer.getLatLngs) {
            if (!layer._origLatLngs) {
                const cloneCoords = (arr) => {
                    if (!arr) return arr;
                    if (arr.lat !== undefined) return { lat: arr.lat, lng: arr.lng };
                    return arr.map(cloneCoords);
                };
                layer._origLatLngs = cloneCoords(layer.getLatLngs());
            }

            const mutateCoords = (orig, curr) => {
                if (!orig || !curr) return;
                if (orig.lat !== undefined && curr.lat !== undefined) {
                    curr.lat = orig.lat + dLat;
                    curr.lng = orig.lng + dLng;
                } else if (Array.isArray(orig) && Array.isArray(curr)) {
                    for (let i = 0; i < orig.length; i++) {
                        mutateCoords(orig[i], curr[i]);
                    }
                }
            };
            mutateCoords(layer._origLatLngs, layer.getLatLngs());
        }

        // 2. Asegurar que layer._bounds exista antes de mutar
        if (!layer._bounds && layer.getBounds) {
            layer.getBounds();
        }

        // 3. Mutar los bounds para que getBounds().getCenter() y intersects() estén 100% sincronizados
        if (layer._bounds) {
            if (!layer._origBounds) {
                layer._origBounds = {
                    southWest: { lat: layer._bounds._southWest.lat, lng: layer._bounds._southWest.lng },
                    northEast: { lat: layer._bounds._northEast.lat, lng: layer._bounds._northEast.lng }
                };
            }
            layer._bounds._southWest.lat = layer._origBounds.southWest.lat + dLat;
            layer._bounds._southWest.lng = layer._origBounds.southWest.lng + dLng;
            layer._bounds._northEast.lat = layer._origBounds.northEast.lat + dLat;
            layer._bounds._northEast.lng = layer._origBounds.northEast.lng + dLng;

            layer._cachedBounds = layer.getBounds ? layer.getBounds() : layer._bounds;
            if (layer._cachedBounds) {
                layer._cachedCenter = layer._cachedBounds.getCenter();
            }
        }

        if (layer.redraw) {
            layer.redraw();
        }
    };

    if (AppState.geojsonLayer) {
        AppState.geojsonLayer.eachLayer(shiftLayer);
        AppState.subsectorBounds = AppState.geojsonLayer.getBounds();
    }

    if (AppState.capaFondoGroup) {
        AppState.capaFondoGroup.eachLayer(shiftLayer);
    }

    // Refrescar etiquetas para posicionarlas en el centro desplazado exacto
    updateLabels();
}

function handleOffsetInput(eastMeters, northMeters) {
    AppState.offsetE = parseFloat(eastMeters) || 0;
    AppState.offsetN = parseFloat(northMeters) || 0;

    const formattedE = (AppState.offsetE >= 0 ? '+' : '') + AppState.offsetE.toFixed(1) + "m";
    const formattedN = (AppState.offsetN >= 0 ? '+' : '') + AppState.offsetN.toFixed(1) + "m";

    const elE = document.getElementById('f-val-east');
    const elN = document.getElementById('f-val-north');
    if (elE) elE.textContent = formattedE;
    if (elN) elN.textContent = formattedN;

    const sE = document.getElementById('f-offset-east');
    const sN = document.getElementById('f-offset-north');
    if (sE && parseFloat(sE.value) !== AppState.offsetE) sE.value = AppState.offsetE;
    if (sN && parseFloat(sN.value) !== AppState.offsetN) sN.value = AppState.offsetN;

    if (!AppState.offsetQueued) {
        AppState.offsetQueued = true;
        requestAnimationFrame(() => {
            AppState.offsetQueued = false;
            applyOffsetInRealTime();
        });
    }
}

// ====== MOTOR DE CAPAS Y ESTILOS DE PREDIOS ======

function getFeatureStyle(feature) {
    const cod = feature.properties.COD_PROMUVI || '';
    const sheetData = AppState.dbSheet.get(cod.toUpperCase()) || AppState.dbSheet.get(normCode(cod));

    let fillColor = AppState.colorOtros;
    let fillOpacity = AppState.opacidad * 0.35;
    let strokeColor = '#ffffff';

    if (AppState.isSingleColor) {
        fillColor = AppState.colorSingle;
        fillOpacity = AppState.opacidad * 0.7;
    } else if (AppState.modo === 'ASOCIACION') {
        const asoc = sheetData ? sheetData.ASOCIACION : 'OTRA';
        fillColor = ASOCIACION_COLORS[asoc] || '#64748b';
        fillOpacity = sheetData ? (AppState.opacidad * 0.75) : (AppState.opacidad * 0.25);
    } else {
        // Modo 'ESTADO'
        if (sheetData) {
            const st = sheetData.ESTADO;
            if (st.includes('SIN CARPETA')) {
                fillColor = AppState.colorSinCarpeta;
                fillOpacity = AppState.opacidad * 0.75;
            } else if (st.includes('NO TITULADO')) {
                fillColor = AppState.colorNoTitulado;
                fillOpacity = AppState.opacidad * 0.75;
            } else {
                fillColor = '#3b82f6';
                fillOpacity = AppState.opacidad * 0.6;
            }
        } else {
            fillColor = AppState.colorOtros;
            fillOpacity = AppState.opacidad * 0.22;
        }
    }

    let dashArray = null;
    if (AppState.strokeStyle === 'dashed') dashArray = '6, 5';
    else if (AppState.strokeStyle === 'dotted') dashArray = '2, 3';

    return {
        fillColor: fillColor,
        fillOpacity: fillOpacity,
        color: strokeColor,
        weight: AppState.strokeWidth,
        dashArray: dashArray,
        opacity: 0.95
    };
}

function buildVectorLayer() {
    if (AppState.geojsonLayer) {
        AppState.map.removeLayer(AppState.geojsonLayer);
    }

    AppState.geojsonLayer = L.geoJSON(AppState.rawGeoJSON, {
        style: (feature) => getFeatureStyle(feature),
        onEachFeature: (feature, layer) => {
            // Pre-cachear bounds y center para acelerar drásticamente updateLabels y recenter en móviles
            if (layer.getBounds) {
                layer._cachedBounds = layer.getBounds();
                layer._cachedCenter = layer._cachedBounds.getCenter();
            }

            const cod = feature.properties.COD_PROMUVI || '';
            const sheetData = AppState.dbSheet.get(cod.toUpperCase()) || AppState.dbSheet.get(normCode(cod));

            if (sheetData) {
                feature.properties.EN_PADRON = true;
                feature.properties.ASOCIACIÓN = sheetData.ASOCIACION;
                feature.properties.ESTADO = sheetData.ESTADO;
                feature.properties.NRO_CARPETA = sheetData.NRO_CARPETA;
                feature.properties.DNI_1 = sheetData.DNI_1;
                feature.properties.TITULAR_1 = sheetData.TITULAR_1;
                feature.properties.DNI_2 = sheetData.DNI_2;
                feature.properties.TITULAR_2 = sheetData.TITULAR_2;
            } else {
                feature.properties.EN_PADRON = false;
                feature.properties.ESTADO = 'NO PRIORIZADO';
            }

            // Tooltip interactivo
            const tooltipContent = `
                <div style="font-size: 11px; line-height: 1.3;">
                    <strong style="color: #38bdf8;">${cod}</strong><br/>
                    ${sheetData ? `<span style="color: #fef08a;">${sheetData.ESTADO}</span><br/><span style="color:#e2e8f0;">${sheetData.TITULAR_1 || sheetData.ASOCIACION}</span>` : '<span style="color:#94a3b8;">Predio 1A-4</span>'}
                </div>
            `;
            layer.bindTooltip(tooltipContent, { sticky: true, opacity: 0.95 });

            layer.on('click', () => {
                highlightLayer(layer);
                openPropertyModal(feature.properties);
            });

            layer.on('dblclick', (e) => {
                L.DomEvent.stop(e);
                const b = layer._cachedBounds || layer.getBounds();
                AppState.map.fitBounds(b, { padding: [50, 50], maxZoom: 20 });
            });
        }
    }).addTo(AppState.map);

    AppState.subsectorBounds = AppState.geojsonLayer.getBounds();
    AppState.map.fitBounds(AppState.subsectorBounds, { 
        padding: isMobile() ? [15, 15] : [30, 30],
        maxZoom: 18,
        animate: false 
    });

    if (AppState.offsetE !== 0 || AppState.offsetN !== 0) {
        applyOffsetInRealTime();
    }
}

function highlightLayer(layer) {
    if (AppState.highlightedLayer && AppState.highlightedLayer !== layer) {
        AppState.geojsonLayer.resetStyle(AppState.highlightedLayer);
    }
    AppState.highlightedLayer = layer;
    layer.setStyle({
        color: '#38bdf8',
        weight: 3.5,
        dashArray: null,
        fillOpacity: Math.min(AppState.opacidad + 0.25, 1)
    });
    if (!L.Browser.ie && !L.Browser.opera && !L.Browser.edge) {
        layer.bringToFront();
    }
}

// ====== FILTRADO Y ACTUALIZACIÓN ======

function aplicarFiltros() {
    if (!AppState.geojsonLayer) return;

    AppState.visibleLayers = [];
    let countOnlyNotificar = 0;
    let countSinCarpeta = 0;
    let countNoTitulado = 0;
    let countOtros = 0;

    AppState.geojsonLayer.eachLayer((layer) => {
        const props = layer.feature.properties;
        const cod = props.COD_PROMUVI || '';
        const sheetData = AppState.dbSheet.get(cod.toUpperCase()) || AppState.dbSheet.get(normCode(cod));

        let visible = true;

        if (AppState.onlyNotificar && !sheetData) {
            visible = false;
        }

        if (visible) {
            if (AppState.modo === 'MASIVO') {
                if (AppState.batchMatches.size > 0) {
                    visible = AppState.batchMatches.has(cod.toUpperCase()) || AppState.batchMatches.has(normCode(cod));
                }
            } else if (AppState.modo === 'ASOCIACION') {
                if (sheetData) {
                    visible = AppState.asociacionesActivas.has(sheetData.ASOCIACION);
                } else {
                    visible = !AppState.onlyNotificar && AppState.showOtros;
                }
            } else {
                if (sheetData) {
                    const st = sheetData.ESTADO;
                    if (st.includes('SIN CARPETA') && !AppState.showSinCarpeta) visible = false;
                    else if (st.includes('NO TITULADO') && !AppState.showNoTitulado) visible = false;
                } else {
                    if (!AppState.showOtros) visible = false;
                }
            }
        }

        if (visible) {
            if (!AppState.map.hasLayer(layer)) {
                AppState.map.addLayer(layer);
            }
            layer.setStyle(getFeatureStyle(layer.feature));
            AppState.visibleLayers.push(layer);

            if (sheetData) {
                countOnlyNotificar++;
                if (sheetData.ESTADO.includes('SIN CARPETA')) countSinCarpeta++;
                else if (sheetData.ESTADO.includes('NO TITULADO')) countNoTitulado++;
            } else {
                countOtros++;
            }
        } else {
            if (AppState.map.hasLayer(layer)) {
                AppState.map.removeLayer(layer);
            }
        }
    });

    updateCounterBadges(countOnlyNotificar, countSinCarpeta, countNoTitulado, countOtros);
    updateLabels();
}

function updateCounterBadges(notif, sinCarp, noTit, otros) {
    const elN = document.getElementById('count-only-notificar');
    const elS = document.getElementById('count-sin-carpeta');
    const elT = document.getElementById('count-no-titulado');
    const elO = document.getElementById('count-otros');
    const elTotal = document.getElementById('total-notificar-badge');

    if (elN) elN.textContent = notif;
    if (elS) elS.textContent = sinCarp;
    if (elT) elT.textContent = noTit;
    if (elO) elO.textContent = otros;
    if (elTotal) elTotal.textContent = `${notif} Visibles`;
}

// ====== ETIQUETAS DINÁMICAS (EN SINCRONIZACIÓN PERFECTA CON POLÍGONOS) ======

let labelRafId = null;

function updateLabels() {
    if (labelRafId) cancelAnimationFrame(labelRafId);
    labelRafId = requestAnimationFrame(doUpdateLabels);
}

function doUpdateLabels() {
    if (!AppState.labelGroup) return;
    AppState.labelGroup.clearLayers();

    if (!AppState.showLabels) return;

    const zoom = AppState.map.getZoom();
    const minZoom = isMobile() ? 17 : 16;
    if (zoom < minZoom) return;

    const mapBounds = AppState.map.getBounds().pad(0.1);

    AppState.visibleLayers.forEach(layer => {
        const lbounds = layer._cachedBounds || (layer.getBounds ? layer.getBounds() : null);
        if (!lbounds || !mapBounds.intersects(lbounds)) return;

        const center = layer._cachedCenter || (lbounds.getCenter ? lbounds.getCenter() : null);
        if (!center || isNaN(center.lat) || isNaN(center.lng)) return;

        const props = layer.feature ? layer.feature.properties : {};
        const cod = props.COD_PROMUVI || '';
        const sheetData = AppState.dbSheet.get(cod.toUpperCase()) || AppState.dbSheet.get(normCode(cod));

        // Siempre por manzana y lote (Ej: Mz B Lt 8)
        let mz = props.MANZANA || '';
        let lt = props.LOTE || '';
        if (!mz || !lt) {
            const parts = (cod || '').split('_');
            if (parts.length >= 3) {
                mz = mz || parts[1];
                lt = lt || parts[2];
            }
        }
        const text = (mz && lt) ? `Mz ${mz} Lt ${lt}` : (cod || 'PREDIO');

        // Estilo fijo en Texto Limpio con Halo (Sin Fondo)
        let className = 'lot-label-marker';

        if (sheetData) {
            if (sheetData.ESTADO.includes('SIN CARPETA')) className += ' sin-carpeta';
            else if (sheetData.ESTADO.includes('NO TITULADO')) className += ' no-titulado';
        } else {
            className += ' otros';
        }

        const icon = L.divIcon({
            className: 'cus-wrapper',
            html: `<div class="${className}">${text}</div>`,
            iconSize: [0, 0]
        });

        L.marker(center, { icon: icon, interactive: false }).addTo(AppState.labelGroup);
    });
}

// ====== MODAL DETALLES DEL PREDIO ======

function openPropertyModal(props) {
    const cod = props.COD_PROMUVI || 'PREDIO';
    const sheetData = AppState.dbSheet.get(cod.toUpperCase()) || AppState.dbSheet.get(normCode(cod));
    const fieldRecord = AppState.fieldRecordsMap.get(cod.toUpperCase()) || AppState.fieldRecordsMap.get(normCode(cod));

    document.getElementById('modal-title').textContent = `Ficha de Predio • ${cod}`;

    let statusBadgeColor = 'background: rgba(100, 116, 139, 0.2); color: #cbd5e1; border: 1px solid #64748b;';
    let statusText = 'NO PRIORIZADO / FUERA DE PADRÓN';

    if (sheetData) {
        if (sheetData.ESTADO.includes('SIN CARPETA')) {
            statusBadgeColor = 'background: rgba(245, 158, 11, 0.2); color: #f59e0b; border: 1px solid #f59e0b;';
            statusText = 'SIN CARPETA (NOTIFICAR)';
        } else if (sheetData.ESTADO.includes('NO TITULADO')) {
            statusBadgeColor = 'background: rgba(236, 72, 153, 0.2); color: #ec4899; border: 1px solid #ec4899;';
            statusText = 'NO TITULADO (NOTIFICAR)';
        }
    }

    const center = props.CENTRO_LAT && props.CENTRO_LNG ? 
        { lat: props.CENTRO_LAT, lng: props.CENTRO_LNG } : 
        (AppState.highlightedLayer ? AppState.highlightedLayer.getBounds().getCenter() : { lat: -17.20158, lng: -70.92828 });

    const googleEarthWebUrl = `https://earth.google.com/web/search/${center.lat},${center.lng}`;

    const html = `
        <div class="property-details-grid">
            <div class="prop-header-card">
                <div>
                    <span style="font-size: 10px; color: var(--text-muted); text-transform: uppercase;">Código Predial Promuvi</span>
                    <div style="font-size: 18px; font-weight: 800; color: #38bdf8;">${cod}</div>
                    <div style="font-size: 12px; color: #e2e8f0;">Manzana ${props.MANZANA || '--'} • Lote ${props.LOTE || '--'} • Subsector 1A-4</div>
                </div>
                <div class="prop-badge-status" style="${statusBadgeColor}">${statusText}</div>
            </div>

            ${sheetData ? `
            <div class="prop-section-card" style="border-color: rgba(56, 189, 248, 0.3);">
                <h4>📋 Información de Notificación & Expediente</h4>
                <div class="info-grid">
                    <div class="info-item">
                        <div class="info-label">Asociación de Vivienda</div>
                        <div class="info-value" style="color: #fcd34d;">${sheetData.ASOCIACION || '--'}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">N° de Carpeta / Expediente</div>
                        <div class="info-value" style="color: #38bdf8;">${sheetData.NRO_CARPETA || 'SIN CARPETA ASIGNADA'}</div>
                    </div>
                    <div class="info-item" style="grid-column: 1 / -1;">
                        <div class="info-label">Titular Principal (1)</div>
                        <div class="info-value">${sheetData.TITULAR_1 || '--'} ${sheetData.DNI_1 ? `• <strong style="color: #34d399;">DNI: ${sheetData.DNI_1}</strong>` : ''}</div>
                    </div>
                    ${sheetData.TITULAR_2 ? `
                    <div class="info-item" style="grid-column: 1 / -1;">
                        <div class="info-label">Cónyuge / Segundo Titular (2)</div>
                        <div class="info-value">${sheetData.TITULAR_2} ${sheetData.DNI_2 ? `• <strong style="color: #34d399;">DNI: ${sheetData.DNI_2}</strong>` : ''}</div>
                    </div>
                    ` : ''}
                </div>
            </div>
            ` : `
            <div class="prop-section-card" style="border: 1px dashed rgba(255,255,255,0.15); text-align: center; padding: 16px;">
                <span style="color: var(--text-muted); font-size: 12px;">ℹ️ Este lote pertenece al Subsector 1A-4 pero <strong>no se encuentra en el padrón priorizado</strong> de las 77 notificaciones.</span>
            </div>
            `}

            <div class="prop-section-card">
                <h4>🏛️ Información Registral (SUNARP) & Física</h4>
                <div class="info-grid">
                    <div class="info-item">
                        <div class="info-label">Partida Registral</div>
                        <div class="info-value">${props.PART_REG || '--'}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">Partida Antecedente</div>
                        <div class="info-value">${props.PART_ANT || '--'}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">Área Gráfica / Registral</div>
                        <div class="info-value">${props.GEO_AREA ? Number(props.GEO_AREA).toFixed(2) + ' m²' : '--'}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">Perímetro</div>
                        <div class="info-value">${props.GEO_LENGTH ? Number(props.GEO_LENGTH).toFixed(2) + ' m' : '--'}</div>
                    </div>
                    <div class="info-item" style="grid-column: 1 / -1;">
                        <div class="info-label">Acto Registral</div>
                        <div class="info-value">${props.ACT_REG || '--'}</div>
                    </div>
                    <div class="info-item" style="grid-column: 1 / -1;">
                        <div class="info-label">Observaciones SUNARP</div>
                        <div class="info-value" style="font-size: 12px; color: #cbd5e1;">${props.OBSERV || '--'}</div>
                    </div>
                    <div class="info-item" style="grid-column: 1 / -1;">
                        <div class="info-label">Dirección Registral</div>
                        <div class="info-value" style="font-size: 12px;">${props.DIR_REG || '--'}</div>
                    </div>
                </div>
            </div>

            ${fieldRecord ? `
            <div class="prop-section-card" style="border: 1.5px solid rgba(16, 185, 129, 0.4); background: rgba(16, 185, 129, 0.08);">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <h4 style="color: #34d399; margin: 0;">📋 Notificación Registrada en Campo</h4>
                    <span class="field-lot-tag" style="background: #10b981; font-size: 11px;">${fieldRecord.estado_notificacion}</span>
                </div>
                <div style="font-size: 11px; color: #cbd5e1; margin-top: 8px; display: flex; flex-direction: column; gap: 4px;">
                    <div>📅 Fecha: <strong style="color: #fff;">${fieldRecord.fecha_notificacion}</strong> ${fieldRecord.notificador ? `• Notificador: <strong style="color: #38bdf8;">${fieldRecord.notificador}</strong>` : ''}</div>
                    ${fieldRecord.observaciones ? `<div>📝 Observaciones: <em>"${fieldRecord.observaciones}"</em></div>` : ''}
                    ${fieldRecord.gps_lat ? `<div>📍 GPS: ${Number(fieldRecord.gps_lat).toFixed(5)}, ${Number(fieldRecord.gps_lng).toFixed(5)} (±${Number(fieldRecord.gps_precision || 0).toFixed(1)}m)</div>` : ''}
                    ${fieldRecord.foto_url ? `
                    <div style="margin-top: 6px;">
                        <a href="${fieldRecord.foto_url}" target="_blank" style="display: inline-flex; align-items: center; gap: 4px; color: #38bdf8; font-weight: 700; text-decoration: underline;">
                            🖼️ Ver Foto de Evidencia en Google Drive ↗
                        </a>
                    </div>
                    ` : ''}
                </div>
            </div>
            ` : ''}

            <div style="display: flex; gap: 10px; margin-top: 4px; flex-wrap: wrap;">
                <button type="button" id="btn-open-field-capture" class="btn-primary" style="flex: 1.2; min-width: 180px; background: linear-gradient(135deg, #10b981 0%, #059669 100%); font-weight: 800; box-shadow: 0 4px 15px rgba(16, 185, 129, 0.45); display: flex; align-items: center; justify-content: center; gap: 6px; border: none; cursor: pointer;">
                    📸 ${fieldRecord ? 'Actualizar Notificación en Campo' : 'Registrar Notificación en Campo'}
                </button>
                <a href="${googleEarthWebUrl}" target="_blank" class="btn-secondary" style="flex: 0.8; min-width: 130px; text-decoration: none; display: flex; align-items: center; justify-content: center; gap: 6px;">
                    🌍 Google Earth
                </a>
            </div>
        </div>
    `;

    document.getElementById('modal-body-content').innerHTML = html;
    document.getElementById('property-modal').classList.remove('hidden');

    document.getElementById('btn-open-field-capture')?.addEventListener('click', () => {
        document.getElementById('property-modal').classList.add('hidden');
        openFieldCaptureModal(props);
    });
}

// ====== MOTOR DE RECOPILACIÓN EN CAMPO (FOTOS A DRIVE & SHEETS) ======

function loadLocalFieldRecords() {
    try {
        const saved = localStorage.getItem('FIELD_NOTIFICACIONES_RECORDS');
        if (saved) {
            const arr = JSON.parse(saved);
            arr.forEach(item => {
                if (item.cod_promuvi) {
                    AppState.fieldRecordsMap.set(item.cod_promuvi.toUpperCase(), item);
                    AppState.fieldRecordsMap.set(normCode(item.cod_promuvi), item);
                }
            });
        }
    } catch(e) {
        console.warn('Error al cargar registros locales de campo:', e);
    }
}

function saveLocalFieldRecord(record) {
    try {
        AppState.fieldRecordsMap.set(record.cod_promuvi.toUpperCase(), record);
        AppState.fieldRecordsMap.set(normCode(record.cod_promuvi), record);
        const existing = JSON.parse(localStorage.getItem('FIELD_NOTIFICACIONES_RECORDS') || '[]');
        const idx = existing.findIndex(e => normCode(e.cod_promuvi) === normCode(record.cod_promuvi));
        if (idx !== -1) existing[idx] = record;
        else existing.push(record);
        localStorage.setItem('FIELD_NOTIFICACIONES_RECORDS', JSON.stringify(existing));
    } catch(e) {
        console.warn('Error al guardar registro local de campo:', e);
    }
}

/**
 * Comprime fotografías de alta resolución en el navegador para subida ultra rápida
 */
function compressImage(file, maxWidth = 1280, maxHeight = 1280, quality = 0.75) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const img = new Image();
            img.onload = () => {
                let width = img.width;
                let height = img.height;
                if (width > height) {
                    if (width > maxWidth) {
                        height = Math.round((height * maxWidth) / width);
                        width = maxWidth;
                    }
                } else {
                    if (height > maxHeight) {
                        width = Math.round((width * maxHeight) / height);
                        height = maxHeight;
                    }
                }
                const canvas = document.createElement('canvas');
                canvas.width = width;
                canvas.height = height;
                const ctx = canvas.getContext('2d');
                ctx.drawImage(img, 0, 0, width, height);
                const dataUrl = canvas.toDataURL('image/jpeg', quality);
                const base64 = dataUrl.split(',')[1];
                resolve({ dataUrl, base64, width, height, sizeBytes: Math.round(base64.length * 0.75) });
            };
            img.onerror = reject;
            img.src = e.target.result;
        };
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

function getDistanceMeters(lat1, lon1, lat2, lon2) {
    const R = 6371e3;
    const phi1 = lat1 * Math.PI / 180;
    const phi2 = lat2 * Math.PI / 180;
    const dPhi = (lat2 - lat1) * Math.PI / 180;
    const dLam = (lon2 - lon1) * Math.PI / 180;
    const a = Math.sin(dPhi / 2) * Math.sin(dPhi / 2) +
              Math.cos(phi1) * Math.cos(phi2) *
              Math.sin(dLam / 2) * Math.sin(dLam / 2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    return R * c;
}

function requestFieldGps(targetLat, targetLng) {
    const gpsStatusEl = document.getElementById('field-gps-status');
    const gpsBox = document.getElementById('field-gps-box');
    if (!gpsStatusEl || !gpsBox) return;

    if (!navigator.geolocation) {
        gpsStatusEl.textContent = '⚠️ Geolocalización GPS no soportada en este navegador.';
        return;
    }

    gpsStatusEl.textContent = '📍 Obteniendo coordenadas satelitales del GPS...';
    gpsBox.className = 'field-gps-display';

    navigator.geolocation.getCurrentPosition(
        (pos) => {
            const lat = pos.coords.latitude;
            const lng = pos.coords.longitude;
            const acc = pos.coords.accuracy || 0;
            AppState.currentFieldGps = { lat, lng, accuracy: acc };

            if (targetLat && targetLng) {
                const dist = getDistanceMeters(lat, lng, targetLat, targetLng);
                if (dist <= 40) {
                    gpsBox.className = 'field-gps-display gps-match-success';
                    gpsStatusEl.innerHTML = `✅ <strong>GPS:</strong> ${lat.toFixed(5)}, ${lng.toFixed(5)} (±${acc.toFixed(1)}m) • <span style="font-weight: 700;">¡Estás en el lote! (~${dist.toFixed(0)}m)</span>`;
                } else {
                    gpsBox.className = 'field-gps-display gps-match-warning';
                    gpsStatusEl.innerHTML = `⚠️ <strong>GPS:</strong> ${lat.toFixed(5)}, ${lng.toFixed(5)} (±${acc.toFixed(1)}m) • Distancia al lote: <strong>~${dist.toFixed(0)}m</strong>`;
                }
            } else {
                gpsStatusEl.innerHTML = `📍 <strong>GPS:</strong> ${lat.toFixed(5)}, ${lng.toFixed(5)} (Precisión: ±${acc.toFixed(1)}m)`;
            }
        },
        (err) => {
            console.warn('Error de GPS:', err);
            gpsStatusEl.textContent = `⚠️ No se pudo obtener GPS (${err.message}). Puedes guardar la notificación sin GPS.`;
        },
        { enableHighAccuracy: true, timeout: 10000, maximumAge: 0 }
    );
}

function openFieldCaptureModal(props) {
    AppState.currentFieldLotProps = props;
    const cod = props.COD_PROMUVI || 'LOTE';
    const sheetData = AppState.dbSheet.get(cod.toUpperCase()) || AppState.dbSheet.get(normCode(cod));

    document.getElementById('field-lot-code').textContent = cod;
    document.getElementById('field-lot-titular').textContent = (sheetData && sheetData.TITULAR_1) ? sheetData.TITULAR_1 : (props.ASOCIACIÓN || 'Subsector 1A-4');
    document.getElementById('field-lot-details').textContent = `Mz: ${props.MANZANA || '-'} • Lote: ${props.LOTE || '-'} • Carpeta: ${sheetData ? sheetData.NRO_CARPETA : 'SIN CARPETA'}`;

    // Resetear foto
    AppState.capturedPhotoBase64 = null;
    AppState.capturedPhotoDataUrl = null;
    document.getElementById('photo-preview-container').classList.add('hidden');
    document.getElementById('photo-upload-placeholder').classList.remove('hidden');
    const camInput = document.getElementById('field-camera-input');
    if (camInput) camInput.value = '';

    // Cargar nombre de notificador guardado previamente
    const savedNotifier = localStorage.getItem('field_notifier_name') || '';
    const notifierInput = document.getElementById('field-notifier-name');
    if (notifierInput && savedNotifier) notifierInput.value = savedNotifier;

    // Resetear observaciones
    const notesInput = document.getElementById('field-notes');
    if (notesInput) notesInput.value = '';

    // URL de Apps Script
    const gasInput = document.getElementById('field-gas-url');
    if (gasInput) gasInput.value = AppState.fieldAppsScriptUrl || '';

    // Obtener centroide del lote para comparar distancia GPS
    let lotLat = null, lotLng = null;
    if (AppState.highlightedLayer && AppState.highlightedLayer.getBounds) {
        const c = AppState.highlightedLayer.getBounds().getCenter();
        lotLat = c.lat;
        lotLng = c.lng;
    }

    requestFieldGps(lotLat, lotLng);

    document.getElementById('field-capture-modal').classList.remove('hidden');
}

async function submitFieldRecord() {
    if (!AppState.currentFieldLotProps) {
        showToast('No hay ningún lote seleccionado.', 'warning');
        return;
    }

    const cod = AppState.currentFieldLotProps.COD_PROMUVI;
    const selectedRadio = document.querySelector('input[name="field-status-radio"]:checked');
    const estado = selectedRadio ? selectedRadio.value : 'NOTIFICADO EN MANO';
    const notifier = (document.getElementById('field-notifier-name')?.value || '').trim();
    const notes = (document.getElementById('field-notes')?.value || '').trim();

    if (notifier) {
        localStorage.setItem('field_notifier_name', notifier);
    }

    const btnSubmit = document.getElementById('btn-submit-field-record');
    if (btnSubmit) {
        btnSubmit.disabled = true;
        btnSubmit.innerHTML = '<span>⏳ Enviando a Google Drive y Sheets...</span>';
    }

    const record = {
        cod_promuvi: cod,
        estado_notificacion: estado,
        fecha_notificacion: new Date().toLocaleString('es-PE'),
        notificador: notifier,
        observaciones: notes,
        gps_lat: AppState.currentFieldGps ? AppState.currentFieldGps.lat : '',
        gps_lng: AppState.currentFieldGps ? AppState.currentFieldGps.lng : '',
        gps_precision: AppState.currentFieldGps ? AppState.currentFieldGps.accuracy : '',
        foto_base64: AppState.capturedPhotoBase64 || '',
        foto_url: ''
    };

    // 1. Guardar copia local de inmediato
    saveLocalFieldRecord(record);

    // 2. Si hay URL de Google Apps Script, enviar vía POST
    if (AppState.fieldAppsScriptUrl) {
        let sentOk = false;
        try {
            // Intento 1: Fetch estándar CORS con redirección automática
            const resp = await fetch(AppState.fieldAppsScriptUrl, {
                method: 'POST',
                mode: 'cors',
                redirect: 'follow',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify(record)
            });

            if (resp.ok) {
                const resJson = await resp.json();
                if (resJson && resJson.foto_url) {
                    record.foto_url = resJson.foto_url;
                    saveLocalFieldRecord(record);
                }
                sentOk = true;
                showToast(`✅ Notificación y foto subidas a Google Drive con éxito (${cod}).`, 'success', 5000);
            } else {
                throw new Error(`HTTP ${resp.status}`);
            }
        } catch (corsErr) {
            console.warn('Fallo intento CORS estándar, probando envío resiliente por modo no-cors...', corsErr);
            try {
                // Intento 2: Fallback resiliente con mode 'no-cors' (Google Apps Script procesa y guarda la foto en Drive y Sheets)
                await fetch(AppState.fieldAppsScriptUrl, {
                    method: 'POST',
                    mode: 'no-cors',
                    redirect: 'follow',
                    headers: { 'Content-Type': 'text/plain' },
                    body: JSON.stringify(record)
                });
                sentOk = true;
                showToast(`✅ Notificación y foto enviadas a Google Drive y Sheets (${cod}).`, 'success', 5000);
            } catch (fallbackErr) {
                console.error('Error total en el envío:', fallbackErr);
                showToast(`⚠️ Guardado en celular. Al reconectar se subirá a Drive: ${corsErr.message}`, 'warning', 6000);
            }
        }
    } else {
        showToast(`✅ Notificación guardada en el dispositivo (${cod}). Pega la URL de Apps Script para subida directa a Drive.`, 'info', 6000);
    }

    if (btnSubmit) {
        btnSubmit.disabled = false;
        btnSubmit.innerHTML = '<span>💾 Guardar y Enviar Notificación a Google Drive</span>';
    }

    document.getElementById('field-capture-modal')?.classList.add('hidden');

    // Refrescar modal de lote para mostrar la nueva información
    openPropertyModal(AppState.currentFieldLotProps);
    showToast(`Estado de ${cod} actualizado a: ${estado}`, 'success');
}

/**
 * Centra el mapa en la ubicación satelital GPS en tiempo real del usuario
 */
function locateUserOnMap() {
    if (!navigator.geolocation) {
        showToast('Geolocalización GPS no soportada en este dispositivo.', 'warning');
        return;
    }

    showToast('Obteniendo tu ubicación satelital GPS...', 'info', 2500);

    navigator.geolocation.getCurrentPosition(
        (pos) => {
            const lat = pos.coords.latitude;
            const lng = pos.coords.longitude;
            const acc = pos.coords.accuracy || 10;

            if (AppState.gpsUserMarker) AppState.map.removeLayer(AppState.gpsUserMarker);
            if (AppState.gpsUserCircle) AppState.map.removeLayer(AppState.gpsUserCircle);

            AppState.gpsUserCircle = L.circle([lat, lng], {
                radius: acc,
                color: '#38bdf8',
                fillColor: '#38bdf8',
                fillOpacity: 0.15,
                weight: 1
            }).addTo(AppState.map);

            AppState.gpsUserMarker = L.circleMarker([lat, lng], {
                radius: 8,
                fillColor: '#0284c7',
                color: '#ffffff',
                weight: 2.5,
                fillOpacity: 1
            }).addTo(AppState.map);

            AppState.map.setView([lat, lng], 19);
            showToast(`Ubicación GPS fijada (Precisión: ±${acc.toFixed(1)}m).`, 'success');
        },
        (err) => {
            showToast(`No se pudo obtener GPS: ${err.message}`, 'error');
        },
        { enableHighAccuracy: true, timeout: 10000, maximumAge: 0 }
    );
}

// ====== DASHBOARD DE MÉTRICAS & PADRÓN ======

function renderDashboard() {
    const totalPredios = 311;
    const totalNotificar = AppState.rawSheetRows.length; // 77
    
    let sinCarpetaCount = 0;
    let noTituladoCount = 0;
    let conDniCount = 0;
    const asocCount = {};

    AppState.rawSheetRows.forEach(r => {
        if (r.ESTADO.includes('SIN CARPETA')) sinCarpetaCount++;
        else if (r.ESTADO.includes('NO TITULADO')) noTituladoCount++;
        if (r.DNI_1 || r.DNI_2) conDniCount++;

        const asoc = r.ASOCIACION || 'OTRA';
        asocCount[asoc] = (asocCount[asoc] || 0) + 1;
    });

    const kpiContainer = document.getElementById('dashboard-kpis');
    if (kpiContainer) {
        kpiContainer.innerHTML = `
            <div class="metric-card blue">
                <span class="metric-number">${totalNotificar}</span>
                <span class="metric-title">En Padrón Notificar</span>
                <span class="metric-sub">${((totalNotificar/totalPredios)*100).toFixed(1)}% del Subsector 1A-4</span>
            </div>
            <div class="metric-card amber">
                <span class="metric-number">${sinCarpetaCount}</span>
                <span class="metric-title">Sin Carpeta</span>
                <span class="metric-sub">Requiere apertura (${((sinCarpetaCount/totalNotificar)*100).toFixed(1)}%)</span>
            </div>
            <div class="metric-card pink">
                <span class="metric-number">${noTituladoCount}</span>
                <span class="metric-title">No Titulados</span>
                <span class="metric-sub">Con expediente en trámite (${((noTituladoCount/totalNotificar)*100).toFixed(1)}%)</span>
            </div>
            <div class="metric-card emerald">
                <span class="metric-number">${conDniCount}</span>
                <span class="metric-title">Titular Identificado</span>
                <span class="metric-sub">Con DNI registrado en hoja</span>
            </div>
            <div class="metric-card slate">
                <span class="metric-number">${totalPredios}</span>
                <span class="metric-title">Total Sector 1A-4</span>
                <span class="metric-sub">Polígonos catastrales totales</span>
            </div>
        `;
    }

    renderAsociacionChart(asocCount, totalNotificar);
    renderEstadoChart(sinCarpetaCount, noTituladoCount, totalNotificar);
    renderDashboardTable(AppState.rawSheetRows);
}

function renderAsociacionChart(asocCount, total) {
    const container = document.getElementById('chart-asociaciones');
    if (!container) return;

    const entries = Object.entries(asocCount).sort((a, b) => b[1] - a[1]);
    const maxVal = Math.max(...entries.map(e => e[1]), 1);

    let html = `<div style="display: flex; flex-direction: column; gap: 10px; width: 100%; padding: 4px;">`;
    entries.forEach(([name, count]) => {
        const pct = ((count / total) * 100).toFixed(1);
        const barWidth = Math.round((count / maxVal) * 100);
        const col = ASOCIACION_COLORS[name] || '#38bdf8';

        html += `
            <div>
                <div style="display: flex; justify-content: space-between; font-size: 11px; margin-bottom: 3px;">
                    <span style="font-weight: 700; color: #f8fafc;">${name}</span>
                    <span style="color: ${col}; font-weight: 800;">${count} predios (${pct}%)</span>
                </div>
                <div style="background: rgba(255,255,255,0.06); height: 8px; border-radius: 4px; overflow: hidden;">
                    <div style="background: ${col}; width: ${barWidth}%; height: 100%; border-radius: 4px;"></div>
                </div>
            </div>
        `;
    });
    html += `</div>`;
    container.innerHTML = html;
}

function renderEstadoChart(sinCarp, noTit, total) {
    const container = document.getElementById('chart-estados');
    if (!container) return;

    const pctSin = ((sinCarp / total) * 100).toFixed(1);
    const pctNoTit = ((noTit / total) * 100).toFixed(1);

    container.innerHTML = `
        <div style="display: flex; flex-direction: column; gap: 16px; width: 100%; padding: 4px;">
            <div>
                <div style="display: flex; justify-content: space-between; font-size: 12px; margin-bottom: 4px;">
                    <strong style="color: var(--status-sin-carpeta);">SIN CARPETA</strong>
                    <span style="font-weight: 800; color: #fff;">${sinCarp} predios (${pctSin}%)</span>
                </div>
                <div style="background: rgba(255,255,255,0.06); height: 12px; border-radius: 6px; overflow: hidden;">
                    <div style="background: var(--status-sin-carpeta); width: ${pctSin}%; height: 100%; border-radius: 6px;"></div>
                </div>
            </div>
            <div>
                <div style="display: flex; justify-content: space-between; font-size: 12px; margin-bottom: 4px;">
                    <strong style="color: var(--status-no-titulado);">NO TITULADO</strong>
                    <span style="font-weight: 800; color: #fff;">${noTit} predios (${pctNoTit}%)</span>
                </div>
                <div style="background: rgba(255,255,255,0.06); height: 12px; border-radius: 6px; overflow: hidden;">
                    <div style="background: var(--status-no-titulado); width: ${pctNoTit}%; height: 100%; border-radius: 6px;"></div>
                </div>
            </div>
        </div>
    `;
}

function renderDashboardTable(rows) {
    const tbody = document.getElementById('dashboard-table-body');
    if (!tbody) return;

    if (!rows || rows.length === 0) {
        tbody.innerHTML = `<tr><td colspan="8" style="text-align:center; padding: 20px; color: var(--text-muted);">Sin resultados.</td></tr>`;
        return;
    }

    let trs = '';
    rows.forEach(r => {
        const isSinCarpeta = r.ESTADO.includes('SIN CARPETA');
        const badgeCol = isSinCarpeta ? 'color: #f59e0b; background: rgba(245,158,11,0.15);' : 'color: #ec4899; background: rgba(236,72,153,0.15);';

        trs += `
            <tr style="border-bottom: 1px solid rgba(255,255,255,0.04);">
                <td style="padding: 8px; font-weight: 700; color: #38bdf8;">${r.COD_PROMUVI}</td>
                <td style="padding: 8px;">${r.ASOCIACION || '--'}</td>
                <td style="padding: 8px;">${r.MANZANA || '-'}-${r.LOTE || '-'}</td>
                <td style="padding: 8px;"><span style="padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 700; ${badgeCol}">${r.ESTADO}</span></td>
                <td style="padding: 8px; font-family: monospace;">${r.NRO_CARPETA || '--'}</td>
                <td style="padding: 8px;">${r.TITULAR_1 || '--'}</td>
                <td style="padding: 8px; font-family: monospace;">${r.DNI_1 || '--'}</td>
                <td style="padding: 8px; text-align: center;">
                    <button class="btn-locate-row btn-secondary" data-cod="${r.COD_PROMUVI}" style="padding: 2px 8px; font-size: 10px;">Ver 🔍</button>
                </td>
            </tr>
        `;
    });

    tbody.innerHTML = trs;

    tbody.querySelectorAll('.btn-locate-row').forEach(btn => {
        btn.addEventListener('click', () => {
            const cod = btn.getAttribute('data-cod');
            document.getElementById('dashboard-modal')?.classList.add('hidden');
            locateFeatureByCode(cod);
        });
    });
}

function renderAsociacionesFilterList() {
    const list = document.getElementById('asociaciones-list');
    if (!list) return;

    list.innerHTML = '';
    Array.from(AppState.asociacionesSet).sort().forEach(asoc => {
        const item = document.createElement('div');
        item.className = 'chk-filter-container';
        item.style.padding = '6px 8px';
        item.style.background = 'rgba(255,255,255,0.02)';

        const col = ASOCIACION_COLORS[asoc] || '#38bdf8';
        item.innerHTML = `
            <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; flex-grow: 1; margin: 0; font-size: 11px;">
                <input type="checkbox" class="chk-asoc-item" data-asoc="${asoc}" checked style="accent-color: ${col};">
                <span style="color: ${col}; font-weight: 700;">●</span>
                <span>${asoc}</span>
            </label>
        `;

        item.querySelector('input').addEventListener('change', (e) => {
            if (e.target.checked) AppState.asociacionesActivas.add(asoc);
            else AppState.asociacionesActivas.delete(asoc);
            aplicarFiltros();
        });

        list.appendChild(item);
    });
}

// ====== BÚSQUEDA Y NAVEGACIÓN ======

function locateFeatureByCode(codeQuery) {
    if (!codeQuery || !AppState.geojsonLayer) return;

    const targetNorm = normCode(codeQuery);
    const targetClean = codeQuery.trim().toUpperCase();

    let targetLayer = null;

    AppState.geojsonLayer.eachLayer(layer => {
        const props = layer.feature.properties;
        const cod = props.COD_PROMUVI || '';
        if (cod.toUpperCase() === targetClean || normCode(cod) === targetNorm) {
            targetLayer = layer;
        }
    });

    if (targetLayer) {
        if (!AppState.map.hasLayer(targetLayer)) {
            AppState.map.addLayer(targetLayer);
        }
        highlightLayer(targetLayer);
        AppState.map.fitBounds(targetLayer.getBounds(), { padding: [80, 80], maxZoom: 20 });
        openPropertyModal(targetLayer.feature.properties);
        if (isMobile()) closeSidebar();
        showToast(`Predio ${targetClean} localizado en mapa.`, 'success');
    } else {
        showToast(`No se encontró el predio "${codeQuery}" en la cartografía.`, 'warning');
    }
}

function handleQuickSearch() {
    const q = document.getElementById('searchInput')?.value.trim();
    if (!q) {
        showToast('Ingrese un código, lote, DNI o nombre para buscar.', 'info');
        return;
    }

    const qUpper = q.toUpperCase();
    let foundLayer = null;

    AppState.geojsonLayer.eachLayer(layer => {
        if (foundLayer) return;
        const p = layer.feature.properties;
        const cod = (p.COD_PROMUVI || '').toUpperCase();
        const mz = (p.MANZANA || '').toUpperCase();
        const lt = (p.LOTE || '').toUpperCase();
        const part = (p.PART_REG || '').toUpperCase();
        const tit = (p.TITULAR_1 || '').toUpperCase();
        const dni = (p.DNI_1 || '').toUpperCase();
        const carp = (p.NRO_CARPETA || '').toUpperCase();

        if (
            cod === qUpper || normCode(cod) === normCode(qUpper) ||
            (mz === qUpper && lt === qUpper) ||
            part.includes(qUpper) ||
            tit.includes(qUpper) ||
            dni === qUpper ||
            carp === qUpper ||
            `${mz}_${lt}` === qUpper
        ) {
            foundLayer = layer;
        }
    });

    if (foundLayer) {
        if (!AppState.map.hasLayer(foundLayer)) {
            AppState.map.addLayer(foundLayer);
        }
        highlightLayer(foundLayer);
        AppState.map.fitBounds(foundLayer.getBounds(), { padding: [80, 80], maxZoom: 20 });
        openPropertyModal(foundLayer.feature.properties);
        if (isMobile()) closeSidebar();
        showToast(`Predio localizado: ${foundLayer.feature.properties.COD_PROMUVI}`, 'success');
    } else {
        showToast(`No se encontraron predios que coincidan con "${q}".`, 'warning');
    }
}

function handleBatchSearch() {
    const text = document.getElementById('batch-input')?.value.trim();
    if (!text) {
        showToast('Pegue una lista de códigos o DNIs en el área de texto.', 'warning');
        return;
    }

    const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
    AppState.batchMatches.clear();

    lines.forEach(l => {
        AppState.batchMatches.add(l.toUpperCase());
        AppState.batchMatches.add(normCode(l));
    });

    aplicarFiltros();

    const resultBox = document.getElementById('batch-results');
    if (resultBox) {
        resultBox.textContent = `Proyectados ${AppState.visibleLayers.length} de ${lines.length} códigos ingresados.`;
    }

    if (AppState.visibleLayers.length > 0) {
        const featureGroup = L.featureGroup(AppState.visibleLayers);
        AppState.map.fitBounds(featureGroup.getBounds(), { padding: [40, 40] });
        if (isMobile()) closeSidebar();
        showToast(`Se proyectaron ${AppState.visibleLayers.length} predios en el mapa.`, 'success');
    } else {
        showToast('No se encontraron coincidencias para los códigos ingresados.', 'warning');
    }
}

// ====== GESTIÓN DE BARRA LATERAL (RESPONSIVE MÓVIL Y ESCRITORIO) ======
function isMobile() {
    return window.innerWidth <= 900;
}

function openSidebar() {
    const sidebar = document.getElementById('sidebar');
    const backdrop = document.getElementById('sidebar-backdrop');
    if (!sidebar) return;
    if (isMobile()) {
        sidebar.classList.add('sidebar-open');
        sidebar.classList.remove('hidden-sidebar');
        if (backdrop) backdrop.classList.add('active');
    } else {
        sidebar.classList.remove('hidden-sidebar');
    }
    setTimeout(() => {
        if (AppState.map) AppState.map.invalidateSize();
    }, 320);
}

function closeSidebar() {
    const sidebar = document.getElementById('sidebar');
    const backdrop = document.getElementById('sidebar-backdrop');
    if (!sidebar) return;
    if (isMobile()) {
        sidebar.classList.remove('sidebar-open');
        if (backdrop) backdrop.classList.remove('active');
    } else {
        sidebar.classList.add('hidden-sidebar');
    }
    setTimeout(() => {
        if (AppState.map) AppState.map.invalidateSize();
    }, 320);
}

function toggleSidebar() {
    const sidebar = document.getElementById('sidebar');
    if (!sidebar) return;
    if (isMobile()) {
        if (sidebar.classList.contains('sidebar-open')) {
            closeSidebar();
        } else {
            openSidebar();
        }
    } else {
        sidebar.classList.toggle('hidden-sidebar');
        setTimeout(() => {
            if (AppState.map) AppState.map.invalidateSize();
        }, 320);
    }
}

// ====== LISTENERS DE EVENTOS Y CONTROLES ======
function setupEventListeners() {

    // 1. Selector de modo
    document.getElementById('mode-selector')?.addEventListener('change', (e) => {
        AppState.modo = e.target.value;
        const boxNormal = document.getElementById('box-busqueda-normal');
        const boxMasivo = document.getElementById('box-busqueda-masiva');
        const boxEstado = document.getElementById('box-filtros-estado');
        const boxAsoc = document.getElementById('box-filtros-asociacion');

        if (AppState.modo === 'MASIVO') {
            boxNormal?.classList.add('hidden');
            boxMasivo?.classList.remove('hidden');
            boxEstado?.classList.add('hidden');
            boxAsoc?.classList.add('hidden');
        } else if (AppState.modo === 'ASOCIACION') {
            boxNormal?.classList.remove('hidden');
            boxMasivo?.classList.add('hidden');
            boxEstado?.classList.add('hidden');
            boxAsoc?.classList.remove('hidden');
        } else {
            boxNormal?.classList.remove('hidden');
            boxMasivo?.classList.add('hidden');
            boxEstado?.classList.remove('hidden');
            boxAsoc?.classList.add('hidden');
        }
        aplicarFiltros();
    });

    // 2. Búsqueda rápida
    document.getElementById('btnSearch')?.addEventListener('click', handleQuickSearch);
    document.getElementById('searchInput')?.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') handleQuickSearch();
    });

    // 3. Búsqueda Masiva
    document.getElementById('btnBatch')?.addEventListener('click', handleBatchSearch);
    document.getElementById('btnBatchClear')?.addEventListener('click', () => {
        const ta = document.getElementById('batch-input');
        if (ta) ta.value = '';
        AppState.batchMatches.clear();
        const res = document.getElementById('batch-results');
        if (res) res.textContent = '';
        aplicarFiltros();
        showToast('Lista masiva limpiada.', 'info');
    });

    // 4. Filtros de Estado
    document.getElementById('chkOnlyNotificar')?.addEventListener('change', (e) => {
        AppState.onlyNotificar = e.target.checked;
        aplicarFiltros();
    });

    document.getElementById('chkSingleColor')?.addEventListener('change', (e) => {
        AppState.isSingleColor = e.target.checked;
        const picker = document.getElementById('colorPickerSingle');
        if (picker) picker.disabled = !AppState.isSingleColor;
        aplicarFiltros();
    });
    document.getElementById('colorPickerSingle')?.addEventListener('input', (e) => {
        AppState.colorSingle = e.target.value;
        if (AppState.isSingleColor) aplicarFiltros();
    });

    document.getElementById('chkSinCarpeta')?.addEventListener('change', (e) => {
        AppState.showSinCarpeta = e.target.checked;
        aplicarFiltros();
    });
    document.getElementById('colorPickerSinCarpeta')?.addEventListener('input', (e) => {
        AppState.colorSinCarpeta = e.target.value;
        aplicarFiltros();
    });

    document.getElementById('chkNoTitulado')?.addEventListener('change', (e) => {
        AppState.showNoTitulado = e.target.checked;
        aplicarFiltros();
    });
    document.getElementById('colorPickerNoTitulado')?.addEventListener('input', (e) => {
        AppState.colorNoTitulado = e.target.value;
        aplicarFiltros();
    });

    document.getElementById('chkOtrosPredios')?.addEventListener('change', (e) => {
        AppState.showOtros = e.target.checked;
        aplicarFiltros();
    });
    document.getElementById('colorPickerOtros')?.addEventListener('input', (e) => {
        AppState.colorOtros = e.target.value;
        aplicarFiltros();
    });

    // 5. Botones Todas / Ninguna Asociación
    document.getElementById('btn-asoc-all')?.addEventListener('click', () => {
        AppState.asociacionesActivas = new Set(AppState.asociacionesSet);
        document.querySelectorAll('.chk-asoc-item').forEach(c => c.checked = true);
        aplicarFiltros();
    });
    document.getElementById('btn-asoc-none')?.addEventListener('click', () => {
        AppState.asociacionesActivas.clear();
        document.querySelectorAll('.chk-asoc-item').forEach(c => c.checked = false);
        aplicarFiltros();
    });

    // 6. SECCIÓN FONDO CARTOGRÁFICO
    document.getElementById('fondo-toggle')?.addEventListener('click', () => {
        const p = document.getElementById('fondo-panel');
        const arrow = document.getElementById('fondo-arrow');
        if (p) {
            p.classList.toggle('hidden');
            if (arrow) arrow.textContent = p.classList.contains('hidden') ? '▶' : '▼';
        }
    });

    document.querySelectorAll('input[name="fondoOption"]').forEach(radio => {
        radio.addEventListener('change', (e) => {
            AppState.fondoActivo = e.target.value;
            loadFondoActivo();
            showToast(`Capa de Fondo cambiada a: ${e.target.value.toUpperCase()}`, 'info');
        });
    });

    document.getElementById('sliderFondoOpacidad')?.addEventListener('input', (e) => {
        AppState.fondoOpacidad = parseInt(e.target.value) / 100;
        const v = document.getElementById('val-fondo-opacidad');
        if (v) v.textContent = e.target.value + '%';
        if (AppState.capaFondoGroup) {
            AppState.capaFondoGroup.eachLayer(layer => {
                if (layer.setStyle) {
                    layer.setStyle({
                        opacity: AppState.fondoOpacidad,
                        fillOpacity: AppState.isFondoInteractive ? (AppState.fondoOpacidad * 0.15) : 0
                    });
                }
            });
        }
    });

    document.getElementById('colorPickerFondo')?.addEventListener('input', (e) => {
        AppState.fondoColor = e.target.value;
        if (AppState.capaFondoGroup) {
            AppState.capaFondoGroup.eachLayer(layer => {
                if (layer.setStyle) {
                    layer.setStyle({
                        color: AppState.fondoColor,
                        fillColor: AppState.fondoColor
                    });
                }
            });
        }
    });

    document.getElementById('chkInfoFondo')?.addEventListener('change', (e) => {
        AppState.isFondoInteractive = e.target.checked;
        loadFondoActivo();
    });

    // 7. CALIBRACIÓN CARTOGRÁFICA EN L CON ÍCONO DE TUERCA (ESTILO SANEAMIENTO)
    const toggleBtn = document.getElementById('toggle-l-calibration');
    const lContainer = document.getElementById('l-axes-container');
    
    toggleBtn?.addEventListener('click', (e) => {
        e.stopPropagation();
        if (lContainer) {
            const isHidden = lContainer.classList.toggle('hidden');
            toggleBtn.classList.toggle('active', !isHidden);
            if (!isHidden) {
                showToast('Calibración cartográfica activa (Ejes L).', 'info', 2000);
            }
        }
    });

    document.getElementById('f-offset-east')?.addEventListener('input', (e) => {
        handleOffsetInput(e.target.value, AppState.offsetN);
    });

    document.getElementById('f-offset-north')?.addEventListener('input', (e) => {
        handleOffsetInput(AppState.offsetE, e.target.value);
    });

    document.getElementById('f-offset-reset')?.addEventListener('click', () => {
        handleOffsetInput(0, 0);
        showToast('Calibración restablecida a 0.0 m.', 'info');
    });

    // 8. Selección Espacial
    document.getElementById('btnDrawStart')?.addEventListener('click', () => {
        AppState.spatialDrawer.startDrawing();
        document.getElementById('btnDrawStart')?.classList.add('hidden');
        document.getElementById('btnDrawFilter')?.classList.remove('hidden');
        document.getElementById('btnDrawClear')?.classList.remove('hidden');
        if (isMobile()) closeSidebar();
    });

    document.getElementById('btnDrawFilter')?.addEventListener('click', () => {
        const points = AppState.spatialDrawer.drawPoints;
        if (!points || points.length < 3) {
            showToast('Dibuje un área con al menos 3 vértices para filtrar.', 'warning');
            return;
        }

        const filtered = [];
        AppState.visibleLayers.forEach(layer => {
            if (isFeatureInsideSelection(layer, points)) {
                filtered.push(layer);
            } else {
                AppState.map.removeLayer(layer);
            }
        });

        AppState.visibleLayers = filtered;
        updateLabels();
        showToast(`Filtro espacial aplicado: ${filtered.length} predios dentro del área.`, 'success');
    });

    document.getElementById('btnDrawClear')?.addEventListener('click', () => {
        AppState.spatialDrawer.clearDrawing();
        aplicarFiltros();
        showToast('Selección espacial restablecida.', 'info');
    });

    // 9. Exportaciones
    document.getElementById('btnExportCAD')?.addEventListener('click', () => {
        if (AppState.visibleLayers.length === 0) {
            showToast('No hay predios visibles para exportar.', 'warning');
            return;
        }
        const dxf = generateDXF(AppState.visibleLayers);
        downloadFile(dxf, `Subsector_1A4_Notificaciones_${Date.now()}.dxf`, 'text/plain');
        showToast(`AutoCAD DXF exportado con ${AppState.visibleLayers.length} predios en UTM 19S.`, 'success');
    });

    document.getElementById('btnExportKML')?.addEventListener('click', () => {
        if (AppState.visibleLayers.length === 0) {
            showToast('No hay predios visibles para exportar.', 'warning');
            return;
        }
        const kml = generateKML(AppState.visibleLayers, {
            colSinCarpeta: AppState.colorSinCarpeta,
            colNoTitulado: AppState.colorNoTitulado,
            colOtros: AppState.colorOtros
        });
        downloadFile(kml, `Subsector_1A4_GoogleEarth_${Date.now()}.kml`, 'application/vnd.google-earth.kml+xml');
        showToast(`Google Earth KML exportado con ${AppState.visibleLayers.length} predios.`, 'success');
    });

    document.getElementById('btnExportExcel')?.addEventListener('click', () => {
        if (AppState.visibleLayers.length === 0) {
            showToast('No hay predios visibles para exportar.', 'warning');
            return;
        }
        const excelArray = generateExcelWorkbook(AppState.visibleLayers);
        const blob = new Blob([excelArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        downloadFile(blob, `Padron_Notificaciones_1A4_${Date.now()}.xlsx`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        showToast(`Reporte Excel exportado con ${AppState.visibleLayers.length} predios y coordenadas.`, 'success');
    });

    // 10. Etiquetas
    document.getElementById('chkShowLabels')?.addEventListener('change', (e) => {
        AppState.showLabels = e.target.checked;
        updateLabels();
    });

    // 11. Opacidad y Ajustes de Polígonos
    const updateOpacity = (val) => {
        AppState.opacidad = val / 100;
        const vFloat = document.getElementById('floatValOpacidad');
        if (vFloat) vFloat.textContent = `${val}%`;
        const sFloat = document.getElementById('floatSliderOpacidad');
        if (sFloat && sFloat.value !== val) sFloat.value = val;
        aplicarFiltros();
    };

    document.getElementById('floatSliderOpacidad')?.addEventListener('input', (e) => updateOpacity(e.target.value));

    document.getElementById('sliderStroke')?.addEventListener('input', (e) => {
        AppState.strokeWidth = parseFloat(e.target.value);
        const v = document.getElementById('val-stroke');
        if (v) v.textContent = `${AppState.strokeWidth} px`;
        aplicarFiltros();
    });

    document.getElementById('selectStrokeStyle')?.addEventListener('change', (e) => {
        AppState.strokeStyle = e.target.value;
        aplicarFiltros();
    });

    // 12. Mapa Base Exclusivo en el Visor
    const basemapToggleBtn = document.getElementById('btn-basemap-toggle');
    const basemapPanel = document.getElementById('basemap-floating-panel');
    const basemapCloseBtn = document.getElementById('btn-close-basemap');

    basemapToggleBtn?.addEventListener('click', (e) => {
        e.stopPropagation();
        if (basemapPanel) {
            const isHidden = basemapPanel.classList.toggle('hidden');
            basemapToggleBtn.classList.toggle('active', !isHidden);
        }
    });

    basemapCloseBtn?.addEventListener('click', (e) => {
        e.stopPropagation();
        basemapPanel?.classList.add('hidden');
        basemapToggleBtn?.classList.remove('active');
    });

    document.querySelectorAll('.basemap-item').forEach(item => {
        item.addEventListener('click', (e) => {
            const key = item.getAttribute('data-map');
            if (key && AppState.baseLayers[key]) {
                AppState.map.removeLayer(AppState.currentBaseLayer);
                AppState.currentBaseLayer = AppState.baseLayers[key];
                AppState.currentBaseLayer.addTo(AppState.map);
                AppState.currentBaseLayer.bringToBack();

                document.querySelectorAll('.basemap-item').forEach(el => el.classList.remove('active'));
                item.classList.add('active');

                const title = item.querySelector('.basemap-title')?.textContent || key;
                showToast(`Mapa base: ${title}`, 'info', 1500);

                if (isMobile()) {
                    basemapPanel?.classList.add('hidden');
                    basemapToggleBtn?.classList.remove('active');
                }
            }
        });
    });

    document.addEventListener('click', (e) => {
        if (basemapPanel && !basemapPanel.classList.contains('hidden')) {
            if (!basemapPanel.contains(e.target) && e.target !== basemapToggleBtn) {
                basemapPanel.classList.add('hidden');
                basemapToggleBtn?.classList.remove('active');
            }
        }
    });

    // 13. Recargar Google Sheets
    document.getElementById('btn-sync-reload')?.addEventListener('click', async () => {
        showToast('Sincronizando con Google Sheets...', 'info');
        try {
            await fetchGoogleSheetTSV();
            buildVectorLayer();
            aplicarFiltros();
            showToast('Datos de Google Sheets actualizados correctamente.', 'success');
        } catch (e) {
            showToast(`Error al sincronizar: ${e.message}`, 'error');
        }
    });

    // 14. Modales
    document.getElementById('btn-open-dashboard')?.addEventListener('click', () => {
        renderDashboard();
        document.getElementById('dashboard-modal')?.classList.remove('hidden');
    });
    document.getElementById('btn-close-dashboard-modal')?.addEventListener('click', () => {
        document.getElementById('dashboard-modal')?.classList.add('hidden');
    });

    document.getElementById('btn-close-prop-modal')?.addEventListener('click', () => {
        document.getElementById('property-modal')?.classList.add('hidden');
    });

    document.getElementById('btn-close-help-modal')?.addEventListener('click', () => {
        document.getElementById('help-modal')?.classList.add('hidden');
    });

    window.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            document.querySelectorAll('.glass-modal').forEach(m => m.classList.add('hidden'));
            basemapPanel?.classList.add('hidden');
            basemapToggleBtn?.classList.remove('active');
        }
    });

    document.querySelectorAll('.glass-modal').forEach(modal => {
        modal.addEventListener('click', (e) => {
            if (e.target === modal) modal.classList.add('hidden');
        });
    });

    // 15. Herramientas flotantes (Ultra fluido para celulares)
    document.getElementById('btn-recenter')?.addEventListener('click', () => {
        if (!AppState.geojsonLayer) return;

        if (!AppState.subsectorBounds) {
            AppState.subsectorBounds = AppState.geojsonLayer.getBounds();
        }

        const isMob = isMobile();
        AppState.map.fitBounds(AppState.subsectorBounds, {
            padding: isMob ? [15, 15] : [30, 30],
            maxZoom: 18,
            animate: true,
            duration: isMob ? 0.35 : 0.6
        });
        showToast('Vista centrada en Subsector 1A-4.', 'info', 1500);
    });

    document.getElementById('btn-toggle-labels-map')?.addEventListener('click', function() {
        AppState.showLabels = !AppState.showLabels;
        this.classList.toggle('active', AppState.showLabels);
        updateLabels();
        showToast(AppState.showLabels ? 'Etiquetas activadas' : 'Etiquetas desactivadas', 'info', 1500);
    });

    // Control de barra lateral responsive
    document.getElementById('btn-toggle-sidebar')?.addEventListener('click', toggleSidebar);
    document.getElementById('btn-close-sidebar-mobile')?.addEventListener('click', closeSidebar);
    document.getElementById('sidebar-backdrop')?.addEventListener('click', closeSidebar);

    window.addEventListener('resize', () => {
        if (!isMobile()) {
            document.getElementById('sidebar-backdrop')?.classList.remove('active');
            document.getElementById('sidebar')?.classList.remove('sidebar-open');
        }
        if (AppState.map) AppState.map.invalidateSize();
    });

    // 16. Filtro de búsqueda en la tabla del Dashboard
    document.getElementById('filter-table-input')?.addEventListener('input', (e) => {
        const val = e.target.value.trim().toUpperCase();
        const filtered = AppState.rawSheetRows.filter(r => 
            (r.COD_PROMUVI || '').toUpperCase().includes(val) ||
            (r.ASOCIACION || '').toUpperCase().includes(val) ||
            (r.TITULAR_1 || '').toUpperCase().includes(val) ||
            (r.DNI_1 || '').toUpperCase().includes(val) ||
            (r.NRO_CARPETA || '').toUpperCase().includes(val) ||
            (r.ESTADO || '').toUpperCase().includes(val)
        );
        renderDashboardTable(filtered);
    });

    document.getElementById('btn-download-table-excel')?.addEventListener('click', () => {
        const excelArray = generateExcelWorkbook(AppState.visibleLayers);
        const blob = new Blob([excelArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        downloadFile(blob, `Padron_Notificaciones_${Date.now()}.xlsx`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    });

    // 17. Recopilación de Campo (Fotos, GPS, Estados)
    document.getElementById('btn-gps-locate')?.addEventListener('click', locateUserOnMap);

    // Abrir selector de cámara al tocar el dropzone
    const photoPlaceholder = document.getElementById('photo-upload-placeholder');
    const cameraInput = document.getElementById('field-camera-input');
    if (photoPlaceholder && cameraInput) {
        photoPlaceholder.addEventListener('click', () => cameraInput.click());
    }

    // Procesamiento y compresión de imagen tomada por el celular
    cameraInput?.addEventListener('change', async (e) => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;

        showToast('Procesando y optimizando fotografía...', 'info', 2000);
        try {
            const compressed = await compressImage(file, 1280, 1280, 0.75);
            AppState.capturedPhotoBase64 = compressed.base64;
            AppState.capturedPhotoDataUrl = compressed.dataUrl;

            const previewImg = document.getElementById('photo-preview-img');
            const previewBox = document.getElementById('photo-preview-container');
            const infoText = document.getElementById('photo-file-info');

            if (previewImg) previewImg.src = compressed.dataUrl;
            if (infoText) {
                const kb = Math.round(compressed.sizeBytes / 1024);
                infoText.textContent = `${compressed.width}x${compressed.height} px • ${kb} KB • Optimizado`;
            }
            if (photoPlaceholder) photoPlaceholder.classList.add('hidden');
            if (previewBox) previewBox.classList.remove('hidden');

            showToast('Foto cargada y comprimida correctamente.', 'success');
        } catch (err) {
            console.error('Error al comprimir imagen:', err);
            showToast('No se pudo procesar la imagen.', 'error');
        }
    });

    // Retomar y Quitar Foto
    document.getElementById('btn-retake-photo')?.addEventListener('click', () => {
        cameraInput?.click();
    });

    document.getElementById('btn-remove-photo')?.addEventListener('click', () => {
        AppState.capturedPhotoBase64 = null;
        AppState.capturedPhotoDataUrl = null;
        if (cameraInput) cameraInput.value = '';
        document.getElementById('photo-preview-container')?.classList.add('hidden');
        document.getElementById('photo-upload-placeholder')?.classList.remove('hidden');
    });

    // Refrescar GPS en modal de campo
    document.getElementById('btn-refresh-field-gps')?.addEventListener('click', () => {
        let lotLat = null, lotLng = null;
        if (AppState.highlightedLayer && AppState.highlightedLayer.getBounds) {
            const c = AppState.highlightedLayer.getBounds().getCenter();
            lotLat = c.lat;
            lotLng = c.lng;
        }
        requestFieldGps(lotLat, lotLng);
    });

    // Enviar registro de campo
    document.getElementById('btn-submit-field-record')?.addEventListener('click', submitFieldRecord);

    // Cerrar modal de campo
    document.getElementById('btn-close-field-modal')?.addEventListener('click', () => {
        document.getElementById('field-capture-modal')?.classList.add('hidden');
    });
    document.getElementById('btn-cancel-field-modal')?.addEventListener('click', () => {
        document.getElementById('field-capture-modal')?.classList.add('hidden');
    });

    // Desplegable de Configuración de Google Apps Script
    document.getElementById('toggle-gas-config')?.addEventListener('click', () => {
        const panel = document.getElementById('gas-config-panel');
        const arrow = document.getElementById('arrow-gas-config');
        if (panel) {
            const isHidden = panel.classList.contains('hidden');
            panel.classList.toggle('hidden', !isHidden);
            if (arrow) arrow.textContent = isHidden ? '▼' : '▶';
        }
    });

    document.getElementById('btn-save-gas-url')?.addEventListener('click', () => {
        const url = (document.getElementById('field-gas-url')?.value || '').trim();
        AppState.fieldAppsScriptUrl = url;
        localStorage.setItem('field_apps_script_url', url);
        showToast(url ? 'URL de Google Apps Script guardada.' : 'URL eliminada.', 'success');
    });

    document.getElementById('btn-test-gas-url')?.addEventListener('click', async () => {
        const url = (document.getElementById('field-gas-url')?.value || '').trim();
        const resEl = document.getElementById('gas-test-result');
        if (!url) {
            showToast('Ingresa una URL antes de probar.', 'warning');
            return;
        }

        if (resEl) {
            resEl.style.display = 'block';
            resEl.style.background = 'rgba(56, 189, 248, 0.15)';
            resEl.style.color = '#38bdf8';
            resEl.innerHTML = '⏳ Probando conexión con Google Apps Script...';
        }

        try {
            const resp = await fetch(url, { method: 'GET', mode: 'cors' });
            const data = await resp.json();
            if (data && data.status === 'OK') {
                if (resEl) {
                    resEl.style.background = 'rgba(16, 185, 129, 0.2)';
                    resEl.style.color = '#34d399';
                    resEl.innerHTML = '✅ <strong>¡Conexión exitosa!</strong> El backend respondió correctamente y está listo para recibir fotos y datos.';
                }
                showToast('Conexión con Apps Script verificada.', 'success');
            } else {
                throw new Error('Respuesta no válida');
            }
        } catch (err) {
            if (resEl) {
                resEl.style.background = 'rgba(239, 68, 68, 0.2)';
                resEl.style.color = '#f87171';
                resEl.innerHTML = '⚠️ <strong>No se pudo conectar:</strong><br>1. Asegúrate de que la URL termine en <code>/exec</code> (no en <code>/dev</code>).<br>2. En Google Apps Script verifica que en "Quién tiene acceso" esté seleccionado <strong>"Cualquier persona" (Anyone)</strong>.<br>3. Autoriza los permisos si te lo solicitó Google.';
            }
            showToast('Error de conexión con Apps Script.', 'error');
        }
    });
}

// ====== ARRANQUE ======
window.addEventListener('DOMContentLoaded', () => {
    loadLocalFieldRecords();
    initMap();
    setupEventListeners();
    loadAllData();
    if (isMobile()) {
        closeSidebar();
    }
});
