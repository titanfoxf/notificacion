/**
 * ExportEngine.js
 * Client-side exports to AutoCAD DXF, Google Earth KML, and Microsoft Excel (.xlsx).
 */
import * as XLSX from 'xlsx';

/**
 * Converts Geographic Lat/Lon (WGS84) to UTM Coordinates (Zone 19S for Moquegua).
 */
export function latLonToUTM(lat, lon) {
    const a = 6378137.0; // Equatorial radius WGS84
    const eccSquared = 0.00669438002290; // Eccentricity squared
    const k0 = 0.9996; // Scale factor
    let zoneNumber = Math.floor((lon + 180) / 6) + 1;
    const latRad = lat * Math.PI / 180.0;
    const lonRad = lon * Math.PI / 180.0;
    const lonOrigin = (zoneNumber - 1) * 6 - 180 + 3;
    const lonOriginRad = lonOrigin * Math.PI / 180.0;
    const eccPrimeSquared = (eccSquared) / (1 - eccSquared);
    const N = a / Math.sqrt(1 - eccSquared * Math.sin(latRad) * Math.sin(latRad));
    const T = Math.tan(latRad) * Math.tan(latRad);
    const C = eccPrimeSquared * Math.cos(latRad) * Math.cos(latRad);
    const A = Math.cos(latRad) * (lonRad - lonOriginRad);
    const M = a * ((1 - eccSquared / 4 - 3 * eccSquared * eccSquared / 64 - 5 * Math.pow(eccSquared, 3) / 256) * latRad
            - (3 * eccSquared / 8 + 3 * eccSquared * eccSquared / 32 + 45 * Math.pow(eccSquared, 3) / 1024) * Math.sin(2 * latRad)
            + (15 * eccSquared * eccSquared / 256 + 45 * Math.pow(eccSquared, 3) / 1024) * Math.sin(4 * latRad)
            - (3 * eccSquared * eccSquared * eccSquared * 35 / 3072) * Math.sin(6 * latRad));
    let UTMEasting = (k0 * N * (A + (1 - T + C) * A * A * A / 6 + (5 - 18 * T + T * T + 72 * C - 58 * eccPrimeSquared) * Math.pow(A, 5) / 120) + 500000.0);
    let UTMNorthing = (k0 * (M + N * Math.tan(latRad) * (A * A / 2 + (5 - T + 9 * C + 4 * C * C) * Math.pow(A, 4) / 24 + (61 - 58 * T + T * T + 600 * C - 330 * eccPrimeSquared) * Math.pow(A, 6) / 720)));
    if (lat < 0) UTMNorthing += 10000000.0; // Southern hemisphere offset
    return { x: UTMEasting, y: UTMNorthing, zone: `${zoneNumber}S` };
}

/**
 * Triggers a browser download of a text/binary blob.
 */
export function downloadFile(content, filename, mimeType) {
    const blob = content instanceof Blob ? content : new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }, 200);
}

/**
 * Formats numbers for CSV / Clean tables.
 */
function fmtNum(n, dec = 2) {
    return (typeof n === 'number' && isFinite(n)) ? n.toFixed(dec) : (n || "");
}

/**
 * Generates an AutoCAD DXF (R12/2000 compatible ASCII format) in UTM Zone 19S coordinates.
 */
export function generateDXF(layers, options = {}) {
    if (!layers || layers.length === 0) return null;

    let dxf = "  0\nSECTION\n  2\nHEADER\n  0\nENDSEC\n";
    dxf += "  0\nSECTION\n  2\nTABLES\n";
    dxf += "  0\nTABLE\n  2\nLAYER\n 70\n3\n";
    dxf += "  0\nLAYER\n  2\nPREDIOS_SIN_CARPETA\n 70\n0\n 62\n2\n  6\nCONTINUOUS\n"; // Yellow
    dxf += "  0\nLAYER\n  2\nPREDIOS_NO_TITULADO\n 70\n0\n 62\n6\n  6\nCONTINUOUS\n";  // Magenta
    dxf += "  0\nLAYER\n  2\nPREDIOS_OTROS\n 70\n0\n 62\n8\n  6\nCONTINUOUS\n";        // Gray
    dxf += "  0\nLAYER\n  2\nTEXTOS_PREDIOS\n 70\n0\n 62\n7\n  6\nCONTINUOUS\n";      // White
    dxf += "  0\nENDTAB\n  0\nENDSEC\n";
    dxf += "  0\nSECTION\n  2\nBLOCKS\n  0\nENDSEC\n";
    dxf += "  0\nSECTION\n  2\nENTITIES\n";

    layers.forEach((layer) => {
        const feature = layer.feature || layer;
        if (!feature || !feature.geometry) return;

        const props = feature.properties || {};
        const cod = props.COD_PROMUVI || `LOTE_${props.OBJECTID || 'S_N'}`;
        const estado = (props.ESTADO || '').toUpperCase();

        let layerName = 'PREDIOS_OTROS';
        if (estado.includes('SIN CARPETA')) layerName = 'PREDIOS_SIN_CARPETA';
        else if (estado.includes('NO TITULADO')) layerName = 'PREDIOS_NO_TITULADO';

        // Extract rings from Leaflet layer or GeoJSON geometry
        let rings = [];
        if (layer.getLatLngs) {
            const raw = layer.getLatLngs();
            rings = Array.isArray(raw[0]) ? raw : [raw];
        } else if (feature.geometry.type === 'Polygon') {
            rings = [feature.geometry.coordinates[0].map(pt => ({ lat: pt[1], lng: pt[0] }))];
        }

        rings.forEach((ring) => {
            if (!ring || ring.length === 0) return;

            // LWPOLYLINE Entity
            dxf += "  0\nLWPOLYLINE\n";
            dxf += `  8\n${layerName}\n`;
            dxf += ` 90\n${ring.length}\n`;
            dxf += " 70\n1\n"; // 1 = Closed loop

            let sumX = 0, sumY = 0, count = 0;
            ring.forEach((latlng) => {
                const utm = latLonToUTM(latlng.lat, latlng.lng);
                dxf += ` 10\n${utm.x.toFixed(3)}\n`;
                dxf += ` 20\n${utm.y.toFixed(3)}\n`;
                sumX += utm.x;
                sumY += utm.y;
                count++;
            });

            // MTEXT Entity for lot code
            if (count > 0 && cod) {
                const cX = (sumX / count).toFixed(3);
                const cY = (sumY / count).toFixed(3);
                const textVal = `${cod}\\P${props.TITULAR_1 || ''}`.trim();

                dxf += "  0\nMTEXT\n";
                dxf += "  8\nTEXTOS_PREDIOS\n";
                dxf += ` 10\n${cX}\n`;
                dxf += ` 20\n${cY}\n`;
                dxf += " 30\n0.0\n";
                dxf += " 40\n1.8\n"; // Text height in meters
                dxf += " 71\n5\n";   // Middle center attachment
                dxf += `  1\n${textVal}\n`;
            }
        });
    });

    dxf += "  0\nENDSEC\n  0\nEOF\n";
    return dxf;
}

/**
 * Converts Hex to KML AABBGGRR color format.
 */
function hexToKmlColor(hex, opacity = 1) {
    hex = hex.replace('#', '');
    if (hex.length === 3) hex = hex.split('').map(c => c + c).join('');
    const r = hex.substring(0, 2);
    const g = hex.substring(2, 4);
    const b = hex.substring(4, 6);
    let alpha = Math.round(opacity * 255).toString(16).padStart(2, '0');
    return `${alpha}${b}${g}${r}`;
}

/**
 * Escapes XML strings.
 */
function escapeXml(str) {
    if (!str) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

/**
 * Generates a Google Earth KML file with styling and ExtendedData.
 */
export function generateKML(layers, options = {}) {
    if (!layers || layers.length === 0) return null;

    const docName = options.docName || "Notificaciones_Subsector_1A4";
    const colSinCarpeta = options.colSinCarpeta || "#f59e0b";
    const colNoTitulado = options.colNoTitulado || "#ec4899";
    const colOtros = options.colOtros || "#64748b";

    let kml = `<?xml version="1.0" encoding="UTF-8"?>\n`;
    kml += `<kml xmlns="http://www.opengis.net/kml/2.2">\n`;
    kml += `  <Document>\n`;
    kml += `    <name>${escapeXml(docName)}</name>\n`;
    kml += `    <open>1</open>\n`;

    // Define Styles
    const defStyle = (id, strokeColor, fillColor, strokeW = 2) => `
    <Style id="${id}">
      <LineStyle>
        <color>${hexToKmlColor(strokeColor, 1.0)}</color>
        <width>${strokeW}</width>
      </LineStyle>
      <PolyStyle>
        <color>${hexToKmlColor(fillColor, 0.45)}</color>
        <fill>1</fill>
        <outline>1</outline>
      </PolyStyle>
    </Style>`;

    kml += defStyle('styleSinCarpeta', colSinCarpeta, colSinCarpeta, 2.5);
    kml += defStyle('styleNoTitulado', colNoTitulado, colNoTitulado, 2.5);
    kml += defStyle('styleOtros', '#94a3b8', colOtros, 1.2);

    layers.forEach(layer => {
        const feature = layer.feature || layer;
        if (!feature || !feature.geometry) return;

        const props = feature.properties || {};
        const cod = props.COD_PROMUVI || `LOTE_${props.OBJECTID || ''}`;
        const estado = (props.ESTADO || '').toUpperCase();

        let styleUrl = '#styleOtros';
        if (estado.includes('SIN CARPETA')) styleUrl = '#styleSinCarpeta';
        else if (estado.includes('NO TITULADO')) styleUrl = '#styleNoTitulado';

        let coordinates = [];
        if (layer.getLatLngs) {
            const raw = layer.getLatLngs();
            const ring = Array.isArray(raw[0]) ? raw[0] : raw;
            coordinates = ring.map(pt => `${pt.lng},${pt.lat},0`);
        } else if (feature.geometry.type === 'Polygon') {
            coordinates = feature.geometry.coordinates[0].map(pt => `${pt[0]},${pt[1]},0`);
        }

        if (coordinates.length === 0) return;

        // Ensure loop closure for Google Earth
        if (coordinates[0] !== coordinates[coordinates.length - 1]) {
            coordinates.push(coordinates[0]);
        }

        kml += `    <Placemark>\n`;
        kml += `      <name>${escapeXml(cod)}</name>\n`;
        kml += `      <styleUrl>${styleUrl}</styleUrl>\n`;
        kml += `      <description><![CDATA[
            <b>Código:</b> ${cod}<br/>
            <b>Estado:</b> ${props.ESTADO || 'NO PRIORIZADO'}<br/>
            <b>Asociación:</b> ${props.ASOCIACIÓN || props.ASOCIACION || '--'}<br/>
            <b>Manzana/Lote:</b> Mz ${props.MANZANA || '--'}, Lt ${props.LOTE || '--'}<br/>
            <b>N° Carpeta:</b> ${props.NRO_CARPETA || props['Nº CARPETA'] || '--'}<br/>
            <b>Titular 1:</b> ${props.TITULAR_1 || '--'} (DNI: ${props.DNI_1 || '--'})<br/>
            <b>Titular 2:</b> ${props.TITULAR_2 || '--'} (DNI: ${props.DNI_2 || '--'})<br/>
            <b>Partida Registral:</b> ${props.PART_REG || '--'}<br/>
            <b>Área Gráfica:</b> ${fmtNum(props.GEO_AREA || props.Shape_Area, 2)} m²
        ]]></description>\n`;
        kml += `      <ExtendedData>\n`;
        for (let k in props) {
            if (typeof props[k] !== 'object' && props[k] !== undefined && props[k] !== null) {
                kml += `        <Data name="${escapeXml(k)}"><value>${escapeXml(props[k])}</value></Data>\n`;
            }
        }
        kml += `      </ExtendedData>\n`;
        kml += `      <Polygon>\n`;
        kml += `        <outerBoundaryIs>\n`;
        kml += `          <LinearRing>\n`;
        kml += `            <coordinates>${coordinates.join(' ')}</coordinates>\n`;
        kml += `          </LinearRing>\n`;
        kml += `        </outerBoundaryIs>\n`;
        kml += `      </Polygon>\n`;
        kml += `    </Placemark>\n`;
    });

    kml += `  </Document>\n`;
    kml += `</kml>`;
    return kml;
}

/**
 * Generates an Excel Workbook (.xlsx) with two sheets: Predios and Coordenadas UTM.
 */
export function generateExcelWorkbook(layers) {
    if (!layers || layers.length === 0) return null;

    const rowsPredios = [];
    const rowsCoords = [];

    layers.forEach((layer, idx) => {
        const feature = layer.feature || layer;
        if (!feature) return;

        const props = feature.properties || {};
        const cod = props.COD_PROMUVI || `PREDIO_${idx + 1}`;

        let centerLat = "", centerLng = "", utmEste = "", utmNorte = "";
        let nVertices = 0;

        let rings = [];
        if (layer.getLatLngs) {
            const raw = layer.getLatLngs();
            rings = Array.isArray(raw[0]) ? raw : [raw];
        } else if (feature.geometry && feature.geometry.type === 'Polygon') {
            rings = [feature.geometry.coordinates[0].map(pt => ({ lat: pt[1], lng: pt[0] }))];
        }

        if (rings.length > 0 && rings[0].length > 0) {
            let sumLat = 0, sumLng = 0, cCount = 0;
            rings[0].forEach(pt => {
                if (pt && typeof pt.lat === 'number' && typeof pt.lng === 'number') {
                    sumLat += pt.lat;
                    sumLng += pt.lng;
                    cCount++;
                }
            });
            if (cCount > 0) {
                centerLat = sumLat / cCount;
                centerLng = sumLng / cCount;
                const utmC = latLonToUTM(centerLat, centerLng);
                utmEste = utmC.x;
                utmNorte = utmC.y;
            }
        } else if (layer.getBounds) {
            const c = layer.getBounds().getCenter();
            centerLat = c.lat;
            centerLng = c.lng;
            const utmC = latLonToUTM(c.lat, c.lng);
            utmEste = utmC.x;
            utmNorte = utmC.y;
        }

        // Add centroid coordinate row
        rowsCoords.push({
            "COD_PROMUVI": cod,
            "TIPO_PUNTO": "CENTROIDE",
            "VERTICE_NUM": 0,
            "LATITUD": fmtNum(centerLat, 6),
            "LONGITUD": fmtNum(centerLng, 6),
            "UTM_ESTE": fmtNum(utmEste, 3),
            "UTM_NORTE": fmtNum(utmNorte, 3),
            "ZONA": "19S",
            "DATUM": "WGS84"
        });

        // Add each vertex coordinate
        rings.forEach(ring => {
            ring.forEach((latlng, vIdx) => {
                nVertices++;
                const utmV = latLonToUTM(latlng.lat, latlng.lng);
                rowsCoords.push({
                    "COD_PROMUVI": cod,
                    "TIPO_PUNTO": "VERTICE",
                    "VERTICE_NUM": vIdx + 1,
                    "LATITUD": fmtNum(latlng.lat, 6),
                    "LONGITUD": fmtNum(latlng.lng, 6),
                    "UTM_ESTE": fmtNum(utmV.x, 3),
                    "UTM_NORTE": fmtNum(utmV.y, 3),
                    "ZONA": "19S",
                    "DATUM": "WGS84"
                });
            });
        });

        // Add record to main summary sheet
        rowsPredios.push({
            "ITEM": idx + 1,
            "COD_PROMUVI": cod,
            "ESTADO_NOTIFICACION": props.ESTADO || "NO PRIORIZADO",
            "ASOCIACION": props.ASOCIACIÓN || props.ASOCIACION || "--",
            "MANZANA": props.MANZANA || "--",
            "LOTE": props.LOTE || "--",
            "DISTRITO": props.DISTRITO || "Moquegua",
            "N_CARPETA": props.NRO_CARPETA || props['Nº CARPETA'] || "--",
            "DNI_1": props.DNI_1 || "--",
            "TITULAR_1": props.TITULAR_1 || "--",
            "DNI_2": props.DNI_2 || "--",
            "TITULAR_2": props.TITULAR_2 || "--",
            "PARTIDA_REGISTRAL": props.PART_REG || "--",
            "PARTIDA_ANTECEDENTE": props.PART_ANT || "--",
            "ACTO_REGISTRAL": props.ACT_REG || "--",
            "OBSERVACIONES_REG": props.OBSERV || "--",
            "DIRECCION_REG": props.DIR_REG || "--",
            "AREA_M2": fmtNum(props.GEO_AREA || props.Shape_Area, 2),
            "PERIMETRO_M": fmtNum(props.GEO_LENGTH || props.Shape_Length, 2),
            "N_VERTICES": nVertices,
            "CENTRO_UTM_ESTE": fmtNum(utmEste, 3),
            "CENTRO_UTM_NORTE": fmtNum(utmNorte, 3),
            "CENTRO_LAT": fmtNum(centerLat, 6),
            "CENTRO_LNG": fmtNum(centerLng, 6)
        });
    });

    const wb = XLSX.utils.book_new();
    const wsPredios = XLSX.utils.json_to_sheet(rowsPredios);
    const wsCoords = XLSX.utils.json_to_sheet(rowsCoords);

    // Auto-adjust column widths
    const setColWidths = (ws, data) => {
        if (!data || data.length === 0) return;
        const keys = Object.keys(data[0]);
        ws['!cols'] = keys.map(k => {
            let maxLen = k.length;
            for (let i = 0; i < Math.min(data.length, 100); i++) {
                const val = data[i][k];
                if (val !== null && val !== undefined) {
                    maxLen = Math.max(maxLen, String(val).length);
                }
            }
            return { wch: Math.min(maxLen + 3, 40) };
        });
    };

    setColWidths(wsPredios, rowsPredios);
    setColWidths(wsCoords, rowsCoords);

    XLSX.utils.book_append_sheet(wb, wsPredios, "Padron_Notificaciones");
    XLSX.utils.book_append_sheet(wb, wsCoords, "Coordenadas_UTM_19S");

    return XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
}
