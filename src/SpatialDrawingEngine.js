/**
 * SpatialDrawingEngine.js
 * Interactive spatial selection with Ray-Casting Point-in-Polygon (PIP) for Leaflet.
 */
import L from 'leaflet';

/**
 * Determines whether a 2D point (x, y) lies inside a polygon ring using Ray-Casting.
 * @param {Object} point - Point object { x: lng, y: lat }.
 * @param {Array<Object>} vs - Array of polygon vertices [{ x: lng, y: lat }, ...].
 * @returns {boolean} True if inside the polygon.
 */
export function pointInPolygon(point, vs) {
    const x = point.x, y = point.y;
    let inside = false;
    for (let i = 0, j = vs.length - 1; i < vs.length; j = i++) {
        const xi = vs[i].x, yi = vs[i].y;
        const xj = vs[j].x, yj = vs[j].y;
        const intersect = ((yi > y) !== (yj > y)) &&
            (x < (xj - xi) * (y - yi) / (yj - yi) + xi);
        if (intersect) inside = !inside;
    }
    return inside;
}

/**
 * Checks if a Leaflet LatLng point lies inside a polygon formed by an array of LatLngs.
 */
export function isLatLngInsidePolygon(latlng, polygonLatLngs) {
    const pt = { x: latlng.lng, y: latlng.lat };
    const vs = polygonLatLngs.map(l => ({ x: l.lng, y: l.lat }));
    return pointInPolygon(pt, vs);
}

/**
 * Evaluates whether a GeoJSON feature (Point, Polygon, MultiPolygon) intersects or is contained within a selection polygon.
 */
export function isFeatureInsideSelection(featureLayer, selectionPolygonLatLngs) {
    const vs = selectionPolygonLatLngs.map(l => ({ x: l.lng, y: l.lat }));

    // 1. Check feature centroid if available
    if (featureLayer.getBounds) {
        const center = featureLayer.getBounds().getCenter();
        if (pointInPolygon({ x: center.lng, y: center.lat }, vs)) {
            return true;
        }
    }

    // 2. Check individual geometry vertices (prefer actual shifted layer coordinates)
    if (featureLayer.getLatLngs) {
        const raw = featureLayer.getLatLngs();
        const rings = Array.isArray(raw[0]) ? (Array.isArray(raw[0][0]) ? raw : [raw[0]]) : [raw];
        for (const ring of rings) {
            for (const pt of ring) {
                if (pt && typeof pt.lng === 'number' && typeof pt.lat === 'number') {
                    if (pointInPolygon({ x: pt.lng, y: pt.lat }, vs)) {
                        return true;
                    }
                }
            }
        }
    }

    const geom = featureLayer.feature ? featureLayer.feature.geometry : featureLayer.geometry;
    if (!geom) return false;

    if (geom.type === 'Point') {
        return pointInPolygon({ x: geom.coordinates[0], y: geom.coordinates[1] }, vs);
    }

    if (geom.type === 'Polygon' || geom.type === 'MultiPolygon') {
        const rings = geom.type === 'Polygon' ? [geom.coordinates[0]] : geom.coordinates.map(p => p[0]);
        for (const ring of rings) {
            for (const pt of ring) {
                if (pointInPolygon({ x: pt[0], y: pt[1] }, vs)) {
                    return true;
                }
            }
        }
    }

    return false;
}

export class LeafletSpatialDrawer {
    constructor(map, callbacks = {}) {
        this.map = map;
        this.isDrawing = false;
        this.drawPoints = [];
        this.drawMarkers = [];
        this.drawPolyline = null;
        this.drawPolygon = null;
        this.callbacks = callbacks; // onPointAdded, onComplete, onClear

        this.handleMapClick = this.handleMapClick.bind(this);
        this.handleDoubleClick = this.handleDoubleClick.bind(this);
    }

    startDrawing(onComplete) {
        this.clearDrawing();
        this.isDrawing = true;
        this.onCompleteCallback = onComplete;

        this.map.doubleClickZoom.disable();
        this.map.getContainer().style.cursor = 'crosshair';

        this.map.on('click', this.handleMapClick);
        this.map.on('dblclick', this.handleDoubleClick);
    }

    handleMapClick(e) {
        if (!this.isDrawing) return;

        const latlng = e.latlng;
        this.drawPoints.push(latlng);

        const marker = L.circleMarker(latlng, {
            radius: 5,
            color: '#ffffff',
            weight: 2,
            fillColor: '#38bdf8',
            fillOpacity: 1,
            interactive: false
        }).addTo(this.map);
        this.drawMarkers.push(marker);

        if (this.drawPolyline) this.map.removeLayer(this.drawPolyline);
        this.drawPolyline = L.polyline(this.drawPoints, {
            color: '#38bdf8',
            weight: 2,
            dashArray: '6,6',
            interactive: false
        }).addTo(this.map);

        if (this.drawPoints.length >= 3) {
            if (this.drawPolygon) this.map.removeLayer(this.drawPolygon);
            this.drawPolygon = L.polygon(this.drawPoints, {
                color: '#38bdf8',
                weight: 2,
                fillColor: '#38bdf8',
                fillOpacity: 0.25,
                interactive: false
            }).addTo(this.map);
        }

        if (this.callbacks.onPointAdded) {
            this.callbacks.onPointAdded(this.drawPoints.length);
        }
    }

    handleDoubleClick(e) {
        L.DomEvent.stop(e);
        if (this.isDrawing && this.drawPoints.length >= 3) {
            this.finishDrawing();
        }
    }

    finishDrawing() {
        this.isDrawing = false;
        this.map.off('click', this.handleMapClick);
        this.map.off('dblclick', this.handleDoubleClick);

        this.map.doubleClickZoom.enable();
        this.map.getContainer().style.cursor = '';

        if (this.onCompleteCallback) {
            this.onCompleteCallback(this.drawPoints);
        }
        if (this.callbacks.onComplete) {
            this.callbacks.onComplete(this.drawPoints);
        }
    }

    clearDrawing() {
        this.isDrawing = false;
        this.map.off('click', this.handleMapClick);
        this.map.off('dblclick', this.handleDoubleClick);

        if (this.drawPolyline) this.map.removeLayer(this.drawPolyline);
        if (this.drawPolygon) this.map.removeLayer(this.drawPolygon);
        this.drawMarkers.forEach(m => this.map.removeLayer(m));

        this.drawPoints = [];
        this.drawMarkers = [];
        this.drawPolyline = null;
        this.drawPolygon = null;

        this.map.doubleClickZoom.enable();
        this.map.getContainer().style.cursor = '';

        if (this.callbacks.onClear) {
            this.callbacks.onClear();
        }
    }
}
