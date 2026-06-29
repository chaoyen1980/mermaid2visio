import puppeteer from 'puppeteer';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export interface GraphNode {
    id: string;
    x: number;
    y: number;
    width: number;
    height: number;
    text?: string;
    type?: string; 
    rounding?: number; // Corner radius
    url?: string; // Hyperlink
    style?: {
        fill?: string;
        stroke?: string;
        strokeWidth?: string;
        strokeDasharray?: string;
        color?: string;
        fontSize?: string;
        fontFamily?: string;
        fontWeight?: string;
        fontStyle?: string;
        textAlign?: string; // start, middle, end
    };
}

export interface GraphEdge {
    d: string;
    startId?: string; 
    endId?: string;
    arrowStart?: boolean;
    arrowEnd?: boolean;
    style?: {
        stroke?: string;
        strokeWidth?: string;
        strokeDasharray?: string;
    };
}

export interface GraphCluster {
    id: string;
    x: number;
    y: number;
    width: number;
    height: number;
    text?: string;
    style?: {
        fill?: string;
        stroke?: string;
        strokeWidth?: string;
        strokeDasharray?: string;
        color?: string;
        fontSize?: string;
    };
}

export interface GraphLabel {
    x: number;
    y: number;
    width: number;
    height: number;
    text: string;
    style?: {
        color?: string;
        fill?: string;
        fontSize?: string;
        fontFamily?: string;
        fontWeight?: string;
        fontStyle?: string;
    };
}

export interface GraphData {
    nodes: GraphNode[];
    edges: GraphEdge[];
    clusters: GraphCluster[];
    labels: GraphLabel[];
    width: number;
    height: number;
}

export async function parseMermaid(definition: string): Promise<GraphData> {
    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();

    // Inject Mermaid from node_modules
    const mermaidPath = path.resolve(__dirname, '../node_modules/mermaid/dist/mermaid.min.js');
    await page.addScriptTag({ path: mermaidPath });

    await page.setContent(`
        <div id="graphDiv"></div>
        <script>
            mermaid.initialize({ startOnLoad: false });
        </script>
    `);

    try {
        const result = await page.evaluate(async (def) => {
            // @ts-ignore
            const { svg } = await mermaid.render('graphDiv', def);
            document.body.innerHTML = svg;

            const svgElement = document.querySelector('svg');
            const viewBox = svgElement?.getAttribute('viewBox')?.split(' ').map(parseFloat) || [0, 0, 0, 0];
            // viewBox: [minX, minY, width, height]. Offsets matter because sequence diagrams
            // emit elements at coordinates outside (0,0) and the bbox window starts at minX/minY.
            const vbX = viewBox[0] || 0;
            const vbY = viewBox[1] || 0;
            const graphWidth = viewBox[2] || parseFloat(svgElement?.getAttribute('width') || '0');
            const graphHeight = viewBox[3] || parseFloat(svgElement?.getAttribute('height') || '0');

            // -------- Sequence diagram branch --------
            // Mermaid emits sequenceDiagram with rect.actor + line.actor-line + line.messageLine0/1
            const actorTops = Array.from(document.querySelectorAll('rect.actor.actor-top')) as SVGGraphicsElement[];
            const lifelines = Array.from(document.querySelectorAll('line.actor-line')) as SVGLineElement[];
            const messageLines = Array.from(document.querySelectorAll('line.messageLine0, line.messageLine1')) as SVGLineElement[];

            if (actorTops.length > 0 && lifelines.length > 0) {
                // Build nodes from actor boxes (top + bottom), edges from message lines.
                const seqNodes: any[] = [];
                const seqEdges: any[] = [];
                const seqLabels: any[] = [];

                const shiftX = (x: number) => x - vbX;
                const shiftY = (y: number) => y - vbY;

                // Helper: locate the text inside an actor group
                const actorText = (rect: SVGGraphicsElement) => {
                    const parent = rect.parentElement;
                    if (!parent) return '';
                    const t = parent.querySelector('text.actor.actor-box, text.actor-box, text');
                    return t?.textContent?.trim() || '';
                };

                // Top actor heads — emit center coords because VsdxGenerator treats node.x/y as the shape center
                for (const rect of actorTops) {
                    const bb = rect.getBBox();
                    const id = (rect.parentElement?.id) || `actor-top-${seqNodes.length}`;
                    const styleR = window.getComputedStyle(rect);
                    const textEl = rect.parentElement?.querySelector('text.actor.actor-box, text.actor-box') as SVGGraphicsElement | null;
                    const styleT = textEl ? window.getComputedStyle(textEl) : null;
                    seqNodes.push({
                        id,
                        x: shiftX(bb.x) + bb.width / 2,
                        y: shiftY(bb.y) + bb.height / 2,
                        width: bb.width,
                        height: bb.height,
                        text: actorText(rect),
                        type: 'rect',
                        rounding: 0,
                        url: undefined,
                        style: {
                            fill: styleR.fill,
                            stroke: styleR.stroke,
                            strokeWidth: styleR.strokeWidth,
                            strokeDasharray: styleR.strokeDasharray,
                            color: styleT ? styleT.color : '#000000',
                            fontSize: styleT ? styleT.fontSize : undefined,
                            fontFamily: styleT ? styleT.fontFamily : undefined,
                            fontWeight: styleT ? styleT.fontWeight : 'bold',
                            fontStyle: styleT ? styleT.fontStyle : undefined,
                            textAlign: 'center',
                        },
                    });
                }

                // Bottom actor heads (mirror the top participants)
                const actorBottoms = Array.from(document.querySelectorAll('rect.actor.actor-bottom')) as SVGGraphicsElement[];
                for (const rect of actorBottoms) {
                    const bb = rect.getBBox();
                    const styleR = window.getComputedStyle(rect);
                    const textEl = rect.parentElement?.querySelector('text.actor.actor-box, text.actor-box') as SVGGraphicsElement | null;
                    const styleT = textEl ? window.getComputedStyle(textEl) : null;
                    seqNodes.push({
                        id: `actor-bot-${seqNodes.length}`,
                        x: shiftX(bb.x) + bb.width / 2,
                        y: shiftY(bb.y) + bb.height / 2,
                        width: bb.width,
                        height: bb.height,
                        text: actorText(rect),
                        type: 'rect',
                        rounding: 0,
                        url: undefined,
                        style: {
                            fill: styleR.fill,
                            stroke: styleR.stroke,
                            strokeWidth: styleR.strokeWidth,
                            strokeDasharray: styleR.strokeDasharray,
                            color: styleT ? styleT.color : '#000000',
                            fontSize: styleT ? styleT.fontSize : undefined,
                            fontFamily: styleT ? styleT.fontFamily : undefined,
                            fontWeight: styleT ? styleT.fontWeight : 'bold',
                            fontStyle: styleT ? styleT.fontStyle : undefined,
                            textAlign: 'center',
                        },
                    });
                }

                // Lifelines as dashed edges (no arrow)
                for (const ll of lifelines) {
                    const x1 = shiftX(parseFloat(ll.getAttribute('x1') || '0'));
                    const y1 = shiftY(parseFloat(ll.getAttribute('y1') || '0'));
                    const x2 = shiftX(parseFloat(ll.getAttribute('x2') || '0'));
                    const y2 = shiftY(parseFloat(ll.getAttribute('y2') || '0'));
                    const s = window.getComputedStyle(ll);
                    seqEdges.push({
                        d: `M ${x1} ${y1} L ${x2} ${y2}`,
                        arrowStart: false,
                        arrowEnd: false,
                        style: {
                            stroke: s.stroke || '#999999',
                            strokeWidth: s.strokeWidth || '1px',
                            strokeDasharray: s.strokeDasharray && s.strokeDasharray !== 'none' ? s.strokeDasharray : '3 3',
                        },
                    });
                }

                // Message arrows
                for (const ml of messageLines) {
                    const x1 = shiftX(parseFloat(ml.getAttribute('x1') || '0'));
                    const y1 = shiftY(parseFloat(ml.getAttribute('y1') || '0'));
                    const x2 = shiftX(parseFloat(ml.getAttribute('x2') || '0'));
                    const y2 = shiftY(parseFloat(ml.getAttribute('y2') || '0'));
                    const s = window.getComputedStyle(ml);
                    const dashed = (ml.getAttribute('class') || '').includes('messageLine1');
                    seqEdges.push({
                        d: `M ${x1} ${y1} L ${x2} ${y2}`,
                        arrowStart: false,
                        arrowEnd: true,
                        style: {
                            stroke: s.stroke || '#333333',
                            strokeWidth: s.strokeWidth || '1.5px',
                            strokeDasharray: dashed ? '4 4' : undefined,
                        },
                    });
                }

                // Message labels (sequence-number + text)
                const messageTexts = Array.from(document.querySelectorAll('text.messageText')) as SVGGraphicsElement[];
                for (const t of messageTexts) {
                    const text = t.textContent?.trim() || '';
                    if (!text) continue;
                    let bb: any;
                    try { bb = t.getBBox(); } catch { bb = { width: 80, height: 14 }; }
                    const x = parseFloat(t.getAttribute('x') || '0');
                    const y = parseFloat(t.getAttribute('y') || '0');
                    const st = window.getComputedStyle(t);
                    seqLabels.push({
                        x: shiftX(x) - bb.width / 2,
                        y: shiftY(y) - bb.height,
                        width: bb.width,
                        height: bb.height,
                        text,
                        style: {
                            color: st.color,
                            fill: 'none',
                            fontSize: st.fontSize,
                            fontFamily: st.fontFamily,
                            fontWeight: st.fontWeight,
                            fontStyle: st.fontStyle,
                        },
                    });
                }

                return {
                    width: graphWidth,
                    height: graphHeight,
                    nodes: seqNodes,
                    edges: seqEdges,
                    clusters: [],
                    labels: seqLabels,
                };
            }
            // -------- End sequence diagram branch --------

            const nodes = Array.from(document.querySelectorAll('.node'));
            // Select paths that are direct children of .edgePaths or have the flowchart-link class
            const edges = Array.from(document.querySelectorAll('.edgePaths path'));
            
            // Clusters (Subgraphs)
            // Render the title back into the cluster shape itself (Visio handles top-align natively).
            // Use the rect's actual bbox (which Mermaid sizes to include the title at top).
            const clusters = Array.from(document.querySelectorAll('.cluster')).map(cluster => {
                const id = cluster.id;
                const rect = cluster.querySelector('rect, polygon, path');
                const bbox = rect ? (rect as SVGGraphicsElement).getBBox() : { width: 0, height: 0, x: 0, y: 0 };

                const transform = cluster.getAttribute('transform');
                const match = /translate\(([^,]+),([^)]+)\)/.exec(transform || '');
                let x = match ? parseFloat(match[1]) : 0;
                let y = match ? parseFloat(match[2]) : 0;

                let rectX = x + bbox.x;
                let rectY = y + bbox.y;
                let width = bbox.width;
                let height = bbox.height;

                const textEl = cluster.querySelector('.nodeLabel, foreignObject, text');
                const text = textEl?.textContent?.trim() || '';
                const textStyle = textEl ? window.getComputedStyle(textEl) : null;
                const computedStyle = window.getComputedStyle(rect || cluster);

                return {
                    id,
                    x: rectX,
                    y: rectY,
                    width,
                    height,
                    text,
                    style: {
                        fill: computedStyle.fill,
                        stroke: computedStyle.stroke,
                        strokeWidth: computedStyle.strokeWidth,
                        strokeDasharray: computedStyle.strokeDasharray,
                        color: textStyle ? textStyle.color : '#000000',
                        fontSize: textStyle ? textStyle.fontSize : undefined,
                    }
                };
            });

            // Edge Labels
            const labels = Array.from(document.querySelectorAll('.edgeLabel')).map(label => {
                // Mermaid puts labels in a div inside a foreignObject, or directly as text
                // The .edgeLabel group usually has the transform
                const transform = label.getAttribute('transform');
                const match = /translate\(([^,]+),([^)]+)\)/.exec(transform || '');
                const x = match ? parseFloat(match[1]) : 0;
                const y = match ? parseFloat(match[2]) : 0;

                const div = label.querySelector('div, foreignObject, text');
                const bbox = div ? (div as SVGGraphicsElement).getBBox() : { width: 0, height: 0 };
                const text = label.textContent?.trim() || '';
                
                // Attempt to get background color if there's a label box
                const bgRect = label.querySelector('.label-container rect'); // Sometimes mermaid wraps it
                const bgStyle = bgRect ? window.getComputedStyle(bgRect) : null;
                const textStyle = window.getComputedStyle(div || label);

                return {
                    x,
                    y,
                    width: bbox.width || 10, // Fallback
                    height: bbox.height || 10,
                    text,
                    style: {
                        color: textStyle.color,
                        fill: bgStyle?.fill !== 'none' ? bgStyle?.fill : undefined,
                        fontSize: textStyle.fontSize,
                        fontFamily: textStyle.fontFamily,
                        fontWeight: textStyle.fontWeight,
                        fontStyle: textStyle.fontStyle
                    }
                };
            }).filter(l => l.text); // Filter empty labels

            return {
                width: graphWidth,
                height: graphHeight,
                nodes: nodes.map(node => {
                    const id = node.id;
                    const nodeClasses = Array.from(node.classList).join(' ');
                    const transform = node.getAttribute('transform');
                    const match = /translate\(([^,]+),([^)]+)\)/.exec(transform || '');
                    const x = match ? parseFloat(match[1]) : 0;
                    const y = match ? parseFloat(match[2]) : 0;
                    
                    const rect = node.querySelector('rect, circle, polygon, path, ellipse') as SVGGraphicsElement;
                    const bbox = rect ? rect.getBBox() : { width: 0, height: 0, x: 0, y: 0 };
                    
                    // Infer shape type from classes and tags
                    let type = 'rectangle';
                    const shapeTag = rect ? rect.tagName.toLowerCase() : 'unknown';
                    const points = rect ? rect.getAttribute('points') || '' : '';
                    const d = rect ? rect.getAttribute('d') || '' : '';
                    const dataShape = node.getAttribute('data-shape') || rect?.getAttribute('data-shape') || '';
                    
                    if (dataShape) {
                        type = dataShape;
                    } else if (nodeClasses.includes('stadium')) {
                        type = 'stadium';
                    } else if (shapeTag === 'polygon') {
                        const pts = points.split(/[\s,]+/).filter(p => p).length / 2;
                        if (pts === 4) {
                             type = 'diamond'; 
                        } else if (pts > 4) {
                             type = 'subroutine';
                        }
                    } else if (shapeTag === 'path') {
                        if (d.includes('a') || d.includes('A')) type = 'cylinder';
                        else if (d.includes('c') || d.includes('C')) type = 'stadium';
                    } else if (shapeTag === 'circle') {
                        type = 'circle';
                    } else if (shapeTag === 'ellipse') {
                        type = 'ellipse';
                    }

                    // Extract styles
                    const computedStyle = window.getComputedStyle(rect || node);
                    const textEl = node.querySelector('div, span, text');
                    const textStyle = textEl ? window.getComputedStyle(textEl) : null;
                    
                    // Rounding (rx/ry)
                    let rounding = 0;
                    if (rect && rect.tagName.toLowerCase() === 'rect') {
                        const rx = parseFloat(rect.getAttribute('rx') || '0');
                        if (rx > 0) rounding = rx;
                    }

                    // Hyperlink
                    const anchor = node.querySelector('a') || node.closest('a');
                    const url = anchor ? (anchor.getAttribute('href') || anchor.getAttribute('xlink:href') || undefined) : undefined;

                    // Improved text extraction (preserve breaks)
                    let text = '';
                    if (textEl) {
                        // Clone to not mess up DOM
                        const clone = textEl.cloneNode(true) as HTMLElement;
                        // Replace <br> with newlines
                        const brs = clone.querySelectorAll('br');
                        brs.forEach(br => br.replaceWith('\n'));
                        text = clone.textContent?.trim() || '';
                    } else {
                        text = node.textContent?.trim() || '';
                    }

                    // Text Alignment
                    let textAlign = 'center';
                    if (textStyle) {
                        const anchor = textStyle.getPropertyValue('text-anchor');
                        if (anchor === 'start') textAlign = 'left';
                        else if (anchor === 'end') textAlign = 'right';
                    }

                    // Ensure shape height fits text. Mermaid sometimes sizes shapes tightly to
                    // the text bounding box, but Visio's vertical centering can clip multi-line
                    // text. Expand height to fit (lineCount * fontSize * lineHeight + padding).
                    const fontPx = textStyle ? parseFloat(textStyle.fontSize || '14') : 14;
                    const lineCount = Math.max(1, text.split('\n').length);
                    const neededHeight = lineCount * fontPx * 1.3 + 8;
                    const finalHeight = Math.max(bbox.height, neededHeight);

                    return {
                        id,
                        x,
                        y,
                        width: bbox.width,
                        height: finalHeight,
                        text,
                        type,
                        rounding,
                        url,
                        style: {
                            fill: computedStyle.fill,
                            stroke: computedStyle.stroke,
                            strokeWidth: computedStyle.strokeWidth,
                            strokeDasharray: computedStyle.strokeDasharray,
                            color: textStyle ? textStyle.color : '#000000',
                            fontSize: textStyle ? textStyle.fontSize : undefined,
                            fontFamily: textStyle ? textStyle.fontFamily : undefined,
                            fontWeight: textStyle ? textStyle.fontWeight : undefined,
                            fontStyle: textStyle ? textStyle.fontStyle : undefined,
                            textAlign
                        }
                    };
                }),
                edges: edges.map(path => {
                   const computedStyle = window.getComputedStyle(path);
                   
                   // Arrow detection
                   const markerStart = path.getAttribute('marker-start');
                   const markerEnd = path.getAttribute('marker-end');

                   // ID Extraction (Start/End Nodes)
                   // Mermaid v10+ puts classes on the group wrapping the path: <g class="edgePaths ... LS-startId LE-endId">
                   let startId, endId;
                   const parentGroup = path.parentElement;
                   if (parentGroup) {
                       const classes = Array.from(parentGroup.classList);
                       // Look for LS-* and LE-*
                       const ls = classes.find(c => c.startsWith('LS-'));
                       const le = classes.find(c => c.startsWith('LE-'));
                       if (ls) startId = ls.replace('LS-', '');
                       if (le) endId = le.replace('LE-', '');
                   }

                   return { 
                       d: path.getAttribute('d') || '',
                       startId,
                       endId,
                       arrowStart: !!markerStart, 
                       arrowEnd: !!markerEnd,
                       style: {
                           stroke: computedStyle.stroke,
                           strokeWidth: computedStyle.strokeWidth,
                           strokeDasharray: computedStyle.strokeDasharray
                       }
                   };
                }),
                clusters: clusters.map(c => {
                    // Re-map cluster styles to include dasharray
                    return {
                        ...c,
                        style: {
                            ...c.style,
                            strokeDasharray: c.style.strokeDasharray // Was missing in previous replacement
                        }
                    };
                }),
                labels
            };
        }, definition);

        // -------- Post-process: de-overlap labels against nodes and each other --------
        // Mermaid edge labels and flowchart nodes use (x, y) = center of bbox.
        // Walk labels in order; if a label's bbox overlaps any node or already-placed label,
        // shift it vertically (down first, then up) until it no longer collides.
        if (result.labels && result.labels.length) {
            const labelBBox = (l: any) => ({
                x1: l.x - (l.width || 10) / 2,
                y1: l.y - (l.height || 10) / 2,
                x2: l.x + (l.width || 10) / 2,
                y2: l.y + (l.height || 10) / 2,
            });
            const nodeBBox = (n: any) => ({
                x1: n.x - n.width / 2,
                y1: n.y - n.height / 2,
                x2: n.x + n.width / 2,
                y2: n.y + n.height / 2,
            });
            const overlap = (a: any, b: any) =>
                !(a.x2 < b.x1 || b.x2 < a.x1 || a.y2 < b.y1 || b.y2 < a.y1);

            const obstacles: Array<{ x1: number, y1: number, x2: number, y2: number }> =
                (result.nodes || []).map(nodeBBox);
            const placed: typeof obstacles = [];

            for (const lb of result.labels as any[]) {
                const tryShift = (deltaY: number): boolean => {
                    const candidate = { ...lb, y: lb.y + deltaY };
                    const bb = labelBBox(candidate);
                    for (const o of obstacles) if (overlap(bb, o)) return false;
                    for (const p of placed) if (overlap(bb, p)) return false;
                    lb.y = candidate.y;
                    placed.push(bb);
                    return true;
                };
                let resolved = false;
                // Try original position first
                if (tryShift(0)) { resolved = true; }
                if (!resolved) {
                    // Try increasing offsets, alternating down/up
                    for (let step = 1; step <= 20 && !resolved; step++) {
                        const dy = step * ((lb.height || 12) + 4);
                        if (tryShift(dy)) { resolved = true; break; }
                        if (tryShift(-dy)) { resolved = true; break; }
                    }
                }
                if (!resolved) placed.push(labelBBox(lb));
            }
        }

        return result;
    } finally {
        await browser.close();
    }
}
