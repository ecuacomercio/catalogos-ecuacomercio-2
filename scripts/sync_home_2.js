const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const fetch = require('node-fetch');

// --- CONFIGURACIÓN (Excel de Líneas) ---
const EXCEL_URL = 'https://drive.google.com/uc?export=download&id=1iDn9HZVd46P8ZGFpNX1vl9s8L7Jk-RKf';
const HOME_HTML_PATH = path.join(__dirname, '..', 'index.html');

// 🔧 Función para normalizar códigos (evita errores por espacios, puntos o guiones)
function normalizarCodigo(codigo) {
    if (!codigo) return '';
    return String(codigo)
        .trim()
        .replace(/\s+/g, '')
        .replace(/\./g, '')
        .replace(/-/g, '')
        .toUpperCase();
}

async function actualizarHomeLineas() {
    try {
        // 1. Leer el HTML actual del Home de Líneas para extraer el search-index existente
        console.log('Leyendo index.html del Home de Líneas...');
        if (!fs.existsSync(HOME_HTML_PATH)) {
            throw new Error(`No se encontró el archivo: ${HOME_HTML_PATH}`);
        }
        let html = fs.readFileSync(HOME_HTML_PATH, 'utf8');

        const regex = /(<script id="search-index" type="application\/json">)(.*?)(<\/script>)/s;
        const match = html.match(regex);

        if (!match) throw new Error('No se encontró el script search-index en index.html');

        const oldSearchIndex = JSON.parse(match[2].trim());
        console.log(`Productos actuales en el search-index del Home: ${oldSearchIndex.length}`);

        // Diccionario por código normalizado para búsqueda rápida
        const oldDataByCode = {};
        oldSearchIndex.forEach(p => {
            if (p.c) {
                oldDataByCode[normalizarCodigo(p.c)] = p;
            }
        });

        // 2. Descargar el Excel de Líneas desde Google Drive
        console.log('Descargando Excel de líneas...');
        const response = await fetch(EXCEL_URL);
        if (!response.ok) throw new Error(`Error al descargar: ${response.statusText}`);
        
        const arrayBuffer = await response.arrayBuffer();
        const buffer = Buffer.from(arrayBuffer);

        const workbook = XLSX.read(buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        console.log(`Productos encontrados en el Excel: ${data.length}`);

        // 3. Fusionar: Mantener la estructura intacta y actualizar SOLO el precio si el código coincide
        let actualizados = 0;
        let sinCambios = 0;
        let noEncontradosEnHtml = 0;

        const newSearchIndex = oldSearchIndex.map(oldProd => {
            const codigoNorm = normalizarCodigo(oldProd.c);
            // Buscar el producto correspondiente en el Excel usando el código normalizado
            const prodExcel = data.find(p => normalizarCodigo(p.c) === codigoNorm);

            if (prodExcel) {
                const precioNuevo = prodExcel.p;
                
                // Si el precio cambió, lo actualizamos conservando todo lo demás intacto
                if (oldProd.p !== precioNuevo) {
                    actualizados++;
                    return {
                        ...oldProd,
                        p: precioNuevo
                    };
                } else {
                    sinCambios++;
                    return oldProd; // Sin cambios, se devuelve tal cual
                }
            } else {
                noEncontradosEnHtml++;
                return oldProd; // Si no está en el Excel, se queda tal como estaba en el HTML
            }
        });

        // 4. Escribir el nuevo search-index en el HTML del Home
        const jsonSearch = JSON.stringify(newSearchIndex);
        html = html.replace(regex, `$1${jsonSearch}$3`);
        fs.writeFileSync(HOME_HTML_PATH, html);

        console.log(`\n🎉 ¡Sincronización de precios completada!`);
        console.log(`   - Precios actualizados: ${actualizados}`);
        console.log(`   - Productos sin cambios de precio: ${sinCambios}`);
        console.log(`   - Productos en el HTML no encontrados en el Excel: ${noEncontradosEnHtml}`);
        console.log(`   - Total en el search-index: ${newSearchIndex.length}`);

    } catch (error) {
        console.error('❌ Error crítico:', error);
        process.exit(1);
    }
}

actualizarHomeLineas();