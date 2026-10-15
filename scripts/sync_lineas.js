const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const fetch = require('node-fetch');

// --- CONFIGURACIÓN ---
const EXCEL_URL = 'https://drive.google.com/uc?export=download&id=1iDn9HZVd46P8ZGFpNX1vl9s8L7Jk-RKf';
const CARPETAS_IGNORADAS = ['scripts', 'node_modules', '.git', '.github', 'pedido', 'tutorial'];

// 🔧 Función para normalizar códigos (quitar espacios, puntos, pasar a mayúsculas)
function normalizarCodigo(codigo) {
    if (!codigo) return '';
    return String(codigo)
        .trim()                    // Quita espacios al inicio y final
        .replace(/\s+/g, '')       // Quita TODOS los espacios internos
        .replace(/\./g, '')        // Quita puntos
        .replace(/-/g, '')         // Quita guiones
        .toUpperCase();            // Todo a mayúsculas
}

async function sincronizarLineas() {
    try {
        // 1. Descargar el Excel
        console.log('Descargando Excel desde Google Drive...');
        const response = await fetch(EXCEL_URL);
        if (!response.ok) throw new Error(`Error al descargar: ${response.statusText}`);
        
        const arrayBuffer = await response.arrayBuffer();
        const buffer = Buffer.from(arrayBuffer);

        const workbook = XLSX.read(buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        console.log(`Productos en el Excel: ${data.length}`);

        // 2. Crear diccionario de productos del Excel (usando código normalizado como clave)
        const excelDataByCode = {};
        data.forEach(prod => {
            if (prod.c) {
                const keyNormalizada = normalizarCodigo(prod.c);
                excelDataByCode[keyNormalizada] = prod;
            }
        });
        console.log(`Productos indexados por código: ${Object.keys(excelDataByCode).length}\n`);

        // 3. Detectar todas las carpetas
        const raiz = path.join(__dirname, '..');
        const carpetas = fs.readdirSync(raiz).filter(item => {
            const rutaCompleta = path.join(raiz, item);
            return fs.statSync(rutaCompleta).isDirectory() && !CARPETAS_IGNORADAS.includes(item);
        });

        console.log(`Carpetas de líneas encontradas: ${carpetas.join(', ')}\n`);

        let totalActualizados = 0;
        let totalSinCambio = 0;
        let totalNoEncontrados = 0;

        // 4. Procesar cada carpeta
        for (const carpeta of carpetas) {
            const carpetaPath = path.join(raiz, carpeta, 'index.html');
            
            if (!fs.existsSync(carpetaPath)) {
                console.warn(`⚠️ No existe: ${carpeta}/index.html. Saltando...`);
                continue;
            }

            let html = fs.readFileSync(carpetaPath, 'utf8');
            const regex = /(<script id="product-data" type="application\/json">)(.*?)(<\/script>)/s;
            const match = html.match(regex);

            if (!match) {
                console.warn(`⚠️ No se encontró 'product-data' en ${carpeta}/index.html. Saltando...`);
                continue;
            }

            const productosHTML = JSON.parse(match[2].trim());
            
            let actualizados = 0;
            let sinCambio = 0;
            let noEncontrados = 0;

            // 5. Actualizar SOLO el precio (usando código normalizado)
            const productosActualizados = productosHTML.map(prodHTML => {
                const codigoNormalizado = normalizarCodigo(prodHTML.c);
                const prodExcel = excelDataByCode[codigoNormalizado];

                if (prodExcel) {
                    const precioNuevo = prodExcel.p;
                    const precioViejo = prodHTML.p;

                    if (precioViejo !== precioNuevo) {
                        actualizados++;
                        return {
                            ...prodHTML,
                            p: precioNuevo
                        };
                    } else {
                        sinCambio++;
                        return prodHTML;
                    }
                } else {
                    noEncontrados++;
                    return prodHTML;
                }
            });

            // 6. Solo escribir si hubo cambios
            if (actualizados > 0) {
                const jsonProductos = JSON.stringify(productosActualizados);
                html = html.replace(regex, `$1${jsonProductos}$3`);
                fs.writeFileSync(carpetaPath, html);
            }

            console.log(`📁 ${carpeta} (${productosHTML.length} productos en HTML)`);
            console.log(`   ✅ Precios actualizados: ${actualizados}`);
            console.log(`   ➖ Sin cambios: ${sinCambio}`);
            console.log(`   ⚠️ No encontrados en Excel: ${noEncontrados}\n`);

            totalActualizados += actualizados;
            totalSinCambio += sinCambio;
            totalNoEncontrados += noEncontrados;
        }

        console.log(`\n🎉 ¡Proceso terminado!`);
        console.log(`   - Total precios actualizados: ${totalActualizados}`);
        console.log(`   - Total productos sin cambios: ${totalSinCambio}`);
        console.log(`   - Total productos en HTML no encontrados en Excel: ${totalNoEncontrados}`);

    } catch (error) {
        console.error('❌ Error crítico:', error);
        process.exit(1);
    }
}

sincronizarLineas();