const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const RAIZ = path.join(__dirname, '..');
const ARCHIVO_SALIDA = path.join(RAIZ, 'productos_lineas_extraidos.xlsx');

// Carpetas que NO son líneas
const CARPETAS_IGNORADAS = ['scripts', 'node_modules', '.git', '.github', 'pedido', 'tutorial', 'original'];

function extraerProductos() {
    console.log('Extrayendo productos del Home 2 (líneas)...');
    let todosLosProductos = [];

    // 1. Detectar todas las carpetas
    const carpetas = fs.readdirSync(RAIZ).filter(item => {
        const rutaCompleta = path.join(RAIZ, item);
        return fs.statSync(rutaCompleta).isDirectory() && !CARPETAS_IGNORADAS.includes(item);
    });

    console.log(`Carpetas encontradas: ${carpetas.join(', ')}\n`);

    carpetas.forEach(linea => {
        const archivoHTML = path.join(RAIZ, linea, 'index.html');
        
        if (fs.existsSync(archivoHTML)) {
            const html = fs.readFileSync(archivoHTML, 'utf8');
            const regex = /<script id="product-data" type="application\/json">(.*?)<\/script>/s;
            const match = html.match(regex);

            if (match && match[1]) {
                try {
                    const productos = JSON.parse(match[1].trim());
                    
                    // Filtrar para quitar la imagen (base64 es demasiado larga)
                    const productosLimpios = productos.map(p => {
                        const productoFiltrado = { linea: linea };
                        
                        for (const [key, value] of Object.entries(p)) {
                            // Si es la imagen base64, la omitimos
                            if (key === 'i' && typeof value === 'string' && value.length > 30000) {
                                continue;
                            }
                            productoFiltrado[key] = value;
                        }
                        return productoFiltrado;
                    });

                    todosLosProductos = todosLosProductos.concat(productosLimpios);
                    console.log(`✅ ${linea}: ${productos.length} productos extraídos.`);
                } catch (e) {
                    console.error(`❌ Error al parsear JSON en ${linea}/index.html:`, e.message);
                }
            } else {
                console.warn(`⚠️ No se encontró 'product-data' en ${linea}/index.html`);
            }
        }
    });

    if (todosLosProductos.length === 0) {
        console.log('No se encontraron productos.');
        return;
    }

    // Crear el Excel
    const hoja = XLSX.utils.json_to_sheet(todosLosProductos);
    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, 'Productos_Lineas');

    XLSX.writeFile(libro, ARCHIVO_SALIDA);

    console.log(`\n🎉 ¡Éxito! Se extrajeron ${todosLosProductos.length} productos de líneas.`);
    console.log(`📁 Archivo creado: ${ARCHIVO_SALIDA}`);
}

extraerProductos();