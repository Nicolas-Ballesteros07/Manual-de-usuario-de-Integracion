/**
 * Convierte una matriz de filas y columnas a texto TXT.
 *
 * Opciones de generarTXT(filas, opciones):
 *  - alinear:         true  -> columnas con ancho fijo (se lee ordenado, ideal para revisar)
 *                     false -> formato compacto, sin espacios extra (ideal para importar)
 *  - separador:       carácter que separa las columnas (por defecto ";")
 *  - lineaSeparadora: true  -> agrega una línea "-----" bajo el encabezado
 *                     (solo aplica cuando alinear = true)
 *
 * NOTA: Este es un formato de prueba. Todavía debemos ajustarlo
 * al formato real de integración de nómina.
 */

function limpiarValor(valor, separador) {
  if (valor === null || valor === undefined) return "";

  return String(valor)
    .replace(/\r?\n/g, " ")
    .split(separador)
    .join(",")
    .trim();
}

export function generarTXT(
  filas,
  { alinear = true, separador = ";", lineaSeparadora = true } = {}
) {
  if (!Array.isArray(filas) || filas.length === 0) {
    throw new Error("No hay datos para generar el TXT.");
  }

  const limpias = filas.map((fila) =>
    fila.map((valor) => limpiarValor(valor, separador))
  );

  // Formato compacto (el original): sin espacios de relleno
  if (!alinear) {
    return limpias.map((fila) => fila.join(separador)).join("\r\n");
  }

  // Ancho de cada columna = el valor más largo de esa columna
  const numColumnas = Math.max(...limpias.map((fila) => fila.length));
  const anchos = Array.from({ length: numColumnas }, (_, i) =>
    Math.max(...limpias.map((fila) => (fila[i] ?? "").length))
  );

  const union = ` ${separador} `;

  const lineas = limpias.map((fila) =>
    anchos
      .map((ancho, i) => (fila[i] ?? "").padEnd(ancho))
      .join(union)
      .trimEnd()
  );

  // Línea separadora debajo del encabezado
  if (lineaSeparadora && lineas.length > 1) {
    const guiones = anchos.map((ancho) => "-".repeat(Math.max(ancho, 1)));
    lineas.splice(1, 0, guiones.join("-+-"));
  }

  return lineas.join("\r\n");
}

export function descargarTXT(contenido, nombreArchivo = "nomina.txt") {
  const blob = new Blob(["\uFEFF", contenido], {
    type: "text/plain;charset=utf-8",
  });

  const url = URL.createObjectURL(blob);
  const enlace = document.createElement("a");

  enlace.href = url;
  enlace.download = nombreArchivo;

  document.body.appendChild(enlace);
  enlace.click();
  enlace.remove();

  URL.revokeObjectURL(url);
}