
/**
 * Convierte una matriz de filas y columnas a texto TXT.
 *
 * NOTA:
 * Este es un formato de prueba separado por punto y coma.
 * Todavía debemos ajustarlo al formato real de integración
 * de nómina.
 */

export function generarTXT(filas) {
  if (!Array.isArray(filas) || filas.length === 0) {
    throw new Error("No hay datos para generar el TXT.");
  }

  return filas
    .map((fila) =>
      fila
        .map((valor) => {
          if (valor === null || valor === undefined) {
            return "";
          }

          return String(valor)
            .replace(/\r?\n/g, " ")
            .replace(/;/g, ",")
            .trim();
        })
        .join(";")
    )
    .join("\r\n");
}

export function descargarTXT(contenido, nombreArchivo = "nomina.txt") {
  const blob = new Blob(
    ["\uFEFF", contenido],
    { type: "text/plain;charset=utf-8" }
  );

  const url = URL.createObjectURL(blob);
  const enlace = document.createElement("a");

  enlace.href = url;
  enlace.download = nombreArchivo;

  document.body.appendChild(enlace);
  enlace.click();
  enlace.remove();

  URL.revokeObjectURL(url);
}
