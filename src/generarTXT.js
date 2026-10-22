import { FORMATO_INTEGRACION } from "./formatoIntegracion";

/**
 * Convierte una matriz de filas y columnas a texto TXT.
 *
 * Opciones de generarTXT(filas, opciones):
 *  - alinear:         false (por defecto) -> formato compacto, sin espacios extra.
 *                     true  -> columnas con ancho fijo (ordenado a la vista).
 *  - separador:       carácter que separa las columnas
 *  - lineaSeparadora: true -> agrega una línea "-----" bajo el encabezado
 *                     (solo aplica cuando alinear = true)
 *  - espaciarSeparador: true (por defecto) -> " ; " con espacios alrededor
 *                     false -> el separador pegado, sin espacios extra
 *                     (solo aplica cuando alinear = true)
 *  - saltoLinea:      "\r\n" o "\n"
 *  - saltoFinal:      true -> salto de línea después de la última fila
 *
 * Los valores por defecto salen de FORMATO_INTEGRACION.
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
  {
    alinear = false,
    separador = FORMATO_INTEGRACION.separador,
    lineaSeparadora = true,
    espaciarSeparador = true,
    saltoLinea = FORMATO_INTEGRACION.saltoLinea,
    saltoFinal = FORMATO_INTEGRACION.saltoFinal,
  } = {}
) {
  if (!Array.isArray(filas) || filas.length === 0) {
    throw new Error("No hay datos para generar el TXT.");
  }

  const limpias = filas.map((fila) =>
    fila.map((valor) => limpiarValor(valor, separador))
  );

  // Formato compacto: sin espacios extra
  if (!alinear) {
    const texto = limpias.map((fila) => fila.join(separador)).join(saltoLinea);
    return saltoFinal ? texto + saltoLinea : texto;
  }

  // Formato alineado: cada columna con el mismo ancho en todas las filas
  const numColumnas = Math.max(...limpias.map((fila) => fila.length));
  const anchos = Array.from({ length: numColumnas }, (_, i) =>
    Math.max(...limpias.map((fila) => (fila[i] ?? "").length))
  );

  const union = espaciarSeparador ? ` ${separador} ` : separador;

  const lineas = limpias.map((fila) =>
    anchos
      .map((ancho, i) => (fila[i] ?? "").padEnd(ancho))
      .join(union)
      .trimEnd()
  );

  if (lineaSeparadora && lineas.length > 1) {
    const guiones = anchos.map((ancho) => "-".repeat(Math.max(ancho, 1)));
    lineas.splice(1, 0, guiones.join("-+-"));
  }

  const texto = lineas.join(saltoLinea);
  return saltoFinal ? texto + saltoLinea : texto;
}

export function descargarTXT(
  contenido,
  nombreArchivo = "nomina.txt",
  { bom = FORMATO_INTEGRACION.bom } = {}
) {
  const partes = bom ? ["\uFEFF", contenido] : [contenido];

  const blob = new Blob(partes, {
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