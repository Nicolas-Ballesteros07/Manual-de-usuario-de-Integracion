/**
 * Formato del TXT de integración, en un solo lugar.
 *
 * Todos estos valores están PENDIENTES de confirmar con Ricardo.
 * Cuando responda, solo se cambia este objeto.
 *
 *  - separador:         carácter entre columnas
 *  - incluirEncabezado: true -> primera fila "tercero;contrato;notas;..."
 *  - bom:               true -> agrega el BOM UTF-8 al inicio del archivo
 *  - saltoLinea:        "\r\n" (Windows) o "\n"
 *  - saltoFinal:        true -> deja un salto de línea tras la última fila
 */
export const FORMATO_INTEGRACION = {
  separador: ";",
  incluirEncabezado: true,
  bom: false,
  saltoLinea: "\r\n",
  saltoFinal: true,
};