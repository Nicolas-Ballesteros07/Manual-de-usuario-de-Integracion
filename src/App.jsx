import { useMemo, useState } from "react";
import JSZip from "jszip";
import "./App.css";

import {
  hojaAObjetos,
  construirPayComponents,
  listarMeses,
  filasParaTXT,
} from "./payComponents";

import { generarTXT, descargarTXT } from "./generarTXT";
import { FORMATO_INTEGRACION } from "./formatoIntegracion";

// true  -> el TXT descargado sale con las columnas alineadas (ordenado).
// false -> formato compacto, sin espacios extra.
const TXT_ALINEADO = true;

// ============================================================
// OBTENER ÍNDICE DE UNA COLUMNA DE EXCEL
// ============================================================

function obtenerIndiceColumna(referencia) {
  const letras = referencia.replace(/[0-9]/g, "");
  let indice = 0;

  for (const letra of letras) {
    indice = indice * 26 + letra.charCodeAt(0) - 64;
  }

  return indice - 1;
}

// ============================================================
// CONVERTIR TEXTO XML A DOCUMENTO
// ============================================================

function convertirXML(textoXML) {
  const documento = new DOMParser().parseFromString(
    textoXML,
    "application/xml"
  );

  if (documento.getElementsByTagName("parsererror").length > 0) {
    throw new Error(
      "No fue posible interpretar un archivo interno del Excel."
    );
  }

  return documento;
}

// ============================================================
// LEER TEXTOS COMPARTIDOS DEL EXCEL
// ============================================================

function leerTextosCompartidos(xml) {
  if (!xml) {
    return [];
  }

  const documento = convertirXML(xml);

  return Array.from(documento.getElementsByTagName("si")).map(
    (elemento) =>
      Array.from(elemento.getElementsByTagName("t"))
        .map((texto) => texto.textContent || "")
        .join("")
  );
}

// ============================================================
// OBTENER RUTA DE UNA HOJA DEL EXCEL
// ============================================================

function obtenerRutaHoja(destino) {
  const ruta = destino.replace(/^\/+/, "");

  if (ruta.startsWith("xl/")) {
    return ruta;
  }

  return `xl/${ruta}`;
}

// ============================================================
// LEER LAS FILAS Y CELDAS DE UNA HOJA
// ============================================================

function leerFilasHoja(xml, textosCompartidos) {
  const documento = convertirXML(xml);
  const sheetData = documento.getElementsByTagName("sheetData")[0];

  if (!sheetData) {
    return [];
  }

  return Array.from(sheetData.children)
    .filter((elemento) => elemento.localName === "row")
    .map((filaXML) => {
      const numeroFila = Number(filaXML.getAttribute("r")) || null;
      const celdas = {};

      Array.from(filaXML.children)
        .filter((elemento) => elemento.localName === "c")
        .forEach((celdaXML) => {
          const referencia = celdaXML.getAttribute("r");

          if (!referencia) {
            return;
          }

          const tipo = celdaXML.getAttribute("t");

          const valorXML = Array.from(celdaXML.children).find(
            (elemento) => elemento.localName === "v"
          );

          let valor = valorXML ? valorXML.textContent || "" : "";

          if (tipo === "s" && valor !== "") {
            valor = textosCompartidos[Number(valor)] ?? valor;
          } else if (tipo === "inlineStr") {
            valor = Array.from(celdaXML.getElementsByTagName("t"))
              .map((texto) => texto.textContent || "")
              .join("");
          } else if (tipo === "b") {
            valor = valor === "1" ? "VERDADERO" : "FALSO";
          }

          celdas[referencia] = valor;
        });

      return {
        numeroFila,
        celdas,
      };
    });
}

// ============================================================
// LEER EL ARCHIVO EXCEL XLSX O XLSM
// ============================================================

async function leerExcel(archivo) {
  const zip = await JSZip.loadAsync(archivo);

  const archivoLibro = zip.file("xl/workbook.xml");
  const archivoRelaciones = zip.file(
    "xl/_rels/workbook.xml.rels"
  );

  if (!archivoLibro || !archivoRelaciones) {
    throw new Error(
      "El archivo no tiene la estructura esperada de un Excel .xlsx o .xlsm."
    );
  }

  const [textoLibro, textoRelaciones] = await Promise.all([
    archivoLibro.async("text"),
    archivoRelaciones.async("text"),
  ]);

  const libroXML = convertirXML(textoLibro);
  const relacionesXML = convertirXML(textoRelaciones);

  const relaciones = {};

  Array.from(
    relacionesXML.getElementsByTagName("Relationship")
  ).forEach((relacion) => {
    relaciones[relacion.getAttribute("Id")] =
      relacion.getAttribute("Target");
  });

  const archivoTextos = zip.file("xl/sharedStrings.xml");

  const textosCompartidos = leerTextosCompartidos(
    archivoTextos ? await archivoTextos.async("text") : null
  );

  const hojasXML = Array.from(
    libroXML.getElementsByTagName("sheet")
  );

  const hojas = [];

  for (const hojaXML of hojasXML) {
    const nombre = hojaXML.getAttribute("name");

    const idRelacion =
      hojaXML.getAttribute("r:id") ||
      hojaXML.getAttributeNS(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "id"
      );

    const destino = relaciones[idRelacion];

    if (!destino) {
      throw new Error(
        `No se encontró la ruta de la hoja "${nombre}".`
      );
    }

    const ruta = obtenerRutaHoja(destino);
    const archivoHoja = zip.file(ruta);

    if (!archivoHoja) {
      throw new Error(
        `No se encontró el contenido de la hoja "${nombre}".`
      );
    }

    const contenido = await archivoHoja.async("text");

    hojas.push({
      nombre,
      filas: leerFilasHoja(contenido, textosCompartidos),
    });
  }

  return hojas;
}

// ============================================================
// ICONO DE DESCARGA
// ============================================================

function IconoDescarga() {
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      width="28"
      height="28"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
      <polyline points="7 10 12 15 17 10" />
      <line x1="12" y1="15" x2="12" y2="3" />
    </svg>
  );
}

// ============================================================
// ICONO DE AYUDA
// ============================================================

function IconoAyuda() {
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      width="18"
      height="18"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <circle cx="12" cy="12" r="10" />
      <path d="M9.1 9a3 3 0 0 1 5.8 1c0 2-3 2.5-3 4.5" />
      <line x1="12" y1="17.8" x2="12" y2="17.8" />
    </svg>
  );
}

// ============================================================
// ICONO DE ARCHIVO
// ============================================================

function IconoArchivo() {
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      width="22"
      height="22"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
      <polyline points="14 2 14 8 20 8" />
      <path d="m9 15 2 2 4-4" />
    </svg>
  );
}

// ============================================================
// APLICACIÓN PRINCIPAL
// ============================================================

export default function App() {
  const [archivo, setArchivo] = useState(null);
  const [hojas, setHojas] = useState([]);
  const [hojaActiva, setHojaActiva] = useState(0);
  const [cargando, setCargando] = useState(false);
  const [error, setError] = useState("");

  // Estados para la conversión a TXT
  const [registros, setRegistros] = useState([]);
  const [meses, setMeses] = useState([]);
  const [mesElegido, setMesElegido] = useState("");
  const [errorConversion, setErrorConversion] = useState("");
  const [claveInput, setClaveInput] = useState(0);

  // ==========================================================
  // REINICIAR PARA CARGAR OTRO ARCHIVO
  // ==========================================================

  function reiniciar() {
    setArchivo(null);
    setHojas([]);
    setHojaActiva(0);
    setCargando(false);
    setError("");
    setRegistros([]);
    setMeses([]);
    setMesElegido("");
    setErrorConversion("");
    setClaveInput((n) => n + 1);

    window.scrollTo({
      top: 0,
      behavior: "smooth",
    });
  }

  // ==========================================================
  // PROCESAR EL ARCHIVO SELECCIONADO
  // ==========================================================

  async function procesarArchivo(evento) {
    const seleccionado = evento.target.files?.[0];

    setArchivo(null);
    setHojas([]);
    setHojaActiva(0);
    setError("");
    setRegistros([]);
    setMeses([]);
    setMesElegido("");
    setErrorConversion("");

    if (!seleccionado) {
      return;
    }

    const nombre = seleccionado.name.toLowerCase();

    if (
      !nombre.endsWith(".xlsx") &&
      !nombre.endsWith(".xlsm")
    ) {
      setError(
        "Selecciona un archivo .xlsx o .xlsm. El formato antiguo .xls no está admitido."
      );

      evento.target.value = "";
      return;
    }

    setCargando(true);

    try {
      const resultado = await leerExcel(seleccionado);

      setArchivo(seleccionado);
      setHojas(resultado);

      if (resultado.length === 0) {
        setError("El archivo no contiene hojas de cálculo.");
      } else {
        // La conversión utiliza la primera hoja del archivo.
        // La inspección de las hojas sigue disponible si falla.

        try {
          const objetos = hojaAObjetos(resultado[0]);
          const mesesDisponibles = listarMeses(objetos);

          setRegistros(objetos);
          setMeses(mesesDisponibles);

          setMesElegido(
            mesesDisponibles[mesesDisponibles.length - 1] || ""
          );
        } catch (errConversion) {
          setErrorConversion(
            errConversion instanceof Error
              ? errConversion.message
              : "No se pudo preparar la hoja para la conversión."
          );
        }
      }
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : "Ocurrió un error al leer el Excel."
      );
    } finally {
      setCargando(false);
    }
  }

  // ==========================================================
  // CONSTRUIR PAY COMPONENTS DEL MES ELEGIDO
  // ==========================================================

  const conversion = useMemo(() => {
    if (!mesElegido || registros.length === 0) {
      return {
        empleados: [],
        errores: [],
        advertencias: [],
      };
    }

    return construirPayComponents(registros, mesElegido);
  }, [registros, mesElegido]);

  // ==========================================================
  // GENERAR Y DESCARGAR EL TXT
  // El formato (separador, encabezado, BOM, etc.) se configura
  // en formatoIntegracion.js
  // ==========================================================

  function descargar() {
    try {
      const contenido = generarTXT(
        filasParaTXT(conversion.empleados, {
          incluirEncabezado: FORMATO_INTEGRACION.incluirEncabezado,
        }),
        {
          alinear: TXT_ALINEADO,
          lineaSeparadora: false,
          espaciarSeparador: false,
        }
      );

      descargarTXT(contenido, `nomina-${mesElegido}.txt`);

      setErrorConversion("");
    } catch (err) {
      setErrorConversion(
        err instanceof Error
          ? err.message
          : "No se pudo generar el TXT."
      );
    }
  }

  // ==========================================================
  // INFORMACIÓN DE LA HOJA SELECCIONADA
  // ==========================================================

  const hojaSeleccionada = hojas[hojaActiva];

  const columnas = hojaSeleccionada
    ? Array.from(
        new Set(
          hojaSeleccionada.filas.flatMap((fila) =>
            Object.keys(fila.celdas).map((referencia) =>
              referencia.replace(/\d+/g, "")
            )
          )
        )
      ).sort(
        (a, b) =>
          obtenerIndiceColumna(a) - obtenerIndiceColumna(b)
      )
    : [];

  // ==========================================================
  // INTERFAZ
  // ==========================================================

  return (
    <main className="contenedor">
      {/* ENCABEZADO: LOGO, TÍTULO Y BIENVENIDA */}

      <header className="encabezado-app">
        <div className="encabezado-izquierda">
          <img
            src="/imagen.png"
            alt="Logo Solutions & Payroll"
            className="encabezado-logo"
          />

          <h1 className="encabezado-titulo">
            Conversor de nómina
          </h1>
        </div>

        <span className="encabezado-bienvenida">
          Bienvenido, usuario
        </span>
      </header>

      {/* ======================================================
          CÓMO SE USA LA PÁGINA
      ====================================================== */}

      <section className="ayuda">
        <details>
          <summary>
            <span className="ayuda-icono">
              <IconoAyuda />
            </span>
            <span className="ayuda-titulo">
              Cómo se usa la página
            </span>
            <span className="ayuda-flecha" aria-hidden="true" />
          </summary>

          <ol className="pasos-uso">
            <li>
              <strong>Selecciona</strong> el Excel de nómina (
              <code>.xlsx</code> o <code>.xlsm</code>).
            </li>
            <li>
              <strong>Revisa</strong> las hojas y columnas. La
              conversión usa la <em>primera hoja</em>.
            </li>
            <li>
              <strong>Elige el mes</strong>. Por defecto es el
              último disponible.
            </li>
            <li>
              <strong>Verifica</strong> los empleados procesados,
              los registros con error y las advertencias.
            </li>
            <li>
              <strong>Descarga</strong> el archivo con{" "}
              <em>Descargar TXT</em>.
            </li>
            <li>
              Usa <em>Añadir nuevo archivo</em> para empezar de
              nuevo.
            </li>
          </ol>
        </details>
      </section>

      {/* ======================================================
          1. SELECCIONAR ARCHIVO
      ====================================================== */}

      {archivo ? (
        <>
          <section className="barra-archivo">
            <span className="barra-archivo-icono">
              <IconoArchivo />
            </span>

            <p className="barra-archivo-texto">
              Archivo seleccionado: <strong>{archivo.name}</strong>
            </p>

            <button
              type="button"
              className="barra-archivo-cambiar"
              onClick={reiniciar}
            >
              Cambiar archivo
            </button>
          </section>

          {error && <p className="mensaje-error">{error}</p>}
        </>
      ) : (
      <section className="panel">
        <h2>1. Seleccionar archivo de nómina</h2>

        <p className="descripcion">
          Carga el Excel para inspeccionar su estructura real.
        </p>

        <label className="selector-archivo">
          <span>Seleccionar archivo Excel</span>

          <input
            key={claveInput}
            type="file"
            accept=".xlsx,.xlsm"
            onChange={procesarArchivo}
          />
        </label>

        {cargando && (
          <p className="nota">
            Leyendo el archivo con JSZip...
          </p>
        )}

        {error && (
          <p className="mensaje-error">{error}</p>
        )}
      </section>
      )}

      {/* ======================================================
          2. ESTRUCTURA DEL EXCEL
      ====================================================== */}

      {hojas.length > 0 && (
        <section className="panel">
          <h2>2. Estructura encontrada</h2>

          <p className="descripcion">
            Selecciona una hoja para revisar sus columnas y datos.
          </p>

          <div className="pestanas">
            {hojas.map((hoja, indice) => (
              <button
                key={`${hoja.nombre}-${indice}`}
                type="button"
                className={
                  hojaActiva === indice
                    ? "pestana activa"
                    : "pestana"
                }
                onClick={() => setHojaActiva(indice)}
              >
                {hoja.nombre}
              </button>
            ))}
          </div>

          {hojaSeleccionada && (
            <>
              <p className="resumen-hoja">
                Hoja:{" "}
                <strong>{hojaSeleccionada.nombre}</strong>
                {" · "}
                Filas:{" "}
                <strong>
                  {hojaSeleccionada.filas.length}
                </strong>
                {" · "}
                Columnas:{" "}
                <strong>{columnas.length}</strong>
              </p>

              <div className="tabla-contenedor">
                <table>
                  <thead>
                    <tr>
                      <th>Fila Excel</th>

                      {columnas.map((columna) => (
                        <th key={columna}>{columna}</th>
                      ))}
                    </tr>
                  </thead>

                  <tbody>
                    {hojaSeleccionada.filas
                      .slice(0, 30)
                      .map((fila, indice) => (
                        <tr
                          key={`${
                            fila.numeroFila ?? indice
                          }-${indice}`}
                        >
                          <td className="numero-fila">
                            {fila.numeroFila ?? indice + 1}
                          </td>

                          {columnas.map((columna) => {
                            const numeroFila =
                              fila.numeroFila ?? indice + 1;

                            const letraColumna =
                              columna.replace(/\d+/g, "");

                            const referencia =
                              `${letraColumna}${numeroFila}`;

                            const valor =
                              fila.celdas[referencia] ?? "";

                            return (
                              <td key={columna}>
                                {String(valor)}
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                  </tbody>
                </table>
              </div>

              {hojaSeleccionada.filas.length > 30 && (
                <p className="nota">
                  Se muestran las primeras 30 filas. La lectura
                  conserva los datos encontrados en las demás filas.
                </p>
              )}
            </>
          )}
        </section>
      )}

      {/* ======================================================
          3. GENERAR TXT DE PAY COMPONENTS
      ====================================================== */}

      {hojas.length > 0 && (
        <section className="panel">
          <h2>3. Generar TXT de Pay Components</h2>

          {errorConversion && (
            <p className="mensaje-error">
              {errorConversion}
            </p>
          )}

          {meses.length > 0 && (
            <>
              <p className="descripcion">
                Elige el mes a transferir. Mapeo propuesto desde
                el manual; solo SUELDO (001) está confirmado.
              </p>

              <label className="selector-mes">
                Mes:{" "}

                <select
                  value={mesElegido}
                  onChange={(evento) =>
                    setMesElegido(evento.target.value)
                  }
                >
                  {meses.map((mes) => (
                    <option key={mes} value={mes}>
                      {mes}
                    </option>
                  ))}
                </select>
              </label>

              {/* RESUMEN DE LA CONVERSIÓN */}

              <p className="resumen-hoja">
                Empleados con Pay Components:{" "}
                <strong>{conversion.empleados.length}</strong>
                {" · "}
                Con error:{" "}
                <strong>{conversion.errores.length}</strong>
                {" · "}
                Con advertencia:{" "}
                <strong>{conversion.advertencias.length}</strong>
              </p>

              {/* ERRORES DE CONVERSIÓN */}

              {conversion.errores.length > 0 && (
                <ul className="mensaje-error">
                  {conversion.errores.map((e, indice) => (
                    <li
                      key={`${e.tercero}-${e.contrato}-${indice}`}
                    >
                      {e.tercero} · {e.nombre} (contrato{" "}
                      {e.contrato}): {e.mensaje}
                    </li>
                  ))}
                </ul>
              )}

              {/* ADVERTENCIAS DE CONVERSIÓN */}

              {conversion.advertencias.length > 0 && (
                <div className="mensaje-advertencia">
                  <p>
                    Estos empleados tienen datos sin reportar. No
                    bloquean la descarga, pero revísalos antes de
                    usar el archivo:
                  </p>

                  <ul>
                    {conversion.advertencias.map((a, indice) => (
                      <li
                        key={`${a.tercero}-${a.contrato}-${a.codigo}-${indice}`}
                      >
                        {a.tercero} · {a.nombre} (contrato{" "}
                        {a.contrato}): {a.mensaje}
                      </li>
                    ))}
                  </ul>
                </div>
              )}

              {/* TABLA DE EMPLEADOS */}

              {conversion.empleados.length > 0 && (
                <div className="tabla-contenedor">
                  <table>
                    <thead>
                      <tr>
                        <th>Tercero</th>
                        <th>Nombre</th>
                        <th>Contrato</th>
                        <th>Cadena</th>
                        <th>Caracteres</th>
                      </tr>
                    </thead>

                    <tbody>
                      {conversion.empleados
                        .slice(0, 20)
                        .map((empleado, indice) => (
                          <tr
                            key={`${empleado.tercero}-${empleado.contrato}-${indice}`}
                          >
                            <td>{empleado.tercero}</td>
                            <td>{empleado.nombre}</td>
                            <td>{empleado.contrato}</td>
                            <td className="celda-cadena">
                              {empleado.cadena}
                            </td>
                            <td>
                              {empleado.cadena.length}
                            </td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                </div>
              )}

              {conversion.empleados.length > 20 && (
                <p className="nota">
                  Se muestran los primeros 20 empleados. El TXT
                  incluye todos.
                </p>
              )}

              {/* DESCARGA DEL ARCHIVO TXT */}

              <div className="descarga-txt">
                <div className="descarga-detalles">
                  <div className="descarga-detalle">
                    <span className="detalle-label">
                      Mes seleccionado
                    </span>

                    <strong>{mesElegido || "—"}</strong>
                  </div>

                  <div className="descarga-detalle">
                    <span className="detalle-label">
                      Empleados procesados
                    </span>

                    <strong>
                      {conversion.empleados.length}
                    </strong>
                  </div>

                  <div className="descarga-detalle">
                    <span className="detalle-label">
                      Registros con errores
                    </span>

                    <strong
                      className={
                        conversion.errores.length > 0
                          ? "texto-error"
                          : ""
                      }
                    >
                      {conversion.errores.length}
                    </strong>
                  </div>
                </div>

                {conversion.errores.length > 0 && (
                  <p className="descarga-advertencia">
                    Hay registros con errores. Revisa el detalle
                    anterior antes de utilizar el archivo generado.
                  </p>
                )}

                <div className="descarga-acciones">
                  <button
                    type="button"
                    className="boton-descargar-txt"
                    onClick={descargar}
                    disabled={
                      conversion.empleados.length === 0 ||
                      !mesElegido
                    }
                  >
                    <IconoDescarga />
                    Descargar TXT
                  </button>

                  <button
                    type="button"
                    className="boton-secundario"
                    onClick={reiniciar}
                  >
                    Añadir nuevo archivo
                  </button>
                </div>

                <p className="descarga-pie">
                  Nombre del archivo:{" "}
                  <strong>
                    nomina-{mesElegido || "mes"}.txt
                  </strong>
                </p>
              </div>
            </>
          )}

          {meses.length === 0 && !errorConversion && (
            <p className="nota">
              No se encontraron meses disponibles para generar
              el archivo TXT.
            </p>
          )}
        </section>
      )}
    </main>
  );
}