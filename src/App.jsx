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

function obtenerIndiceColumna(referencia) {
  const letras = referencia.replace(/[0-9]/g, "");
  let indice = 0;

  for (const letra of letras) {
    indice = indice * 26 + letra.charCodeAt(0) - 64;
  }

  return indice - 1;
}

function convertirXML(textoXML) {
  const documento = new DOMParser().parseFromString(
    textoXML,
    "application/xml"
  );

  if (documento.getElementsByTagName("parsererror").length > 0) {
    throw new Error("No fue posible interpretar un archivo interno del Excel.");
  }

  return documento;
}

function leerTextosCompartidos(xml) {
  if (!xml) return [];

  const documento = convertirXML(xml);

  return Array.from(documento.getElementsByTagName("si")).map((elemento) =>
    Array.from(elemento.getElementsByTagName("t"))
      .map((texto) => texto.textContent || "")
      .join("")
  );
}

function obtenerRutaHoja(destino) {
  const ruta = destino.replace(/^\/+/, "");

  if (ruta.startsWith("xl/")) {
    return ruta;
  }

  return `xl/${ruta}`;
}

function leerFilasHoja(xml, textosCompartidos) {
  const documento = convertirXML(xml);
  const sheetData = documento.getElementsByTagName("sheetData")[0];

  if (!sheetData) return [];

  return Array.from(sheetData.children)
    .filter((elemento) => elemento.localName === "row")
    .map((filaXML) => {
      const numeroFila = Number(filaXML.getAttribute("r")) || null;
      const celdas = {};

      Array.from(filaXML.children)
        .filter((elemento) => elemento.localName === "c")
        .forEach((celdaXML) => {
          const referencia = celdaXML.getAttribute("r");

          if (!referencia) return;

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

      return { numeroFila, celdas };
    });
}

async function leerExcel(archivo) {
  const zip = await JSZip.loadAsync(archivo);

  const archivoLibro = zip.file("xl/workbook.xml");
  const archivoRelaciones = zip.file("xl/_rels/workbook.xml.rels");

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

  Array.from(relacionesXML.getElementsByTagName("Relationship")).forEach(
    (relacion) => {
      relaciones[relacion.getAttribute("Id")] = relacion.getAttribute("Target");
    }
  );

  const archivoTextos = zip.file("xl/sharedStrings.xml");

  const textosCompartidos = leerTextosCompartidos(
    archivoTextos ? await archivoTextos.async("text") : null
  );

  const hojasXML = Array.from(libroXML.getElementsByTagName("sheet"));
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
      throw new Error(`No se encontró la ruta de la hoja "${nombre}".`);
    }

    const ruta = obtenerRutaHoja(destino);
    const archivoHoja = zip.file(ruta);

    if (!archivoHoja) {
      throw new Error(`No se encontró el contenido de la hoja "${nombre}".`);
    }

    const contenido = await archivoHoja.async("text");

    hojas.push({
      nombre,
      filas: leerFilasHoja(contenido, textosCompartidos),
    });
  }

  return hojas;
}

export default function App() {
  const [archivo, setArchivo] = useState(null);
  const [hojas, setHojas] = useState([]);
  const [hojaActiva, setHojaActiva] = useState(0);
  const [cargando, setCargando] = useState(false);
  const [error, setError] = useState("");

  // Estados nuevos para la conversión a TXT
  const [registros, setRegistros] = useState([]);
  const [meses, setMeses] = useState([]);
  const [mesElegido, setMesElegido] = useState("");
  const [errorConversion, setErrorConversion] = useState("");

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

    if (!seleccionado) return;

    const nombre = seleccionado.name.toLowerCase();

    if (!nombre.endsWith(".xlsx") && !nombre.endsWith(".xlsm")) {
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
        // La conversión usa la primera hoja. Si falla, la inspección sigue funcionando.
        try {
          const objetos = hojaAObjetos(resultado[0]);
          const mesesDisponibles = listarMeses(objetos);

          setRegistros(objetos);
          setMeses(mesesDisponibles);
          setMesElegido(mesesDisponibles[mesesDisponibles.length - 1] || "");
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
        err instanceof Error ? err.message : "Ocurrió un error al leer el Excel."
      );
    } finally {
      setCargando(false);
    }
  }

  // Cadenas de Pay Components del mes elegido
  const conversion = useMemo(() => {
    if (!mesElegido || registros.length === 0) {
      return { empleados: [], errores: [] };
    }

    return construirPayComponents(registros, mesElegido);
  }, [registros, mesElegido]);

  function descargar() {
    try {
      const contenido = generarTXT(filasParaTXT(conversion.empleados));
      descargarTXT(contenido, `nomina-${mesElegido}.txt`);
    } catch (err) {
      setErrorConversion(
        err instanceof Error ? err.message : "No se pudo generar el TXT."
      );
    }
  }

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
      ).sort((a, b) => obtenerIndiceColumna(a) - obtenerIndiceColumna(b))
    : [];

  return (
    <main className="contenedor">
      <header className="encabezado">
        <div className="logo">N</div>

        <div>
          <h1>Conversor de nómina</h1>
          <p>Excel a TXT · Integración de nómina</p>
        </div>
      </header>

      <section className="panel">
        <h2>1. Seleccionar archivo de nómina</h2>

        <p className="descripcion">
          Carga el Excel para inspeccionar su estructura real.
        </p>

        <label className="selector-archivo">
          <span>Seleccionar archivo Excel</span>

          <input type="file" accept=".xlsx,.xlsm" onChange={procesarArchivo} />
        </label>

        {archivo && (
          <p className="archivo-seleccionado">
            Archivo seleccionado: <strong>{archivo.name}</strong>
          </p>
        )}

        {cargando && <p>Leyendo el archivo con JSZip...</p>}

        {error && <p className="mensaje-error">{error}</p>}
      </section>

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
                className={hojaActiva === indice ? "pestana activa" : "pestana"}
                onClick={() => setHojaActiva(indice)}
              >
                {hoja.nombre}
              </button>
            ))}
          </div>

          {hojaSeleccionada && (
            <>
              <p className="resumen-hoja">
                Hoja: <strong>{hojaSeleccionada.nombre}</strong>
                {" · "}
                Filas: <strong>{hojaSeleccionada.filas.length}</strong>
                {" · "}
                Columnas: <strong>{columnas.length}</strong>
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
                    {hojaSeleccionada.filas.slice(0, 30).map((fila, indice) => (
                      <tr key={`${fila.numeroFila ?? indice}-${indice}`}>
                        <td className="numero-fila">
                          {fila.numeroFila ?? indice + 1}
                        </td>

                        {columnas.map((columna) => {
                          const numeroFila = fila.numeroFila ?? indice + 1;
                          const letraColumna = columna.replace(/\d+/g, "");
                          const referencia = `${letraColumna}${numeroFila}`;
                          const valor = fila.celdas[referencia] ?? "";

                          return <td key={columna}>{String(valor)}</td>;
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {hojaSeleccionada.filas.length > 30 && (
                <p className="nota">
                  Se muestran las primeras 30 filas. La lectura conserva los
                  datos encontrados en las demás filas.
                </p>
              )}
            </>
          )}
        </section>
      )}

      {hojas.length > 0 && (
        <section className="panel">
          <h2>3. Generar TXT de Pay Components</h2>

          {errorConversion && (
            <p className="mensaje-error">{errorConversion}</p>
          )}

          {meses.length > 0 && (
            <>
              <p className="descripcion">
                Elige el mes a transferir. Por ahora solo está mapeado el
                concepto SUELDO (001) como 9GBA.
              </p>

              <label>
                Mes:{" "}
                <select
                  value={mesElegido}
                  onChange={(evento) => setMesElegido(evento.target.value)}
                >
                  {meses.map((mes) => (
                    <option key={mes} value={mes}>
                      {mes}
                    </option>
                  ))}
                </select>
              </label>

              <p className="resumen-hoja">
                Empleados con Pay Components:{" "}
                <strong>{conversion.empleados.length}</strong>
                {" · "}
                Con error: <strong>{conversion.errores.length}</strong>
              </p>

              {conversion.errores.length > 0 && (
                <ul className="mensaje-error">
                  {conversion.errores.map((e) => (
                    <li key={`${e.tercero}-${e.contrato}`}>
                      {e.tercero} · {e.nombre} (contrato {e.contrato}):{" "}
                      {e.mensaje}
                    </li>
                  ))}
                </ul>
              )}

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
                      {conversion.empleados.slice(0, 20).map((e) => (
                        <tr key={`${e.tercero}-${e.contrato}`}>
                          <td>{e.tercero}</td>
                          <td>{e.nombre}</td>
                          <td>{e.contrato}</td>
                          <td>{e.cadena}</td>
                          <td>{e.cadena.length}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              {conversion.empleados.length > 20 && (
                <p className="nota">
                  Se muestran los primeros 20 empleados. El TXT incluye todos.
                </p>
              )}

              <button
                type="button"
                className="pestana activa"
                onClick={descargar}
                disabled={conversion.empleados.length === 0}
              >
                Descargar TXT
              </button>
            </>
          )}
        </section>
      )}
    </main>
  );
}