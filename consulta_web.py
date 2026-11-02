import io
import os
import re
import random
import string
import time
from flask import Flask, render_template_string, request, jsonify, send_file
import requests
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# =====================================================================
# LÓGICA DE EXTRACCIÓN DIRECTA SUNAT (MÉTODO sunatrecaptcha3.js)
# =====================================================================
def generar_key_sunat(longitud=34):
    caracteres = string.ascii_letters + string.digits
    return "".join(random.choices(caracteres, k=longitud))


def consultar_ruc_sunat(ruc):
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp",
        "Origin": "https://e-consultaruc.sunat.gob.pe",
        "Accept-Language": "es-PE,es;q=0.9",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    })

    url_base = "https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp"
    url_alias = "https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias"

    try:
        session.get(url_base, timeout=10)
        token_recaptcha = generar_key_sunat(34)

        payload = {
            "accion": "consPorRuc",
            "nroRuc": str(ruc).strip(),
            "contexto": "ti-it",
            "modo": "1",
            "rbtnTipo": "1",
            "search1": str(ruc).strip(),
            "codigo": "",
            "tipdoc": "1",
            "token": token_recaptcha
        }

        resp = session.post(url_alias, data=payload, timeout=12)
        if resp.encoding != 'utf-8':
            resp.encoding = resp.apparent_encoding or 'iso-8859-1'
        html = resp.text

        if "Resultado de la Búsqueda" not in html and "desRuc" not in html:
            return {
                "ruc": str(ruc),
                "razon_social": "NO ENCONTRADO O ERROR EN SUNAT",
                "condicion": "-",
                "fecha_inicio": "-",
                "direccion": "-"
            }

        # 1. Razón Social
        des_match = re.search(r'<input\s+type="hidden"\s+name="desRuc"\s+value="([^"]+)"', html)
        if des_match:
            razon_social = des_match.group(1).strip()
        else:
            h4_match = re.search(r'<h4 class="list-group-item-heading">\d{11}\s*-\s*([^<]+)</h4>', html)
            razon_social = h4_match.group(1).strip() if h4_match else "-"

        # 2. Condición
        c_match = re.search(r'Condici[oó\W]+n.*?<p class="list-group-item-text">\s*([^<]+?)\s*</p>', html, re.DOTALL | re.IGNORECASE)
        if c_match:
            texto_cond = re.sub(r'\s+', ' ', c_match.group(1)).strip().upper()
            if "HABIDO" in texto_cond:
                condicion = "HABIDO"
            elif "NO HABIDO" in texto_cond:
                condicion = "NO HABIDO"
            elif "PENDIENTE" in texto_cond:
                condicion = "PENDIENTE"
            else:
                condicion = texto_cond
        else:
            if "NO HABIDO" in html.upper():
                condicion = "NO HABIDO"
            elif "HABIDO" in html.upper():
                condicion = "HABIDO"
            else:
                condicion = "-"

        # 3. Fecha Inicio Actividades
        f_match = re.search(r'Fecha de Inicio de Actividades:.*?<p class="list-group-item-text">\s*([0-9/]+)\s*</p>', html, re.DOTALL | re.IGNORECASE)
        fecha_inicio = f_match.group(1).strip() if f_match else "-"

        # 4. Dirección / Domicilio Fiscal
        d_match = re.search(r'Domicilio Fiscal:.*?<p class="list-group-item-text">\s*([^<]+?)\s*</p>', html, re.DOTALL | re.IGNORECASE)
        direccion = re.sub(r'\s+', ' ', d_match.group(1)).strip() if d_match else "-"

        return {
            "ruc": str(ruc),
            "razon_social": razon_social,
            "condicion": condicion,
            "fecha_inicio": fecha_inicio,
            "direccion": direccion
        }
    except Exception as ex:
        return {
            "ruc": str(ruc),
            "razon_social": "ERROR DE CONEXIÓN",
            "condicion": "-",
            "fecha_inicio": "-",
            "direccion": str(ex)
        }


# =====================================================================
# INTERFAZ WEB (HTML/CSS/JS INTEGRADO)
# =====================================================================
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Consulta RUC SUNAT - Plataforma Web</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        body { background-color: #f3f4f6; color: #1f2937; padding: 30px 15px; }
        .container { max-width: 1050px; margin: 0 auto; background: #ffffff; border-radius: 10px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); padding: 25px; }
        h1 { font-size: 24px; color: #1e3a8a; margin-bottom: 20px; display: flex; align-items: center; gap: 10px; }
        .card { background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 8px; padding: 18px; margin-bottom: 20px; }
        .card-title { font-weight: 600; font-size: 16px; margin-bottom: 12px; color: #374151; }
        .grid-inputs { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        @media(max-width: 768px) { .grid-inputs { grid-template-columns: 1fr; } }
        textarea { width: 100%; height: 95px; border: 1px solid #d1d5db; border-radius: 6px; padding: 10px; font-family: monospace; resize: vertical; }
        input[type="file"] { width: 100%; padding: 10px; border: 1px dashed #9ca3af; border-radius: 6px; background: #fff; cursor: pointer; }
        .btn-group { display: flex; gap: 12px; margin-bottom: 20px; align-items: center; }
        button { cursor: pointer; border: none; padding: 10px 20px; border-radius: 6px; font-weight: 600; font-size: 14px; transition: all 0.2s; }
        .btn-primary { background: #2563eb; color: #fff; }
        .btn-primary:hover { background: #1d4ed8; }
        .btn-primary:disabled { background: #93c5fd; cursor: not-allowed; }
        .btn-success { background: #059669; color: #fff; }
        .btn-success:hover { background: #047857; }
        .btn-success:disabled { background: #a7f3d0; cursor: not-allowed; }
        #status { font-style: italic; color: #4b5563; font-size: 14px; }
        .table-responsive { overflow-x: auto; margin-top: 15px; border-radius: 6px; border: 1px solid #e5e7eb; }
        table { width: 100%; border-collapse: collapse; text-align: left; font-size: 13px; }
        th { background: #1e3a8a; color: #fff; padding: 12px; text-transform: uppercase; font-size: 11px; letter-spacing: 0.5px; }
        td { padding: 10px 12px; border-bottom: 1px solid #e5e7eb; }
        tr:nth-child(even) { background-color: #f9fafb; }
        .badge { display: inline-block; padding: 3px 8px; border-radius: 4px; font-weight: 700; font-size: 11px; }
        .badge-habido { background: #dcfce7; color: #166534; }
        .badge-nohabido { background: #fee2e2; color: #991b1b; }
    </style>
</head>
<body>
    <div class="container">
        <h1>🇵🇪 Consulta RUC SUNAT (Plataforma Web)</h1>

        <div class="card">
            <div class="card-title">1. Origen de RUCs</div>
            <div class="grid-inputs">
                <div>
                    <label style="display:block; font-size: 13px; margin-bottom: 6px; font-weight: 500;">Ingreso Manual (uno por línea o separados por comas):</label>
                    <textarea id="txtManual" placeholder="Ej: 20100130204, 20100070970"></textarea>
                </div>
                <div>
                    <label style="display:block; font-size: 13px; margin-bottom: 6px; font-weight: 500;">O Cargar Archivo Excel (.xlsx / .xls):</label>
                    <input type="file" id="fileExcel" accept=".xlsx, .xls">
                    <p style="font-size: 11px; color: #6b7280; margin-top: 6px;">Lee automáticamente cualquier columna o fila con 11 dígitos, con o sin encabezados.</p>
                </div>
            </div>
        </div>

        <div class="btn-group">
            <button class="btn-primary" id="btnConsultar" onclick="iniciarProceso()">▶ Iniciar Consulta Automática</button>
            <button class="btn-success" id="btnDescargar" onclick="descargarExcel()" disabled>💾 Descargar Reporte Excel</button>
            <span id="status">Listo para consultar.</span>
        </div>

        <div class="card">
            <div class="card-title">2. Resultados Extraídos en Vivo (<span id="count">0</span>)</div>
            <div class="table-responsive">
                <table>
                    <thead>
                        <tr>
                            <th style="width: 120px; text-align: center;">RUC</th>
                            <th>Razón Social</th>
                            <th style="width: 100px; text-align: center;">Condición</th>
                            <th style="width: 110px; text-align: center;">F. Inicio Act.</th>
                            <th>Dirección Fiscal</th>
                        </tr>
                    </thead>
                    <tbody id="tbodyResultados">
                        <tr><td colspan="5" style="text-align: center; color: #9ca3af; padding: 25px;">No hay datos para mostrar</td></tr>
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <script>
        let resultadosGlobales = [];

        async function iniciarProceso() {
            const manualText = document.getElementById("txtManual").value;
            const fileInput = document.getElementById("fileExcel");
            const btnConsultar = document.getElementById("btnConsultar");
            const btnDescargar = document.getElementById("btnDescargar");
            const status = document.getElementById("status");
            const tbody = document.getElementById("tbodyResultados");

            const formData = new FormData();
            formData.append("manual", manualText);
            if (fileInput.files.length > 0) {
                formData.append("excel", fileInput.files[0]);
            }

            btnConsultar.disabled = true;
            btnDescargar.disabled = true;
            status.innerText = "Extrayendo RUCs de entrada...";
            tbody.innerHTML = "";
            resultadosGlobales = [];
            document.getElementById("count").innerText = "0";

            try {
                // Paso 1: Obtener lista consolidada de RUCs
                const respRucs = await fetch("/extraer-rucs", { method: "POST", body: formData });
                const dataRucs = await respRucs.json();

                if (!dataRucs.rucs || dataRucs.rucs.length === 0) {
                    alert("No se detectaron RUCs válidos de 11 dígitos.");
                    btnConsultar.disabled = false;
                    status.innerText = "Listo.";
                    return;
                }

                const rucs = dataRucs.rucs;
                const total = rucs.length;

                // Paso 2: Consultar uno por uno e ir agregando a la tabla web en tiempo real
                for (let i = 0; i < total; i++) {
                    const ruc = rucs[i];
                    status.innerText = `Consultando [${i + 1}/${total}]: ${ruc}...`;

                    const rConsulta = await fetch(`/consultar-uno?ruc=${ruc}`);
                    const item = await rConsulta.json();
                    resultadosGlobales.push(item);

                    // Insertar en tabla
                    const tr = document.createElement("tr");
                    const condClass = item.condicion === "HABIDO" ? "badge-habido" : (item.condicion.includes("NO") ? "badge-nohabido" : "");
                    tr.innerHTML = `
                        <td style="text-align: center; font-weight: 600;">${item.ruc}</td>
                        <td>${item.razon_social}</td>
                        <td style="text-align: center;"><span class="badge ${condClass}">${item.condicion}</span></td>
                        <td style="text-align: center;">${item.fecha_inicio}</td>
                        <td>${item.direccion}</td>
                    `;
                    tbody.appendChild(tr);
                    document.getElementById("count").innerText = resultadosGlobales.length;
                }

                status.innerText = `Completado: ${total} RUCs procesados con éxito.`;
                btnDescargar.disabled = false;
            } catch (err) {
                alert("Ocurrió un error al procesar las consultas.");
                status.innerText = "Error.";
            } finally {
                btnConsultar.disabled = false;
            }
        }

        async function descargarExcel() {
            if (resultadosGlobales.length === 0) return;
            const resp = await fetch("/exportar-excel", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ datos: resultadosGlobales })
            });
            const blob = await resp.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = "Reporte_Consulta_RUC_SUNAT.xlsx";
            document.body.appendChild(a);
            a.click();
            a.remove();
        }
    </script>
</body>
</html>
"""

# =====================================================================
# RUTAS DE FLASK
# =====================================================================
@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route("/extraer-rucs", methods=["POST"])
def extraer_rucs():
    rucs = []
    vistos = set()

    # Manual
    manual_text = request.form.get("manual", "")
    for r in re.findall(r"\b\d{11}\b", manual_text):
        if r not in vistos:
            vistos.add(r)
            rucs.append(r)

    # Archivo Excel
    if "excel" in request.files:
        archivo = request.files["excel"]
        if archivo.filename:
            try:
                df = pd.read_excel(archivo, header=None, dtype=str)
                for col in df.columns:
                    for celda in df[col].dropna():
                        for r in re.findall(r"\b\d{11}\b", str(celda).strip()):
                            if r not in vistos:
                                vistos.add(r)
                                rucs.append(r)
            except Exception as e:
                print(f"Error leyendo excel: {e}")

    return jsonify({"rucs": rucs})


@app.route("/consultar-uno")
def consultar_uno():
    ruc = request.args.get("ruc", "")
    dato = consultar_ruc_sunat(ruc)
    return jsonify(dato)


@app.route("/exportar-excel", methods=["POST"])
def exportar_excel():
    datos = request.json.get("datos", [])
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Consulta_RUC"

    fill_hdr = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    font_hdr = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    borde = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    headers = ["N°", "RUC", "Razón Social", "Condición Contribuyente", "Fecha Inicio Actividades", "Dirección Fiscal"]
    ws.append(headers)

    for cell in ws[1]:
        cell.fill = fill_hdr
        cell.font = font_hdr
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for idx, r in enumerate(datos, start=1):
        ws.append([
            idx,
            str(r["ruc"]),
            r["razon_social"],
            r["condicion"],
            r["fecha_inicio"],
            r["direccion"]
        ])
        fila = ws[idx + 1]
        fila[0].alignment = Alignment(horizontal="center")
        fila[1].alignment = Alignment(horizontal="center")
        fila[1].number_format = '@'
        fila[2].alignment = Alignment(horizontal="left")
        fila[3].alignment = Alignment(horizontal="center")
        fila[4].alignment = Alignment(horizontal="center")
        fila[5].alignment = Alignment(horizontal="left")

        for cell in fila:
            cell.border = borde

    for col in ws.columns:
        max_len = max(len(str(c.value or '')) for c in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 65)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        download_name="Reporte_Consulta_RUC_SUNAT.xlsx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    print("Iniciando servidor web...")
    print("Abre en tu navegador: http://127.0.0.1:5000")
    app.run(debug=True, port=5000)