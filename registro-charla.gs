const SPREADSHEET_ID = '1bW58IcIZLdlb7qDfsUZLx4LDKjwep3hn44xqblaoIVY'
const SHEET_NAME = 'Respuestas'

const ORIGEN_PARAM = 'qXZ9p0'

function doPost(e) {
  try {
    const data = e.parameter

    const origen = data[ORIGEN_PARAM] || ''
    data.origen = origen

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID)
    const sheet = spreadsheet.getSheetByName(SHEET_NAME)
    const id = Utilities.getUuid()

    const fechaHora = new Date().toLocaleString('es-ES', {
      timeZone: 'America/Lima',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
    })

    const rowData = [
      id,
      fechaHora,
      data.nombres || '',
      data.apellidos || '',
      data.documento || '',
      data.celular || '',
      data.correo || '',
      data.programa || '',
      origen, // nueva columna al final
    ]

    sheet.appendRow(rowData)

    try {
      const htmlBody = `
 <!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Confirmación de Registro</title>

<style>
    body {
        margin: 0;
        padding: 0;
        background-color: #f4f6f9;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        color: #1f2937;
    }

    .wrapper {
        width: 100%;
        padding: 20px 12px;
    }

    .card {
        max-width: 600px;
        margin: 0 auto;
        background: #ffffff;
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 8px 20px rgba(0,0,0,0.08);
    }

    .header {
        background: linear-gradient(135deg, #030d30 0%, #143293 100%);
        color: white;
        padding: 24px 18px;
        text-align: center;
    }

    .header h1 {
        margin: 0;
        font-size: 20px;
        font-weight: 600;
    }

    .header p {
        margin: 6px 0 0 0;
        font-size: 13px;
        opacity: 0.9;
    }

    .hero {
        padding: 22px 20px 6px 20px;
    }

    .hero h2 {
        margin: 0 0 6px 0;
        font-size: 18px;
        color: #030d30;
    }

    .hero p {
        margin: 0;
        font-size: 14px;
        color: #4b5563;
    }

    .event-box {
        margin: 14px 20px;
        padding: 14px;
        border-radius: 10px;
        background: linear-gradient(135deg, #030d30 0%, #143293 100%);
        color: white;
    }

    .event-title {
        font-size: 15px;
        font-weight: 600;
        margin-bottom: 6px;
    }

    .event-details {
        font-size: 13px;
        line-height: 1.5;
        opacity: 0.95;
    }

    .message-box {
        margin: 12px 20px;
        padding: 14px;
        border-radius: 10px;
        background-color: #e8f0ff;
        border: 1px solid #143293;
        color: #0f1f50;
        font-size: 13.5px;
        line-height: 1.5;
    }

    .section {
        padding: 0 20px 20px 20px;
    }

    .section-title {
        font-size: 12px;
        font-weight: 700;
        color: #143293;
        margin-bottom: 10px;
        letter-spacing: 0.4px;
        text-transform: uppercase;
    }

    .info-grid {
        display: grid;
        gap: 8px;
    }

    .info-item {
        background: #f9fafb;
        padding: 10px 12px;
        border-radius: 8px;
    }

    .label {
        font-size: 11px;
        color: #6b7280;
        margin-bottom: 2px;
        display: block;
    }

    .value {
        font-size: 13.5px;
        font-weight: 600;
        color: #111827;
    }

    .divider {
        height: 1px;
        background: #e5e7eb;
        margin: 6px 20px 14px 20px;
    }

    .closing {
        padding: 0 20px 20px 20px;
        font-size: 13.5px;
        color: #4b5563;
        line-height: 1.5;
    }

    .footer {
        background-color: #030d30;
        color: white;
        padding: 16px;
        text-align: center;
    }

    .footer p {
        margin: 4px 0;
        font-size: 12px;
    }

</style>
</head>

<body>

<div class="wrapper">
    <div class="card">

        <div class="header">
            <h1>Blackwell Global University</h1>
            <p>Confirmación de registro a charla informativa</p>
        </div>

        <div class="hero">
            <h2>Hola ${data.nombres},</h2>
            <p>Tu registro para la charla informativa ha sido confirmado.</p>
        </div>

        <div class="event-box">
            <div class="event-title">Charla Informativa BGU</div>
            <div class="event-details">
                📅 Fecha: Lunes 06 de Abril<br>
                ⏰ Hora: 07:30 pm<br>
                📍 Modalidad: ??
            </div>
        </div>

        <div class="message-box">
            Bienvenido. Desde hoy descubrirás cómo llevar tu perfil profesional a nivel internacional con Blackwell Global University.
        </div>

        <div class="section">
            <div class="section-title">Datos registrados</div>

            <div class="info-grid">
                <div class="info-item">
                    <span class="label">Programa de interés</span>
                    <span class="value">${data.programa}</span>
                </div>

                <div class="info-item">
                    <span class="label">Documento de Identidad</span>
                    <span class="value">${data.documento}</span>
                </div>

                <div class="info-item">
                    <span class="label">Celular</span>
                    <span class="value">${data.celular}</span>
                </div>
            </div>
        </div>

        <div class="divider"></div>

        <div class="closing">
            Recibirás próximamente el enlace de acceso a la charla en este mismo correo.
            <br><br>
            Si tienes alguna consulta, nuestro equipo de admisión estará encantado de ayudarte.
            <br><br>
            <strong>Equipo de Admisión</strong><br>
            Blackwell Global University
        </div>

        <div class="footer">
            <p><strong>Blackwell Global University</strong></p>
            <p>Formando líderes del mañana</p>
            <p style="font-size:11px; margin-top:6px;">
                Este es un mensaje automático. Por favor, no respondas directamente a este correo.
            </p>
        </div>

    </div>
</div>

</body>
</html>
`

      MailApp.sendEmail({
        to: data.correo,
        subject: 'Confirmación de Registro - Charla Informativa BGU',
        htmlBody: htmlBody,
      })
    } catch (emailError) {
      console.error('Error enviando correo:', emailError.message)
    }

    return ContentService.createTextOutput(
      JSON.stringify({ success: true }),
    ).setMimeType(ContentService.MimeType.JSON)
  } catch (error) {
    // Respuesta de error
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, error: error.message }),
    ).setMimeType(ContentService.MimeType.JSON)
  }
}
