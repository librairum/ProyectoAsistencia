const XLSX = require('xlsx-style');



function calcularTotal(entrada, salida) {
    if (!entrada || !salida) return '';
    const [h1, m1] = entrada.split(':').map(Number);
    const [h2, m2] = salida.split(':').map(Number);
    let minutos = (h2 * 60 + m2) - (h1 * 60 + m1);
    if (minutos < 0) return '';
    const horas = Math.floor(minutos / 60);
    const mins = minutos % 60;
    return `${String(horas).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
}

exports.exportarExcelPorPagina = (req, res) => {

    try {
   
    // Recibe empleadosPagina y fechas desde el frontend (puedes ajustar esto)
    const { empleadosPagina, fechas } = req.body;

    const wb = { SheetNames: [], Sheets: {} };

    empleadosPagina.forEach(emp => {
        const ws = {};

        const ws_data = [];
        ws_data[1] = [];
        ws_data[1][2] = emp.emp_pin || '';
        ws_data[1][3] = `${emp.emp_firstname} ${emp.emp_lastname}`;
        ws_data[2] = [];
        ws_data[2][2] = 'dni';
        ws_data[2][3] = 'fecha';
        ws_data[2][4] = 'nombreDia';
        ws_data[2][5] = 'HoraEntrada';
        ws_data[2][6] = 'HoraSalida';
        ws_data[2][7] = 'Total';

        let rowIndex = 3;


        fechas.forEach(f => {
            const { nombre } = getDiaSemana(f);
            const marcacion = emp.fechas && emp.fechas[f];
            let entrada = '', salida = '', total = '';
            if (marcacion && marcacion.includes('-')) {
                [entrada, salida] = marcacion.split('-').map(x => x.trim());
                total = calcularTotal(entrada, salida);
            } else if (marcacion && marcacion.trim() !== '-' && marcacion.trim() !== '') {
                entrada = marcacion;
                salida = '';
                total = '';
            }
            ws_data[rowIndex] = [];
            ws_data[rowIndex][2] = emp.emp_pin || '';
            const [yyyy, mm, dd] = f.split('-');
            const fechaFormateada = `${dd}/${mm}/${yyyy}`;
            ws_data[rowIndex][3] = fechaFormateada;
            ws_data[rowIndex][4] = nombre;
            ws_data[rowIndex][5] = entrada || '';
            ws_data[rowIndex][6] = salida || '';
            ws_data[rowIndex][7] = total || '';
            rowIndex++;
        });

        /*
        let nombreHoja = `${emp.emp_firstname} ${emp.emp_lastname}`.trim();
        if (nombreHoja.length > 28) nombreHoja = emp.emp_pin || 'Empleado';

        // Crea la hoja vacía
        const ws = {};
        // Agrega los datos con sheet_add_aoa
        XLSX.utils.sheet_add_aoa(ws, ws_data);
        //const ws = XLSX.utils.aoa_to_sheet(ws_data);
        */


        for (let r = 1; r < ws_data.length; r++) {
        for (let c = 2; c < ws_data[r].length; c++) {
            const cellAddress = XLSX.utils.encode_cell({ c, r });
            let cellValue = ws_data[r][c] || '';
            let cell = {
                v: cellValue,
                s: {
                    border: {
                        top:    { style: "thin", color: { rgb: "000000" } },
                        bottom: { style: "thin", color: { rgb: "000000" } },
                        left:   { style: "thin", color: { rgb: "000000" } },
                        right:  { style: "thin", color: { rgb: "000000" } }
                    }
                }
            };

            // Formato por columna
            if (c === 2 && ((r > 2 && !isNaN(Number(cellValue))) || (r === 1 && !isNaN(Number(cellValue)))) ) { // DNI en C2 y en datos
                cell.t = 'n';
            } else if (c === 3 && r > 2 && cellValue) { // Fecha
                // Convierte dd/mm/yyyy a Date
                const [dd, mm, yyyy] = cellValue.split('/');
                cell.t = 'd';
                cell.v = new Date(`${yyyy}-${mm}-${dd}`);
                cell.z = 'dd/mm/yyyy';
            } else {
                cell.t = 's';
            }

            ws[cellAddress] = cell;
        }
    }


        // Aplica bordes a las celdas con datos
        /*
        for (let r = 1; r < ws_data.length; r++) {
            for (let c = 2; c < ws_data[r].length; c++) {
                const cellAddress = XLSX.utils.encode_cell({ c, r });
                ws[cellAddress] = {
                    v: ws_data[r][c] || '',
                    s: {
                        border: {
                            top:    { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } },
                            left:   { style: "thin", color: { rgb: "000000" } },
                            right:  { style: "thin", color: { rgb: "000000" } }
                        }
                    }
                };
            }
        }
        */


        // Define el rango de la hoja
        ws['!ref'] = `C2:H${ws_data.length}`;

        let nombreHoja = `${emp.emp_firstname} ${emp.emp_lastname}`.trim();
        if (nombreHoja.length > 28) nombreHoja = emp.emp_pin || 'Empleado';

        wb.SheetNames.push(nombreHoja || 'Empleado');
        wb.Sheets[nombreHoja || 'Empleado'] = ws;
    });

    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Disposition', 'attachment; filename=asistencia_por_pagina.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
    }
    catch (err) {
        console.error('Error al generar Excel:', err);
        res.status(500).send('Error interno al generar el Excel');
    }
};

// Helper para día de semana
function getDiaSemana(fechaISO) {
    const dias = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
    const [anio, mes, dia] = fechaISO.split('-').map(Number);
    const d = new Date(anio, mes - 1, dia);
    return { nombre: dias[d.getDay()], numero: d.getDay() };
}

