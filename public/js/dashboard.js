document.addEventListener('DOMContentLoaded', () => {
    cargarEmpresa();
    cargarPersonas();
    cargarEmpleadosPorDepartamento();
    
    const hoy = new Date();
    const fin = hoy.toISOString().split('T')[0];
    const inicioDate = new Date(hoy);
    inicioDate.setDate(hoy.getDate() - 6);
    const inicio = inicioDate.toISOString().split('T')[0];

    cargarAsistenciaDetallada(inicio, fin);

    const btnHoy = document.getElementById('btnHoy');
    const btn7dias = document.getElementById('btn7dias');
    const btnMes = document.getElementById('btnMes');
    const btnPersonalizado = document.getElementById('btnPersonalizado');
    const btnOk = document.getElementById('btnOk');
    const inputFechaInicio = document.getElementById('fechaInicio');
    const inputFechaFin = document.getElementById('fechaFin');

    if (btnHoy) {
        btnHoy.onclick = () => {
            ocultarPersonalizado();
            const hoy = new Date();
            const fecha = hoy.toISOString().split('T')[0];
            cargarAsistenciaDetallada(fecha, fecha);
        };
    }
    if (btn7dias) {
        btn7dias.onclick = () => {
            ocultarPersonalizado();
            const hoy = new Date();
            const fin = hoy.toISOString().split('T')[0];
            const inicioDate = new Date(hoy);
            inicioDate.setDate(hoy.getDate() - 6);
            const inicio = inicioDate.toISOString().split('T')[0];
            cargarAsistenciaDetallada(inicio, fin);
        };
    }
    if (btnMes) {
        btnMes.onclick = () => {
            ocultarPersonalizado();
            const hoy = new Date();
            const fin = hoy.toISOString().split('T')[0];
            const inicio = fin.slice(0, 8) + '01';
            cargarAsistenciaDetallada(inicio, fin);
        };
    }
    if (btnPersonalizado && btnOk && inputFechaInicio && inputFechaFin) {
        btnPersonalizado.onclick = () => mostrarPersonalizado();
        btnOk.onclick = () => {
            const inicio = inputFechaInicio.value;
            const fin = inputFechaFin.value;
            if (inicio && fin) {
                cargarAsistenciaDetallada(inicio, fin);
            }
        };
    }

    document.getElementById('btnPrev').onclick = () => {
        if (paginaActual > 1) {
            paginaActual--;
            renderizarTablaAsistencia();
            actualizarPaginacion();
        }
    };
    document.getElementById('btnNext').onclick = () => {
        const totalPaginas = Math.max(1, Math.ceil(empleadosFiltrados.length / empleadosPorPagina));
        if (paginaActual < totalPaginas) {
            paginaActual++;
            renderizarTablaAsistencia();
            actualizarPaginacion();
        }
    };

    document.getElementById('filtroEmpleado').addEventListener('change', () => {
        paginaActual = 1;
        renderizarTablaAsistencia();
        actualizarPaginacion();
    });

    // Descargar PDF
    document.getElementById('btnDescargarPDF').onclick = () => {
        descargarPDF();
    };

    // Descargar Excel (por página, por empleado, cada uno en una hoja)
    document.getElementById('btnDescargarExcel').onclick = () => {
        // Requiere SheetJS (xlsx) en tu HTML:
        // <script src="https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js"></script>
        exportarExcelPorPagina();
    };

    // Filtro de filas por página
    document.getElementById('filasPorPagina').addEventListener('change', function() {
        empleadosPorPagina = parseInt(this.value, 10);
        paginaActual = 1;
        renderizarTablaAsistencia();
        actualizarPaginacion();
    });
});

function mostrarPersonalizado() {
    document.getElementById('fechaInicio').style.display = '';
    document.getElementById('fechaFin').style.display = '';
    document.getElementById('btnOk').style.display = '';
}
function ocultarPersonalizado() {
    document.getElementById('fechaInicio').style.display = 'none';
    document.getElementById('fechaFin').style.display = 'none';
    document.getElementById('btnOk').style.display = 'none';
}

async function cargarEmpresa() {
    const res = await fetch('/api/empresa');
    const data = await res.json();
    document.getElementById('empresaNombre').textContent = data.Empresa || '---';
}

async function cargarPersonas() {
    const res = await fetch('/api/personas');
    const data = await res.json();
    document.getElementById('cantidadPersonas').textContent = data.CantidadPersonas || '---';
}

let chartMarcaciones, chartTopEmpleados, chartEmpleadosDepartamento;

async function cargarMarcaciones(filtro, inicio, fin) {
    let url = '/api/marcaciones?filtro=' + filtro;
    if (filtro === 'personalizado' && inicio && fin) {
        url += `&inicio=${inicio}&fin=${fin}`;
    }

    const res = await fetch(url);
    const data = await res.json();

    const labels = data.map(d => d.Fecha);
    const values = data.map(d => d.TotalMarcaciones);

    const ctx = document.getElementById('chartMarcaciones').getContext('2d');
    if (chartMarcaciones) chartMarcaciones.destroy();

    chartMarcaciones = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'Marcaciones',
                data: values,
                backgroundColor: 'rgba(75, 192, 192, 0.6)'
            }]
        },
        options: { responsive: true, scales: { y: { beginAtZero: true } } }
    });
}

async function cargarTopEmpleados() {
    const res = await fetch('/api/top-empleados');
    const data = await res.json();

    const labels = data.map(d => d.Empleado);
    const values = data.map(d => d.TotalMarcaciones);

    const ctx = document.getElementById('chartTopEmpleados').getContext('2d');
    if (chartTopEmpleados) chartTopEmpleados.destroy();

    chartTopEmpleados = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'Top 5 Empleados',
                data: values,
                backgroundColor: 'rgba(255, 99, 132, 0.6)'
            }]
        },
        options: { responsive: true, scales: { y: { beginAtZero: true } } }
    });
}

async function cargarEmpleadosPorDepartamento() {
    const res = await fetch('/api/empleados-por-departamento');
    const data = await res.json();

    const labels = data.map(d => d.Departamento);
    const values = data.map(d => d.TotalEmpleados);

    // Paleta de colores variada
    const colores = [
        '#4a90e2', '#f5a623', '#50e3c2', '#b8e986', '#e94e77',
        '#7ed6df', '#e17055', '#fdcb6e', '#00b894', '#636e72'
    ];
    const backgroundColors = labels.map((_, i) => colores[i % colores.length]);

    const ctx = document.getElementById('chartEmpleadosDepartamento').getContext('2d');
    if (chartEmpleadosDepartamento) chartEmpleadosDepartamento.destroy();

    chartEmpleadosDepartamento = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'Empleados por Departamento',
                data: values,
                backgroundColor: backgroundColors
            }]
        },
        options: {
            responsive: true,
            scales: {
                x: {
                    title: {
                        display: true,
                        text: 'Departamento',
                        font: { size: 18, weight: 'bold' }
                    },
                    ticks: {
                        font: { size: 16 }
                    }
                },
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Cantidad de Empleados',
                        font: { size: 18, weight: 'bold' }
                    },
                    ticks: {
                        font: { size: 16 }
                    }
                }
            },
            plugins: {
                legend: {
                    labels: {
                        font: { size: 17 }
                    }
                }
            }
        }
    });
}

let datosAsistencia = [];
let paginaActual = 1;
let empleadosPorPagina = 15;
let empleadosCombo = [];
let empleadosFiltrados = [];
let fechasTabla = [];

async function cargarAsistenciaDetallada(inicio, fin) {
    if (!inicio || !fin) {
        console.warn('Fechas inválidas para asistencia detallada');
        return;
    }

    try {
        const url = `/api/asistencia-detallada?inicio=${inicio}&fin=${fin}`;
        const res = await fetch(url);
        const data = await res.json();

        if (!Array.isArray(data)) {
            console.error('Respuesta inesperada:', data);
            return;
        }

        datosAsistencia = data;
        paginaActual = 1;

        cargarComboEmpleados();
        renderizarTablaAsistencia();
        actualizarPaginacion();
    } catch (error) {
        console.error('Error cargando asistencia detallada:', error);
    }
}

function cargarComboEmpleados() {
    const select = document.getElementById('filtroEmpleado');
    // Obtener empleados únicos
    const empleadosUnicos = {};
    datosAsistencia.forEach(row => {
        empleadosUnicos[row.idempleado] = {
            nombre: `${row.emp_firstname} ${row.emp_lastname}`,
            dni: row.emp_pin || ''
        };
    });
    empleadosCombo = Object.entries(empleadosUnicos).map(([id, obj]) => ({ id, ...obj }));

    // Limpiar y cargar opciones
    select.innerHTML = `<option value="">Todos</option>`;
    empleadosCombo.forEach(emp => {
        const option = document.createElement('option');
        option.value = emp.id;
        // Mostrar nombre y DNI juntos
        option.textContent = `${emp.nombre}  |  ${emp.dni}`;
        select.appendChild(option);
    });
}

function getDiaSemana(fechaISO) {
    // Corrige el reconocimiento de días para fechas en formato yyyy-MM-dd
    // y ajusta a zona local para evitar desfases por UTC
    const dias = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
    // Parse fechaISO como yyyy-MM-dd
    const [anio, mes, dia] = fechaISO.split('-').map(Number);
    // new Date(año, mes-1, día) crea la fecha en local
    const d = new Date(anio, mes - 1, dia);
    return { nombre: dias[d.getDay()], numero: d.getDay() };
}

function renderizarTablaAsistencia() {
    const select = document.getElementById('filtroEmpleado');
    const filtroId = select.value;
    // Agrupar por empleado y fechas
    const fechasSet = new Set();
    const empleadosMap = {};

    datosAsistencia.forEach(row => {
        if (filtroId && row.idempleado != filtroId) return;
        fechasSet.add(row.FechaAsistencia);
        const key = row.idempleado;
        if (!empleadosMap[key]) {
            empleadosMap[key] = {
                // idempleado: row.idempleado, // ID ya no se usa en la tabla
                emp_pin: row.emp_pin,
                emp_firstname: row.emp_firstname,
                emp_lastname: row.emp_lastname,
                dept_name: row.dept_name,
                fechas: {}
            };
        }
        empleadosMap[key].fechas[row.FechaAsistencia] = `${row.entrada || ''} - ${row.salida || ''}`;
    });

    // Ordenar fechas por día-mes-año
    fechasTabla = Array.from(fechasSet).sort((a, b) => {
        return new Date(a) - new Date(b);
    });
    empleadosFiltrados = Object.values(empleadosMap);

    // Paginación de empleados
    const totalPaginas = Math.max(1, Math.ceil(empleadosFiltrados.length / empleadosPorPagina));
    if (paginaActual > totalPaginas) paginaActual = totalPaginas;
    const inicio = (paginaActual - 1) * empleadosPorPagina;
    const fin = inicio + empleadosPorPagina;
    const empleadosPagina = empleadosFiltrados.slice(inicio, fin);

    // Renderizar tabla dinámica
    const thead = document.querySelector('#tablaAsistencia thead');
    const tbody = document.querySelector('#tablaAsistencia tbody');
    thead.innerHTML = '';
    tbody.innerHTML = '';

    let ths = `<tr>
        <th>N° Doc</th>
        <th>Nombre y Apellido</th>
        <th>Departamento</th>`;
    fechasTabla.forEach(f => {
        const { nombre } = getDiaSemana(f);
        const [yyyy, mm, dd] = f.split('-');
        const fechaFormateada = `${dd}-${mm}-${yyyy}`;
        ths += `<th>${nombre}<br>${fechaFormateada}</th>`;
    });
    ths += `</tr>`;
    thead.innerHTML = ths;

    // Filas
    empleadosPagina.forEach(emp => {
        let tds = `<td>${emp.emp_pin || ''}</td>
            <td>${emp.emp_firstname} ${emp.emp_lastname}</td>
            <td>${emp.dept_name}</td>`;
        fechasTabla.forEach(f => {
            const { numero } = getDiaSemana(f);
            let cellColor = '';
            let cellText = '';
            const marcacion = emp.fechas[f];
            let entrada = '', salida = '';
            if (numero === 6 || numero === 0) {
                // Sábado o domingo: mostrar datos pero sin color
                if (marcacion && marcacion.includes('-')) {
                    cellText = marcacion;
                } else if (marcacion && marcacion.trim() !== '-' && marcacion.trim() !== '') {
                    cellText = marcacion;
                } else {
                    cellText = '';
                }
                tds += `<td>${cellText}</td>`;
                return;
            }
            if (marcacion && marcacion.includes('-')) {
                [entrada, salida] = marcacion.split('-').map(x => x.trim());
                let cumpleHorario = false;
                let cumpleHoras = false;
                let horasTrabajadas = 0;
                if (entrada && salida) {
                    cumpleHorario = (entrada === '08:00' && salida === '17:00');
                    const entradaMin = parseInt(entrada.substr(0,2),10)*60 + parseInt(entrada.substr(3,2),10);
                    const salidaMin = parseInt(salida.substr(0,2),10)*60 + parseInt(salida.substr(3,2),10);
                    horasTrabajadas = (salidaMin - entradaMin);
                    if (entrada <= '13:00' && salida >= '14:00') {
                        horasTrabajadas -= 60;
                    }
                    cumpleHoras = horasTrabajadas >= 8*60;
                    if (cumpleHoras && salida > '17:00') {
                        cellColor = 'background:#d4f8d4;';
                    } else if (cumpleHorario && !cumpleHoras) {
                        cellColor = 'background:#ffe5b4;';
                    } else if (!cumpleHorario) {
                        cellColor = 'background:#ffcccc;';
                    }
                }
                cellText = marcacion;
            } else if (marcacion && marcacion.trim() !== '-' && marcacion.trim() !== '') {
                cellText = marcacion;
                cellColor = '';
            } else {
                cellText = '';
                cellColor = '';
            }
            tds += `<td style="${cellColor}">${cellText}</td>`;
        });
        const tr = document.createElement('tr');
        tr.innerHTML = tds;
        tbody.appendChild(tr);
    });

    // Mostrar paginación
    const paginacion = document.querySelector('.tabla-paginacion');
    if (paginacion) paginacion.style.display = '';
}

function actualizarPaginacion() {
    const totalPaginas = Math.max(1, Math.ceil(empleadosFiltrados.length / empleadosPorPagina));
    document.getElementById('paginaActual').textContent = paginaActual;
    document.getElementById('totalPaginas').textContent = totalPaginas;
    document.getElementById('btnPrev').disabled = paginaActual === 1;
    document.getElementById('btnNext').disabled = paginaActual === totalPaginas;
}

function descargarPDF() {
    const filtroId = document.getElementById('filtroEmpleado').value;

    function crearTablaPDFVertical(empleadosPagina, fechasTabla) {
        const tempTable = document.createElement('table');
        tempTable.style.width = '100%';
        tempTable.style.borderCollapse = 'collapse';
        tempTable.style.fontSize = '0.98rem';
        tempTable.style.marginBottom = '1.5rem';

        // Cabecera
        const thead = document.createElement('thead');
        let thRow = document.createElement('tr');
        ['N° DOC', 'NOMBRE Y APELLIDO', 'DEPARTAMENTO', 'DIA', 'FECHA', 'HORA DE ENTRADA', 'HORA DE SALIDA', ''].forEach((txt, idx, arr) => {
            const th = document.createElement('th');
            th.innerHTML = txt;
            th.style.padding = '7px 8px';
            th.style.background = '#eaf1fb';
            th.style.border = '1px solid #e0e7ef';
            th.style.fontWeight = 'bold';
            th.style.textAlign = 'center';
            th.style.fontSize = '1.05rem';
            // Si es la última columna (la raya), dale mismo grosor que las otras rayas laterales
            if (idx === arr.length - 1) {
                th.style.width = '1px';
                th.style.background = '#b6c6e0';
                th.style.borderLeft = '1px solid #e0e7ef';
                th.style.borderRight = 'none';
                th.innerHTML = '';
                th.style.padding = '0';
            }
            thRow.appendChild(th);
        });
        thead.appendChild(thRow);
        tempTable.appendChild(thead);

        // Filas
        const tbody = document.createElement('tbody');
        empleadosPagina.forEach(emp => {
            fechasTabla.forEach(f => {
                const { nombre } = getDiaSemana(f);
                const marcacion = emp.fechas && emp.fechas[f];
                let entrada = '', salida = '';
                if (marcacion && marcacion.includes('-')) {
                    [entrada, salida] = marcacion.split('-').map(x => x.trim());
                } else if (marcacion && marcacion.trim() !== '-' && marcacion.trim() !== '') {
                    entrada = marcacion;
                    salida = '';
                } else {
                    entrada = '';
                    salida = '';
                }
                const { numero } = getDiaSemana(f);
                if ((numero === 6 || numero === 0) && (!marcacion || marcacion.trim() === '' || marcacion.trim() === '-')) {
                    return;
                }
                const [yyyy, mm, dd] = f.split('-');
                const fechaFormateada = `${dd}-${mm}-${yyyy}`;
                const tr = document.createElement('tr');
                // N° DOC
                let td = document.createElement('td');
                td.textContent = emp.emp_pin || '';
                td.style.padding = '7px 8px';
                td.style.border = '1px solid #e0e7ef';
                td.style.textAlign = 'center';
                tr.appendChild(td);
                // NOMBRE Y APELLIDO
                td = document.createElement('td');
                td.textContent = `${emp.emp_firstname} ${emp.emp_lastname}`;
                td.style.padding = '7px 8px';
                td.style.border = '1px solid #e0e7ef';
                td.style.textAlign = 'left';
                tr.appendChild(td);
                // DEPARTAMENTO
                td = document.createElement('td');
                td.textContent = emp.dept_name;
                td.style.padding = '7px 8px';
                td.style.border = '1px solid #e0e7ef';
                td.style.textAlign = 'left';
                tr.appendChild(td);
                // DIA
                td = document.createElement('td');
                td.textContent = nombre;
                td.style.padding = '7px 8px';
                td.style.border = '1px solid #e0e7ef';
                td.style.textAlign = 'center';
                tr.appendChild(td);
                // FECHA
                td = document.createElement('td');
                td.textContent = fechaFormateada;
                td.style.padding = '7px 8px';
                td.style.border = '1px solid #e0e7ef';
                td.style.textAlign = 'center';
                tr.appendChild(td);
                // HORA DE ENTRADA
                td = document.createElement('td');
                td.textContent = entrada || '';
                td.style.padding = '7px 8px';
                td.style.border = '1px solid #e0e7ef';
                td.style.textAlign = 'center';
                tr.appendChild(td);
                // HORA DE SALIDA
                td = document.createElement('td');
                td.textContent = salida || '';
                td.style.padding = '7px 8px';
                td.style.border = '1px solid #e0e7ef';
                td.style.textAlign = 'center';
                tr.appendChild(td);
                // Raya lateral
                td = document.createElement('td');
                td.style.border = '1px solid #e0e7ef';
                td.style.background = '#b6c6e0';
                td.style.width = '1px';
                td.style.padding = '0';
                td.innerHTML = '';
                tr.appendChild(td);

                tbody.appendChild(tr);
            });
        });
        tempTable.appendChild(tbody);
        return tempTable;
    }

    // Determinar el rango de fechas seleccionado
    let fechaInicio = document.getElementById('fechaInicio')?.value;
    let fechaFin = document.getElementById('fechaFin')?.value;
    if (!fechaInicio || !fechaFin) {
        if (fechasTabla.length > 0) {
            fechaInicio = fechasTabla[0];
            fechaFin = fechasTabla[fechasTabla.length - 1];
        } else {
            fechaInicio = '';
            fechaFin = '';
        }
    }
    function formatFecha(fecha) {
        if (!fecha) return '';
        const [yyyy, mm, dd] = fecha.split('-');
        return `${dd}-${mm}-${yyyy}`;
    }
    const subtituloRango = (fechaInicio && fechaFin)
        ? `<div style="text-align:center;font-size:1.1rem;margin-bottom:0.7rem;">Del ${formatFecha(fechaInicio)} al ${formatFecha(fechaFin)}</div>`
        : '';

    function getTituloPDF(empleadosPagina, empleadosFiltrados) {
        if (filtroId) return 'Lista de asistencia Personal';
        const docs = new Set();
        (empleadosPagina || []).forEach(emp => docs.add(emp.emp_pin));
        if (docs.size === 1) return 'Lista de asistencia Personal';
        const allDocs = new Set();
        (empleadosFiltrados || []).forEach(emp => allDocs.add(emp.emp_pin));
        return allDocs.size > 1 ? 'Lista de asistencia del Personal GM' : 'Lista de asistencia Personal';
    }

    if (!filtroId) {
        // Si está en "Todos", mostrar todos los empleados y días en formato vertical
        const paginaActualOriginal = paginaActual;
        const totalPaginas = Math.max(1, Math.ceil(empleadosFiltrados.length / empleadosPorPagina));
        let htmlCompleto = '';

        for (let pag = 1; pag <= totalPaginas; pag++) {
            paginaActual = pag;
            const inicio = (pag - 1) * empleadosPorPagina;
            const fin = inicio + empleadosPorPagina;
            const empleadosPagina = empleadosFiltrados.slice(inicio, fin);

            const titulo = getTituloPDF(empleadosPagina, empleadosFiltrados);
            htmlCompleto += `<h2 style="text-align:center;font-family:'Segoe UI',Arial,sans-serif;margin-bottom:0.5rem;">${titulo}</h2>`;
            htmlCompleto += subtituloRango;
            htmlCompleto += `<div style="page-break-inside:avoid;">`;
            // Eliminado el texto de página
            // htmlCompleto += `<h4 style="margin-top:1.2rem;text-align:center;">Página ${pag} de ${totalPaginas}</h4>`;
            htmlCompleto += crearTablaPDFVertical(empleadosPagina, fechasTabla).outerHTML;
            htmlCompleto += `</div>`;
        }

        const contenedor = document.createElement('div');
        contenedor.innerHTML = htmlCompleto;
        document.body.appendChild(contenedor);

        const opt = {
            margin:       0.3,
            filename:     'asistencia_detallada.pdf',
            image:        { type: 'jpeg', quality: 0.98 },
            html2canvas:  { scale: 2 },
            jsPDF:        { unit: 'in', format: 'a4', orientation: 'landscape' }
        };
        html2pdf().set(opt).from(contenedor).save().then(() => {
            document.body.removeChild(contenedor);
            paginaActual = paginaActualOriginal;
            renderizarTablaAsistencia();
            actualizarPaginacion();
        });
    } else {
        const empleadosPagina = empleadosFiltrados.slice((paginaActual - 1) * empleadosPorPagina, paginaActual * empleadosPorPagina);
        const titulo = getTituloPDF(empleadosPagina, empleadosFiltrados);
        const contenedor = document.createElement('div');
        contenedor.innerHTML = `<h2 style="text-align:center;font-family:'Segoe UI',Arial,sans-serif;margin-bottom:0.5rem;">${titulo}</h2>${subtituloRango}`;
        contenedor.appendChild(crearTablaPDFVertical(empleadosPagina, fechasTabla));
        document.body.appendChild(contenedor);

        const opt = {
            margin:       0.3,
            filename:     'asistencia_detallada.pdf',
            image:        { type: 'jpeg', quality: 0.98 },
            html2canvas:  { scale: 2 },
            jsPDF:        { unit: 'in', format: 'a4', orientation: 'landscape' }
        };
        html2pdf().set(opt).from(contenedor).save().then(() => {
            document.body.removeChild(contenedor);
        });
    }
}

// Reemplaza el evento del botón Excel por la exportación real por página
document.getElementById('btnDescargarExcel').onclick = () => {
    exportarExcelPorPagina();
};

function exportarExcelPorPagina() {
    // Obtén los empleados de la página actual
    const inicio = (paginaActual - 1) * empleadosPorPagina;
    const fin = inicio + empleadosPorPagina;
    const empleadosPagina = empleadosFiltrados.slice(inicio, fin);

    if (!empleadosPagina.length) return;

    const fechas = fechasTabla;
    const wb = XLSX.utils.book_new();

    empleadosPagina.forEach(emp => {
        const ws_data = [
            ['N° DOC', emp.emp_pin || ''],
            ['NOMBRE Y APELLIDO', `${emp.emp_firstname} ${emp.emp_lastname}`],
            ['DEPARTAMENTO', emp.dept_name || ''],
            [],
            ['DIA', 'FECHA', 'HORA DE ENTRADA', 'HORA DE SALIDA']
        ];

        fechas.forEach(f => {
            const { nombre } = getDiaSemana(f);
            const marcacion = emp.fechas && emp.fechas[f];
            let entrada = '', salida = '';
            if (marcacion && marcacion.includes('-')) {
                [entrada, salida] = marcacion.split('-').map(x => x.trim());
            } else if (marcacion && marcacion.trim() !== '-' && marcacion.trim() !== '') {
                entrada = marcacion;
                salida = '';
            }
            const { numero } = getDiaSemana(f);
            if ((numero === 6 || numero === 0) && (!marcacion || marcacion.trim() === '' || marcacion.trim() === '-')) {
                return;
            }
            const [yyyy, mm, dd] = f.split('-');
            const fechaFormateada = `${dd}-${mm}-${yyyy}`;
            ws_data.push([
                nombre,
                fechaFormateada,
                entrada || '',
                salida || ''
            ]);
        });

        let nombreHoja = `${emp.emp_firstname} ${emp.emp_lastname}`.trim();
        if (nombreHoja.length > 28) nombreHoja = emp.emp_pin || 'Empleado';

        const ws = XLSX.utils.aoa_to_sheet(ws_data);
        XLSX.utils.book_append_sheet(wb, ws, nombreHoja || 'Empleado');
    });

    XLSX.writeFile(wb, 'asistencia_por_pagina.xlsx');
}

