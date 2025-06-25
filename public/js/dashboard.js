document.addEventListener('DOMContentLoaded', () => {
    cargarEmpresa();
    cargarPersonas();
    cargarMarcaciones('7dias');
    cargarTopEmpleados();
    cargarEmpleadosPorDepartamento();

    const hoy = new Date();
    const fin = hoy.toISOString().split('T')[0]; // yyyy-MM-dd
    const inicioDate = new Date(hoy);
    inicioDate.setDate(hoy.getDate() - 6);
    const inicio = inicioDate.toISOString().split('T')[0];

    cargarAsistenciaDetallada(inicio, fin);


    document.getElementById('btnHoy').onclick = () => {
        ocultarPersonalizado();
        cargarMarcaciones('hoy');
    };
    document.getElementById('btn7dias').onclick = () => {
        ocultarPersonalizado();
        cargarMarcaciones('7dias');
    };
    document.getElementById('btn30dias').onclick = () => {
        ocultarPersonalizado();
        cargarMarcaciones('30dias');
    };
    document.getElementById('btnMes').onclick = () => {
        ocultarPersonalizado();
        cargarMarcaciones('mes');
    };
    document.getElementById('btnPersonalizado').onclick = () => mostrarPersonalizado();
    document.getElementById('btnOk').onclick = () => {
        const inicio = document.getElementById('fechaInicio').value;
        const fin = document.getElementById('fechaFin').value;
        if (inicio && fin) {
            cargarMarcaciones('personalizado', inicio, fin);
            cargarAsistenciaDetallada(inicio, fin); // ← llamado con fechas
        }
    };

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

    const ctx = document.getElementById('chartEmpleadosDepartamento').getContext('2d');
    if (chartEmpleadosDepartamento) chartEmpleadosDepartamento.destroy();

    chartEmpleadosDepartamento = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'Empleados por Departamento',
                data: values,
                backgroundColor: 'rgba(255, 206, 86, 0.6)'
            }]
        },
        options: { responsive: true, scales: { y: { beginAtZero: true } } }
    });
}

let datosAsistencia = [];
let paginaActual = 1;
const empleadosPorPagina = 5;
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
        empleadosUnicos[row.idempleado] = `${row.emp_firstname} ${row.emp_lastname}`;
    });
    empleadosCombo = Object.entries(empleadosUnicos).map(([id, nombre]) => ({ id, nombre }));

    // Limpiar y cargar opciones
    select.innerHTML = `<option value="">Todos</option>`;
    empleadosCombo.forEach(emp => {
        const option = document.createElement('option');
        option.value = emp.id;
        option.textContent = emp.nombre;
        select.appendChild(option);
    });
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
                idempleado: row.idempleado,
                emp_firstname: row.emp_firstname,
                emp_lastname: row.emp_lastname,
                dept_name: row.dept_name,
                fechas: {}
            };
        }
        empleadosMap[key].fechas[row.FechaAsistencia] = `${row.entrada || ''} - ${row.salida || ''}`;
    });

    fechasTabla = Array.from(fechasSet).sort();
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

    // Cabecera
    let ths = `<tr>
        <th>ID</th>
        <th>Empleado</th>
        <th>Depto</th>`;
    fechasTabla.forEach(f => {
        ths += `<th>${f}</th>`;
    });
    ths += `</tr>`;
    thead.innerHTML = ths;

    // Filas
    empleadosPagina.forEach(emp => {
        let tds = `<td>${emp.idempleado}</td>
            <td>${emp.emp_firstname} ${emp.emp_lastname}</td>
            <td>${emp.dept_name}</td>`;
        fechasTabla.forEach(f => {
            tds += `<td>${emp.fechas[f] || '-'}</td>`;
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

document.addEventListener('DOMContentLoaded', () => {
    // ...existing code...

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

    // ...existing code...
});


