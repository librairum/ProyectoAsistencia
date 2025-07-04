const express = require('express');
const { loginUser } = require('./controllers/loginController');
const sql = require('mssql');
const path = require('path');

const app = express();
app.use(express.json());

const dbConfig = {
    user: 'sa',
    password: 'admin123456',
    server: 'localhost',
    database: 'Asistencia',
    options: { encrypt: false }
};

async function getConnection() {
    if (!sql.pool) {
        sql.pool = await sql.connect(dbConfig);
    }
    return sql.pool;
}

app.post('/login', loginUser);

const PORT = 3000;
app.listen(PORT, () => {
    console.log(`🚀 Servidor corriendo en http://localhost:${PORT}`);
});

app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

app.get('/api/asistencia-detallada', async (req, res) => {
    try {
        const { inicio, fin } = req.query;
        if (!inicio || !fin) {
            return res.status(400).json({ error: 'Fechas requeridas' });
        }

        const pool = await getConnection();
        const query = `
            SELECT 
                CONVERT(varchar(5), punch_time, 108) AS hora,
                CONVERT(varchar(10), punch_time, 120) AS FechaAsistencia,
                emp_id AS idempleado
            INTO #tblmarcaciones
            FROM dbo.att_punches;

            WITH cte AS (
                SELECT  
                    idempleado,
                    FechaAsistencia,
                    MIN(hora) AS entrada,
                    MAX(hora) AS salida
                FROM #tblmarcaciones
                GROUP BY idempleado, FechaAsistencia
            )
            SELECT 
                cte.idempleado,
                personal.emp_pin, -- DNI agregado
                personal.emp_firstname,
                personal.emp_lastname,
                dept.dept_name,
                cte.FechaAsistencia,
                cte.entrada,
                cte.salida
            FROM cte
            INNER JOIN hr_employee personal ON personal.id = cte.idempleado
            INNER JOIN hr_department dept ON dept.id = personal.emp_dept
            WHERE CAST(cte.FechaAsistencia AS DATE) BETWEEN '${inicio}' AND '${fin}'
            ORDER BY idempleado ASC, cte.FechaAsistencia ASC;
        `;

        const result = await pool.request().query(query);
        res.json(result.recordset);
    } catch (err) {
        console.error('Error en asistencia detallada:', err);
        res.status(500).json({ error: 'Error al obtener asistencia detallada' });
    }
});


// EXISTENTE
app.get('/api/empresa', async (req, res) => {
    try {
        const pool = await getConnection();
        const result = await pool.request().query("SELECT TOP 1 cmp_name AS Empresa FROM hr_company;");
        res.json(result.recordset[0] || { Empresa: null });
    } catch (err) {
        res.status(500).json({ error: 'Error al obtener empresa' });
    }
});

app.get('/api/personas', async (req, res) => {
    try {
        const pool = await getConnection();
        const result = await pool.request().query("SELECT COUNT(*) AS CantidadPersonas FROM hr_employee WHERE emp_active = 1;");
        res.json(result.recordset[0] || { CantidadPersonas: 0 });
    } catch (err) {
        res.status(500).json({ error: 'Error al obtener cantidad de personas' });
    }
});
// 

app.get('/api/marcaciones', async (req, res) => {
    try {
        const pool = await getConnection();
        let query = '';
        const filtro = req.query.filtro;

        if (filtro === 'hoy') {
            query = `
                SELECT FORMAT(punch_time, 'yyyy-MM-dd') AS Fecha, COUNT(*) AS TotalMarcaciones
                FROM att_punches
                WHERE CAST(punch_time AS DATE) = CAST(GETDATE() AS DATE)
                GROUP BY FORMAT(punch_time, 'yyyy-MM-dd')
                ORDER BY Fecha;
            `;
        } else if (filtro === '7dias') {
            query = `
                SELECT FORMAT(punch_time, 'yyyy-MM-dd') AS Fecha, COUNT(*) AS TotalMarcaciones
                FROM att_punches
                WHERE punch_time >= DATEADD(DAY, -6, CAST(GETDATE() AS DATE))
                GROUP BY FORMAT(punch_time, 'yyyy-MM-dd')
                ORDER BY Fecha;
            `;
        } else if (filtro === '30dias') {
            query = `
                SELECT FORMAT(punch_time, 'yyyy-MM-dd') AS Fecha, COUNT(*) AS TotalMarcaciones
                FROM att_punches
                WHERE punch_time >= DATEADD(DAY, -29, CAST(GETDATE() AS DATE))
                GROUP BY FORMAT(punch_time, 'yyyy-MM-dd')
                ORDER BY Fecha;
            `;
        } else if (filtro === 'mes') {
            query = `
                SELECT FORMAT(punch_time, 'yyyy-MM-dd') AS Fecha, COUNT(*) AS TotalMarcaciones
                FROM att_punches
                WHERE MONTH(punch_time) = MONTH(GETDATE()) AND YEAR(punch_time) = YEAR(GETDATE())
                GROUP BY FORMAT(punch_time, 'yyyy-MM-dd')
                ORDER BY Fecha;
            `;
        } else if (filtro === 'personalizado') {
            const inicio = req.query.inicio;
            const fin = req.query.fin;
            if (!inicio || !fin) {
                return res.status(400).json({ error: 'Fechas requeridas' });
            }
            query = `
                SELECT FORMAT(punch_time, 'yyyy-MM-dd') AS Fecha, COUNT(*) AS TotalMarcaciones
                FROM att_punches
                WHERE punch_time BETWEEN '${inicio}' AND '${fin}'
                GROUP BY FORMAT(punch_time, 'yyyy-MM-dd')
                ORDER BY Fecha;
            `;
        } else {
            return res.json([]);
        }

        const result = await pool.request().query(query);
        res.json(result.recordset);
    } catch (err) {
        res.status(500).json({ error: 'Error al obtener marcaciones' });
    }
});

app.get('/api/top-empleados', async (req, res) => {
    try {
        const pool = await getConnection();
        const result = await pool.request().query(`
            SELECT TOP 5 
                e.emp_firstname + ' ' + e.emp_lastname AS Empleado,
                COUNT(*) AS TotalMarcaciones
            FROM att_punches p
            INNER JOIN hr_employee e ON p.emp_id = e.id
            WHERE MONTH(p.punch_time) = MONTH(GETDATE()) AND YEAR(p.punch_time) = YEAR(GETDATE())
            GROUP BY e.emp_firstname, e.emp_lastname
            ORDER BY TotalMarcaciones DESC;
        `);
        res.json(result.recordset);
    } catch (err) {
        res.status(500).json({ error: 'Error al obtener top de empleados' });
    }
});

app.get('/api/empleados-por-departamento', async (req, res) => {
    try {
        const pool = await getConnection();
        const result = await pool.request().query(`
            SELECT 
                d.dept_name AS Departamento,
                STRING_AGG(e.emp_firstname + ' ' + e.emp_lastname, ', ') AS Empleados,
                COUNT(*) AS TotalEmpleados
            FROM hr_employee e
            INNER JOIN hr_department d ON e.emp_dept = d.id
            WHERE e.emp_active = 1
            GROUP BY d.dept_name
            ORDER BY d.dept_name;
        `);
        res.json(result.recordset);
    } catch (err) {
        res.status(500).json({ error: 'Error al obtener empleados por departamento' });
    }
});
