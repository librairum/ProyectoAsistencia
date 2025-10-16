const { poolPromise, sql } = require('../db/conexion');

const loginUser = async (req, res) => {
    // Acceso permitido sin autenticación
    res.json({ success: true, message: '🔐 Acceso permitido' });
};

module.exports = {
    loginUser
};
/*
async function loginUser(req, res) {
    const { username, password } = req.body;

    try {
        const pool = await poolPromise;
        const result = await pool.request()
            .input('username', sql.VarChar, username)
            .input('password', sql.VarChar, password)
            .output('is_authenticated', sql.Int)
            .execute('sp_login_user');

        const auth = result.output.is_authenticated;
        if (auth === 1) {
            res.json({ success: true, message: '🔐 Acceso permitido' });
        } else {
            res.status(401).json({ success: false, message: '❗ Usuario o contraseña incorrectos' });
        }

    } catch (err) {
        console.error('❌ Error en login:', err);
        res.status(500).json({ success: false, error: 'Error interno' });
    }
}

module.exports = {
    loginUser
};*/
