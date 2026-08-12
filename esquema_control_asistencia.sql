-- =========================================================
-- ESQUEMA DE BASE DE DATOS: CONTROL DE ASISTENCIA / PLANILLA
-- Basado en la planilla "PLANILLA_16_31" (marcado, roles de turno,
-- reporte de marcaciones y tabla dinámica)
-- Motor objetivo: PostgreSQL (ajustar tipos si usas MySQL/SQL Server)
-- =========================================================

-- 1. CATÁLOGOS BÁSICOS -----------------------------------

CREATE TABLE edificios (
    id              SERIAL PRIMARY KEY,
    codigo          VARCHAR(20) UNIQUE,        -- ej: 'E_031'
    nombre          VARCHAR(150) NOT NULL,     -- ej: 'HOSPITAL MUNICIPAL EL BAJIO'
    codigo_partida  VARCHAR(60)                -- ej: '001-009-003-004-001-004-000-000-000-000'
);

CREATE TABLE cargos (
    id           SERIAL PRIMARY KEY,
    nombre_cargo VARCHAR(150) NOT NULL UNIQUE  -- ej: 'MEDICO ESPECIALISTA EN UROLOGIA'
);

-- 2. EMPLEADOS --------------------------------------------

CREATE TABLE empleados (
    ci               BIGINT PRIMARY KEY,        -- carnet de identidad, viene tal cual de la planilla
    nombre_completo  VARCHAR(200) NOT NULL,
    tipo_contrato    VARCHAR(60),               -- 'Estructura', 'Contrato', 'Contrato(Salud)'
    cargo_id         INT REFERENCES cargos(id),
    edificio_id      INT REFERENCES edificios(id),
    servicio         VARCHAR(100),              -- ej: 'UROLOGIA', 'GOB', 'CONSULT.EXT'
    observaciones    TEXT,                      -- notas tipo "CORREGIR C.I. ..."
    activo           BOOLEAN NOT NULL DEFAULT TRUE
);

-- 3. TURNOS (para no repetir texto libre en cada fila) -----

-- Catálogo maestro de patrones de horario (uno por cada texto distinto de "ROL")
CREATE TABLE turnos (
    id                    SERIAL PRIMARY KEY,
    nombre                VARCHAR(150) NOT NULL,   -- nombre corto, ej: 'Mañana Lun-Vie 8-16'
    descripcion_original  TEXT NOT NULL            -- el texto crudo tal cual venía, ej:
                                                     -- 'MAR 19:30-7:30 VIE(MENOS 3) 7:30-19:30'
);

-- Desglose día a día de cada turno (evita parsear texto en cada consulta)
CREATE TABLE turno_horarios (
    id                 SERIAL PRIMARY KEY,
    turno_id           INT NOT NULL REFERENCES turnos(id) ON DELETE CASCADE,
    dia_semana         SMALLINT NOT NULL,   -- 1=lunes ... 7=domingo
    hora_entrada       TIME NOT NULL,
    hora_salida        TIME NOT NULL,
    cruza_medianoche   BOOLEAN NOT NULL DEFAULT FALSE,
    horas_lactancia    NUMERIC(4,2)         -- si aplica (columna "LAC")
);

-- Qué turno tiene cada empleado y en qué periodo (permite historial de cambios)
CREATE TABLE asignaciones_turno (
    id            SERIAL PRIMARY KEY,
    ci            BIGINT NOT NULL REFERENCES empleados(ci),
    turno_id      INT NOT NULL REFERENCES turnos(id),
    fecha_inicio  DATE NOT NULL,
    fecha_fin     DATE,                      -- NULL = turno vigente
    UNIQUE (ci, fecha_inicio)
);

-- 4. MARCACIONES (registro crudo del reloj biométrico) -----
-- Corresponde a las hojas "reporte" / "reporte (2)" / "TABLA"

CREATE TABLE marcaciones (
    id            BIGSERIAL PRIMARY KEY,
    ci            BIGINT NOT NULL REFERENCES empleados(ci),
    edificio_id   INT REFERENCES edificios(id),
    fecha         DATE NOT NULL,
    hora          TIME NOT NULL,
    tipo          CHAR(1) CHECK (tipo IN ('E','S')),  -- Entrada / Salida
    origen        VARCHAR(30) DEFAULT 'BIOMETRICO',   -- por si luego hay otras fuentes
    UNIQUE (ci, fecha, hora, tipo)
);

CREATE INDEX idx_marcaciones_ci_fecha ON marcaciones (ci, fecha);

-- 5. PLANILLAS (el reporte final impreso, uno por quincena/edificio) ---

CREATE TABLE planillas (
    id               SERIAL PRIMARY KEY,
    numero_planilla  VARCHAR(30),           -- ej: '026_5'
    edificio_id      INT REFERENCES edificios(id),
    fecha_inicio     DATE NOT NULL,
    fecha_fin        DATE NOT NULL,
    fecha_generacion TIMESTAMP DEFAULT NOW()
);

-- Detalle día a día por empleado, tal como se imprime en la planilla
-- (se puede recalcular desde marcaciones + asignaciones_turno,
--  pero conviene guardarlo como "foto" histórica de lo entregado)
CREATE TABLE planilla_detalle (
    id               BIGSERIAL PRIMARY KEY,
    planilla_id      INT NOT NULL REFERENCES planillas(id) ON DELETE CASCADE,
    ci               BIGINT NOT NULL REFERENCES empleados(ci),
    fecha            DATE NOT NULL,
    entrada_manana   TIME,
    salida_manana    TIME,
    entrada_tarde    TIME,
    salida_tarde     TIME,
    observacion      VARCHAR(200),          -- ej: 'LUN-VIE 8:00-16:00'
    UNIQUE (planilla_id, ci, fecha)
);

-- =========================================================
-- VISTA DE APOYO: reconstruye el detalle de planilla
-- directamente desde las marcaciones crudas (opcional,
-- útil si no quieres materializar planilla_detalle)
-- =========================================================

CREATE VIEW v_marcaciones_dia AS
SELECT
    m.ci,
    e.nombre_completo,
    m.fecha,
    MIN(CASE WHEN m.tipo = 'E' AND m.hora < TIME '13:00' THEN m.hora END) AS entrada_manana,
    MIN(CASE WHEN m.tipo = 'S' AND m.hora < TIME '13:00' THEN m.hora END) AS salida_manana,
    MIN(CASE WHEN m.tipo = 'E' AND m.hora >= TIME '13:00' THEN m.hora END) AS entrada_tarde,
    MIN(CASE WHEN m.tipo = 'S' AND m.hora >= TIME '13:00' THEN m.hora END) AS salida_tarde
FROM marcaciones m
JOIN empleados e ON e.ci = m.ci
GROUP BY m.ci, e.nombre_completo, m.fecha;
