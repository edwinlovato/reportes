<?php
include("php/check_sess.php");
include('adodb/adodb.inc.php'); 
include_once('adodb/adodb-pager.inc.php');
include_once('conf/dbconn_param.php');

$db = ADONewConnection($dbdriver);
$db->Connect($server, $user, $password, $database);
$db->SetFetchMode(ADODB_FETCH_ASSOC);

// Procesar filtros
$fecha_inicio = isset($_POST['fecha_inicio']) ? $_POST['fecha_inicio'] : date('Y-m-01');
$fecha_fin = isset($_POST['fecha_fin']) ? $_POST['fecha_fin'] : date('Y-m-d');
$cconsigna = isset($_POST['cconsigna']) ? trim($_POST['cconsigna']) : '';
$cnombre = isset($_POST['cnombre']) ? trim($_POST['cnombre']) : '';
$fue_manifiesto = isset($_POST['fue_manifiesto']) ? trim($_POST['fue_manifiesto']) : '';
$invoice = isset($_POST['invoice']) ? trim($_POST['invoice']) : '';
$accion = isset($_POST['accion']) ? $_POST['accion'] : '';

// Parámetros de paginación
$registros_por_pagina = isset($_POST['registros_por_pagina']) ? (int)$_POST['registros_por_pagina'] : 20;
$pagina_actual = isset($_POST['pagina']) ? (int)$_POST['pagina'] : 1;
$offset = ($pagina_actual - 1) * $registros_por_pagina;

// Construir consulta con filtros
$sql_base = "SELECT DATE_FORMAT(facturas.`Fecha_facturacion`,'%d/%m/%Y') AS fecha, 
    DATE_FORMAT(DATE_ADD(Fecha_facturacion,INTERVAL 1 DAY),'%d/%m/%Y') AS fecha_vuelo, 
    facturas.invoice, facturas.`NUM_PACK`, 
    CONCAT('CE',clientes.tar) AS cod_ventas, 
    '' AS cod_contabilidad, 
    clientes.`CCONSIGNA` AS cliente,
    clientes.`CNOMBRE` AS subcliente, 
    facturas.`tipoventa`, 
    clientes.`CZONA` AS mercado, 
    fue.`fue_ingreso`, 
    fue.`fue_caduca`,
    fue.`fue_manifiesto` AS dae, 
    SUBSTRING_INDEX(facturas.`guia_aerea`,'/',1) AS guia_madre,
    SUBSTRING_INDEX(facturas.`guia_aerea`,'/',-1) AS guia_hija, 
    facturas.`NUM_FACTURA` AS fac_sri, 
    facturas.autorizacion, 
    fue_pais AS pais, 
    carguera, 
    facturas.`lineaaerea`,
    facturas.usd, 
    SUM(tallos) AS tallos, 
    SUM(ramos) AS ramos,
    SUM(cajas) AS cajas, 
    SUM(hb)*2 AS hb,
    SUM(qb)*4 AS qb,
    SUM(ob)*8 AS ob, 
    SUM(hb+qb+ob) AS fulles, 
    peso_neto, 
    peso_bruto
FROM facturas
LEFT JOIN fue ON fue.`fue_numero`=facturas.`FUE`
INNER JOIN (
    SELECT NUM_PACK, caja, tipo_caja, SUM(num_tallos) AS tallos, COUNT(*) AS ramos,
    COUNT(DISTINCT caja) AS cajas, 
    IF(MID(`tipo_caja`,1,2)='HB',COUNT(DISTINCT caja)*0.500,0) AS HB, 
    IF(MID(`tipo_caja`,1,2)='QB',COUNT(DISTINCT caja)*0.250,0) AS QB, 
    IF(MID(`tipo_caja`,1,2)='EB',COUNT(DISTINCT caja)*0.125,0) AS OB	
    FROM bunche WHERE estado='V' GROUP BY NUM_PACK, caja, tipo_caja	
) AS bunche ON bunche.`NUM_PACK`=facturas.`NUM_PACK`
INNER JOIN clientes ON clientes.`COD_CLIENT`=facturas.`cod_client`
INNER JOIN (SELECT cod_prove, nombre AS carguera FROM proveedores) AS proveedores 
    ON proveedores.`COD_PROVE`=facturas.`cod_proveedor`
WHERE Fecha_facturacion>=? AND Fecha_facturacion<=?";

$params = array($fecha_inicio, $fecha_fin);

// Agregar filtros adicionales
if (!empty($cconsigna)) {
    $sql_base .= " AND clientes.`CCONSIGNA` LIKE ?";
    $params[] = "%$cconsigna%";
}
if (!empty($cnombre)) {
    $sql_base .= " AND clientes.`CNOMBRE` LIKE ?";
    $params[] = "%$cnombre%";
}
if (!empty($fue_manifiesto)) {
    $sql_base .= " AND fue.`fue_manifiesto` LIKE ?";
    $params[] = "%$fue_manifiesto%";
}
if (!empty($invoice)) {
    $sql_base .= " AND facturas.invoice LIKE ?";
    $params[] = "%$invoice%";
}

$sql_base .= " GROUP BY invoice ORDER BY facturas.Fecha_facturacion DESC";

// Consulta para contar total de registros
$sql_count = str_replace("SELECT DATE_FORMAT(facturas.`Fecha_facturacion`,'%%d/%m/%Y') AS fecha, 
    DATE_FORMAT(DATE_ADD(Fecha_facturacion,INTERVAL 1 DAY),'%d/%m/%Y') AS fecha_vuelo, 
    facturas.invoice, facturas.`NUM_PACK`, 
    CONCAT('CE',clientes.tar) AS cod_ventas, 
    clientes.id_nif AS cod_contabilidad, 
    clientes.`CCONSIGNA` AS cliente,
    clientes.`CNOMBRE` AS subcliente, 
    facturas.`tipoventa`, 
    clientes.`CZONA` AS mercado, 
    fue.`fue_ingreso`, 
    fue.`fue_caduca`,
    fue.`fue_manifiesto` AS dae, 
    SUBSTRING_INDEX(facturas.`guia_aerea`,'/',1) AS guia_madre,
    SUBSTRING_INDEX(facturas.`guia_aerea`,'/',-1) AS guia_hija, 
    facturas.`NUM_FACTURA` AS fac_sri, 
    facturas.autorizacion, 
    fue_pais AS pais, 
    carguera, 
    facturas.`lineaaerea`,
    facturas.usd, 
    SUM(tallos) AS tallos, 
    SUM(ramos) AS ramos,
    SUM(cajas) AS cajas, 
    SUM(hb)*2 AS hb,
    SUM(qb)*4 AS qb,
    SUM(ob)*8 AS ob, 
    SUM(hb+qb+ob) AS fulles, 
    peso_neto, 
    peso_bruto", "COUNT(DISTINCT facturas.invoice) as total", $sql_base);

$rs_count = $db->Execute($sql_count, $params);
$total_registros = $rs_count ? $rs_count->fields['total'] : 0;
$total_paginas = ceil($total_registros / $registros_por_pagina);

// Agregar LIMIT a la consulta principal
$sql_paginado = $sql_base . " LIMIT $offset, $registros_por_pagina";

// Ejecutar consulta con manejo de errores
$rs = $db->Execute($sql_paginado, $params);
if (!$rs) {
    echo "<div class='notification is-danger'>Error en la consulta: " . $db->ErrorMsg() . "</div>";
    $rs = false;
}

// Funciones para exportación
function exportarExcel($db, $sql, $params, $filename) {
    header('Content-Type: application/vnd.ms-excel');
    header('Content-Disposition: attachment;filename="' . $filename . '.xls"');
    header('Cache-Control: max-age=0');
    
    echo "<table border='1'>";
    echo "<tr>";
    echo "<th>Fecha</th><th>Fecha Vuelo</th><th>Invoice</th><th>Num Pack</th>";
    echo "<th>Cod Ventas</th><th>Cod Contabilidad</th><th>Cliente</th><th>Subcliente</th>";
    echo "<th>Tipo Venta</th><th>Mercado</th><th>DAE</th><th>Guía Madre</th>";
    echo "<th>Guía Hija</th><th>Fac SRI</th><th>País</th><th>Carguera</th>";
    echo "<th>Línea Aérea</th><th>USD</th><th>Tallos</th><th>Ramos</th>";
    echo "<th>Cajas</th><th>HB</th><th>QB</th><th>OB</th><th>Fulles</th>";
    echo "<th>Peso Neto</th><th>Peso Bruto</th>";
    echo "</tr>";
    
    $rs = $db->Execute($sql, $params);
    if ($rs) {
        while (!$rs->EOF) {
            echo "<tr>";
            foreach ($rs->fields as $field) {
                echo "<td>" . htmlspecialchars($field ?? '') . "</td>";
            }
            echo "</tr>";
            $rs->MoveNext();
        }
    }
    echo "</table>";
    exit;
}

function generarGraficosData($db, $sql, $params, $tipo) {
    $rs = $db->Execute($sql, $params);
    $data = array();
    
    // Verificar si la consulta fue exitosa
    if (!$rs) {
        error_log("Error en consulta SQL: " . $db->ErrorMsg());
        return $data;
    }
    
    // Convertir recordset a array para usar foreach
    $resultados = array();
    while (!$rs->EOF) {
        $resultados[] = $rs->fields;
        $rs->MoveNext();
    }
    
    // Procesar datos con foreach
    foreach ($resultados as $row) {
        $key = '';
        switch($tipo) {
            case 'czona':
                $key = isset($row['mercado']) ? $row['mercado'] : 'Sin clasificar';
                break;
            case 'cconsigna':
                $key = isset($row['cliente']) ? $row['cliente'] : 'Sin cliente';
                break;
            case 'fue_manifiesto':
                $key = isset($row['dae']) ? $row['dae'] : 'Sin DAE';
                break;
        }
        
        if (empty($key)) $key = 'Sin datos';
        
        if (!isset($data[$key])) {
            $data[$key] = array('tallos' => 0, 'valor' => 0);
        }
        $data[$key]['tallos'] += (int)($row['tallos'] ?? 0);
        $data[$key]['valor'] += (float)($row['usd'] ?? 0);
    }
    
    return $data;
}

// Procesar acciones
if ($accion == 'exportar_excel') {
    exportarExcel($db, $sql_base, $params, 'reporte_facturas_' . date('Y-m-d'));
}

// Obtener datos para gráficos
$grafico_czona = generarGraficosData($db, $sql_base, $params, 'czona');
$grafico_cconsigna = generarGraficosData($db, $sql_base, $params, 'cconsigna');
$grafico_manifiesto = generarGraficosData($db, $sql_base, $params, 'fue_manifiesto');

?>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistema de Reportes</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bulma@0.9.4/css/bulma.min.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>

<style>
        .chart-container {
            position: relative;
            width: 100%; /* Ensure it takes full width of its parent box */
            transition: height 0.3s ease; /* Smooth height transition */
        }
        .chart-size-small { height: 250px; } /* Adjusted for better visibility */
        .chart-size-medium { height: 350px; } /* Adjusted for better visibility */
        .chart-size-large { height: 450px; } /* Adjusted for better visibility */
        
        .chart-controls {
            margin-bottom: 1rem;
            display: flex;
            gap: 0.5rem;
            align-items: center;
            flex-wrap: wrap;
        }
        
        .control-group {
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }
        
        .control-label {
            font-size: 0.875rem;
            font-weight: 500;
            color: #4a4a4a;
        }
        
        @media (max-width: 768px) {
            .chart-controls {
                flex-direction: column;
                align-items: stretch;
            }
            
            .control-group {
                justify-content: space-between;
            }
        }
    </style>
	
</head>
<body>
    <div class="container is-fluid">
        <div class="section">
            <h1 class="title">Sistema de Reportes de Facturas</h1>
            
            <!-- Formulario de Filtros -->
            <div class="box">
                <h2 class="subtitle">Filtros de Búsqueda</h2>
                <form method="POST" action="">
                    <div class="columns">
                        <div class="column">
                            <div class="field">
                                <label class="label">Fecha Inicio</label>
                                <div class="control">
                                    <input class="input" type="date" name="fecha_inicio" value="<?= $fecha_inicio ?>">
                                </div>
                            </div>
                        </div>
                        <div class="column">
                            <div class="field">
                                <label class="label">Fecha Fin</label>
                                <div class="control">
                                    <input class="input" type="date" name="fecha_fin" value="<?= $fecha_fin ?>">
                                </div>
                            </div>
                        </div>
                        <div class="column">
                            <div class="field">
                                <label class="label">Cliente (CCONSIGNA)</label>
                                <div class="control">
                                    <input class="input" type="text" name="cconsigna" value="<?= htmlspecialchars($cconsigna) ?>" placeholder="Buscar cliente...">
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="columns">
                        <div class="column">
                            <div class="field">
                                <label class="label">Subcliente (CNOMBRE)</label>
                                <div class="control">
                                    <input class="input" type="text" name="cnombre" value="<?= htmlspecialchars($cnombre) ?>" placeholder="Buscar subcliente...">
                                </div>
                            </div>
                        </div>
                        <div class="column">
                            <div class="field">
                                <label class="label">DAE (Manifiesto)</label>
                                <div class="control">
                                    <input class="input" type="text" name="fue_manifiesto" value="<?= htmlspecialchars($fue_manifiesto) ?>" placeholder="Buscar DAE...">
                                </div>
                            </div>
                        </div>
                        <div class="column">
                            <div class="field">
                                <label class="label">Invoice</label>
                                <div class="control">
                                    <input class="input" type="text" name="invoice" value="<?= htmlspecialchars($invoice) ?>" placeholder="Buscar invoice...">
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="field is-grouped">
                        <div class="control">
                            <button class="button is-primary" type="submit">
                                <span class="icon">
                                    <i class="fas fa-search"></i>
                                </span>
                                <span>Buscar</span>
                            </button>
                        </div>
                        <div class="control">
                            <button class="button is-success" type="submit" name="accion" value="exportar_excel">
                                <span class="icon">
                                    <i class="fas fa-file-excel"></i>
                                </span>
                                <span>Exportar Excel</span>
                            </button>
                        </div>
                        <div class="control">
                            <button class="button is-info" type="button" onclick="exportarPDF()" id="btn-pdf">
                                <span class="icon">
                                    <i class="fas fa-file-pdf"></i>
                                </span>
                                <span>Exportar PDF</span>
                            </button>
                        </div>
                        <div class="control">
                            <div class="select">
                                <select name="registros_por_pagina" onchange="this.form.submit()">
                                    <option value="20" <?= $registros_por_pagina == 20 ? 'selected' : '' ?>>20 por página</option>
                                    <option value="50" <?= $registros_por_pagina == 50 ? 'selected' : '' ?>>50 por página</option>
                                    <option value="100" <?= $registros_por_pagina == 100 ? 'selected' : '' ?>>100 por página</option>
                                </select>
                            </div>
                        </div>
                    </div>
                </form>
            </div>

            <!-- Gráficos -->
            <div class="columns">
                <div class="column">
                    <div class="box">
                        <h3 class="subtitle">Tallos por Mercado (CZONA)</h3>
                        <div class="chart-controls field is-grouped is-grouped-right mb-0">
                            <label class="label control-label mr-2">Tamaño:</label>
                            <div class="control">
                                <div class="buttons has-addons">
                                    <button class="button is-small chart-size-btn" data-target="graficoCzona" data-size="small">Pequeño</button>
                                    <button class="button is-small chart-size-btn is-selected" data-target="graficoCzona" data-size="medium">Mediano</button>
                                    <button class="button is-small chart-size-btn" data-target="graficoCzona" data-size="large">Grande</button>
                                </div>
                            </div>
                            <label class="label control-label mr-2 ml-3">Tipo:</label>
                            <div class="control">
                                <div class="select is-small">
                                    <select class="chart-type-select" data-target="graficoCzona">
                                        <option value="pie" selected>Circular</option>
                                        <option value="doughnut">Dona</option>
                                        <option value="bar">Barras</option>
                                        <option value="line">Lineal</option>
                                    </select>
                                </div>
                            </div>
                        </div>
                        <div class="chart-container chart-size-medium mt-3">
                            <canvas id="graficoCzona"></canvas>
                        </div>
                    </div>
                </div>
                <div class="column">
                    <div class="box">
                        <h3 class="subtitle">Tallos por Cliente</h3>
                        <div class="chart-controls field is-grouped is-grouped-right mb-0">
                            <label class="label control-label mr-2">Tamaño:</label>
                            <div class="control">
                                <div class="buttons has-addons">
                                    <button class="button is-small chart-size-btn" data-target="graficoCliente" data-size="small">Pequeño</button>
                                    <button class="button is-small chart-size-btn is-selected" data-target="graficoCliente" data-size="medium">Mediano</button>
                                    <button class="button is-small chart-size-btn" data-target="graficoCliente" data-size="large">Grande</button>
                                </div>
                            </div>
                            <label class="label control-label mr-2 ml-3">Tipo:</label>
                            <div class="control">
                                <div class="select is-small">
                                    <select class="chart-type-select" data-target="graficoCliente">
                                        <option value="pie" selected>Circular</option>
                                        <option value="doughnut">Dona</option>
                                        <option value="bar">Barras</option>
                                        <option value="line">Lineal</option>
                                    </select>
                                </div>
                            </div>
                        </div>
                        <div class="chart-container chart-size-medium mt-3">
                            <canvas id="graficoCliente"></canvas>
                        </div>
                    </div>
                </div>
            </div>

            <div class="columns">
                <div class="column">
                    <div class="box">
                        <h3 class="subtitle">Valor USD por Mercado</h3>
                        <div class="chart-controls field is-grouped is-grouped-right mb-0">
                            <label class="label control-label mr-2">Tamaño:</label>
                            <div class="control">
                                <div class="buttons has-addons">
                                    <button class="button is-small chart-size-btn" data-target="graficoValorCzona" data-size="small">Pequeño</button>
                                    <button class="button is-small chart-size-btn is-selected" data-target="graficoValorCzona" data-size="medium">Mediano</button>
                                    <button class="button is-small chart-size-btn" data-target="graficoValorCzona" data-size="large">Grande</button>
                                </div>
                            </div>
                            <label class="label control-label mr-2 ml-3">Tipo:</label>
                            <div class="control">
                                <div class="select is-small">
                                    <select class="chart-type-select" data-target="graficoValorCzona">
                                        <option value="pie" selected>Circular</option>
                                        <option value="doughnut">Dona</option>
                                        <option value="bar">Barras</option>
                                        <option value="line">Lineal</option>
                                    </select>
                                </div>
                            </div>
                        </div>
                        <div class="chart-container chart-size-medium mt-3">
                            <canvas id="graficoValorCzona"></canvas>
                        </div>
                    </div>
                </div>
                <div class="column">
                    <div class="box">
                        <h3 class="subtitle">Tallos por DAE</h3>
                        <div class="chart-controls field is-grouped is-grouped-right mb-0">
                            <label class="label control-label mr-2">Tamaño:</label>
                            <div class="control">
                                <div class="buttons has-addons">
                                    <button class="button is-small chart-size-btn" data-target="graficoManifiesto" data-size="small">Pequeño</button>
                                    <button class="button is-small chart-size-btn is-selected" data-target="graficoManifiesto" data-size="medium">Mediano</button>
                                    <button class="button is-small chart-size-btn" data-target="graficoManifiesto" data-size="large">Grande</button>
                                </div>
                            </div>
                            <label class="label control-label mr-2 ml-3">Tipo:</label>
                            <div class="control">
                                <div class="select is-small">
                                    <select class="chart-type-select" data-target="graficoManifiesto">
                                        <option value="pie" selected>Circular</option>
                                        <option value="doughnut">Dona</option>
                                        <option value="bar">Barras</option>
                                        <option value="line">Lineal</option>
                                    </select>
                                </div>
                            </div>
                        </div>
                        <div class="chart-container chart-size-medium mt-3">
                            <canvas id="graficoManifiesto"></canvas>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Tabla de Resultados -->
            <div class="box" id="tabla-resultados">
                <div class="level">
                    <div class="level-left">
                        <div class="level-item">
                            <h2 class="subtitle">Resultados (<?= number_format($total_registros) ?> registros encontrados)</h2>
                        </div>
                    </div>
                    <div class="level-right">
                        <div class="level-item">
                            <span class="tag is-info">Página <?= $pagina_actual ?> de <?= $total_paginas ?></span>
                        </div>
                    </div>
                </div>

                <!-- Paginación Superior -->
                <?php if ($total_paginas > 1): ?>
                <nav class="pagination is-centered" role="navigation">
                    <form method="POST" style="display: contents;">
                        <!-- Mantener filtros en paginación -->
                        <input type="hidden" name="fecha_inicio" value="<?= $fecha_inicio ?>">
                        <input type="hidden" name="fecha_fin" value="<?= $fecha_fin ?>">
                        <input type="hidden" name="cconsigna" value="<?= htmlspecialchars($cconsigna) ?>">
                        <input type="hidden" name="cnombre" value="<?= htmlspecialchars($cnombre) ?>">
                        <input type="hidden" name="fue_manifiesto" value="<?= htmlspecialchars($fue_manifiesto) ?>">
                        <input type="hidden" name="invoice" value="<?= htmlspecialchars($invoice) ?>">
                        <input type="hidden" name="registros_por_pagina" value="<?= $registros_por_pagina ?>">
                        
                        <?php if ($pagina_actual > 1): ?>
                            <button type="submit" name="pagina" value="<?= $pagina_actual - 1 ?>" class="pagination-previous">Anterior</button>
                        <?php else: ?>
                            <span class="pagination-previous" disabled>Anterior</span>
                        <?php endif; ?>
                        
                        <?php if ($pagina_actual < $total_paginas): ?>
                            <button type="submit" name="pagina" value="<?= $pagina_actual + 1 ?>" class="pagination-next">Siguiente</button>
                        <?php else: ?>
                            <span class="pagination-next" disabled>Siguiente</span>
                        <?php endif; ?>
                        
                        <ul class="pagination-list">
                            <?php 
                            $inicio_pag = max(1, $pagina_actual - 2);
                            $fin_pag = min($total_paginas, $pagina_actual + 2);
                            
                            if ($inicio_pag > 1): ?>
                                <li><button type="submit" name="pagina" value="1" class="pagination-link">1</button></li>
                                <?php if ($inicio_pag > 2): ?>
                                    <li><span class="pagination-ellipsis">&hellip;</span></li>
                                <?php endif; ?>
                            <?php endif; ?>
                            
                            <?php for ($i = $inicio_pag; $i <= $fin_pag; $i++): ?>
                                <li>
                                    <?php if ($i == $pagina_actual): ?>
                                        <span class="pagination-link is-current"><?= $i ?></span>
                                    <?php else: ?>
                                        <button type="submit" name="pagina" value="<?= $i ?>" class="pagination-link"><?= $i ?></button>
                                    <?php endif; ?>
                                </li>
                            <?php endfor; ?>
                            
                            <?php if ($fin_pag < $total_paginas): ?>
                                <?php if ($fin_pag < $total_paginas - 1): ?>
                                    <li><span class="pagination-ellipsis">&hellip;</span></li>
                                <?php endif; ?>
                                <li><button type="submit" name="pagina" value="<?= $total_paginas ?>" class="pagination-link"><?= $total_paginas ?></button></li>
                            <?php endif; ?>
                        </ul>
                    </form>
                </nav>
                <?php endif; ?>

                <div class="table-container">
                    <table class="table is-fullwidth is-striped is-hoverable">
                        <thead>
                            <tr>
                                <th>Fecha</th>
                                <th>Invoice</th>
                                <th>Cliente</th>
                                <th>Subcliente</th>
                                <th>Mercado</th>
                                <th>DAE</th>
                                <th>Carguera</th>
                                <th>USD</th>
                                <th>Tallos</th>
                                <th>Cajas</th>
                                <th>Peso Neto</th>
                            </tr>
                        </thead>
                        <tbody>
                            <?php
                            $total_usd = 0;
                            $total_tallos = 0;
                            $total_cajas = 0;
                            
                            if ($rs && !$rs->EOF) {
                                while (!$rs->EOF) {
                                    $total_usd += $rs->fields['usd'];
                                    $total_tallos += $rs->fields['tallos'];
                                    $total_cajas += $rs->fields['cajas'];
                                    ?>
                                    <tr>
                                        <td><?= $rs->fields['fecha'] ?></td>
                                        <td><?= $rs->fields['invoice'] ?></td>
                                        <td><?= $rs->fields['cliente'] ?></td>
                                        <td><?= $rs->fields['subcliente'] ?></td>
                                        <td><?= $rs->fields['mercado'] ?></td>
                                        <td><?= $rs->fields['dae'] ?></td>
                                        <td><?= $rs->fields['carguera'] ?></td>
                                        <td>$<?= number_format($rs->fields['usd'], 2) ?></td>
                                        <td><?= number_format($rs->fields['tallos']) ?></td>
                                        <td><?= $rs->fields['cajas'] ?></td>
                                        <td><?= $rs->fields['peso_neto'] ?></td>
                                    </tr>
                                    <?php
                                    $rs->MoveNext();
                                }
                            } else {
                                echo "<tr><td colspan='11' class='has-text-centered'>No se encontraron resultados</td></tr>";
                            }
                            ?>
                        </tbody>
                        <tfoot>
                            <tr class="has-background-primary-light">
                                <th colspan="7">TOTALES</th>
                                <th>$<?= number_format($total_usd, 2) ?></th>
                                <th><?= number_format($total_tallos) ?></th>
                                <th><?= number_format($total_cajas) ?></th>
                                <th>-</th>
                            </tr>
                        </tfoot>
                    </table>
                </div>

                <!-- Paginación Inferior -->
                <?php if ($total_paginas > 1): ?>
                <nav class="pagination is-centered" role="navigation" style="margin-top: 1rem;">
                    <form method="POST" style="display: contents;">
                        <!-- Mantener filtros en paginación -->
                        <input type="hidden" name="fecha_inicio" value="<?= $fecha_inicio ?>">
                        <input type="hidden" name="fecha_fin" value="<?= $fecha_fin ?>">
                        <input type="hidden" name="cconsigna" value="<?= htmlspecialchars($cconsigna) ?>">
                        <input type="hidden" name="cnombre" value="<?= htmlspecialchars($cnombre) ?>">
                        <input type="hidden" name="fue_manifiesto" value="<?= htmlspecialchars($fue_manifiesto) ?>">
                        <input type="hidden" name="invoice" value="<?= htmlspecialchars($invoice) ?>">
                        <input type="hidden" name="registros_por_pagina" value="<?= $registros_por_pagina ?>">
                        
                        <?php if ($pagina_actual > 1): ?>
                            <button type="submit" name="pagina" value="<?= $pagina_actual - 1 ?>" class="pagination-previous">Anterior</button>
                        <?php else: ?>
                            <span class="pagination-previous" disabled>Anterior</span>
                        <?php endif; ?>
                        
                        <?php if ($pagina_actual < $total_paginas): ?>
                            <button type="submit" name="pagina" value="<?= $pagina_actual + 1 ?>" class="pagination-next">Siguiente</button>
                        <?php else: ?>
                            <span class="pagination-next" disabled>Siguiente</span>
                        <?php endif; ?>
                        
                        <ul class="pagination-list">
                            <?php 
                            $inicio_pag = max(1, $pagina_actual - 2);
                            $fin_pag = min($total_paginas, $pagina_actual + 2);
                            
                            if ($inicio_pag > 1): ?>
                                <li><button type="submit" name="pagina" value="1" class="pagination-link">1</button></li>
                                <?php if ($inicio_pag > 2): ?>
                                    <li><span class="pagination-ellipsis">&hellip;</span></li>
                                <?php endif; ?>
                            <?php endif; ?>
                            
                            <?php for ($i = $inicio_pag; $i <= $fin_pag; $i++): ?>
                                <li>
                                    <?php if ($i == $pagina_actual): ?>
                                        <span class="pagination-link is-current"><?= $i ?></span>
                                    <?php else: ?>
                                        <button type="submit" name="pagina" value="<?= $i ?>" class="pagination-link"><?= $i ?></button>
                                    <?php endif; ?>
                                </li>
                            <?php endfor; ?>
                            
                            <?php if ($fin_pag < $total_paginas): ?>
                                <?php if ($fin_pag < $total_paginas - 1): ?>
                                    <li><span class="pagination-ellipsis">&hellip;</span></li>
                                <?php endif; ?>
                                <li><button type="submit" name="pagina" value="<?= $total_paginas ?>" class="pagination-link"><?= $total_paginas ?></button></li>
                            <?php endif; ?>
                        </ul>
                    </form>
                </nav>
                <?php endif; ?>
            </div>
        </div>
    </div>

    <script>

		// Variables globales para los gráficos
        let charts = {
            czona: null,
            cliente: null,
            valorCzona: null,
            manifiesto: null
        };
		
        // Datos para gráficos
        const dataCzona = <?= json_encode($grafico_czona) ?>;
        const dataCliente = <?= json_encode($grafico_cconsigna) ?>;
        const dataManifiesto = <?= json_encode($grafico_manifiesto) ?>;

        // Paleta de colores comunes
        const chartColors = [
            '#DC2626', '#EF4444', '#F87171', '#FB923C', '#FBBF24',
            '#FDE047', '#A3E635', '#65A30D', '#16A34A', '#15803D'
        ];

        function actualizarOGenerarGrafico(canvasId, chartType, sourceData, titulo, campoY) {
            const ctx = document.getElementById(canvasId).getContext('2d');
            const chartKey = canvasId.replace('grafico', '').toLowerCase();

            // Destruir gráfico existente si existe
            if (charts[chartKey]) {
                charts[chartKey].destroy();
            }

            const labels = Object.keys(sourceData).slice(0, 10);
            const valores = labels.map(key => sourceData[key][campoY]);

            let chartDataConfig = {
                labels: labels,
                datasets: [{
                    label: titulo, // Usado en tooltips y leyendas para bar/line
                    data: valores,
                    backgroundColor: chartColors, // Para pie, doughnut, bar
                    borderColor: chartColors[0], // Para line
                    borderWidth: (chartType === 'line' || chartType === 'bar') ? 2 : 2, // Borde de dataset para line/bar, borde de segmento para pie/doughnut
                    hoverBorderWidth: 3,
                    hoverBorderColor: '#374151',
                    fill: chartType === 'line' ? false : true, // No rellenar área bajo la línea por defecto
                    tension: chartType === 'line' ? 0.1 : 0 // Curvatura de la línea
                }]
            };
            
            if (chartType === 'bar' || chartType === 'line') {
                 // Para bar y line, es común tener un color por dataset, no por cada valor.
                 // Si se desea un color por barra, se mantiene la lista de colores en backgroundColor.
                 // Para line chart, usualmente se usa borderColor para el color de la línea.
                chartDataConfig.datasets[0].backgroundColor = chartType === 'bar' ? chartColors : chartColors[0]; // Un color para line, muchos para bar
                chartDataConfig.datasets[0].borderColor = chartType === 'line' ? chartColors[0] : chartColors; // Un color para line, muchos para bar
            }


            let options = {
                responsive: true,
                maintainAspectRatio: false, // Importante para que el chart-container controle la altura sin distorsión
                plugins: {
                    title: {
                        display: true,
                        text: titulo,
                        font: { size: 16, weight: 'bold' }
                    },
                    legend: {
                        position: (chartType === 'pie' || chartType === 'doughnut') ? 'bottom' : 'top',
                        labels: { padding: 15, usePointStyle: true, font: { size: 12 } }
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                let label = context.label || '';
                                let value = context.parsed.y !== undefined ? context.parsed.y : context.parsed;
                                
                                if (chartType === 'pie' || chartType === 'doughnut') {
                                    const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                    const percentage = total > 0 ? ((value / total) * 100).toFixed(1) : 0;
                                    return `${label}: ${value.toLocaleString()} (${percentage}%)`;
                                }
                                return `${label}: ${value.toLocaleString()}`;
                            }
                        }
                    }
                },
                animation: {
                    duration: 1000 // Mantener animación general
                }
            };

            if (chartType === 'bar' || chartType === 'line') {
                options.scales = {
                    y: {
                        beginAtZero: true,
                        title: { display: true, text: campoY.toUpperCase() }
                    },
                    x: {
                        title: { display: false, text: 'Categorías' } // Título X puede ser redundante si las etiquetas son claras
                    }
                };
            } else {
                delete options.scales; // Pie y Doughnut no usan escalas
            }
            
            charts[chartKey] = new Chart(ctx, {
                type: chartType,
                data: chartDataConfig,
                options: options
            });
        }

        // Crear gráficos iniciales
        actualizarOGenerarGrafico('graficoCzona', 'pie', dataCzona, 'Tallos por Mercado (CZONA)', 'tallos');
        actualizarOGenerarGrafico('graficoCliente', 'pie', dataCliente, 'Tallos por Cliente', 'tallos');
        actualizarOGenerarGrafico('graficoValorCzona', 'pie', dataCzona, 'Valor USD por Mercado (CZONA)', 'valor');
        actualizarOGenerarGrafico('graficoManifiesto', 'pie', dataManifiesto, 'Tallos por DAE', 'tallos');

        // Event listeners para los botones de tamaño de gráfico
        document.querySelectorAll('.chart-size-btn').forEach(button => {
            button.addEventListener('click', function() {
                const targetCanvasId = this.dataset.target;
                const newSize = this.dataset.size;
                const chartContainer = document.getElementById(targetCanvasId).parentElement;

                chartContainer.classList.remove('chart-size-small', 'chart-size-medium', 'chart-size-large');
                chartContainer.classList.add(`chart-size-${newSize}`);

                this.parentElement.querySelectorAll('.chart-size-btn').forEach(btn => btn.classList.remove('is-selected'));
                this.classList.add('is-selected');
            });
        });

        // Event listeners para los selectores de tipo de gráfico
        document.querySelectorAll('.chart-type-select').forEach(select => {
            select.addEventListener('change', function() {
                const canvasId = this.dataset.target;
                const newType = this.value;
                let sourceData, titulo, campoY;

                switch (canvasId) {
                    case 'graficoCzona':
                        sourceData = dataCzona;
                        titulo = 'Tallos por Mercado (CZONA)';
                        campoY = 'tallos';
                        break;
                    case 'graficoCliente':
                        sourceData = dataCliente;
                        titulo = 'Tallos por Cliente';
                        campoY = 'tallos';
                        break;
                    case 'graficoValorCzona':
                        sourceData = dataCzona; // Usa dataCzona pero para el campo 'valor'
                        titulo = 'Valor USD por Mercado (CZONA)';
                        campoY = 'valor';
                        break;
                    case 'graficoManifiesto':
                        sourceData = dataManifiesto;
                        titulo = 'Tallos por DAE';
                        campoY = 'tallos';
                        break;
                }
                actualizarOGenerarGrafico(canvasId, newType, sourceData, titulo, campoY);
            });
        });

        // Función para exportar PDF
        function exportarPDF() {
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF('l', 'mm', 'a4');
            
            // Título
            doc.setFontSize(16);
            doc.text('Reporte de Facturas', 14, 22);
            
            // Fecha del reporte
            doc.setFontSize(10);
            doc.text('Generado el: ' + new Date().toLocaleDateString(), 14, 30);
            
            // Capturar tabla
            html2canvas(document.getElementById('tabla-resultados')).then(canvas => {
                const imgData = canvas.toDataURL('image/png');
                const imgWidth = 270;
                const pageHeight = 190;
                const imgHeight = (canvas.height * imgWidth) / canvas.width;
                let heightLeft = imgHeight;
                let position = 40;

                doc.addImage(imgData, 'PNG', 14, position, imgWidth, imgHeight);
                heightLeft -= pageHeight;

                while (heightLeft >= 0) {
                    position = heightLeft - imgHeight;
                    doc.addPage();
                    doc.addImage(imgData, 'PNG', 14, position, imgWidth, imgHeight);
                    heightLeft -= pageHeight;
                }

                doc.save('reporte_facturas_' + new Date().toISOString().split('T')[0] + '.pdf');
            });
        }
    </script>
</body>
</html>
