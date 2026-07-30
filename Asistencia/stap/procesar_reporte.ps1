# ==============================================================================
# Script de Automatizacion: Limpieza de Caracteres y Exportacion a Excel
# Autor: Experto en Base de Datos y Programacion (Antigravity AI)
# Descripcion: Lee un archivo CSV, corrige codificaciones incorrectas (mojibake),
#              reemplaza vocales con acentos/tildes y la letra Ñ por sus equivalentes
#              limpios en ASCII (a, e, i, o, u, n), simplifica los encabezados de las
#              columnas para una mejor presentacion, y exporta la informacion formateada
#              con un estilo profesional a un archivo Excel (.xlsx).
# ==============================================================================

Param(
    [string]$InputFile
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptDir) { $scriptDir = Get-Location }

# Extensiones de Excel soportadas
$excelExts = @(".xlsx", ".xlsb", ".xlsm", ".xls", ".xltx", ".xltm", ".xlt")

# 1. Buscar el archivo CSV o de Excel si no se especifico como argumento
if (-not $InputFile) {
    Write-Host "Buscando archivo de reporte mas reciente en: $scriptDir..." -ForegroundColor Cyan
    $allFiles = Get-ChildItem -Path $scriptDir -File | Where-Object { $_.Extension -eq ".csv" -or ($excelExts -contains $_.Extension) }
    
    $candidateFiles = $allFiles | Where-Object {
        $name = $_.Name
        $ext = $_.Extension
        
        # Excluir archivos temporales y generados
        if ($name -like "temp_*" -or $name -like "*_Corregido*" -or $name -like "*_Procesado*" -or $name -like "*_Limpio*") {
            return $false
        }
        
        # Si es un archivo de Excel, excluir si existe un .csv con el mismo nombre (el .csv es el origen)
        if ($excelExts -contains $ext) {
            $baseName = $_.BaseName
            $csvPath = Join-Path $_.DirectoryName "$baseName.csv"
            if (Test-Path $csvPath) {
                return $false
            }
        }
        
        return $true
    } | Sort-Object LastWriteTime -Descending
    
    if ($candidateFiles.Count -eq 0) {
        Write-Host "ERROR: No se encontro ningun archivo .csv o de Excel valido en el directorio." -ForegroundColor Red
        Read-Host "Presione Enter para salir"
        exit
    }
    $InputFile = $candidateFiles[0].FullName
}

if (-not (Test-Path $InputFile)) {
    Write-Host "ERROR: El archivo especificado no existe: $InputFile" -ForegroundColor Red
    Read-Host "Presione Enter para salir"
    exit
}
# Convertir a ruta absoluta para evitar problemas con rutas relativas
$InputFile = (Resolve-Path $InputFile).Path

$inputExt = [System.IO.Path]::GetExtension($InputFile)
$isExcelInput = $excelExts -contains $inputExt

if ($isExcelInput) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
    $dirName = Split-Path -Parent $InputFile
    $xlsxPath = Join-Path $dirName "$baseName`_Procesado.xlsx"
} else {
    $xlsxPath = [System.IO.Path]::ChangeExtension($InputFile, ".xlsx")
}
$fileName = Split-Path -Leaf $InputFile

Write-Host "Procesando archivo: $fileName" -ForegroundColor Green
Write-Host "Salida programada: $(Split-Path -Leaf $xlsxPath)" -ForegroundColor Green

# Mapeo de reemplazos utilizando codigos de caracteres para evitar problemas de codificacion de script
$reps = New-Object System.Collections.Generic.List[PSCustomObject]
function Add-Rep ($target, $replacement) {
    $reps.Add([PSCustomObject]@{ Target = $target; Replacement = $replacement })
}

# Reemplazos para tildes y ñ con doble codificacion (mojibake de UTF-8 leido como Windows-1252)
Add-Rep ([string]([char]0x00C3) + [char]0x00C2 + [char]0x0081) "A"
Add-Rep ([string]([char]0x00C3) + [char]0x00C2 + [char]0x008D) "I"
Add-Rep ([string]([char]0x00C3) + [char]0x00C2 + [char]0x00AD) "i"

Add-Rep ([string]([char]0x00C3) + [char]0x00A1) "a"
Add-Rep ([string]([char]0x00C3) + [char]0x00A9) "e"
Add-Rep ([string]([char]0x00C3) + [char]0x00AD) "i"
Add-Rep ([string]([char]0x00C3) + [char]0x00B3) "o"
Add-Rep ([string]([char]0x00C3) + [char]0x00BA) "u"
Add-Rep ([string]([char]0x00C3) + [char]0x00B1) "n"

Add-Rep ([string]([char]0x00C3) + [char]0x0081) "A"
Add-Rep ([string]([char]0x00C3) + [char]0x0089) "E"
Add-Rep ([string]([char]0x00C3) + [char]0x2030) "E"
Add-Rep ([string]([char]0x00C3) + [char]0x008D) "I"
Add-Rep ([string]([char]0x00C3) + [char]0x0152) "I"
Add-Rep ([string]([char]0x00C3) + [char]0x008C) "I"
Add-Rep ([string]([char]0x00C3) + [char]0x0093) "O"
Add-Rep ([string]([char]0x00C3) + [char]0x201C) "O"
Add-Rep ([string]([char]0x00C3) + [char]0x009A) "U"
Add-Rep ([string]([char]0x00C3) + [char]0x0161) "U"
Add-Rep ([string]([char]0x00C3) + [char]0x0091) "N"
Add-Rep ([string]([char]0x00C3) + [char]0x2018) "N"

# Reemplazos para tildes y ñ simples (codificacion UTF-8 correcta)
Add-Rep ([string][char]0x00E1) "a"
Add-Rep ([string][char]0x00E8) "e"
Add-Rep ([string][char]0x00E9) "e"
Add-Rep ([string][char]0x00ED) "i"
Add-Rep ([string][char]0x00F3) "o"
Add-Rep ([string][char]0x00FA) "u"
Add-Rep ([string][char]0x00F1) "n"
Add-Rep ([string][char]0x00FC) "u"
Add-Rep ([string][char]0x00BF) "" # Elimina el caracter de interrogacion invertido '¿'

Add-Rep ([string][char]0x00C1) "A"
Add-Rep ([string][char]0x00C9) "E"
Add-Rep ([string][char]0x00CD) "I"
Add-Rep ([string][char]0x00D3) "O"
Add-Rep ([string][char]0x00DA) "U"
Add-Rep ([string][char]0x00D1) "N"
Add-Rep ([string][char]0x00DC) "U"

function Clean-Text ($txt) {
    if (-not $txt) { return "" }
    $cleanedStr = $txt.ToString()
    foreach ($rep in $reps) {
        $cleanedStr = $cleanedStr.Replace($rep.Target, $rep.Replacement)
    }
    return $cleanedStr
}

# 2. Leer contenido y aplicar reemplazos
$tempInputFile = $null
if ($isExcelInput) {
    $tempInputFile = Join-Path $env:TEMP "temp_input_$pid$inputExt"
    if (Test-Path $tempInputFile) { Remove-Item $tempInputFile -Force }
    Copy-Item $InputFile -Destination $tempInputFile -Force
    
    Write-Host "1/6. Convirtiendo archivo Excel a CSV temporal..." -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    try {
        $workbook = $excel.Workbooks.Open($tempInputFile)
        $tempCsv = Join-Path $env:TEMP "temp_converted_$pid.csv"
        if (Test-Path $tempCsv) { Remove-Item $tempCsv -Force }
        $workbook.SaveAs($tempCsv, 6) # 6 = xlCSV
        $workbook.Close($false)
    }
    catch {
        Write-Host "ERROR al procesar el archivo Excel: $_" -ForegroundColor Red
        if ($excel) { $excel.Quit() }
        if ($tempInputFile -and (Test-Path $tempInputFile)) { Remove-Item $tempInputFile -Force }
        Read-Host "Presione Enter para salir"
        exit
    }
    finally {
        if ($excel) {
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
    
    if ($tempInputFile -and (Test-Path $tempInputFile)) { Remove-Item $tempInputFile -Force }
    
    $csvToProcess = $tempCsv
    $csvDelimiter = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ListSeparator
} else {
    $csvToProcess = $InputFile
    # Detectar delimitador leyendo la primera linea del archivo CSV de forma segura contra bloqueos
    $csvDelimiter = ';'
    try {
        $stream = New-Object System.IO.FileStream($csvToProcess, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        $firstLine = $reader.ReadLine()
        $reader.Close()
        $stream.Close()
        
        if ($null -ne $firstLine) {
            $commaCount = ($firstLine.ToCharArray() | Where-Object { $_ -eq ',' }).Count
            $semicolonCount = ($firstLine.ToCharArray() | Where-Object { $_ -eq ';' }).Count
            if ($commaCount -gt $semicolonCount) {
                $csvDelimiter = ','
            }
        }
    } catch {
        # En caso de error, mantener el punto y coma como delimitador predeterminado
    }
}

# Leer y limpiar caracteres especiales del CSV de forma segura contra bloqueos
Write-Host "1/6. Leyendo y limpiando caracteres especiales..." -ForegroundColor Yellow
$content = $null
try {
    $stream = New-Object System.IO.FileStream($csvToProcess, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    if ($isExcelInput) {
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::GetEncoding(1252))
    } else {
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
    }
    $content = $reader.ReadToEnd()
    $reader.Close()
    $stream.Close()
} catch {
    Write-Host "ERROR: No se pudo leer el archivo. Asegurese de que no este bloqueado por otra aplicacion de forma exclusiva." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Read-Host "Presione Enter para salir"
    exit
}

$cleaned = $content
foreach ($rep in $reps) {
    $cleaned = $cleaned.Replace($rep.Target, $rep.Replacement)
}

$tempCleanedCsv = Join-Path $env:TEMP "temp_cleaned_$pid.csv"
[System.IO.File]::WriteAllText($tempCleanedCsv, $cleaned, [System.Text.Encoding]::UTF8)

Write-Host "2/6. Estructurando informacion..." -ForegroundColor Yellow
$data = Import-Csv -Path $tempCleanedCsv -Delimiter $csvDelimiter

# Limpiar CSV temporal de conversion
if ($isExcelInput -and (Test-Path $csvToProcess)) {
    Start-Sleep -Milliseconds 200
    Remove-Item $csvToProcess -Force
}

Write-Host "3/6. Simplificando titulos de columnas..." -ForegroundColor Yellow
$tempImportCsv = Join-Path $env:TEMP "temp_import_$pid.csv"

# Funcion para escapar campos CSV de manera segura (maneja comillas, punto y coma y saltos de linea)
function Escape-CSVField ($val) {
    if ($null -eq $val) { return '""' }
    $strVal = $val.ToString()
    return '"' + $strVal.Replace('"', '""') + '"'
}

# Nuevos titulos para presentar de manera ejecutiva y profesional
$newHeaders = @("Cedula", "Nombre Completo", "Perfil", "Satisfaccion General", "Calidad Transmision", "Relevancia Rol", "Observaciones", "Fecha Registro")

# Usar directiva sep=; para asegurar que Excel separe por columnas en cualquier configuracion regional
Set-Content -Path $tempImportCsv -Value "sep=;" -Encoding utf8
$escapedHeaders = $newHeaders | ForEach-Object { Escape-CSVField $_ }
Add-Content -Path $tempImportCsv -Value ($escapedHeaders -join ";") -Encoding utf8

foreach ($row in $data) {
    $cedulaVal = $row."Cedula"
    if ($null -eq $cedulaVal -or $cedulaVal -eq "") { $cedulaVal = $row."CEDULA" }
    if ($null -eq $cedulaVal -or $cedulaVal -eq "") { $cedulaVal = $row."Numero de Cedula" }
    if ($null -eq $cedulaVal -or $cedulaVal -eq "") { $cedulaVal = $row."Numero de cedula" }

    $line = @(
        (Escape-CSVField $cedulaVal),
        (Escape-CSVField $row."Nombre completo"),
        (Escape-CSVField $row."Perfil"),
        (Escape-CSVField $row."Que tan satisfecho estas con la organizacion general de este espacio?"),
        (Escape-CSVField $row."Califica, por favor, la calidad de transmision de este espacio."),
        (Escape-CSVField $row."Consideras que lo aprendido es relevante para tu rol?"),
        (Escape-CSVField $row."Observaciones: Dejanos saber si tienes algun comentario positivo o para mejorar sobre la formacion recibida el dia de hoy."),
        (Escape-CSVField $row."Added Time")
    ) -join ";"
    Add-Content -Path $tempImportCsv -Value $line -Encoding utf8
}

# 4. Crear el archivo Excel utilizando COM
Write-Host "4/6. Creando libro de Excel..." -ForegroundColor Yellow
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Declarar variables COM para liberar despues
$srcWb = $null
$copiedSheet = $null

# Cargar el archivo CSV temporal
$workbook = $excel.Workbooks.Open($tempImportCsv)
$worksheet = $workbook.Worksheets.Item(1)

# Copiar hoja de activos y realizar buscarv (VLOOKUP) desde la carpeta 'planta'
Write-Host "Buscando informacion de planta activa para cruzar datos..." -ForegroundColor Yellow
$plantaDir = Join-Path $scriptDir "planta"
$plantaFiles = @()
if (Test-Path $plantaDir) {
    $plantaFiles = Get-ChildItem -Path $plantaDir -File | Where-Object { $excelExts -contains $_.Extension }
}

if ($plantaFiles.Count -gt 0) {
    $plantaFile = ($plantaFiles | Sort-Object LastWriteTime -Descending)[0]
    try {
        $srcWb = $excel.Workbooks.Open($plantaFile.FullName)
        
        # Buscar la hoja de activos (coincidencia parcial insensible a mayusculas)
        $activosSheet = $null
        foreach ($sh in $srcWb.Worksheets) {
            if ($sh.Name -like "*activos*" -or $sh.Name -like "*Activos*") {
                $activosSheet = $sh
                break
            }
        }
        
        if ($null -ne $activosSheet) {
            Write-Host "Copiando hoja '$($activosSheet.Name)' al libro de destino..." -ForegroundColor Yellow
            # Copiar la hoja ACTIVOS al libro procesado (se inserta antes de la primera hoja)
            $activosSheet.Copy($worksheet)
            
            # Obtener la referencia de la hoja copiada en el libro de destino
            $copiedSheet = $workbook.Worksheets.Item(1)
            $copiedSheet.Name = "ACTIVOS"
            
            # Mover la hoja copiada despues de la hoja de datos principal (que ahora esta en posicion 2)
            $copiedSheet.Move([System.Type]::Missing, $workbook.Worksheets.Item(2))
            
            # Volver a referenciar la hoja de datos principal (que vuelve a ser la numero 1)
            $worksheet = $workbook.Worksheets.Item(1)
            
            # Obtener los nombres de las cabeceras de las columnas E, F, G, H de la hoja ACTIVOS
            $headerCargo = $copiedSheet.Cells.Item(1, 5).Text
            $headerNivel1 = $copiedSheet.Cells.Item(1, 6).Text
            $headerCentroCosto = $copiedSheet.Cells.Item(1, 7).Text
            $headerNivel2 = $copiedSheet.Cells.Item(1, 8).Text
            
            # Valores por defecto en caso de estar vacios
            if (-not $headerCargo) { $headerCargo = "Descripción Cargo" }
            if (-not $headerNivel1) { $headerNivel1 = "Nombre Nivel 1" }
            if (-not $headerCentroCosto) { $headerCentroCosto = "Nombre Centro Costo" }
            if (-not $headerNivel2) { $headerNivel2 = "Nombre Nivel 2" }
            
            # Escribir cabeceras en columnas I, J, K, L (9, 10, 11, 12)
            $worksheet.Cells.Item(1, 9).Value2 = $headerCargo
            $worksheet.Cells.Item(1, 10).Value2 = $headerNivel1
            $worksheet.Cells.Item(1, 11).Value2 = $headerCentroCosto
            $worksheet.Cells.Item(1, 12).Value2 = $headerNivel2
            
            # Aplicar formulas buscarv (VLOOKUP) en un solo paso de COM para maximo rendimiento
            $lastRow = $worksheet.UsedRange.Rows.Count
            if ($lastRow -gt 1) {
                Write-Host "Aplicando formulas VLOOKUP para $lastRow registros..." -ForegroundColor Yellow
                $worksheet.Range("I2:I$lastRow").Formula = "=IFERROR(VLOOKUP(A2, ACTIVOS!`$A:`$H, 5, FALSE), """")"
                $worksheet.Range("J2:J$lastRow").Formula = "=IFERROR(VLOOKUP(A2, ACTIVOS!`$A:`$H, 6, FALSE), """")"
                $worksheet.Range("K2:K$lastRow").Formula = "=IFERROR(VLOOKUP(A2, ACTIVOS!`$A:`$H, 7, FALSE), """")"
                $worksheet.Range("L2:L$lastRow").Formula = "=IFERROR(VLOOKUP(A2, ACTIVOS!`$A:`$H, 8, FALSE), """")"
            }
        } else {
            Write-Host "ADVERTENCIA: No se encontro la hoja 'ACTIVOS' en el archivo de planta." -ForegroundColor Yellow
        }
        
        $srcWb.Close($false)
    }
    catch {
        Write-Host "ADVERTENCIA: Error al cruzar los datos de planta activa: $_" -ForegroundColor Yellow
        if ($null -ne $srcWb) {
            try { $srcWb.Close($false) } catch {}
        }
    }
} else {
    Write-Host "ADVERTENCIA: No se encontro ningun archivo Excel de planta en la carpeta '$plantaDir'." -ForegroundColor Yellow
}

# 5. Aplicar diseño y formato premium
Write-Host "5/6. Aplicando formato de diseño premium..." -ForegroundColor Yellow
$lastRow = $worksheet.UsedRange.Rows.Count
$lastCol = $worksheet.UsedRange.Columns.Count
$range = $worksheet.Range("A1", $worksheet.Cells.Item($lastRow, $lastCol))

# Fuente y tamaño general
$range.Font.Name = "Calibri"
$range.Font.Size = 11

# Estilo para la fila de encabezados (dinamico hasta el ultimo elemento de la cabecera)
$headerRange = $worksheet.Range("A1", $worksheet.Cells.Item(1, $lastCol))
$headerRange.Font.Bold = $true
$headerRange.Font.Color = 16777215 # Color de fuente Blanco (Hex: #FFFFFF)
$headerRange.Interior.Color = 7884319 # Fondo Azul Marino Oscuro (#1F4E78 en formato BGR)
$headerRange.RowHeight = 28
$headerRange.VerticalAlignment = -4108 # Centrado vertical

# Alineacion de datos por columna
# Centrado para Cedula, Perfil, Satisfaccion, Calidad, Relevancia, Fecha
# Izquierda para Nombre Completo, Observaciones, y los nuevos campos cruzados de Planta
$worksheet.Range("A2", "A$lastRow").HorizontalAlignment = -4108
$worksheet.Range("B2", "B$lastRow").HorizontalAlignment = -4131
$worksheet.Range("C2", "C$lastRow").HorizontalAlignment = -4108
$worksheet.Range("D2", "D$lastRow").HorizontalAlignment = -4108
$worksheet.Range("E2", "E$lastRow").HorizontalAlignment = -4108
$worksheet.Range("F2", "F$lastRow").HorizontalAlignment = -4108
$worksheet.Range("G2", "G$lastRow").HorizontalAlignment = -4131
$worksheet.Range("H2", "H$lastRow").HorizontalAlignment = -4108

# Alinear a la izquierda las columnas de planta si existen
if ($lastCol -ge 12) {
    $worksheet.Range("I2", "I$lastRow").HorizontalAlignment = -4131
    $worksheet.Range("J2", "J$lastRow").HorizontalAlignment = -4131
    $worksheet.Range("K2", "K$lastRow").HorizontalAlignment = -4131
    $worksheet.Range("L2", "L$lastRow").HorizontalAlignment = -4131
}

# Formato de celda especial
# Evitar notacion cientifica en Cedula
$worksheet.Range("A2", "A$lastRow").NumberFormat = "0"
# Formato de Fecha limpio
$worksheet.Range("H2", "H$lastRow").NumberFormat = "yyyy-mm-dd hh:mm"

# Bordes finos gris claro para todas las celdas de la cuadricula
# xlEdgeLeft=7, xlEdgeTop=8, xlEdgeBottom=9, xlEdgeRight=10, xlInsideVertical=11, xlInsideHorizontal=12
$borderTypes = @(7, 8, 9, 10, 11, 12)
foreach ($bt in $borderTypes) {
    $border = $range.Borders.Item($bt)
    $border.LineStyle = 1 # xlContinuous
    $border.Weight = 2 # xlThin
    $border.Color = 14277081 # Gris Claro (#D9D9D9 en BGR)
}

# Autoajustar anchos de columnas
$worksheet.UsedRange.Columns.AutoFit()

# 6. Guardar archivo final
Write-Host "6/6. Guardando como libro de Excel (.xlsx)..." -ForegroundColor Yellow
$tempXlsxPath = Join-Path $env:TEMP "temp_final_output_$pid.xlsx"
if (Test-Path $tempXlsxPath) {
    Remove-Item $tempXlsxPath -Force
}
$workbook.SaveAs($tempXlsxPath, 51) # 51 representa el formato de archivo XLSX (.xlsx)
$workbook.Close($false)
$excel.Quit()

# Liberar los objetos COM antes de copiar el archivo para liberar cualquier bloqueo
if ($null -ne $copiedSheet) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($copiedSheet) | Out-Null
}
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
if ($null -ne $srcWb) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($srcWb) | Out-Null
}
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

# Copiar el archivo generado a su destino final (bypasseando bloqueos de OneDrive)
try {
    if (Test-Path $xlsxPath) {
        Remove-Item $xlsxPath -Force
    }
    Copy-Item $tempXlsxPath -Destination $xlsxPath -Force
} catch {
    Write-Host "ADVERTENCIA: No se pudo sobrescribir el archivo de destino porque esta abierto o bloqueado." -ForegroundColor Yellow
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($xlsxPath)
    $dirName = Split-Path -Parent $xlsxPath
    $alternativePath = Join-Path $dirName "$baseName`_$timestamp.xlsx"
    
    Write-Host "Guardando una copia alternativa con fecha y hora en: $(Split-Path -Leaf $alternativePath)" -ForegroundColor Yellow
    try {
        Copy-Item $tempXlsxPath -Destination $alternativePath -Force
        $xlsxPath = $alternativePath
    } catch {
        Write-Host "ERROR: No se pudo guardar el archivo. Por favor, asegurese de tener permisos y cierre Excel si esta abierto." -ForegroundColor Red
        Read-Host "Presione Enter para salir"
        exit
    }
}

# Limpiar archivos temporales
if (Test-Path $tempXlsxPath) {
    Remove-Item $tempXlsxPath -Force
}
if ($tempCleanedCsv -and (Test-Path $tempCleanedCsv)) {
    Remove-Item $tempCleanedCsv -Force
}
if ($tempImportCsv -and (Test-Path $tempImportCsv)) {
    Remove-Item $tempImportCsv -Force
}

Write-Host "PROCESO COMPLETADO EXITOSAMENTE!" -ForegroundColor Green
Write-Host "Archivo generado en: $xlsxPath" -ForegroundColor Cyan
Read-Host "Presione Enter para finalizar"
