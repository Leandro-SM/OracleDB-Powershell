## executar "notepad $Profile" e adicionar as duas linhas abaixo no profile do Powershell
# $env:TNS_ADMIN = "C:\app\product\21c\network\admin"
# $env:ORACLE_HOME = "C:\app\product\21c"


[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$Username,
    
    [Parameter(Mandatory = $false)]
    [Security.SecureString]$Password,
    
    [Parameter(Mandatory = $true)]
    [string]$DataSource,
    
    [Parameter(Mandatory = $false)]
    [string]$Query,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputFile,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("WindowsAuth", "OracleAuth", "Wallet")]
    [string]$AuthMethod = "OracleAuth",
    
    [Parameter(Mandatory = $false)]
    [string]$WalletLocation,
    
    [Parameter(Mandatory = $false)]
    [int]$ConnectionTimeout = 30,
    
    [Parameter(Mandatory = $false)]
    [int]$CommandTimeout = 120
)
$ScriptVersion = "2.0"
$ExecutionStart = Get-Date

$LogFile = "$PSScriptRoot\OracleDB_Connection_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp [$Level] $Message"
    Add-Content -Path $LogFile -Value $logMessage
    switch ($Level) {
        "ERROR" { Write-Error $Message }
        "WARN" { Write-Warning $Message }
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
        default { Write-Host $Message }
    }
}

Write-Log "Iniciando script de conexão OracleDB - Versão $ScriptVersion"
Write-Log "Método de autenticação selecionado: $AuthMethod"

try {
)
    $possiblePaths = @(
        "C:\Oracle\product\*\client_*\odp.net\managed\common\Oracle.ManagedDataAccess.dll",
        "C:\app\*\product\*\client_*\odp.net\managed\common\Oracle.ManagedDataAccess.dll",
        "C:\Oracle\ODP.NET\managed\common\Oracle.ManagedDataAccess.dll",
        "$env:USERPROFILE\Oracle.ManagedDataAccess.dll",
        "$PSScriptRoot\Oracle.ManagedDataAccess.dll"
    )
    
    $dllLoaded = $false
    foreach ($pathPattern in $possiblePaths) {
        $resolvedPath = Get-Item -Path $pathPattern -ErrorAction SilentlyContinue | 
                        Select-Object -First 1 -ExpandProperty FullName
        if ($resolvedPath -and (Test-Path $resolvedPath)) {
            try {
                Add-Type -Path $resolvedPath -ErrorAction Stop
                Write-Log "DLL Oracle carregada com sucesso: $resolvedPath" "SUCCESS"
                $dllLoaded = $true
                break
            } catch {
                Write-Log "Falha ao carregar DLL de $resolvedPath: $($_.Exception.Message)" "WARN"
            }
        }
    }
    
    if (-not $dllLoaded) {
        throw "Oracle.ManagedDataAccess.dll não encontrada. Instale o Oracle Data Provider for .NET"
    }
} catch {
    Write-Log "Erro crítico ao carregar DLL Oracle: $($_.Exception.Message)" "ERROR"
    exit 1
}


function New-OracleConnectionString {
    param(
        [string]$Username,
        [string]$PlainPassword,
        [string]$DataSource,
        [string]$AuthMethod,
        [string]$WalletLocation,
        [int]$ConnectionTimeout
    )
    

    $connStringBuilder = New-Object Oracle.ManagedDataAccess.Client.OracleConnectionStringBuilder
    
    $connStringBuilder.DataSource = $DataSource
    $connStringBuilder.ConnectionTimeout = $ConnectionTimeout
    $connStringBuilder.Pooling = $true
    $connStringBuilder.MinPoolSize = 1
    $connStringBuilder.MaxPoolSize = 10
    
    switch ($AuthMethod) {
        "WindowsAuth" {
            $connStringBuilder.UserID = "/"
            Write-Log "Usando autenticação Windows integrada"
        }
        "OracleAuth" {
            if ($Username -and $PlainPassword) {
                $connStringBuilder.UserID = $Username
                $connStringBuilder.Password = $PlainPassword
                Write-Log "Usando autenticação Oracle tradicional"
            } else {
                throw "Usuário e senha são obrigatórios para autenticação Oracle"
            }
        }
        "Wallet" {
            if (-not $WalletLocation) {
                throw "Localização do wallet é obrigatória para autenticação via wallet"
            }
            $connStringBuilder.UserID = "/"
            $connStringBuilder.WalletLocation = $WalletLocation
            Write-Log "Usando autenticação via Oracle Wallet: $WalletLocation"
        }
    }
    
    return $connStringBuilder.ConnectionString
}

function Invoke-OracleQuery {
    param(
        [Oracle.ManagedDataAccess.Client.OracleConnection]$Connection,
        [string]$Query,
        [int]$CommandTimeout
    )
    
    try {
        $command = $Connection.CreateCommand()
        $command.CommandText = $Query
        $command.CommandTimeout = $CommandTimeout
        
        Write-Log "Executando query: $(if($Query.Length -gt 100){"$($Query.Substring(0,100))..."} else {$Query})"
        
        if ($Query.TrimStart().StartsWith("SELECT", [System.StringComparison]::OrdinalIgnoreCase)) {
            
            $dataAdapter = New-Object Oracle.ManagedDataAccess.Client.OracleDataAdapter($command)
            $dataTable = New-Object System.Data.DataTable
            $recordCount = $dataAdapter.Fill($dataTable)
            Write-Log "Query executada com sucesso. $recordCount registros retornados." "SUCCESS"
            return $dataTable
        } else {
            
            $affectedRows = $command.ExecuteNonQuery()
            Write-Log "Comando executado. $affectedRows linhas afetadas." "SUCCESS"
            return $affectedRows
        }
    } catch {
        Write-Log "Erro ao executar query: $($_.Exception.Message)" "ERROR"
        throw
    }
}

try {
    
    if ($AuthMethod -eq "OracleAuth" -and -not $Password) {
        $Password = Read-Host "Digite a senha para o usuário $Username" -AsSecureString
    }
    
    
    $plainPassword = ""
    if ($Password -and $AuthMethod -eq "OracleAuth") {
        $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
        $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    }
    
    
    $connectionString = New-OracleConnectionString `
        -Username $Username `
        -PlainPassword $plainPassword `
        -DataSource $DataSource `
        -AuthMethod $AuthMethod `
        -WalletLocation $WalletLocation `
        -ConnectionTimeout $ConnectionTimeout
    
    Write-Log "String de conexão criada (senha ocultada)"
    
    
    $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection($connectionString)
    
    Write-Log "Abrindo conexão com OracleDB..."
    $connection.Open()
    Write-Log "Conexão estabelecida com sucesso!" "SUCCESS"
    Write-Log "Server Version: $($connection.ServerVersion)"
    Write-Log "Database Name: $($connection.Database)"
    
    
    if ($Query) {
        $results = Invoke-OracleQuery `
            -Connection $connection `
            -Query $Query `
            -CommandTimeout $CommandTimeout
        
      
        if ($results -is [System.Data.DataTable]) {
            Write-Log "Exibindo primeiras 5 linhas dos resultados:"
            $results | Select-Object -First 5 | Format-Table -AutoSize
            
      
            if ($OutputFile) {
                $results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
                Write-Log "Resultados exportados para: $OutputFile" "SUCCESS"
            }
            
            
            Write-Log "Total de registros: $($results.Rows.Count)"
            Write-Log "Colunas: $($results.Columns.ColumnName -join ', ')"
        }
    } else {
        Write-Log "Nenhuma query fornecida. Conexão testada com sucesso." "SUCCESS"
    }
    
    
} catch {
    Write-Log "ERRO na conexão/execução: $($_.Exception.Message)" "ERROR"
    Write-Log "Detalhes: $($_.Exception.InnerException.Message)" "ERROR"
    
    
    if ($DataSource -notmatch ":") {
        Write-Log "AVISO: DataSource não contém ':' - verifique formato Easy Connect (host:port/service)" "WARN"
        Write-Log "Formato Easy Connect: host:port/service_name"
        Write-Log "Formato TNS: alias_configurado_no_tnsnames.ora"
    }
    
    exit 1
} finally {
    
    if ($plainPassword) {
        
        $plainPassword = $null
        [GC]::Collect()
    }
    
    if ($connection -and $connection.State -eq 'Open') {
        $connection.Close()
        $connection.Dispose()
        Write-Log "Conexão fechada e recursos liberados."
    }
    
    $executionTime = (Get-Date) - $ExecutionStart
    Write-Log "Tempo total de execução: $($executionTime.TotalSeconds.ToString('0.00')) segundos"
    Write-Log "Script finalizado. Log disponível em: $LogFile"
    
}
