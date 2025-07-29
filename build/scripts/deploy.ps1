# 🚀 Discord AI Assistant Bot - PowerShell скрипт розгортання
# Версія: 2.3.0

param(
    [switch]$Dev,
    [switch]$Test,
    [switch]$Systemd,
    [switch]$PM2,
    [switch]$Docker
)

# Функції логування
function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message" -ForegroundColor Blue
}

function Write-Success {
    param([string]$Message)
    Write-Host "[SUCCESS] $Message" -ForegroundColor Green
}

function Write-Warning {
    param([string]$Message)
    Write-Host "[WARNING] $Message" -ForegroundColor Yellow
}

function Write-Error {
    param([string]$Message)
    Write-Host "[ERROR] $Message" -ForegroundColor Red
}

# Перевірка наявності необхідних команд
function Test-Requirements {
    Write-Info "Перевірка необхідних залежностей..."
    
    $missingDeps = @()
    
    # Перевірка Node.js
    try {
        $nodeVersion = node --version
        if ($LASTEXITCODE -ne 0) {
            $missingDeps += "Node.js"
        }
    } catch {
        $missingDeps += "Node.js"
    }
    
    # Перевірка npm
    try {
        $npmVersion = npm --version
        if ($LASTEXITCODE -ne 0) {
            $missingDeps += "npm"
        }
    } catch {
        $missingDeps += "npm"
    }
    
    # Перевірка Git
    try {
        $gitVersion = git --version
        if ($LASTEXITCODE -ne 0) {
            $missingDeps += "git"
        }
    } catch {
        $missingDeps += "git"
    }
    
    if ($missingDeps.Count -gt 0) {
        Write-Error "Відсутні залежності: $($missingDeps -join ', ')"
        Write-Info "Встановіть їх перед продовженням"
        exit 1
    }
    
    Write-Success "Всі залежності встановлені"
}

# Перевірка версії Node.js
function Test-NodeVersion {
    Write-Info "Перевірка версії Node.js..."
    
    $nodeVersion = node --version
    $version = $nodeVersion.TrimStart('v')
    $majorVersion = [int]($version.Split('.')[0])
    
    if ($majorVersion -lt 18) {
        Write-Error "Потрібна Node.js версія 18 або вище. Поточна версія: $version"
        exit 1
    }
    
    Write-Success "Node.js версія $version підходить"
}

# Створення .env файлу
function Set-Environment {
    Write-Info "Налаштування змінних середовища..."
    
    if (-not (Test-Path ".env")) {
        if (Test-Path "env.example") {
            Copy-Item "env.example" ".env"
            Write-Warning "Створено .env файл з прикладу. Відредагуйте його!"
        } else {
            Write-Error "Файл env.example не знайдено"
            exit 1
        }
    } else {
        Write-Info "Файл .env вже існує"
    }
}

# Встановлення залежностей
function Install-Dependencies {
    Write-Info "Встановлення npm залежностей..."
    
    if (Test-Path "package-lock.json") {
        npm ci --production
    } else {
        npm install --production
    }
    
    if ($LASTEXITCODE -eq 0) {
        Write-Success "Залежності встановлені"
    } else {
        Write-Error "Помилка встановлення залежностей"
        exit 1
    }
}

# Встановлення dev залежностей
function Install-DevDependencies {
    if ($Dev) {
        Write-Info "Встановлення dev залежностей..."
        npm install
        if ($LASTEXITCODE -eq 0) {
            Write-Success "Dev залежності встановлені"
        } else {
            Write-Error "Помилка встановлення dev залежностей"
            exit 1
        }
    }
}

# Створення необхідних директорій
function New-Directories {
    Write-Info "Створення необхідних директорій..."
    
    $directories = @("logs", "metrics", "tmp", "config")
    
    foreach ($dir in $directories) {
        if (-not (Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
    }
    
    Write-Success "Директорії створені"
}

# Налаштування логування
function Set-Logging {
    Write-Info "Налаштування логування..."
    
    # Створення файлів логів
    if (-not (Test-Path "logs/bot.log")) {
        New-Item -ItemType File -Path "logs/bot.log" -Force | Out-Null
    }
    
    if (-not (Test-Path "logs/error.log")) {
        New-Item -ItemType File -Path "logs/error.log" -Force | Out-Null
    }
    
    Write-Success "Логування налаштовано"
}

# Реєстрація Discord команд
function Deploy-Commands {
    Write-Info "Реєстрація Discord команд..."
    
    node deploy-commands.js
    if ($LASTEXITCODE -eq 0) {
        Write-Success "Команди зареєстровані"
    } else {
        Write-Error "Помилка реєстрації команд"
        exit 1
    }
}

# Запуск тестів
function Invoke-Tests {
    if ($Test) {
        Write-Info "Запуск тестів..."
        
        npm test
        if ($LASTEXITCODE -eq 0) {
            Write-Success "Всі тести пройшли"
        } else {
            Write-Error "Тести не пройшли"
            exit 1
        }
    }
}

# Перевірка конфігурації
function Test-Configuration {
    Write-Info "Перевірка конфігурації..."
    
    if (-not (Test-Path ".env")) {
        Write-Error "Файл .env не знайдено"
        exit 1
    }
    
    # Перевірка обов'язкових змінних
    $requiredVars = @("DISCORD_TOKEN", "CLIENT_ID", "GUILD_ID")
    $envContent = Get-Content ".env"
    
    foreach ($var in $requiredVars) {
        if (-not ($envContent -match "^${var}=")) {
            Write-Warning "Змінна $var не знайдена в .env"
        }
    }
    
    Write-Success "Конфігурація перевірена"
}

# Створення Windows сервісу
function New-WindowsService {
    if ($Systemd) {
        Write-Info "Створення Windows сервісу..."
        
        $currentDir = Get-Location
        $serviceScript = @"
# Windows Service для Discord Bot
# Запуск: node index.js

Set-Location "$currentDir"
node index.js
"@
        
        $serviceScript | Out-File -FilePath "start-bot.ps1" -Encoding UTF8
        
        Write-Success "Скрипт запуску створено: start-bot.ps1"
        Write-Info "Для створення Windows сервісу використовуйте:"
        Write-Info "sc create DiscordBot binPath= `"powershell.exe -File $currentDir\start-bot.ps1`""
        Write-Info "sc start DiscordBot"
    }
}

# Створення PM2 конфігурації
function New-PM2Config {
    if ($PM2) {
        Write-Info "Створення PM2 конфігурації..."
        
        $pm2Config = @"
module.exports = {
  apps: [{
    name: 'discord-bot',
    script: 'index.js',
    instances: 'max',
    exec_mode: 'cluster',
    env: {
      NODE_ENV: 'production'
    },
    error_file: './logs/err.log',
    out_file: './logs/out.log',
    log_file: './logs/combined.log',
    time: true,
    max_memory_restart: '500M',
    restart_delay: 4000,
    max_restarts: 10
  }]
};
"@
        
        $pm2Config | Out-File -FilePath "ecosystem.config.js" -Encoding UTF8
        
        Write-Success "PM2 конфігурація створена"
        Write-Info "Для запуску виконайте: pm2 start ecosystem.config.js"
    }
}

# Створення Docker конфігурації
function Test-DockerConfig {
    if ($Docker) {
        Write-Info "Перевірка Docker конфігурації..."
        
        if (-not (Test-Path "Dockerfile")) {
            Write-Error "Dockerfile не знайдено"
            exit 1
        }
        
        if (-not (Test-Path "docker-compose.yml")) {
            Write-Error "docker-compose.yml не знайдено"
            exit 1
        }
        
        Write-Success "Docker конфігурація готова"
        Write-Info "Для запуску виконайте: docker-compose up -d"
    }
}

# Фінальна перевірка
function Test-FinalCheck {
    Write-Info "Фінальна перевірка..."
    
    $checks = @("package.json", "index.js", ".env", "logs", "node_modules")
    
    foreach ($check in $checks) {
        if (Test-Path $check) {
            Write-Success "✓ $check"
        } else {
            Write-Error "✗ $check не знайдено"
        }
    }
}

# Головна функція
function Main {
    Write-Host "🚀 Discord AI Assistant Bot - Розгортання v2.3.0" -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan
    
    # Виконання кроків розгортання
    Test-Requirements
    Test-NodeVersion
    Set-Environment
    New-Directories
    Set-Logging
    
    if ($Dev) {
        Install-DevDependencies
    } else {
        Install-Dependencies
    }
    
    Test-Configuration
    
    if ($Test) {
        Invoke-Tests
    }
    
    Deploy-Commands
    
    if ($Systemd) {
        New-WindowsService
    }
    
    if ($PM2) {
        New-PM2Config
    }
    
    if ($Docker) {
        Test-DockerConfig
    }
    
    Test-FinalCheck
    
    Write-Host ""
    Write-Host "🎉 Розгортання завершено успішно!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Наступні кроки:" -ForegroundColor Yellow
    Write-Host "1. Відредагуйте .env файл з вашими налаштуваннями"
    Write-Host "2. Запустіть бота: node index.js"
    Write-Host "3. Перевірте логи в папці logs/"
    Write-Host ""
    Write-Host "Документація: README.md" -ForegroundColor Cyan
    Write-Host "Підтримка: FAQ_SUPPORT.md" -ForegroundColor Cyan
}

# Запуск головної функції
Main 