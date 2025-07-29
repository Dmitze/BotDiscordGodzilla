#!/bin/bash

# 🚀 Discord AI Assistant Bot - Скрипт розгортання
# Версія: 2.3.0

set -e  # Зупинка при помилці

# Кольори для виводу
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Функції логування
log_info() {
    echo -e "${BLUE}[INFO]${NC} $1"
}

log_success() {
    echo -e "${GREEN}[SUCCESS]${NC} $1"
}

log_warning() {
    echo -e "${YELLOW}[WARNING]${NC} $1"
}

log_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

# Перевірка наявності необхідних команд
check_requirements() {
    log_info "Перевірка необхідних залежностей..."
    
    local missing_deps=()
    
    if ! command -v node &> /dev/null; then
        missing_deps+=("Node.js")
    fi
    
    if ! command -v npm &> /dev/null; then
        missing_deps+=("npm")
    fi
    
    if ! command -v git &> /dev/null; then
        missing_deps+=("git")
    fi
    
    if [ ${#missing_deps[@]} -ne 0 ]; then
        log_error "Відсутні залежності: ${missing_deps[*]}"
        log_info "Встановіть їх перед продовженням"
        exit 1
    fi
    
    log_success "Всі залежності встановлені"
}

# Перевірка версії Node.js
check_node_version() {
    log_info "Перевірка версії Node.js..."
    
    local node_version=$(node -v | cut -d'v' -f2)
    local major_version=$(echo $node_version | cut -d'.' -f1)
    
    if [ "$major_version" -lt 18 ]; then
        log_error "Потрібна Node.js версія 18 або вище. Поточна версія: $node_version"
        exit 1
    fi
    
    log_success "Node.js версія $node_version підходить"
}

# Створення .env файлу
setup_env() {
    log_info "Налаштування змінних середовища..."
    
    if [ ! -f .env ]; then
        if [ -f env.example ]; then
            cp env.example .env
            log_warning "Створено .env файл з прикладу. Відредагуйте його!"
        else
            log_error "Файл env.example не знайдено"
            exit 1
        fi
    else
        log_info "Файл .env вже існує"
    fi
}

# Встановлення залежностей
install_dependencies() {
    log_info "Встановлення npm залежностей..."
    
    if [ -f package-lock.json ]; then
        npm ci --production
    else
        npm install --production
    fi
    
    log_success "Залежності встановлені"
}

# Встановлення dev залежностей (якщо потрібно)
install_dev_dependencies() {
    if [ "$1" = "--dev" ]; then
        log_info "Встановлення dev залежностей..."
        npm install
        log_success "Dev залежності встановлені"
    fi
}

# Створення необхідних директорій
create_directories() {
    log_info "Створення необхідних директорій..."
    
    mkdir -p logs
    mkdir -p metrics
    mkdir -p tmp
    mkdir -p config
    
    log_success "Директорії створені"
}

# Налаштування логування
setup_logging() {
    log_info "Налаштування логування..."
    
    # Створення файлу логів якщо не існує
    touch logs/bot.log
    touch logs/error.log
    
    # Встановлення прав доступу
    chmod 644 logs/*.log
    
    log_success "Логування налаштовано"
}

# Реєстрація Discord команд
deploy_commands() {
    log_info "Реєстрація Discord команд..."
    
    if node deploy-commands.js; then
        log_success "Команди зареєстровані"
    else
        log_error "Помилка реєстрації команд"
        exit 1
    fi
}

# Запуск тестів
run_tests() {
    if [ "$1" = "--test" ]; then
        log_info "Запуск тестів..."
        
        if npm test; then
            log_success "Всі тести пройшли"
        else
            log_error "Тести не пройшли"
            exit 1
        fi
    fi
}

# Перевірка конфігурації
validate_config() {
    log_info "Перевірка конфігурації..."
    
    if [ ! -f .env ]; then
        log_error "Файл .env не знайдено"
        exit 1
    fi
    
    # Перевірка обов'язкових змінних
    local required_vars=("DISCORD_TOKEN" "CLIENT_ID" "GUILD_ID")
    
    for var in "${required_vars[@]}"; do
        if ! grep -q "^${var}=" .env; then
            log_warning "Змінна $var не знайдена в .env"
        fi
    done
    
    log_success "Конфігурація перевірена"
}

# Створення systemd сервісу
create_systemd_service() {
    if [ "$1" = "--systemd" ]; then
        log_info "Створення systemd сервісу..."
        
        local current_dir=$(pwd)
        local user=$(whoami)
        
        cat > discord-bot.service << EOF
[Unit]
Description=Discord AI Assistant Bot
After=network.target

[Service]
Type=simple
User=$user
WorkingDirectory=$current_dir
ExecStart=/usr/bin/node index.js
Restart=always
RestartSec=10
Environment=NODE_ENV=production

[Install]
WantedBy=multi-user.target
EOF
        
        log_success "Файл discord-bot.service створено"
        log_info "Для встановлення виконайте:"
        log_info "sudo cp discord-bot.service /etc/systemd/system/"
        log_info "sudo systemctl enable discord-bot"
        log_info "sudo systemctl start discord-bot"
    fi
}

# Створення PM2 конфігурації
create_pm2_config() {
    if [ "$1" = "--pm2" ]; then
        log_info "Створення PM2 конфігурації..."
        
        cat > ecosystem.config.js << EOF
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
EOF
        
        log_success "PM2 конфігурація створена"
        log_info "Для запуску виконайте: pm2 start ecosystem.config.js"
    fi
}

# Створення Docker конфігурації
create_docker_config() {
    if [ "$1" = "--docker" ]; then
        log_info "Створення Docker конфігурації..."
        
        if [ ! -f Dockerfile ]; then
            log_error "Dockerfile не знайдено"
            exit 1
        fi
        
        if [ ! -f docker-compose.yml ]; then
            log_error "docker-compose.yml не знайдено"
            exit 1
        fi
        
        log_success "Docker конфігурація готова"
        log_info "Для запуску виконайте: docker-compose up -d"
    fi
}

# Фінальна перевірка
final_check() {
    log_info "Фінальна перевірка..."
    
    local checks=(
        "package.json"
        "index.js"
        ".env"
        "logs/"
        "node_modules/"
    )
    
    for check in "${checks[@]}"; do
        if [ -e "$check" ]; then
            log_success "✓ $check"
        else
            log_error "✗ $check не знайдено"
        fi
    done
}

# Головна функція
main() {
    echo "🚀 Discord AI Assistant Bot - Розгортання v2.3.0"
    echo "================================================"
    
    # Парсинг аргументів
    local dev_mode=false
    local test_mode=false
    local systemd_mode=false
    local pm2_mode=false
    local docker_mode=false
    
    while [[ $# -gt 0 ]]; do
        case $1 in
            --dev)
                dev_mode=true
                shift
                ;;
            --test)
                test_mode=true
                shift
                ;;
            --systemd)
                systemd_mode=true
                shift
                ;;
            --pm2)
                pm2_mode=true
                shift
                ;;
            --docker)
                docker_mode=true
                shift
                ;;
            *)
                log_error "Невідомий аргумент: $1"
                exit 1
                ;;
        esac
    done
    
    # Виконання кроків розгортання
    check_requirements
    check_node_version
    setup_env
    create_directories
    setup_logging
    
    if [ "$dev_mode" = true ]; then
        install_dev_dependencies --dev
    else
        install_dependencies
    fi
    
    validate_config
    
    if [ "$test_mode" = true ]; then
        run_tests --test
    fi
    
    deploy_commands
    
    if [ "$systemd_mode" = true ]; then
        create_systemd_service --systemd
    fi
    
    if [ "$pm2_mode" = true ]; then
        create_pm2_config --pm2
    fi
    
    if [ "$docker_mode" = true ]; then
        create_docker_config --docker
    fi
    
    final_check
    
    echo ""
    echo "🎉 Розгортання завершено успішно!"
    echo ""
    echo "Наступні кроки:"
    echo "1. Відредагуйте .env файл з вашими налаштуваннями"
    echo "2. Запустіть бота: node index.js"
    echo "3. Перевірте логи в папці logs/"
    echo ""
    echo "Документація: README.md"
    echo "Підтримка: FAQ_SUPPORT.md"
}

# Запуск головної функції
main "$@" 