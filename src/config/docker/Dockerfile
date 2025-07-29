# Указываем базовый образ
FROM node:14

# Устанавливаем рабочую директорию
WORKDIR /usr/src/app

# Копируем package.json и package-lock.json
COPY package*.json ./

# Устанавливаем зависимости
RUN npm install

# Копируем все файлы приложения
COPY . .

# Открываем порт, если приложение его использует
EXPOSE 3000

# Команда для запуска приложения
CMD ["node", "index.js"]

# Для логирования добавим следующее:
# Включаем цветной вывод логов
ENV NODE_ENV=production