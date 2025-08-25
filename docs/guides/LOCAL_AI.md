# 🦙 Локальний AI з Ollama

Цей посібник описує, як налаштувати та використовувати локальні AI-моделі за допомогою Ollama у вашому боті.

## 📋 Передумови

- Встановлений [Docker](https://www.docker.com/)
- Принаймні 8GB вільної оперативної пам'яті (рекомендовано 16GB+)
- 10GB вільного місця на диску

## 🚀 Швидкий старт

### 1. Встановлення Ollama

```bash
# Запуск Ollama через Docker
docker run -d -v ollama:/root/.ollama -p 11434:11434 --name ollama ollama/ollama

# Перевірка роботи
curl http://localhost:11434/api/tags
```

### 2. Завантаження моделі

```bash
# Завантаження моделі (наприклад, llama3)
docker exec -it ollama ollama pull llama3

# Доступні моделі: llama3, mistral, codellama, phi3, тощо
```

### 3. Налаштування бота

Додайте до вашого `.env` файлу:

```env
# Ollama налаштування
OLLAMA_BASE_URL=http://localhost:11434
OLLAMA_MODEL=llama3  # або інша завантажена модель
OLLAMA_TIMEOUT=30000  # мс

# Вимкнення хмарних AI-провайдерів
AI_PROVIDER=local
OPENAI_API_KEY=  # залиште порожнім
```

## 🛠️ Використання

### Перевірка підключення

```bash
# Перевірка роботи Ollama
curl http://localhost:11434/api/version

# Перевірка доступних моделей
curl http://localhost:11434/api/tags
```

### Приклад запиту до API

```bash
curl http://localhost:11434/api/generate -d '{
  "model": "llama3",
  "prompt": "Привіт, як справи?",
  "stream": false
}'
```

## 🔧 Розширені налаштування

### Використання GPU (якщо доступно)

```bash
docker run -d --gpus=all -v ollama:/root/.ollama -p 11434:11434 --name ollama ollama/ollama
```

### Налаштування параметрів моделі

Створіть файл `Modelfile` для кастомних налаштувань:

```dockerfile
FROM llama3

# Параметри генерації
PARAMETER num_ctx 4096
PARAMETER temperature 0.7
PARAMETER top_k 50
PARAMETER top_p 0.9
```

Потім створіть кастомну модель:

```bash
docker exec -it ollama ollama create my-model -f /path/to/Modelfile
```

## 🔄 Інтеграція з RAG

Для використання з RAG додайте до конфігурації:

```typescript
// У конфігурації RAG-сервісу
const ragConfig = {
  retriever: {
    type: 'hybrid',
    options: {
      k: 5,
      alpha: 0.5
    }
  },
  generator: {
    provider: 'ollama',
    model: 'llama3',
    temperature: 0.7,
    maxTokens: 2000
  }
};
```

## ⚠️ Обмеження

- Швидкість відповіді залежить від апаратних можливостей
- Великі моделі можуть вимагати багато оперативної пам'яті
- Деякі моделі можуть мати обмеження контексту

## 🔍 Діагностика проблем

### Перевірка використання пам'яті

```bash
docker stats ollama
```

### Перезапуск сервісу

```bash
docker restart ollama
```

### Перегляд логів

```bash
docker logs ollama
```

## 📚 Додаткові ресурси

- [Офіційна документація Ollama](https://ollama.ai/)
- [Доступні моделі](https://ollama.ai/library)
- [Приклади використання API](https://github.com/ollama/ollama/blob/main/docs/api.md)

## 🤖 Приклад використання в коді

```typescript
import { Ollama } from 'ollama';

const ollama = new Ollama({
  host: process.env.OLLAMA_BASE_URL || 'http://localhost:11434',
});

async function generateResponse(prompt: string) {
  try {
    const response = await ollama.generate({
      model: process.env.OLLAMA_MODEL || 'llama3',
      prompt: prompt,
      stream: false,
      options: {
        temperature: 0.7,
        top_p: 0.9,
      },
    });
    
    return response.response;
  } catch (error) {
    console.error('Помилка генерації відповіді:', error);
    throw error;
  }
}
```

## 🔄 Оновлення моделей

Для оновлення моделі до останньої версії:

```bash
docker exec -it ollama ollama pull llama3:latest
```

## 🚨 Важливі зауваження

1. Завжди перевіряйте ліцензію моделі перед використанням у продакшені
2. Деякі моделі можуть вимагати додаткових дозволів або мати обмеження на комерційне використання
3. Рекомендується використовувати VPN або захищене з'єднання при роботі з конфіденційними даними
