const { Client, GatewayIntentBits, Routes, REST } = require('discord.js');
const fetch = require('node-fetch');

// Створюємо клієнта бота
const client = new Client({
  intents: [
    GatewayIntentBits.Guilds,
    GatewayIntentBits.GuildMessages,
    GatewayIntentBits.MessageContent
  ]
});

// URL до твого Google Apps Script
const API_URL = ''; 

// Команди
const commands = [
  {
    name: 'дані',
    description: 'Отримати дані з таблиці',
  },
];

// Коли бот запущений
client.once('ready', async () => {
  console.log(`Бот ${client.user.tag} онлайн!`);

  // Реєструємо команди
  const rest = new REST({ version: '10' }).setToken('ВАШ_ТОКЕН');
  await rest.put(Routes.applicationCommands(client.user.id), { body: commands });

  console.log('Команди зареєстровані!');
});

// Обробка slash-команд
client.on('interactionCreate', async (interaction) => {
  if (!interaction.isChatInputCommand()) return;

  if (interaction.commandName === 'дані') {
    try {
      const res = await fetch(API_URL);
      const data = await res.json();

      const headers = data[0]; // перший рядок - заголовки
      const rows = data.slice(1); // решта - записи

      if (rows.length === 0) {
        await interaction.reply('Немає доступних даних.');
        return;
      }

      let output = '**Дані з таблиці:**\n';
      for (let i = 0; i < Math.min(5, rows.length); i++) {
        const row = rows[i];
        output += `\n• ${headers[6]}: ${row[6]} | Ціна: ${row[4]} | Кількість: ${row[3]}`;
      }

      await interaction.reply(output);
    } catch (err) {
      console.error(err);
      await interaction.reply('Помилка при отриманні даних.');
    }
  }
});

// Логін бота
client.login('');
