<<<<<<< HEAD
const { Client, GatewayIntentBits, Routes } = require('discord.js');
const fetch = (...args) => import('node-fetch').then(({ default: f }) => f(...args));
const XLSX = require('xlsx');
require('dotenv').config();

// Зчитуємо змінні з .env
const SHEET_ID = process.env.SHEET_ID;
const SHEET_NAME = 'Аркуш1';
const GOOGLE_API_KEY = process.env.GOOGLE_API_KEY;
const APP_SCRIPT_URL = process.env.APP_SCRIPT_URL;
const API_URL = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${SHEET_NAME}?key=${GOOGLE_API_KEY}`;
const CELLS_URL = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/C964:E964?key=${GOOGLE_API_KEY}`;
const client = new Client({
  intents: [
    GatewayIntentBits.Guilds,
    GatewayIntentBits.GuildMessages,
    GatewayIntentBits.MessageContent,
    GatewayIntentBits.GuildMessageReactions
  ]
});

// Реєстрація слеш-команд
const { REST } = require('@discordjs/rest');
const { version } = require('discord.js').Constants;
const commands = [
  {
    name: 'залишки',
    description: 'Показує загальну кількість товарів'
  },
  {
    name: 'оновити',
    description: 'Принудово оновлює дані'
  },
  {
    name: 'порожні',
    description: 'Показує товари з мінімальною кількістю'
  },
  {
    name: 'пошук',
    description: 'Пошук за полями таблиці',
    options: [
      {
        name: 'поле',
        description: 'За яким полем шукати',
        type: 3,
        required: true,
        choices: [
          { name: 'Найменування', value: 'назва' },
          { name: 'Серійний номер', value: 'серія' },
          { name: 'Контрагент', value: 'контрагент' },
          { name: 'Кількість', value: 'кількість' },
          { name: 'Ціна', value: 'ціна' }
        ]
      },
      {
        name: 'запит',
        description: 'Що шукати (наприклад: "стол", "5")',
        type: 3,
        required: true
      }
    ]
  },
  {
    name: 'розумний-пошук',
    description: 'Пошук за кількома полями',
    options: [
      { name: 'номенклатура', type: 3, description: 'Шукати за назвою товару' },
      { name: 'контрагент', type: 3, description: 'Шукати за контрагентом' },
      { name: 'серія', type: 3, description: 'Шукати за серійним номером' },
      { name: 'ціна_вище', type: 10, description: 'Показувати товари дорожче цього значення' },
      { name: 'кількість_вище', type: 10, description: 'Показувати товари з кількістю більше' }
    ]
  },
  {
    name: 'пошук-експортовано',
    description: 'Експортує результати пошуку в Excel (.xlsx)',
    options: [
      {
        name: 'поле',
        description: 'За яким полем шукати',
        type: 3,
        required: true,
        choices: [
          { name: 'Найменування', value: 'назва' },
          { name: 'Серійний номер', value: 'серія' },
          { name: 'Контрагент', value: 'контрагент' },
          { name: 'Кількість', value: 'кількість' },
          { name: 'Ціна', value: 'ціна' }
        ]
      },
      {
        name: 'запит',
        description: 'Що шукати (наприклад: "стол", "5")',
        type: 3,
        required: true
      }
    ]
  },
  {
    name: 'експорт',
    description: 'Експортує всю таблицю в Excel (.xlsx)'
  },
  {
    name: 'help',
    description: 'Показує список усіх доступних команд'
  }
];
const rest = new REST({ version: '10' }).setToken(process.env.BOT_TOKEN);

// ────────────────
// 🧠 Функція: автоматичне сповіщення при зміні в таблиці
// ────────────────
let previousData = null;
async function checkForChanges(botClient) {
  try {
    const res = await fetch(API_URL);
    const data = await res.json();
    const currentRows = data.values;
    if (!previousData) {
      previousData = currentRows;
      return;
    }
    const changedCells = [];
    for (let i = 0; i < Math.min(currentRows.length, previousData.length); i++) {
      const oldRow = previousData[i];
      const newRow = currentRows[i];
      if (!oldRow || !newRow) continue;
      for (let j = 0; j < Math.min(oldRow.length, newRow.length); j++) {
        if (oldRow[j] !== newRow[j]) {
          changedCells.push({
            row: i + 1,
            column: j + 1,
            from: oldRow[j],
            to: newRow[j]
          });
        }
      }
    }
    if (changedCells.length > 0) {
      const channel = botClient.channels.cache.find(ch => ch.name === 'склад' && ch.type === 0);
      if (!channel) return;
      let message = `🔔 Виявлено зміни в таблиці:
`;
      changedCells.forEach(change => {
        const colLetter = String.fromCharCode(64 + change.column); // A=1 → @CharCode(65)
        message += `Клітинка ${colLetter}${change.row}:
  Було: \`${change.from}\`
  Стало: \`${change.to}\`
`;
      });
      const embed = new EmbedBuilder()
        .setTitle('🔔 Виявлено зміни в таблиці')
        .setDescription(message)
        .setColor(3447003)
        .setTimestamp();
      await channel.send({ embeds: [embed] });
    }
    previousData = currentRows;
  } catch (err) {
    console.error('❌ Не вдалося перевірити зміни:', err);
  }
}

// ────────────────
// 🚀 Логін і реєстрація команд
// ────────────────
client.once('ready', async () => {
  console.log(`Бот ${client.user.tag} онлайн!`);
  try {
    await rest.put(Routes.applicationCommands(client.user.id), { body: commands });
    console.log('Slash-команди зареєстровані!');
  } catch (error) {
    console.error('Не вдалося зареєструвати команди:', error);
  }
  // Автоматична перевірка кожні 5 хвилин
  setInterval(() => checkForChanges(client), 300000); // 5 хвилин
});

// ────────────────
// 📊 Обробка слеш-команд
// ────────────────
client.on('interactionCreate', async interaction => {
  if (!interaction.isChatInputCommand()) return;
  try {
    switch (interaction.commandName) {
      case 'залишки':
        const res = await fetch(CELLS_URL);
        const data = await res.json();
        const cellValues = data.values?.flat() || [];
        const totalValue = Number(cellValues[0]) || 0;
        const totalQuantity = Number(cellValues[1]) || 0;
        const avgPrice = Number(cellValues[2]) || 0;
        const embed = new EmbedBuilder()
          .setTitle('📊 Загальні залишки')
          .addFields(
            { name: 'Загальна сума', value: `${totalValue.toFixed(2)} грн`, inline: true },
            { name: 'Кількість', value: `${totalQuantity} шт.`, inline: true },
            { name: 'Середня ціна', value: `${avgPrice.toFixed(2)} грн`, inline: true }
          )
          .setColor(5763719)
          .setFooter({ text: 'Фінансова служба' })
          .setTimestamp();
        await interaction.reply({ embeds: [embed], ephemeral: false });
        break;
      case 'оновити':
        const resUpdate = await fetch(API_URL);
        const dataUpdate = await resUpdate.json();
        const rowsUpdate = dataUpdate.values?.slice(1) || [];
        const headersUpdate = dataUpdate.values?.[0] || [];
        if (rowsUpdate.length === 0) {
          await interaction.reply({ content: '⚠️ Таблиця порожня.', ephemeral: false });
          return;
        }
        let output = '| Назва       | Кількість | Ціна |
|--------------|------------|--------|
';
        for (let i = Math.max(0, rowsUpdate.length - 10); i < rowsUpdate.length; i++) {
          const row = rowsUpdate[i];
          const name = row[6] || '—';
          const quantity = row[3] || '—';
          const price = row[4] || '—';
          output += `| ${name.padEnd(13).slice(0, 13)} | ${quantity.toString().padStart(9)} | ${price.toString().padStart(6)} |
`;
        }
        const embedUpdate = new EmbedBuilder()
          .setTitle('🔄 Останні записи')
          .setDescription(`\`\`\`md
${output}\`\`\``)
          .setColor(3066993)
          .setFooter({ text: 'Виведено останні 10 записів' });
        await interaction.reply({ embeds: [embedUpdate], ephemeral: false });
        break;
      case 'порожні':
        const resLowStock = await fetch(API_URL);
        const dataLowStock = await resLowStock.json();
        const rowsLowStock = dataLowStock.values?.slice(1) || [];
        const headersLowStock = dataLowStock.values?.[0] || [];
        const lowStock = rowsLowStock.filter(row => Number(row[3] || 0) <= 5);
        if (lowStock.length === 0) {
          await interaction.reply({ content: '🟢 Усі товари в наявності.', ephemeral: false });
          return;
        }
        output = '';
        for (let i = 0; i < Math.min(10, lowStock.length); i++) {
          const row = lowStock[i];
          const name = row[6] || '—';
          const quantity = row[3] || '—';
          output += `
• ${name} | Кількість: ${quantity}`;
        }
        const embedLowStock = new EmbedBuilder()
          .setTitle('⚠️ Мало товару')
          .setDescription(output)
          .setColor(15158332);
        const rowButtons = new ActionRowBuilder()
          .addComponents(
            new ButtonBuilder()
              .setCustomId('download_excel_low_stock')
              .setLabel('Завантажити Excel')
              .setStyle(ButtonStyle.Success)
          );
        await interaction.reply({ embeds: [embedLowStock], components: [rowButtons], ephemeral: false });
        break;
      case 'пошук':
      case 'пошук-експортовано':
        const field = interaction.options.getString('поле');
        const query = interaction.options.getString('запит').toLowerCase();
        const resSearch = await fetch(API_URL);
        const dataSearch = await resSearch.json();
        const rowsSearch = dataSearch.values?.slice(1) || [];
        const headersSearch = dataSearch.values?.[0] || [];
        let colIndex = -1;
        switch (field) {
          case 'назва': colIndex = 6; break;
          case 'серія': colIndex = 7; break;
          case 'контрагент': colIndex = 5; break;
          case 'кількість': colIndex = 3; break;
          case 'ціна': colIndex = 4; break;
        }
        if (colIndex === -1) {
          await interaction.reply({ content: '❌ Невідоме поле для пошуку.', ephemeral: false });
          return;
        }
        const isNumericField = ['кількість', 'ціна'].includes(field);
        const results = rowsSearch.filter(row => {
          const value = row[colIndex]?.toString().toLowerCase() || '';
          return isNumericField ? Number(value) >= Number(query) : value.includes(query);
        });
        if (results.length === 0) {
          await interaction.reply({ content: '🔍 Нічого не знайдено.', ephemeral: false });
          return;
        }
        output = '| Назва       | Кількість | Ціна |
|--------------|------------|--------|
';
        for (let i = 0; i < Math.min(10, results.length); i++) {
          const row = results[i];
          const name = row[6] || '—';
          const quantity = row[3] || '—';
          const price = row[4] || '—';
          output += `| ${name.padEnd(13).slice(0, 13)} | ${quantity} | ${price} |
`;
        }
        const embedSearch = new EmbedBuilder()
          .setTitle(`🔍 Результати пошуку (${results.length})`)
          .setDescription(`\`\`\`md
${output}\`\`\``)
          .setColor(3066993);
        if (interaction.commandName === 'пошук') {
          const rowButtons = new ActionRowBuilder()
            .addComponents(
              new ButtonBuilder()
                .setCustomId('download_excel_search')
                .setLabel('Завантажити Excel')
                .setStyle(ButtonStyle.Primary)
            );
          await interaction.reply({ embeds: [embedSearch], components: [rowButtons], ephemeral: false });
        } else if (interaction.commandName === 'пошук-експортовано') {
          const exportData = [headersSearch, ...results]; // додаємо заголовки + результати
          const worksheet = XLSX.utils.aoa_to_sheet(exportData);
          const workbook = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(workbook, worksheet, 'Результати пошуку');
          const filePath = './search_results.xlsx';
          XLSX.writeFile(workbook, filePath);
          await interaction.reply({
            content: '📊 Експортуємо файл...',
            files: [filePath],
            ephemeral: false
          });
        }
        break;
      case 'розумний-пошук':
        const resSmartSearch = await fetch(API_URL);
        const dataSmartSearch = await resSmartSearch.json();
        const rowsSmartSearch = dataSmartSearch.values?.slice(1) || [];
        const headersSmartSearch = dataSmartSearch.values?.[0] || [];
        const filters = {
          name: interaction.options.getString('номенклатура'),
          client: interaction.options.getString('контрагент'),
          series: interaction.options.getString('серія'),
          priceMin: interaction.options.getNumber('ціна_вище'),
          quantityMin: interaction.options.getNumber('кількість_вище')
        };
        const smartResults = rowsSmartSearch.filter(row => {
          const nameMatch = !filters.name || row[6]?.toLowerCase().includes(filters.name.toLowerCase());
          const clientMatch = !filters.client || row[5]?.toLowerCase().includes(filters.client.toLowerCase());
          const seriesMatch = !filters.series || row[7]?.toLowerCase().includes(filters.series.toLowerCase());
          const priceMatch = !filters.priceMin || Number(row[4] || 0) >= filters.priceMin;
          const quantityMatch = !filters.quantityMin || Number(row[3] || 0) >= filters.quantityMin;
          return nameMatch && clientMatch && seriesMatch && priceMatch && quantityMatch;
        });
        if (smartResults.length === 0) {
          await interaction.reply({ content: '🔍 Нічого не знайдено.', ephemeral: false });
          return;
        }
        output = '| Назва       | Кількість | Ціна |
|--------------|------------|--------|
';
        for (let i = 0; i < Math.min(10, smartResults.length); i++) {
          const row = smartResults[i];
          const name = row[6] || '—';
          const quantity = row[3] || '—';
          const price = row[4] || '—';
          output += `| ${name.padEnd(13).slice(0, 13)} | ${quantity} | ${price} |
`;
        }
        const embedSmartSearch = new EmbedBuilder()
          .setTitle(`🔍 Результати розумного пошуку (${smartResults.length})`)
          .setDescription(`\`\`\`md
${output}\`\`\``)
          .setColor(3066993);
        const rowSmartExport = new ActionRowBuilder()
          .addComponents(
            new ButtonBuilder()
              .setCustomId('download_excel_smart')
              .setLabel('Завантажити Excel')
              .setStyle(ButtonStyle.Success)
          );
        await interaction.reply({ embeds: [embedSmartSearch], components: [rowSmartExport], ephemeral: false });
        break;
      case 'експорт':
        const resExport = await fetch(API_URL);
        const dataExport = await resExport.json();
        const exportRows = dataExport.values || [];
        const worksheet = XLSX.utils.aoa_to_sheet(exportRows);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Дані');
        const filePathExport = './table.xlsx';
        XLSX.writeFile(workbook, filePathExport);
        await interaction.reply({
          content: '📎 Експортуємо всю таблицю...',
          files: [filePathExport],
          ephemeral: false
        });
        break;
      case 'help':
        const helpEmbed = new EmbedBuilder()
          .setTitle('📚 Допомога')
          .setDescription('Ось усі доступні команди:')
          .addFields([
            { name: '/залишки', value: 'Показує загальну кількість і суму товарів', inline: false },
            { name: '/оновити', value: 'Показує останні 10 записів', inline: false },
            { name: '/порожні', value: 'Показує товари, де кількість ≤ 5', inline: false },
            { name: '/пошук [поле] [запит]', value: 'Пошук за полями: назва, серія, контрагент', inline: false },
            { name: '/розумний-пошук', value: 'Пошук за кількома полями', inline: false },
            { name: '/пошук-експортовано [поле] [запит]', value: 'Експортує результати пошуку в Excel', inline: false },
            { name: '!додати [назва] [кількість]', value: 'Додає новий запис через Google Apps Script', inline: false },
            { name: '!експорт', value: 'Експортується вся таблиця', inline: false }
          ])
          .setColor(5763719)
          .setTimestamp();
        const rowHelp = new ActionRowBuilder()
          .addComponents(
            new ButtonBuilder()
              .setLabel('Документація')
              .setURL('https://your-docs-link-here')
              .setStyle(ButtonStyle.Link)
          );
        await interaction.reply({ embeds: [helpEmbed], components: [rowHelp], ephemeral: false });
        break;
    }
  } catch (err) {
    console.error(err);
    await interaction.reply({ content: '❌ Помилка при завантаженні даних.', ephemeral: false });
  }
});

// ────────────────
// 💬 Текстові команди
// ────────────────
client.on('messageCreate', async msg => {
  if (msg.author.bot) return;
  const args = msg.content.split(' ');
  // !додати [назва] [кількість]
  if (args[0] === '!додати') {
    if (args.length < 3) {
      msg.reply('Використання: `!додати [назва] [кількість]`');
      return;
    }
    const name = args.slice(1, -1).join(' ');
    const quantity = parseInt(args[args.length - 1]);
    if (!name || isNaN(quantity)) {
      msg.reply('❌ Неправильний формат. Приклад: `!додати ноутбук 5`');
      return;
    }
    try {
      const response = await fetch(APP_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, quantity })
      });
      const text = await response.text();
      if (text.trim() === 'OK') {
        msg.reply(`✅ Додано: "${name}" × ${quantity}`);
      } else {
        msg.reply('❌ Не вдалося додати запис.');
      }
    } catch (err) {
      console.error(err);
      msg.reply('❌ Не вдалося відправити запит до Google Apps Script.');
    }
  }
  // !експорт
  if (msg.content === '!експорт') {
    try {
      const res = await fetch(API_URL);
      const data = await res.json();
      const exportRows = data.values || [];
      const worksheet = XLSX.utils.aoa_to_sheet(exportRows);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Дані');
      const filePath = './table.xlsx';
      XLSX.writeFile(workbook, filePath);
      await msg.reply({
        content: '📊 Дані експортовано:',
        files: [filePath]
      });
    } catch (err) {
      console.error(err);
      msg.reply('❌ Не вдалося згенерувати файл.');
    }
  }
});

// ────────────────
// ⚙️ Логін бота
// ────────────────
client.login(process.env.BOT_TOKEN);
=======
const { Client, GatewayIntentBits, Routes } = require('discord.js');
const fetch = (...args) => import('node-fetch').then(({ default: f }) => f(...args));
const XLSX = require('xlsx');
require('dotenv').config();

// Зчитуємо змінні з .env
const SHEET_ID = process.env.SHEET_ID;
const SHEET_NAME = 'Аркуш1';
const GOOGLE_API_KEY = process.env.GOOGLE_API_KEY;
const APP_SCRIPT_URL = process.env.APP_SCRIPT_URL;
if (!SHEET_ID || !GOOGLE_API_KEY || !APP_SCRIPT_URL) {
  console.error("❗ Один з обов'язкових ENV-параметрів відсутній");
  process.exit(1);
}
const API_URL = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${SHEET_NAME}?key=${GOOGLE_API_KEY}`;
const CELLS_URL = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/C964:E964?key=${GOOGLE_API_KEY}`;
const client = new Client({
  intents: [
    GatewayIntentBits.Guilds,
    GatewayIntentBits.GuildMessages,
    GatewayIntentBits.MessageContent,
    GatewayIntentBits.GuildMessageReactions
  ]
});

// Реєстрація слеш-команд
const { REST } = require('@discordjs/rest');
const { version } = require('discord.js').Constants;
const commands = [
  {
    name: 'залишки',
    description: 'Показує загальну кількість товарів'
  },
  {
    name: 'оновити',
    description: 'Принудово оновлює дані'
  },
  {
    name: 'порожні',
    description: 'Показує товари з мінімальною кількістю'
  },
  {
    name: 'пошук',
    description: 'Пошук за полями таблиці',
    options: [
      {
        name: 'поле',
        description: 'За яким полем шукати',
        type: 3,
        required: true,
        choices: [
          { name: 'Найменування', value: 'назва' },
          { name: 'Серійний номер', value: 'серія' },
          { name: 'Контрагент', value: 'контрагент' },
          { name: 'Кількість', value: 'кількість' },
          { name: 'Ціна', value: 'ціна' }
        ]
      },
      {
        name: 'запит',
        description: 'Що шукати (наприклад: "стол", "5")',
        type: 3,
        required: true
      }
    ]
  },
  {
    name: 'розумний-пошук',
    description: 'Пошук за кількома полями',
    options: [
      { name: 'номенклатура', type: 3, description: 'Шукати за назвою товару' },
      { name: 'контрагент', type: 3, description: 'Шукати за контрагентом' },
      { name: 'серія', type: 3, description: 'Шукати за серійним номером' },
      { name: 'ціна_вище', type: 10, description: 'Показувати товари дорожче цього значення' },
      { name: 'кількість_вище', type: 10, description: 'Показувати товари з кількістю більше' }
    ]
  },
  {
    name: 'пошук-експортовано',
    description: 'Експортує результати пошуку в Excel (.xlsx)',
    options: [
      {
        name: 'поле',
        description: 'За яким полем шукати',
        type: 3,
        required: true,
        choices: [
          { name: 'Найменування', value: 'назва' },
          { name: 'Серійний номер', value: 'серія' },
          { name: 'Контрагент', value: 'контрагент' },
          { name: 'Кількість', value: 'кількість' },
          { name: 'Ціна', value: 'ціна' }
        ]
      },
      {
        name: 'запит',
        description: 'Що шукати (наприклад: "стол", "5")',
        type: 3,
        required: true
      }
    ]
  },
  {
    name: 'експорт',
    description: 'Експортує всю таблицю в Excel (.xlsx)'
  },
  {
    name: 'help',
    description: 'Показує список усіх доступних команд'
  }
];
const rest = new REST({ version: '10' }).setToken(process.env.BOT_TOKEN);

// ────────────────
// 🧠 Функція: автоматичне сповіщення при зміні в таблиці
// ────────────────
let previousData = null;
async function checkForChanges(botClient) {
  try {
    const res = await fetch(API_URL);
    if (!SHEET_ID || !GOOGLE_API_KEY) {
      return interaction.reply({ content: '❌ Не встановлено ключі API в .env', ephemeral: true });
    }
    const data = await res.json();
    const currentRows = data.values;
    if (!previousData) {
      previousData = currentRows;
      return;
    }
    const changedCells = [];
    for (let i = 0; i < Math.min(currentRows.length, previousData.length); i++) {
      const oldRow = previousData[i];
      const newRow = currentRows[i];
      if (!oldRow || !newRow) continue;
      for (let j = 0; j < Math.min(oldRow.length, newRow.length); j++) {
        if (oldRow[j] !== newRow[j]) {
          changedCells.push({
            row: i + 1,
            column: j + 1,
            from: oldRow[j],
            to: newRow[j]
          });
        }
      }
    }
    if (changedCells.length > 0) {
      const channel = botClient.channels.cache.find(ch => ch.name === 'склад' && ch.type === 0);
      if (!channel) return;
      let message = `🔔 Виявлено зміни в таблиці:
`;
      changedCells.forEach(change => {
        const colLetter = String.fromCharCode(64 + change.column); // A=1 → @CharCode(65)
        message += `Клітинка ${colLetter}${change.row}:
  Було: \`${change.from}\`
  Стало: \`${change.to}\`
`;
      });
      const embed = new EmbedBuilder()
        .setTitle('🔔 Виявлено зміни в таблиці')
        .setDescription(message)
        .setColor(3447003)
        .setTimestamp();
      await channel.send({ embeds: [embed] });
    }
    previousData = currentRows;
  } catch (err) {
    console.error('❌ Не вдалося перевірити зміни:', err);
  }
}

// ────────────────
// 🚀 Логін і реєстрація команд
// ────────────────
let retryCount = 0;
async function safeCheck() {
  try {
    await checkForChanges(client);
    retryCount = 0; // Сбрасываем счётчик ошибок
  } catch (err) {
    console.error(`⚠️ Помилка при перевірці змін: ${err.message}`);
    if (retryCount < 3) {
      retryCount++;
      setTimeout(safeCheck, 10000); // Повтор через 10 секунд
    }
  }
}

// Запускаємо кожні 5 хв
setInterval(safeCheck, 300000);

// ────────────────
// 📊 Обробка слеш-команд
// ────────────────
client.on('interactionCreate', async interaction => {
  if (!interaction.isChatInputCommand()) return;
  try {
    switch (interaction.commandName) {
      case 'залишки':
        if (!res.ok) {
          throw new Error(`HTTP-помилка: ${res.status}`);
        }
        const res = await fetch(CELLS_URL);
        if (!SHEET_ID || !GOOGLE_API_KEY) {
          return interaction.reply({ content: '❌ Не встановлено ключі API в .env', ephemeral: true });
        }
        const data = await res.json();
        const cellValues = data.values?.flat() || [];
const totalValue = Number(cellValues[0]) || 0;
const totalQuantity = Number(cellValues[1]) || 0;
const avgPrice = Number(cellValues[2]) || 0;
        const embed = new EmbedBuilder()
          .setTitle('📊 Загальні залишки')
          .addFields(
            { name: 'Загальна сума', value: `${totalValue.toFixed(2)} грн`, inline: true },
            { name: 'Кількість', value: `${totalQuantity} шт.`, inline: true },
            { name: 'Середня ціна', value: `${avgPrice.toFixed(2)} грн`, inline: true }
          )
          .setColor(5763719)
          .setFooter({ text: 'Фінансова служба' })
          .setTimestamp();
        await interaction.reply({ embeds: [embed], ephemeral: false});
        break;
      case 'оновити':
        const resUpdate = await fetch(API_URL);
        if (!SHEET_ID || !GOOGLE_API_KEY) {
          return interaction.reply({ content: '❌ Не встановлено ключі API в .env', ephemeral: true });
        }
        const dataUpdate = await resUpdate.json();
        const rowsUpdate = dataUpdate.values?.slice(1) || [];
        const headersUpdate = dataUpdate.values?.[0] || [];
        if (rowsUpdate.length === 0) {
          await interaction.reply({ content: '⚠️ Таблиця порожня.', ephemeral: false });
          return;
        }
        let output = '| Назва       | Кількість | Ціна |\n';
        output += '|--------------|------------|--------|\n';
        for (let i = 0; i < Math.min(10, results.length); i++) {
          const row = results[i];
          const name = row[6] || '—';
          const quantity = row[3] || '—';
          const price = row[4] || '—';
          output += `| ${name.padEnd(13).slice(0, 13)} | ${quantity} | ${price} |\n`;
        }
        const embedUpdate = new EmbedBuilder()
          .setTitle('🔄 Останні записи')
          .setDescription(`\`\`\`md
${output}\`\`\``)
          .setColor(3066993)
          .setFooter({ text: 'Виведено останні 10 записів' });
        await interaction.reply({ embeds: [embedUpdate], ephemeral: false });
        break;
      case 'порожні':
        const resLowStock = await fetch(API_URL);
        if (!SHEET_ID || !GOOGLE_API_KEY) {
          return interaction.reply({ content: '❌ Не встановлено ключі API в .env', ephemeral: true });
        }
        const dataLowStock = await resLowStock.json();
        const rowsLowStock = dataLowStock.values?.slice(1) || [];
        const headersLowStock = dataLowStock.values?.[0] || [];
        const lowStock = rowsLowStock.filter(row => Number(row[3] || 0) <= 5);
        if (lowStock.length === 0) {
          await interaction.reply({ content: '🟢 Усі товари в наявності.', ephemeral: false });
          return;
        }
        output = '';
        for (let i = 0; i < Math.min(10, lowStock.length); i++) {
          const row = lowStock[i];
          const name = row[6] || '—';
          const quantity = row[3] || '—';
          output += `
• ${name} | Кількість: ${quantity}`;
        }
        const embedLowStock = new EmbedBuilder()
          .setTitle('⚠️ Мало товару')
          .setDescription(output)
          .setColor(15158332);
        const rowButtons = new ActionRowBuilder()
          .addComponents(
            new ButtonBuilder()
              .setCustomId('download_excel_low_stock')
              .setLabel('Завантажити Excel')
              .setStyle(ButtonStyle.Success)
          );
        await interaction.reply({ embeds: [embedLowStock], components: [rowButtons], ephemeral: false });
        break;
      case 'пошук':
      case 'пошук-експортовано':
        const field = interaction.options.getString('поле');
        const query = interaction.options.getString('запит').toLowerCase();
        const value = row[colIndex]?.toString().toLowerCase() || '';
const isMatch = isNumericField 
  ? !isNaN(Number(value)) && Number(value) >= Number(query)
  : value.includes(query);
        const resSearch = await fetch(API_URL);
        if (!SHEET_ID || !GOOGLE_API_KEY) {
          return interaction.reply({ content: '❌ Не встановлено ключі API в .env', ephemeral: true });
        }
        const dataSearch = await resSearch.json();
        const rowsSearch = dataSearch.values?.slice(1) || [];
        const headersSearch = dataSearch.values?.[0] || [];
        let colIndex = -1;
        switch (field) {
          case 'назва': colIndex = 6; break;
          case 'серія': colIndex = 7; break;
          case 'контрагент': colIndex = 5; break;
          case 'кількість': colIndex = 3; break;
          case 'ціна': colIndex = 4; break;
        }
        if (colIndex === -1) {
          await interaction.reply({ content: '❌ Невідоме поле для пошуку.', ephemeral: false });
          return;
        }
        const isNumericField = ['кількість', 'ціна'].includes(field);
        const results = rowsSearch.filter(row => {
          const value = row[colIndex]?.toString().toLowerCase() || '';
          return isNumericField ? Number(value) >= Number(query) : value.includes(query);
        });
        if (results.length === 0) {
          await interaction.reply({ content: '🔍 Нічого не знайдено.', ephemeral: false });
          return;
        }
      output = '| Назва       | Кількість | Ціна |\n';
        output += '|--------------|------------|--------|\n';
        for (let i = 0; i < Math.min(10, results.length); i++) {
          const row = results[i];
          const name = row[6] || '—';
          const quantity = row[3] || '—';
          const price = row[4] || '—';
          output += `| ${name.padEnd(13).slice(0, 13)} | ${quantity} | ${price} |\n`;
        }
        const embedSearch = new EmbedBuilder()
          .setTitle(`🔍 Результати пошуку (${results.length})`)
          .setDescription(`\`\`\`md
${output}\`\`\``)
          .setColor(3066993);
        if (interaction.commandName === 'пошук') {
          const rowButtons = new ActionRowBuilder()
            .addComponents(
              new ButtonBuilder()
                .setCustomId('download_excel_search')
                .setLabel('Завантажити Excel')
                .setStyle(ButtonStyle.Primary)
            );
          await interaction.reply({ embeds: [embedSearch], components: [rowButtons], ephemeral: false });
        } else if (interaction.commandName === 'пошук-експортовано') {
          const exportData = [headersSearch, ...results]; // додаємо заголовки + результати
          const worksheet = XLSX.utils.aoa_to_sheet(exportData);
          if (!headers || !Array.isArray(headers)) {
            headers = ['ID', 'Дата', 'Контрагент', 'Кількість', 'Ціна', 'Контрагент', 'Назва', 'Серія'];
          }
          const workbook = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(workbook, worksheet, 'Результати пошуку');
          const filePath = './search_results.xlsx';
          XLSX.writeFile(workbook, filePath);
          await interaction.reply({
            content: '📊 Експортуємо файл...',
            files: [filePath],
            ephemeral: false
          });
        }
        break;
      case 'розумний-пошук':
        const resSmartSearch = await fetch(API_URL);
        if (!SHEET_ID || !GOOGLE_API_KEY) {
          return interaction.reply({ content: '❌ Не встановлено ключі API в .env', ephemeral: true });
        }
        const dataSmartSearch = await resSmartSearch.json();
        const rowsSmartSearch = dataSmartSearch.values?.slice(1) || [];
        const headersSmartSearch = dataSmartSearch.values?.[0] || [];
        const filters = {
          name: interaction.options.getString('номенклатура'),
          client: interaction.options.getString('контрагент'),
          series: interaction.options.getString('серія'),
          priceMin: interaction.options.getNumber('ціна_вище'),
          quantityMin: interaction.options.getNumber('кількість_вище')
        };
        const smartResults = rowsSmartSearch.filter(row => {
          const nameMatch = !filters.name || row[6]?.toLowerCase().includes(filters.name.toLowerCase());
          const clientMatch = !filters.client || row[5]?.toLowerCase().includes(filters.client.toLowerCase());
          const seriesMatch = !filters.series || row[7]?.toLowerCase().includes(filters.series.toLowerCase());
          const priceMatch = !filters.priceMin || Number(row[4] || 0) >= filters.priceMin;
          const quantityMatch = !filters.quantityMin || Number(row[3] || 0) >= filters.quantityMin;
          return nameMatch && clientMatch && seriesMatch && priceMatch && quantityMatch;
        });
        if (smartResults.length === 0) {
          await interaction.reply({ content: '🔍 Нічого не знайдено.', ephemeral: false });
          return;
        }
        output = '| Назва       | Кількість | Ціна |\n';
        output += '|--------------|------------|--------|\n';
        for (let i = 0; i < Math.min(10, results.length); i++) {
          const row = results[i];
          const name = row[6] || '—';
          const quantity = row[3] || '—';
          const price = row[4] || '—';
          output += `| ${name.padEnd(13).slice(0, 13)} | ${quantity} | ${price} |\n`;
        }
        const embedSmartSearch = new EmbedBuilder()
          .setTitle(`🔍 Результати розумного пошуку (${smartResults.length})`)
          .setDescription(`\`\`\`md
${output}\`\`\``)
          .setColor(3066993);
        const rowSmartExport = new ActionRowBuilder()
          .addComponents(
            new ButtonBuilder()
              .setCustomId('download_excel_smart')
              .setLabel('Завантажити Excel')
              .setStyle(ButtonStyle.Success)
          );
        await interaction.reply({ embeds: [embedSmartSearch], components: [rowSmartExport], ephemeral: false });
        break;
      case 'експорт':
        const resExport = await fetch(API_URL);
        if (!SHEET_ID || !GOOGLE_API_KEY) {
          return interaction.reply({ content: '❌ Не встановлено ключі API в .env', ephemeral: true });
        }
        const dataExport = await resExport.json();
        const exportRows = dataExport.values || [];
        const worksheet = XLSX.utils.aoa_to_sheet(exportRows);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Дані');
        const filePathExport = './table.xlsx';
        XLSX.writeFile(workbook, filePathExport);
        await interaction.reply({
          content: '📎 Експортуємо всю таблицю...',
          files: [filePathExport],
          ephemeral: false
        });
        break;
      case 'help':
        const helpEmbed = new EmbedBuilder()
          .setTitle('📚 Допомога')
          .setDescription('Ось усі доступні команди:')
          .addFields([
            { name: '/залишки', value: 'Показує загальну кількість і суму товарів', inline: false },
            { name: '/оновити', value: 'Показує останні 10 записів', inline: false },
            { name: '/порожні', value: 'Показує товари, де кількість ≤ 5', inline: false },
            { name: '/пошук [поле] [запит]', value: 'Пошук за полями: назва, серія, контрагент', inline: false },
            { name: '/розумний-пошук', value: 'Пошук за кількома полями', inline: false },
            { name: '/пошук-експортовано [поле] [запит]', value: 'Експортує результати пошуку в Excel', inline: false },
            { name: '!додати [назва] [кількість]', value: 'Додає новий запис через Google Apps Script', inline: false },
            { name: '!експорт', value: 'Експортується вся таблиця', inline: false }
          ])
          .setColor(5763719)
          .setTimestamp();
        const rowHelp = new ActionRowBuilder()
          .addComponents(
            new ButtonBuilder()
              .setLabel('Документація')
              .setURL('https://your-docs-link-here')
              .setStyle(ButtonStyle.Link)
          );
        await interaction.reply({ embeds: [helpEmbed], components: [rowHelp], ephemeral: false });
        break;
    }
  } catch (err) {
    console.error(err);
    await interaction.reply({ content: '❌ Помилка при завантаженні даних.', ephemeral: false });
  }
});

// ────────────────
// 💬 Текстові команди
// ────────────────
client.on('messageCreate', async msg => {
  if (msg.author.bot) return;
  const args = msg.content.split(' ');
  // !додати [назва] [кількість]
  if (args[0] === '!додати') {
    if (args.length < 3) {
      msg.reply('Використання: `!додати [назва] [кількість]`');
      return;
    }
    const name = args.slice(1, -1).join(' ');
    const quantity = parseInt(args[args.length - 1]);
    if (quantity <= 0) {
      msg.reply('❌ Кількість має бути більше 0.');
      return;
    } 
    if (!name || isNaN(quantity)) {
      msg.reply('❌ Неправильний формат. Приклад: `!додати ноутбук 5`');
      return;
    }
    try {
      const response = await fetch(APP_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, quantity })
      });
      const text = await response.text();
      if (text.trim() === 'OK') {
        msg.reply(`✅ Додано: "${name}" × ${quantity}`);
      } else {
        msg.reply('❌ Не вдалося додати запис.');
      }
    } catch (err) {
      console.error(err);
      msg.reply('❌ Не вдалося відправити запит до Google Apps Script.');
  }
  }
  // !експорт
  if (msg.content === '!експорт') {
    try {
      const res = await fetch(API_URL);
      if (!SHEET_ID || !GOOGLE_API_KEY) {
        return interaction.reply({ content: '❌ Не встановлено ключі API в .env', ephemeral: true });
      }
      const data = await res.json();
      const exportRows = data.values || [];
      const worksheet = XLSX.utils.aoa_to_sheet(exportRows);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Дані');
      const filePath = './table.xlsx';
      XLSX.writeFile(workbook, filePath);
      await msg.reply({
        content: '📊 Дані експортовано:',
        files: [filePath]
      });
    } catch (err) {
      console.error(err);
      msg.reply('❌ Не вдалося згенерувати файл.');
    }
  }
});

// ────────────────
// ⚙️ Логін бота
// ────────────────
client.login(process.env.BOT_TOKEN);
>>>>>>> master
