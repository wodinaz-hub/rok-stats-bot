import pandas as pd
import discord
from discord.ext import commands
import matplotlib.pyplot as plt
import os
from dotenv import load_dotenv

load_dotenv()  # Загружаем переменные из .env
TOKEN = os.getenv("DISCORD_TOKEN")  # Получаем токен



# === Чтение и подготовка данных ===
def load_and_prepare_data(before_file, after_file, requirements_file):
    # Чтение файлов
    before = pd.read_excel(before_file, header=0)
    after = pd.read_excel(after_file, header=0)
    requirements = pd.read_excel(requirements_file, header=0)

    # Проверка на пустые файлы
    if before.empty:
        print(f"Ошибка: Файл '{before_file}' пуст или отсутствуют данные.")
        exit()
    if after.empty:
        print(f"Ошибка: Файл '{after_file}' пуст или отсутствуют данные.")
        exit()
    if requirements.empty:
        print(f"Ошибка: Файл '{requirements_file}' пуст или отсутствуют данные.")
        exit()

    # Очистка названий столбцов
    before.columns = before.columns.str.strip()
    after.columns = after.columns.str.strip()
    requirements.columns = requirements.columns.str.strip()

    # Проверка наличия столбца 'Governor ID'
    if 'Governor ID' not in before.columns:
        print(f"Ошибка: В файле '{before_file}' отсутствует столбец 'Governor ID'.")
        exit()

    # Приведение данных к строковому типу
    before['Governor ID'] = before['Governor ID'].fillna('').astype(str).str.strip()
    after['Governor ID'] = after['Governor ID'].fillna('').astype(str).str.strip()
    requirements['Governor ID'] = requirements['Governor ID'].fillna('').astype(str).str.strip()

    return before, after, requirements


# === Расчёт изменений статистики ===
def calculate_stats(before, after, requirements):
    # Объединение данных
    result = before.merge(after, on='Governor ID', suffixes=('_before', '_after'))
    result = result.merge(requirements, on='Governor ID')

    # Вычисление изменений
    result['Kills Change'] = result['Kill Points_after'] - result['Kill Points_before']
    result['Deads Change'] = result['Deads_after'] - result['Deads_before']  # Использование 'Deads'
    result['Kills Completion (%)'] = (result['Kill Points_after'] / result['Required Kills']) * 100
    result['Deaths Completion (%)'] = (result['Deads_after'] / result['Required Deaths']) * 100  # Использование 'Deads'

    # Сохранение результата
    result.to_excel('results.xlsx', index=False)
    print("Результаты сохранены в 'results.xlsx'")
    return result


# === Создание круговых диаграмм ===
def create_pie_chart(player_name, kills_completion, deaths_completion):
    # Диаграмма для убийств
    labels_kills = ['Выполнено', 'Не выполнено']
    data_kills = [kills_completion, max(0, 100 - kills_completion)]
    colors = ['green', 'red']

    plt.figure(figsize=(6, 6))
    plt.pie(data_kills, labels=labels_kills, colors=colors, autopct='%1.1f%%', startangle=140)
    plt.title(f'{player_name} - Убийства')
    plt.savefig(f'{player_name}_kills.png')
    plt.close()

    # Диаграмма для смертей
    labels_deaths = ['Выполнено', 'Не выполнено']
    data_deaths = [deaths_completion, max(0, 100 - deaths_completion)]

    plt.figure(figsize=(6, 6))
    plt.pie(data_deaths, labels=labels_deaths, colors=colors, autopct='%1.1f%%', startangle=140)
    plt.title(f'{player_name} - Смерти')
    plt.savefig(f'{player_name}_deaths.png')
    plt.close()


# === Discord-бот ===
intents = discord.Intents.default()  # Создание объекта intents с настройками по умолчанию
intents.message_content = True      # Включение доступа к содержимому сообщений

bot = commands.Bot(command_prefix='!', intents=intents)  # Передача intents в бота


@bot.event
async def on_ready():
    print(f'Бот {bot.user} успешно запущен!')

@bot.command()
async def commands(ctx):
    await ctx.send(
        "Список доступных команд:\n"
        "!stats <Governor ID> - Получить статистику игрока.\n"
        "!overview - Получить общую сводку по всем игрокам.\n"
        "!requirements - Проверить выполнение требований.\n"
    )

@bot.command()
async def req(ctx):
    try:
        # Загрузка данных
        result = pd.read_excel('results.xlsx')

        # Фильтрация игроков, которые не выполнили требования
        not_completed = result[
            (result['Kills Completion (%)'] < 100) | (result['Deaths Completion (%)'] < 100)
        ]

        if not not_completed.empty:
            await ctx.send("📋 Игроки, не выполнившие требования:")
            for index, row in not_completed.iterrows():
                await ctx.send(
                    f"{row['Governor Name']} (ID: {row['Governor ID']}) - "
                    f"Убийства: {row['Kills Completion (%)']:.2f}%, "
                    f"Смерти: {row['Deaths Completion (%)']:.2f}%"
                )
        else:
            await ctx.send("Все игроки выполнили требования!")
    except Exception as e:
        await ctx.send(f"Произошла ошибка: {str(e)}")

@bot.command()
async def overview(ctx):
    try:
        # Загрузка данных
        result = pd.read_excel('results.xlsx')

        # Расчёт сводных показателей
        average_kills = result['Kills Change'].mean()
        average_deads = result['Deads Change'].mean()
        average_completion_kills = result['Kills Completion (%)'].mean()
        average_completion_deads = result['Deaths Completion (%)'].mean()

        # Отправка данных в Discord
        await ctx.send(
            f"📊 Общая статистика:\n"
            f"Средние изменения в убийствах: {average_kills:.2f}\n"
            f"Средние изменения в смертях: {average_deads:.2f}\n"
            f"Средний процент выполнения убийств: {average_completion_kills:.2f}%\n"
            f"Средний процент выполнения смертей: {average_completion_deads:.2f}%"
        )
    except Exception as e:
        await ctx.send(f"Произошла ошибка: {str(e)}")

@bot.command()
async def stats(ctx, player_id):
    try:
        # Load data
        result = pd.read_excel('results.xlsx')
        result['Governor ID'] = result['Governor ID'].astype(str).str.strip()
        player_id = str(player_id).strip()

        # Calculate DKP for all players
        result['DKP'] = (
            (result['Deads_after'] - result['Deads_before']) * 15 +
            (result['Tier 5 Kills_after'] - result['Tier 5 Kills_before']) * 10 +
            (result['Tier 4 Kills_after'] - result['Tier 4 Kills_before']) * 4
        )

        # Rank players by DKP
        result['Rank'] = result['DKP'].rank(ascending=False, method='min')

        # Search for player data
        player_data = result[result['Governor ID'] == player_id]

        if player_data.empty:
            await ctx.send(f"Player with ID {player_id} not found.")
            return

        # Calculate changes
        matchmaking_power = player_data['Power_before'].values[0]
        power_change = player_data['Power_after'].values[0] - player_data['Power_before'].values[0]
        tier4_kills_change = (
            player_data['Tier 4 Kills_after'].values[0] - player_data['Tier 4 Kills_before'].values[0]
        )
        tier5_kills_change = (
            player_data['Tier 5 Kills_after'].values[0] - player_data['Tier 5 Kills_before'].values[0]
        )
        kills_change = tier4_kills_change + tier5_kills_change
        kill_points_change = player_data['Kill Points_after'].values[0] - player_data['Kill Points_before'].values[0]
        deads_change = player_data['Deads_after'].values[0] - player_data['Deads_before'].values[0]

        # Initialize progress variables with default values (to avoid errors)
        kills_completion = 0
        deads_completion = 0

        # Calculate progress
        if kills_change > 0:
            kills_completion = (kills_change / player_data['Required Kills'].values[0]) * 100
        if deads_change > 0:
            deads_completion = (deads_change / player_data['Required Deaths'].values[0]) * 100

        # Round all numeric values and format with thousand separators
        matchmaking_power = f"{round(matchmaking_power):,}".replace(",", ".")
        power_change = f"{round(power_change):,}".replace(",", ".")
        tier4_kills_change = f"{round(tier4_kills_change):,}".replace(",", ".")
        tier5_kills_change = f"{round(tier5_kills_change):,}".replace(",", ".")
        kills_change = f"{round(kills_change):,}".replace(",", ".")
        kill_points_change = f"{round(kill_points_change):,}".replace(",", ".")
        deads_change = f"{round(deads_change):,}".replace(",", ".")
        dkp = f"{round(player_data['DKP'].values[0]):,}".replace(",", ".")
        kills_completion = round(kills_completion)
        deads_completion = round(deads_completion)
        required_kills = f"{round(player_data['Required Kills'].values[0]):,}".replace(",", ".")
        required_deaths = f"{round(player_data['Required Deaths'].values[0]):,}".replace(",", ".")
        rank = int(player_data['Rank'].values[0])

        # Send player statistics
        await ctx.send(f"📊 **Player Statistics:** {player_data['Governor Name'].values[0]} (ID: {player_id})\n"
                       f"🔹 **Matchmaking Power:** {matchmaking_power}\n"
                       f"🔹 **Power Change:** {power_change}\n"
                       f"🔹 **Kill Points (Gained):** {kill_points_change}\n"
                       f"🔹 **Tier 4 Kills (Gained):** {tier4_kills_change}\n"
                       f"🔹 **Tier 5 Kills (Gained):** {tier5_kills_change}\n"
                       f"🔹 **Total Kills (T4 + T5):** {kills_change}\n"
                       f"🔹 **Required Kills:** {required_kills}\n"
                       f"🔹 **Kill Progress (%):** {kills_completion}%\n"
                       f"🔹 **Deaths (Gained):** {deads_change}\n"
                       f"🔹 **Required Deaths:** {required_deaths}\n"
                       f"🔹 **Death Progress (%):** {deads_completion}%\n"
                       f"🔹 **DKP:** {dkp}\n"
                       f"🔹 **DKP Rank:** #{rank}")
    except Exception as e:
        await ctx.send(f"An error occurred: {str(e)}")


# === Основной код ===
def main():
    # Укажите пути к файлам
    before_file = 'start_kvk.xlsx'
    after_file = 'pass4.xlsx'
    requirements_file = 'required.xlsx'

    # Загрузка и обработка данных
    before, after, requirements = load_and_prepare_data(before_file, after_file, requirements_file)

    # Расчёт статистики
    calculate_stats(before, after, requirements)

    # Запуск бота
    bot.run(TOKEN)


# Запуск программы
if __name__ == "__main__":
    main()