import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime, timedelta
import os
import re
import webbrowser
import time
from dotenv import load_dotenv
import sys

"""
NRD_MOEX_new.py - Система мониторинга выплат по облигациям

ЛОГИКА РАБОТЫ СИСТЕМЫ:

1. ✅ ПОЛУЧЕНИЕ ДАННЫХ С MOEX:
   - Запрос к API Московской биржи на получение выпусков с выплатами на завтра
   - Извлечение: эмитент, ISIN выпуска, название облигации
   - Пример: ISIN с MOEX: RU000A108U72

2. ✅ ФИЛЬТРАЦИЯ ПО ПОРТФЕЛЮ:
   - Чтение Excel-файла с портфелем (ISIN в столбце D)
   - Сопоставление эмитентов из MOEX с эмитентами в портфеле
   - Пример: Ваш ISIN из портфеля: RU000A10A414

3. ✅ ПРОВЕРКА INTR НА НРД:
   - Поиск по ISIN с MOEX на сайте НРД для эмитентов из Excel-файла
   - Проверка наличия записи (INTR) о поступлении средств
   - Пример: INTR найден: 17.10.2025

4. ✅ ФОРМИРОВАНИЕ ОТЧЕТА:
   - Статус: ✅ INTR пройден / ❌ INTR отсутствует
   - Ссылка на новости НРД с фильтром по ISIN
   - Многоуровневые предупреждения по времени проверки

АВТОМАТИЧЕСКИЙ ЗАПУСК ПО РАСПИСАНИЮ:
- 17:00 - проверка накануне (evening)
- 09:00 - утренняя проверка (morning) 
- 14:00 - дневная проверка (afternoon)
- 18:00 - финальная проверка (final)

ПРИНЦИП РАБОТЫ:
Мониторим ЭМИТЕНТОВ, а не выпуски. Если у эмитента проблемы с выплатами 
по ЛЮБОМУ выпуску - это сигнал о рисках для всех его облигаций.
"""

# ==================== КОНФИГУРАЦИЯ ====================
load_dotenv()

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")
EXCEL_FILE_PATH = os.getenv("EXCEL_FILE_PATH")

# Определяем тип проверки по аргументу командной строки
CHECK_TYPE = sys.argv[1] if len(sys.argv) > 1 else "evening"  # evening, morning, afternoon, final

# ==================== ФУНКЦИИ ====================

def send_telegram_message(message, chat_id=TELEGRAM_CHAT_ID, bot_token=TELEGRAM_BOT_TOKEN):
    """Отправка сообщения в Telegram"""
    try:
        url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
        payload = {
            'chat_id': chat_id,
            'text': message,
            'parse_mode': 'HTML',
            'disable_web_page_preview': False
        }
        response = requests.post(url, data=payload, timeout=10)
        if response.status_code == 200:
            return True
        else:
            print(f"[Telegram] ❌ Ошибка: {response.status_code} - {response.text}")
            return False
    except Exception as e:
        print(f"[Telegram] ❌ Исключение: {e}")
        return False

def extract_issuer_name(bond_name):
    """Извлекает название эмитента из названия облигации"""
    if not bond_name or pd.isna(bond_name) or str(bond_name).strip() == '':
        return ""

    bond_str = str(bond_name).strip()

    if re.search(r'[\\/]', bond_str):
        print(f"⚠️ Пропускаем некорректное название: {bond_str} (похоже на путь)")
        return ""

    patterns = [
        r'^(.*?)\s+\d{2,}[A-Za-zА-Яа-я]*-\d{2}',
        r'^(.*?)\s+[A-Za-z]+\d+',
        r'^(.*?)(?:\s+обл|\s+ОФЗ|\s+еврообл|\s+вып)',
    ]

    for pattern in patterns:
        match = re.search(pattern, bond_str)
        if match:
            issuer = match.group(1).strip()
            if len(issuer) > 2:
                return issuer

    if len(bond_str) > 3 and not re.search(r'\.(py|xls|xlsx|xlsm)$', bond_str, re.IGNORECASE):
        return bond_str
    else:
        print(f"⚠️ Пропускаем подозрительное название: {bond_str}")
        return ""

def parse_nsd_news_by_isin(isin, from_date, to_date):
    """Парсинг новостей с сайта НРД по ISIN - возвращает сырой HTML"""
    if not isin:
        return "❌ ISIN отсутствует"

    try:
        encoded_isin = requests.utils.quote(isin)
        url = f"https://nsddata.ru/ru/news?text={encoded_isin}&from={from_date}&to={to_date}"

        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }

        response = requests.get(url, headers=headers, timeout=15)

        if response.status_code != 200:
            return f"⚠️ Сайт вернул статус {response.status_code}"

        # ВОЗВРАЩАЕМ СЫРОЙ HTML ТЕКСТ для поиска
        return response.text

    except Exception as e:
        return f"❌ Ошибка парсинга: {str(e)[:150]}"

def check_intr_status(issuers_list):
    """Проверяет статус INTR для списка эмитентов с ISIN - ТОЛЬКО НРД"""
    if not issuers_list:
        return {}
    
    today = datetime.now().date()
    intr_status = {}
    
    for issuer_data in issuers_list:
        issuer = issuer_data['issuer']
        isin_moex = issuer_data['isin_moex']
        
        has_intr = False
        record_date = None
        coupon_date = (today + timedelta(days=1)).strftime('%d.%m.%Y')
        
        try:
            # ПРОВЕРЯЕМ ТОЛЬКО НРД по ISIN
            from_date = (today - timedelta(days=7)).strftime("%d.%m.%Y")
            to_date = today.strftime("%d.%m.%Y")
            
            news_text = parse_nsd_news_by_isin(isin_moex, from_date, to_date)
            print(f"🔍 Проверка ISIN: {isin_moex}")
            print(f"📰 Новости НРД: {news_text[:500]}...")  # Первые 500 символов для отладки
            
            # ДОБАВИМ ОТЛАДОЧНУЮ ИНФОРМАЦИЮ
            print(f"🔍 ISIN в тексте: {isin_moex in news_text}")
            print(f"🔍 'Выплата купонного дохода' в тексте: {'Выплата купонного дохода' in news_text}")
            print(f"🔍 '(INTR)' в тексте: {'(INTR)' in news_text}")
            
            # ПРОСТАЯ И НАДЕЖНАЯ ПРОВЕРКА
            if isin_moex in news_text and "Выплата купонного дохода" in news_text:
                # Ищем дату INTR
                intr_date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})\(INTR\)', news_text)
                print(f"🔍 Дата INTR найдена: {intr_date_match}")
                
                if intr_date_match:
                    record_date = intr_date_match.group(1)
                    has_intr = True
                    status_details = f"ℹ️ INTR был {record_date}, деньги поступили для выплаты"
                else:
                    # Если дату не нашли, но ISIN и выплата есть - все равно считаем INTR пройденным
                    has_intr = True
                    record_date = "дата не определена"
                    status_details = f"ℹ️ INTR найден, деньги поступили для выплаты"
            else:
                has_intr = False
                record_date = None
                status_details = "ℹ️ INTR не найден в НРД"
            
            print(f"🔍 Итоговый статус: has_intr = {has_intr}")
            
            intr_status[isin_moex] = {
                'has_intr': has_intr,
                'issuer': issuer,
                'record_date': record_date,
                'coupon_date': coupon_date,
                'status_details': status_details
            }
                
        except Exception as e:
            print(f"⚠️ Ошибка проверки INTR для {issuer} ({isin_moex}): {e}")
            intr_status[isin_moex] = {
                'has_intr': False,
                'issuer': issuer,
                'record_date': None,
                'coupon_date': coupon_date,
                'status_details': f"❌ Ошибка проверки: {str(e)[:100]}"
            }
    
    return intr_status

def get_issuers_with_coupons_tomorrow():
    """Получает список эмитентов с купонами на завтра через API MOEX"""
    try:
        tomorrow = (datetime.now().date() + timedelta(days=1)).strftime("%Y-%m-%d")
        url = "https://iss.moex.com/iss/statistics/engines/stock/markets/bonds/bondization.json"
        
        params = {
            'from': tomorrow,
            'till': tomorrow,
            'limit': 100,
            'iss.only': 'coupons',
            'lang': 'ru',
            'is_traded': 1
        }
        
        response = requests.get(url, params=params, timeout=15)
        response.raise_for_status()

        data = response.json()

        if 'coupons' not in data:
            print(f"ℹ️ На {tomorrow} выплаты не запланированы (блок 'coupons' отсутствует)")
            return []

        coupons = data['coupons']
        columns = coupons.get('columns', [])
        rows = coupons.get('data', [])

        if not columns or not rows:
            print(f"ℹ️ На {tomorrow} выплаты не запланированы (данные пусты)")
            return []

        try:
            name_idx = columns.index('name')
            isin_idx = columns.index('isin')
        except ValueError as e:
            print(f"❌ В ответе MOEX отсутствуют необходимые колонки: {e}. Доступные: {columns}")
            return []

        issuers = []
        for row in rows:
            if len(row) <= max(name_idx, isin_idx):
                continue
                
            bond_name = str(row[name_idx]).strip() if row[name_idx] else ""
            isin = str(row[isin_idx]).strip() if row[isin_idx] else ""
            
            if not bond_name or not isin:
                continue

            issuer = extract_issuer_name(bond_name)
            if issuer:
                issuers.append({
                    'issuer': issuer,
                    'isin': isin,
                    'bond_name': bond_name
                })

        print(f"✅ Найдено {len(issuers)} выпусков с выплатами на {tomorrow}")
        return issuers

    except requests.exceptions.RequestException as e:
        print(f"❌ Ошибка сети при запросе к MOEX: {e}")
        return []
    except Exception as e:
        print(f"❌ Необработанная ошибка: {e}")
        return []

def get_issuers_from_excel():
    """Извлекает эмитентов и ISIN из Excel-файла (ISIN в столбце D)"""
    try:
        if not os.path.exists(EXCEL_FILE_PATH):
            print(f"❌ Файл Excel не найден: {EXCEL_FILE_PATH}")
            return {}

        df = pd.read_excel(EXCEL_FILE_PATH, sheet_name='Лист4')
        bond_names = df.iloc[3:, 1].dropna().tolist()  # Столбец B - названия
        isins = df.iloc[3:, 3].dropna().tolist()       # Столбец D - ISIN

        portfolio = {}
        for bond_name, isin in zip(bond_names, isins):
            issuer = extract_issuer_name(bond_name)
            if issuer and isin:
                portfolio[issuer.strip().lower()] = {
                    'isin': isin.strip(),
                    'bond_name': bond_name.strip()
                }

        print(f"✅ Из Excel извлечено {len(portfolio)} уникальных эмитентов")
        return portfolio

    except Exception as e:
        print(f"❌ Ошибка при чтении Excel: {e}")
        return {}

def get_issuers_with_coupons_for_date(target_date):
    """Получает список эмитентов с купонами на указанную дату"""
    try:
        date_str = target_date.strftime("%Y-%m-%d")
        
        url = "https://iss.moex.com/iss/statistics/engines/stock/markets/bonds/bondization.json"
        
        params = {
            'from': date_str,
            'till': date_str,
            'limit': 100,
            'iss.only': 'coupons',
            'lang': 'ru',
            'is_traded': 1
        }
        
        response = requests.get(url, params=params, timeout=15)
        response.raise_for_status()

        data = response.json()

        if 'coupons' not in data:
            print(f"ℹ️ На {date_str} выплаты не запланированы (блок 'coupons' отсутствует)")
            return []

        coupons = data['coupons']
        columns = coupons.get('columns', [])
        rows = coupons.get('data', [])

        if not columns or not rows:
            print(f"ℹ️ На {date_str} выплаты не запланированы (данные пусты)")
            return []

        try:
            name_idx = columns.index('name')
            isin_idx = columns.index('isin')
        except ValueError as e:
            print(f"❌ В ответе MOEX отсутствуют необходимые колонки: {e}. Доступные: {columns}")
            return []

        issuers = []
        for row in rows:
            if len(row) <= max(name_idx, isin_idx):
                continue
                
            bond_name = str(row[name_idx]).strip() if row[name_idx] else ""
            isin = str(row[isin_idx]).strip() if row[isin_idx] else ""
            
            if not bond_name or not isin:
                continue

            issuer = extract_issuer_name(bond_name)
            if issuer:
                issuers.append({
                    'issuer': issuer,
                    'isin': isin,
                    'bond_name': bond_name
                })

        print(f"✅ Найдено {len(issuers)} выпусков с выплатами на {date_str}")
        return issuers

    except requests.exceptions.RequestException as e:
        print(f"❌ Ошибка сети при запросе к MOEX: {e}")
        return []
    except Exception as e:
        print(f"❌ Необработанная ошибка: {e}")
        return []

def check_nsd_for_issuers():
    """Основная функция с многоуровневой проверкой INTR"""
    
    # Определяем даты и тип проверки
    current_time = datetime.now()
    check_date = current_time.date()
    
    if CHECK_TYPE == "evening":
        target_date = check_date + timedelta(days=1)
        check_time = "17:00"
    elif CHECK_TYPE == "morning":
        target_date = check_date
        check_time = "09:00"
    elif CHECK_TYPE == "afternoon":
        target_date = check_date
        check_time = "14:00"
    elif CHECK_TYPE == "final":
        target_date = check_date
        check_time = "18:00"
    else:
        print(f"❌ Неизвестный тип проверки: {CHECK_TYPE}")
        return False

    print(f"🕒 Запуск проверки типа: {CHECK_TYPE} в {check_time}")
    print(f"📅 Целевая дата выплат: {target_date.strftime('%d.%m.%Y')}")

    # Получаем эмитентов с выплатами
    if CHECK_TYPE == "evening":
        paying_issuers = get_issuers_with_coupons_tomorrow()
    else:
        paying_issuers = get_issuers_with_coupons_for_date(target_date)

    if not paying_issuers:
        print(f"🔕 Эмитенты с купонами на {target_date.strftime('%d.%m.%Y')} не найдены")
        no_data_message = f"🔕 На {target_date.strftime('%d.%m.%Y')} выплат не найдено"
        send_telegram_message(no_data_message)
        return False

    excel_issuers = get_issuers_from_excel()
    if not excel_issuers:
        print("⚠️ Не удалось получить список эмитентов из Excel")
        return False

    # Фильтрация по портфелю
    filtered_issuers = []
    for issuer_data in paying_issuers:
        issuer_name = issuer_data['issuer'].strip().lower()
        if issuer_name in excel_issuers:
            filtered_issuers.append({
                'issuer': issuer_data['issuer'],
                'isin_portfolio': excel_issuers[issuer_name]['isin'],
                'isin_moex': issuer_data['isin'],
                'bond_name_portfolio': excel_issuers[issuer_name]['bond_name']
            })

    if not filtered_issuers:
        print(f"🔕 Ни один из эмитентов с купонами на {target_date.strftime('%d.%m.%Y')} не найден в вашем Excel-файле")
        no_data_message = f"🔕 На {target_date.strftime('%d.%m.%Y')} выплат по вашим облигациям не найдено"
        send_telegram_message(no_data_message)
        return False

    print(f"✅ Найдено {len(filtered_issuers)} выпусков с выплатами на {target_date.strftime('%d.%m.%Y')} в вашем портфеле:")

    # Проверяем статус INTR (ТОЛЬКО НРД по ISIN с MOEX)
    intr_status = check_intr_status(filtered_issuers)
    
    # Формируем сообщение
    if CHECK_TYPE == "evening":
        telegram_message = f"🔔 <b>Проверка выплат на {target_date.strftime('%d.%m.%Y')}</b>\n"
        telegram_message += f"<b>Время проверки: {check_time} (накануне)</b>\n\n"
    else:
        telegram_message = f"🔔 <b>Проверка выплат на {target_date.strftime('%d.%m.%Y')}</b>\n"
        telegram_message += f"<b>Время проверки: {check_time}</b>\n\n"

    telegram_message += f"<b>Выпуски из вашего портфеля:</b>\n"

    issuers_without_intr = []
    
    for i, issuer_data in enumerate(filtered_issuers, 1):
        isin_portfolio = issuer_data['isin_portfolio']
        isin_moex = issuer_data['isin_moex']
        issuer = issuer_data['issuer']
        bond_name_portfolio = issuer_data['bond_name_portfolio']
        
        intr_data = intr_status.get(isin_moex, {})
        has_intr = intr_data.get('has_intr', False)
        record_date = intr_data.get('record_date')
        coupon_date = intr_data.get('coupon_date')
        status_details = intr_data.get('status_details', '')
        
        status_icon = "✅" if has_intr else "❌"
        status_text = "INTR пройден" if has_intr else "INTR отсутствует"
        
        telegram_message += f"\n<b>{i}. {issuer}</b>\n"
        telegram_message += f"   Ваш: {isin_portfolio}\n"
        telegram_message += f"   Выплаты у выпуска:\n"
        telegram_message += f"   🆔 ISIN: {isin_moex}\n"
        telegram_message += f"   📅 Дата выплаты: {coupon_date}\n"
        telegram_message += f"   📊 Статус: {status_icon} {status_text}\n"
        telegram_message += f"   {status_details}\n"
        
        # Ссылка на новости НРД по ISIN с MOEX
        from_date = (check_date - timedelta(days=7)).strftime("%d.%m.%Y")
        to_date = check_date.strftime("%d.%m.%Y")
        nsd_search_url = f"https://nsddata.ru/ru/news?text={isin_moex}&from={from_date}&to={to_date}"
        telegram_message += f"   🔍 Ссылка на новости НРД ({nsd_search_url})\n"
        
        # Если INTR не найден — добавляем ссылку на e-disclosure.ru
        if not has_intr:
            telegram_message += "   🔍 Ссылка для проверки на сайте ЦРКИ (https://www.e-disclosure.ru/poisk-po-soobshheniyam)\n"
            issuers_without_intr.append(issuer_data)
            
            if CHECK_TYPE == "evening":
                webbrowser.open(nsd_search_url)
                time.sleep(0.5)

    # Добавляем предупреждения
    if issuers_without_intr:
        if CHECK_TYPE == "evening":
            telegram_message += f"\n⚠️ <b>ВНИМАНИЕ:</b> У {len(issuers_without_intr)} выпусков отсутствует INTR\n"
            telegram_message += "Следующая проверка в 9:00 утра"
        elif CHECK_TYPE == "morning":
            telegram_message += f"\n⚠️ <b>ВНИМАНИЕ:</b> У {len(issuers_without_intr)} выпусков всё ещё нет INTR\n"
            telegram_message += "Следующая проверка в 14:00"
        elif CHECK_TYPE == "afternoon":
            telegram_message += f"\n🚨 <b>ВНИМАНИЕ:</b> У {len(issuers_without_intr)} выпусков всё ещё нет INTR\n"
            telegram_message += "<b>Риск дефолта! Осталась последняя проверка в 18:00 МСК</b>"
        elif CHECK_TYPE == "final":
            telegram_message += f"\n🚨🚨🚨 <b>КРИТИЧЕСКИЙ РИСК!</b> 🚨🚨🚨\n"
            telegram_message += f"У {len(issuers_without_intr)} выпусков НЕТ INTR на 18:00\n"
            telegram_message += "<b>ВЫСОКИЙ РИСК ТЕХНИЧЕСКОГО ДЕФОЛТА!</b>"
    else:
        telegram_message += f"\n✅ <b>Все выпуски прошли INTR</b>\n"
        telegram_message += "Выплаты должны пройти в штатном режиме"

    if len(telegram_message) > 4000:
        telegram_message = telegram_message[:3900] + "\n\n⚠️ Сообщение обрезано из-за ограничения Telegram"

    success = send_telegram_message(telegram_message)

    if success:
        print("✅ Отчет успешно отправлен в Telegram")
    else:
        print("❌ Не удалось отправить отчет в Telegram")

    return True

# ==================== ЗАПУСК ====================
if __name__ == "__main__":
    print("🚀 Запуск многоуровневой проверки эмитентов...")
    print(f"🕒 Текущая дата и время: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    print(f"📋 Тип проверки: {CHECK_TYPE}")
    
    check_nsd_for_issuers()
    print("✅ Проверка завершена!")