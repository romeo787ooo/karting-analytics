import streamlit as st
import pandas as pd
import re
import calendar
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import numpy as np
from plotly.subplots import make_subplots

st.set_page_config(page_title="Аналитика картинг-центра", layout="wide")

def analyze_karting_data_sheet(df_raw, sheet_name):
    """
    Анализирует данные картинг-центра из одного листа Excel-файла.
    
    Args:
        df_raw: DataFrame с данными листа Excel.
        sheet_name (str): Название листа для информационных целей.

    Returns:
        list: Список обработанных данных.
    """
    # Определяем списки дней недели (и в верхнем, и в нижнем регистре)
    russian_day_names_display = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье"]
    russian_day_names_lower = [day.lower() for day in russian_day_names_display]
    
    # Получаем месяц и год из названия листа
    month_year_match = re.search(r'(\w+)\s+(\d{4})', sheet_name, re.IGNORECASE)
    if month_year_match:
        month_name, year = month_year_match.groups()
        month_names = {
            'январь': 1, 'февраль': 2, 'март': 3, 'апрель': 4, 'май': 5, 'июнь': 6,
            'июль': 7, 'август': 8, 'сентябрь': 9, 'октябрь': 10, 'ноябрь': 11, 'декабрь': 12
        }
        # Приводим название месяца к нижнему регистру для поиска
        month_num = month_names.get(month_name.lower(), 5)  # По умолчанию май, если не распознано
        year = int(year)
    else:
        # Если не удалось распознать, используем текущий месяц и год
        current_date = datetime.now()
        month_num = current_date.month
        year = current_date.year
    
    # Определяем индексы строк, где начинаются дни недели
    day_start_indices = []
    day_names_found = []  # Для хранения найденных названий дней недели

    for i, row in df_raw.iterrows():
        if i < len(df_raw) - 1:  # Проверяем, что не выходим за пределы DataFrame
            # Проверяем, содержит ли ячейка в столбце C (индекс 2) название дня недели
            cell_value = str(row.iloc[2]).strip().lower() if not pd.isna(row.iloc[2]) else ""
            for idx, day_lower in enumerate(russian_day_names_lower):
                if day_lower in cell_value:
                    day_start_indices.append(i)
                    day_names_found.append(russian_day_names_display[idx])  # Сохраняем корректное название
                    break
    
    # Список для хранения обработанных данных
    processed_data = []
    
    # Перебираем каждый блок дня недели
    for i, start_idx in enumerate(day_start_indices):
        # Определяем конец блока данных для текущего дня
        end_idx = day_start_indices[i + 1] if i < len(day_start_indices) - 1 else len(df_raw)
        
        # Получаем название дня недели из сохраненного списка
        if i < len(day_names_found):
            day_of_week = day_names_found[i]
        else:
            continue  # Пропускаем, если индекс выходит за границы
        
        # Находим строку заголовка с "Количество машин в заезде"
        header_idx = None
        for j in range(start_idx, min(start_idx + 5, end_idx)):  # Ищем в первых 5 строках блока
            if isinstance(df_raw.iloc[j, 2], str) and "количество машин в заезде" in df_raw.iloc[j, 2].lower():
                header_idx = j
                break
        
        if header_idx is None:
            continue  # Пропускаем, если не нашли заголовок
        
        # Получаем числа из первого столбца (календарные дни)
        calendar_day = None
        for j in range(start_idx, min(start_idx + 3, end_idx)):
            try:
                if isinstance(df_raw.iloc[j, 0], (int, float)) and 1 <= float(df_raw.iloc[j, 0]) <= 31:
                    calendar_day = int(df_raw.iloc[j, 0])
                    break
            except (ValueError, TypeError):
                continue
        
        # Устанавливаем индекс первой строки с данными (после заголовков)
        data_start_idx = header_idx + 1
        
        # Создаем полную дату для записи
        try:
            full_date = datetime(year, month_num, calendar_day if calendar_day else 1)
            date_str = full_date.strftime('%Y-%m-%d')
        except (ValueError, TypeError):
            # Если не удалось создать дату, используем просто месяц и год
            date_str = f"{year}-{month_num:02d}-01"
        
        # Перебираем строки с данными
        for data_idx in range(data_start_idx, end_idx):
            # Получаем время заезда из столбца B (индекс 1)
            time_slot = str(df_raw.iloc[data_idx, 1]) if not pd.isna(df_raw.iloc[data_idx, 1]) else ""
            
            # Извлекаем час из времени
            hour = None
            hour_match = re.search(r'(\d{1,2}):', time_slot)
            if hour_match:
                hour = int(hour_match.group(1))
            else:
                continue  # Пропускаем строку без времени
            
            # Подсчитываем количество и типы машин в заезде
            machine_count = 0
            machine_types = {'прокат': 0, 'мини': 0, 'серт': 0, 'твин': 0, 'клуб': 0, 
                            'электро': 0, 'тест': 0, 'бронь': 0, 'wowlife': 0, 'другой': 0}
            
            # Перебираем столбцы с C по N (индексы от 2 до 13), которые содержат данные о машинах
            for col_idx in range(2, 14):  # C до N
                if col_idx < len(df_raw.columns):
                    cell_value = str(df_raw.iloc[data_idx, col_idx]).lower().strip() if not pd.isna(df_raw.iloc[data_idx, col_idx]) else ""
                    if cell_value:  # Если ячейка не пустая
                        machine_count += 1
                        
                        # Определяем тип машины
                        cell_value = re.sub(r'\.+$', '', cell_value)  # удаляет точки в конце строки
                        machine_type = 'другой'  # По умолчанию
                        
                        if 'прокат' in cell_value:
                            machine_type = 'прокат'
                        elif 'мини' in cell_value:
                            machine_type = 'мини'
                        elif 'серт' in cell_value or 'cert' in cell_value:
                            machine_type = 'серт'
                        elif 'твин' in cell_value or 'twin' in cell_value:
                            machine_type = 'твин'
                        elif 'клуб' in cell_value:
                            machine_type = 'клуб'
                        elif 'электро' in cell_value:
                            machine_type = 'электро'
                        elif 'тест' in cell_value:
                            machine_type = 'тест'
                        elif 'бронь' in cell_value:
                            machine_type = 'бронь'
                        elif 'wowlife' in cell_value:
                            machine_type = 'wowlife'
                        
                        machine_types[machine_type] += 1
            
            # Добавляем данные о заезде в список
            if machine_count > 0:
                processed_data.append({
                    'День недели': day_of_week,
                    'Календарный день': calendar_day if calendar_day else 1,
                    'Месяц': int(month_num),  # Убедимся, что месяц - целое число
                    'Год': int(year),         # Убедимся, что год - целое число
                    'Дата': date_str,
                    'Час': hour,
                    'Количество машин': machine_count,
                    'Прокат': machine_types['прокат'],
                    'Мини': machine_types['мини'],
                    'Серт': machine_types['серт'],
                    'Твин': machine_types['твин'],
                    'Клуб': machine_types['клуб'],
                    'Электро': machine_types['электро'],
                    'Тест': machine_types['тест'],
                    'Бронь': machine_types['бронь'],
                    'WowLife': machine_types['wowlife'],
                    'Другой': machine_types['другой']
                })
    
    return processed_data


def analyze_karting_data(uploaded_file, sheet_name=None, all_sheets=False):
    """
    Анализирует данные картинг-центра из Excel-файла.
    
    Args:
        uploaded_file: Загруженный файл Excel.
        sheet_name (str, optional): Название конкретного листа для анализа.
        all_sheets (bool): Если True, анализируются все листы.

    Returns:
        dict: Словарь с данными для построения графиков.
    """
    # Определяем список дней недели
    russian_day_names_display = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье"]
    
    try:
        # Получаем список всех листов в Excel
        xl = pd.ExcelFile(uploaded_file)
        all_sheet_names = xl.sheet_names
        
        # Определяем, какие листы анализировать
        sheets_to_analyze = all_sheet_names if all_sheets else [sheet_name]
        
        # Объединенный список обработанных данных
        all_processed_data = []
        
        # Анализируем каждый лист
        for current_sheet in sheets_to_analyze:
            try:
                # Загружаем данные листа
                df_raw = pd.read_excel(uploaded_file, sheet_name=current_sheet, header=None)
                
                # Анализируем данные листа
                sheet_data = analyze_karting_data_sheet(df_raw, current_sheet)
                
                # Добавляем название листа для отслеживания
                for item in sheet_data:
                    item['Лист'] = current_sheet
                
                # Добавляем в общий список
                all_processed_data.extend(sheet_data)
                
            except Exception as e:
                st.warning(f"Не удалось обработать лист '{current_sheet}': {str(e)}")
                continue
        
        # Если нет данных, возвращаем None
        if not all_processed_data:
            st.error("Не удалось извлечь данные. Проверьте формат файла.")
            return None
        
        # Создаем DataFrame из собранных данных
        data_df = pd.DataFrame(all_processed_data)
        
        # Подготовка данных для графиков
        
        # 1. Данные для стекового бара по дням недели и типам машин
        stacked_data = []
        for machine_type in ['Прокат', 'Мини', 'Серт', 'Твин', 'Клуб', 'Электро', 'Тест', 'Бронь', 'WowLife', 'Другой']:
            for day in russian_day_names_display:
                day_data = data_df[data_df['День недели'] == day]
                if not day_data.empty:
                    # Используем регистронезависимый подсчет (приводим всё к нижнему регистру)
                    machine_type_lower = machine_type.lower()
                    count = 0
                    if machine_type_lower == 'прокат':
                        count = day_data['Прокат'].sum()
                    elif machine_type_lower == 'мини':
                        count = day_data['Мини'].sum()
                    elif machine_type_lower == 'серт':
                        count = day_data['Серт'].sum()
                    elif machine_type_lower == 'твин':
                        count = day_data['Твин'].sum()
                    elif machine_type_lower == 'клуб':
                        count = day_data['Клуб'].sum()
                    elif machine_type_lower == 'электро':
                        count = day_data['Электро'].sum()
                    elif machine_type_lower == 'тест':
                        count = day_data['Тест'].sum()
                    elif machine_type_lower == 'бронь':
                        count = day_data['Бронь'].sum()
                    elif machine_type_lower == 'wowlife':
                        count = day_data['WowLife'].sum()
                    elif machine_type_lower == 'другой':
                        count = day_data['Другой'].sum()
                    
                    if count > 0:
                        stacked_data.append({
                            'День недели': day,
                            'Тип машины': machine_type,
                            'Количество': count
                        })
        
        stacked_df = pd.DataFrame(stacked_data)
        
        # 2. Данные для круговой диаграммы типов машин
        pie_data = []
        for machine_type in ['Прокат', 'Мини', 'Серт', 'Твин', 'Клуб', 'Электро', 'Тест', 'Бронь', 'WowLife', 'Другой']:
            machine_type_lower = machine_type.lower()
            count = 0
            if machine_type_lower == 'прокат':
                count = data_df['Прокат'].sum()
            elif machine_type_lower == 'мини':
                count = data_df['Мини'].sum()
            elif machine_type_lower == 'серт':
                count = data_df['Серт'].sum()
            elif machine_type_lower == 'твин':
                count = data_df['Твин'].sum()
            elif machine_type_lower == 'клуб':
                count = data_df['Клуб'].sum()
            elif machine_type_lower == 'электро':
                count = data_df['Электро'].sum()
            elif machine_type_lower == 'тест':
                count = data_df['Тест'].sum()
            elif machine_type_lower == 'бронь':
                count = data_df['Бронь'].sum()
            elif machine_type_lower == 'wowlife':
                count = data_df['WowLife'].sum()
            elif machine_type_lower == 'другой':
                count = data_df['Другой'].sum()
            
            if count > 0:
                pie_data.append({
                    'Тип машины': machine_type,
                    'Количество': count
                })
        
        pie_df = pd.DataFrame(pie_data)
        
        # 3. Анализ загруженности по часам - среднее значение
        # Группируем данные по дню недели и часу, вычисляем среднее количество машин
        avg_hourly_load = data_df.groupby(['День недели', 'Час'])['Количество машин'].mean().reset_index()
        
        # Создаем матрицу загруженности по часам и дням недели
        day_hour_pivot = avg_hourly_load.pivot(index='День недели', columns='Час', values='Количество машин')
        
        # Обеспечиваем правильный порядок дней недели
        day_hour_pivot = day_hour_pivot.reindex(russian_day_names_display)
        
        # 4. Находим наименее загруженные часы для каждого дня недели
        least_busy_hours = {}
        for day in russian_day_names_display:
            if day in day_hour_pivot.index:
                day_data = day_hour_pivot.loc[day].dropna()
                if not day_data.empty:
                    min_load = day_data.min()
                    least_busy_hour = day_data[day_data == min_load].index.tolist()
                    least_busy_hours[day] = {
                        'часы': least_busy_hour,
                        'загруженность': min_load
                    }
        
        # 5. Находим средние значения загруженности по дням недели
        day_avg_load = data_df.groupby('День недели')['Количество машин'].mean()
        day_avg_load = day_avg_load.reindex(russian_day_names_display)
        least_busy_days = day_avg_load.sort_values().head(3)
        
        # 6. Находим средние значения загруженности по часам (для всех дней)
        hour_avg_load = data_df.groupby('Час')['Количество машин'].mean()
        least_busy_hours_overall = hour_avg_load.sort_values().head(5)
        
        # 7. Данные для столбчатой диаграммы по дням недели
        daily_total = data_df.groupby('День недели')['Количество машин'].sum().reindex(russian_day_names_display)
        daily_df = pd.DataFrame({
            'День недели': daily_total.index,
            'Количество машин': daily_total.values
        })
        
        # Подготовка данных для графиков загруженности по часам для каждого дня недели
        days_hourly_data = {}
        for day in russian_day_names_display:
            day_data = avg_hourly_load[avg_hourly_load['День недели'] == day]
            if not day_data.empty:
                days_hourly_data[day] = day_data
        
        # 8. Если анализируем все листы, добавляем тренды по месяцам
        if all_sheets:
            # Добавляем данные по месяцам
            data_df['Месяц-Год'] = data_df.apply(lambda row: f"{int(row['Месяц']):02d}-{int(row['Год'])}", axis=1)
            
            monthly_data = data_df.groupby(['Год', 'Месяц']).agg({
                'Количество машин': ['sum', 'mean', 'count'],
                'Прокат': 'sum',
                'Мини': 'sum',
                'Серт': 'sum',
                'Твин': 'sum',
                'Клуб': 'sum',
                'Электро': 'sum',
                'Тест': 'sum',
                'Бронь': 'sum',
                'WowLife': 'sum',
                'Другой': 'sum'
            }).reset_index()
            
            # Преобразуем MultiIndex в обычные столбцы
            monthly_data.columns = ['Год', 'Месяц', 'Всего машин', 'Среднее кол-во машин', 'Количество заездов', 
                                   'Прокат', 'Мини', 'Серт', 'Твин', 'Клуб', 'Электро', 
                                   'Тест', 'Бронь', 'WowLife', 'Другой']
            
            # Создаем строковое представление месяца для сортировки и отображения
            monthly_data['Месяц-Год'] = monthly_data.apply(lambda row: f"{int(row['Месяц']):02d}-{int(row['Год'])}", axis=1)
            monthly_data = monthly_data.sort_values(['Год', 'Месяц'])
            
            # Создаем данные для сравнения месяцев по типам машин
            monthly_machine_types = []
            for _, row in monthly_data.iterrows():
                for machine_type in ['Прокат', 'Мини', 'Серт', 'Твин', 'Клуб', 'Электро', 'Тест', 'Бронь', 'WowLife', 'Другой']:
                    if row[machine_type] > 0:
                        monthly_machine_types.append({
                            'Месяц-Год': row['Месяц-Год'],
                            'Тип машины': machine_type,
                            'Количество': row[machine_type]
                        })
            
            monthly_machine_types_df = pd.DataFrame(monthly_machine_types)
        else:
            monthly_data = None
            monthly_machine_types_df = None
        
        # Возвращаем обработанные данные
        return {
            'raw_data': data_df,
            'stacked_df': stacked_df,
            'pie_df': pie_df,
            'avg_hourly_load': avg_hourly_load,
            'day_hour_pivot': day_hour_pivot,
            'least_busy_hours': least_busy_hours,
            'least_busy_days': least_busy_days,
            'least_busy_hours_overall': least_busy_hours_overall,
            'daily_df': daily_df,
            'days_hourly_data': days_hourly_data,
            'monthly_data': monthly_data,
            'monthly_machine_types_df': monthly_machine_types_df,
            'russian_day_names_display': russian_day_names_display,
            'all_sheets': all_sheets,
            'analyzed_sheets': sheets_to_analyze
        }
        
    except Exception as e:
        st.error(f"Произошла ошибка при анализе данных: {str(e)}")
        return None


# Основной код приложения Streamlit
st.title("Аналитика картинг-центра")

# Загрузка файла
uploaded_file = st.file_uploader("Выберите Excel файл с данными", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Получаем список всех листов в Excel-файле
    xl = pd.ExcelFile(uploaded_file)
    sheet_names = xl.sheet_names
    
    # Выбор режима анализа
    analysis_mode = st.radio(
        "Выберите режим анализа",
        ["Отдельный месяц", "Все месяцы (суммарно)"]
    )
    
    # В зависимости от выбранного режима
    if analysis_mode == "Отдельный месяц":
        # Выбор листа
        selected_sheet = st.selectbox("Выберите лист (месяц)", sheet_names)
        all_sheets = False
    else:
        selected_sheet = None
        all_sheets = True
        st.info("Будет выполнен анализ всех листов в файле")
    
    # Анализируем данные
    with st.spinner('Анализируем данные...'):
        results = analyze_karting_data(uploaded_file, selected_sheet, all_sheets)
    
    if results:
        if all_sheets:
            st.success(f'Анализ данных для всех листов выполнен успешно! Обработано {len(results["analyzed_sheets"])} листов.')
        else:
            st.success(f'Анализ данных для листа "{selected_sheet}" выполнен успешно!')
        
        # Сохраняем список дней недели для использования в графиках
        russian_day_names_display = results['russian_day_names_display']
        
        # Создаем вкладки для разных графиков
        if all_sheets:
            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                "Распределение по дням недели", 
                "Типы машин", 
                "Загруженность по часам", 
                "Тепловая карта загруженности", 
                "Общее количество",
                "Анализ по месяцам"
            ])
        else:
            tab1, tab2, tab3, tab4, tab5 = st.tabs([
                "Распределение по дням недели", 
                "Типы машин", 
                "Загруженность по часам", 
                "Тепловая карта загруженности", 
                "Общее количество"
            ])
        
        with tab1:
            st.subheader("Распределение типов машин по дням недели")
            if not results['stacked_df'].empty:
                # Используем Plotly для создания стекового бар-чарта
                fig = px.bar(
                    results['stacked_df'],
                    x='День недели',
                    y='Количество',
                    color='Тип машины',
                    barmode='stack',
                    title='Распределение типов машин по дням недели',
                    category_orders={"День недели": russian_day_names_display}
                )
                fig.update_layout(
                    xaxis_title='День недели',
                    yaxis_title='Количество машин',
                    legend_title='Тип машины'
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Нет данных для отображения")
            
        with tab2:
            st.subheader("Соотношение типов машин")
            if not results['pie_df'].empty:
                # Используем Plotly для круговой диаграммы
                fig = px.pie(
                    results['pie_df'], 
                    values='Количество', 
                    names='Тип машины',
                    title='Соотношение типов машин',
                    hover_data=['Количество'],
                    labels={'Количество': 'Количество машин'}
                )
                fig.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Нет данных для отображения")
            
        with tab3:
            st.subheader("Средняя загруженность в течение дня по дням недели")
            if results['days_hourly_data']:
                # Создаем график для каждого дня недели - один под другим
                for day in russian_day_names_display:
                    if day in results['days_hourly_data']:
                        day_data = results['days_hourly_data'][day]
                        
                        # Используем Plotly для создания столбчатой диаграммы
                        fig = px.bar(
                            day_data, 
                            x='Час', 
                            y='Количество машин',
                            title=f'Средняя загруженность в течение дня ({day})',
                            text_auto='.1f'  # Показываем значения на столбцах с 1 знаком после запятой
                        )
                        fig.update_traces(textposition='outside')
                        fig.update_layout(
                            xaxis_title='Час',
                            yaxis_title='Среднее количество машин',
                            xaxis=dict(tickmode='linear', tick0=0, dtick=1),
                            height=400,  # Уменьшаем высоту для лучшего отображения нескольких графиков
                            margin=dict(l=10, r=10, t=50, b=50)  # Уменьшаем отступы
                        )
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.warning(f"Нет данных для {day}")
            else:
                st.warning("Нет данных о загруженности")
        
        with tab4:
            st.subheader("Тепловая карта загруженности по дням недели и часам")
            if results['day_hour_pivot'] is not None and not results['day_hour_pivot'].empty:
                # Создаем тепловую карту загруженности
                # Подготавливаем данные для тепловой карты
                pivot_df = results['day_hour_pivot'].copy()
                
                # Убедимся, что порядок дней недели правильный
                pivot_df = pivot_df.reindex(russian_day_names_display)
                
                # Создаем тепловую карту с Plotly
                fig = go.Figure(data=go.Heatmap(
                    z=pivot_df.values,
                    x=pivot_df.columns,
                    y=pivot_df.index,
                    colorscale='Blues',
                    hoverongaps=False,
                    text=np.round(pivot_df.values, 1),
                    hovertemplate='День: %{y}<br>Час: %{x}<br>Среднее кол-во машин: %{z:.1f}<extra></extra>',
                ))
                
                fig.update_layout(
                    title='Средняя загруженность по дням недели и часам',
                    xaxis_title='Час',
                    yaxis_title='День недели',
                    height=600
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
                # Суммарная информация о наименее загруженных днях и часах
                st.subheader("Наименее загруженные дни и часы")
                
                # Наименее загруженные дни
                st.markdown("### Наименее загруженные дни недели:")
                least_busy_days_df = pd.DataFrame({
                    'День недели': results['least_busy_days'].index,
                    'Среднее количество машин': results['least_busy_days'].values
                })
                st.dataframe(least_busy_days_df.style.format({'Среднее количество машин': '{:.1f}'}))
                
                # Наименее загруженные часы (в среднем по всем дням)
                st.markdown("### Наименее загруженные часы (в среднем по всем дням):")
                least_busy_hours_df = pd.DataFrame({
                    'Час': results['least_busy_hours_overall'].index,
                    'Среднее количество машин': results['least_busy_hours_overall'].values
                })
                st.dataframe(least_busy_hours_df.style.format({'Среднее количество машин': '{:.1f}'}))
                
                # Наименее загруженные часы для каждого дня недели
                st.markdown("### Наименее загруженные часы по дням недели:")
                for day, data in results['least_busy_hours'].items():
                    hours_str = ", ".join([str(h) for h in data['часы']])
                    st.markdown(f"**{day}**: часы {hours_str} — в среднем {data['загруженность']:.1f} машин")
                
            else:
                st.warning("Недостаточно данных для построения тепловой карты")
                
        with tab5:
            st.subheader("Общее количество машин по дням недели")
            if not results['daily_df'].empty:
                # Используем Plotly для столбчатой диаграммы
                fig = px.bar(
                    results['daily_df'], 
                    x='День недели', 
                    y='Количество машин',
                    title='Общее количество машин по дням недели',
                    text='Количество машин',  # Показываем значения на столбцах
                    category_orders={"День недели": russian_day_names_display}
                )
                fig.update_traces(texttemplate='%{text}', textposition='outside')
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Нет данных для отображения")
        
        # Добавляем вкладку с анализом по месяцам (если анализируем все листы)
        if all_sheets and 'tab6' in locals():
            with tab6:
                st.subheader("Анализ данных по месяцам")
                
                if results['monthly_data'] is not None and not results['monthly_data'].empty:
                    # Показываем таблицу с данными по месяцам
                    st.markdown("### Сводные данные по месяцам")
                    monthly_display = results['monthly_data'][['Месяц-Год', 'Всего машин', 'Среднее кол-во машин', 'Количество заездов']]
                    st.dataframe(monthly_display.style.format({
                        'Среднее кол-во машин': '{:.1f}'
                    }))
                    
                    # График количества машин по месяцам
                    st.markdown("### Динамика количества машин по месяцам")
                    fig = px.line(
                        results['monthly_data'], 
                        x='Месяц-Год', 
                        y='Всего машин',
                        markers=True,
                        title='Общее количество машин по месяцам',
                    )
                    fig.update_layout(
                        xaxis_title='Месяц-Год',
                        yaxis_title='Количество машин',
                    )
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # График среднего количества машин
                    st.markdown("### Динамика среднего количества машин в заездах")
                    fig = px.line(
                        results['monthly_data'], 
                        x='Месяц-Год', 
                        y='Среднее кол-во машин',
                        markers=True,
                        title='Среднее количество машин в заезде по месяцам',
                    )
                    fig.update_layout(
                        xaxis_title='Месяц-Год',
                        yaxis_title='Среднее количество машин',
                    )
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # График количества заездов
                    st.markdown("### Динамика количества заездов")
                    fig = px.line(
                        results['monthly_data'], 
                        x='Месяц-Год', 
                        y='Количество заездов',
                        markers=True,
                        title='Количество заездов по месяцам',
                    )
                    fig.update_layout(
                        xaxis_title='Месяц-Год',
                        yaxis_title='Количество заездов',
                    )
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # График распределения типов машин по месяцам
                    if results['monthly_machine_types_df'] is not None and not results['monthly_machine_types_df'].empty:
                        st.markdown("### Распределение типов машин по месяцам")
                        fig = px.bar(
                            results['monthly_machine_types_df'],
                            x='Месяц-Год',
                            y='Количество',
                            color='Тип машины',
                            barmode='stack',
                            title='Распределение типов машин по месяцам'
                        )
                        fig.update_layout(
                            xaxis_title='Месяц-Год',
                            yaxis_title='Количество машин',
                            legend_title='Тип машины'
                        )
                        st.plotly_chart(fig, use_container_width=True)
                        
                        # Тренды по типам машин
                        st.markdown("### Тренды по типам машин")
                        machine_types_list = results['monthly_machine_types_df']['Тип машины'].unique()
                        selected_types = st.multiselect(
                            "Выберите типы машин для отображения",
                            options=machine_types_list,
                            default=machine_types_list[:3] if len(machine_types_list) >= 3 else machine_types_list  # По умолчанию показываем первые три типа
                        )
                        
                        if selected_types:
                            # Фильтруем данные по выбранным типам
                            filtered_data = results['monthly_machine_types_df'][
                                results['monthly_machine_types_df']['Тип машины'].isin(selected_types)
                            ]
                            
                            # Создаем линейный график для выбранных типов
                            fig = px.line(
                                filtered_data,
                                x='Месяц-Год',
                                y='Количество',
                                color='Тип машины',
                                markers=True,
                                title='Тренды использования типов машин по месяцам'
                            )
                            fig.update_layout(
                                xaxis_title='Месяц-Год',
                                yaxis_title='Количество машин',
                                legend_title='Тип машины'
                            )
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.warning("Выберите хотя бы один тип машин для отображения")
                else:
                    st.warning("Недостаточно данных для анализа по месяцам")
        
        # Добавляем таблицу с сырыми данными
        with st.expander("Просмотр обработанных данных"):
            st.dataframe(results['raw_data'])
            
    else:
        st.error("Не удалось проанализировать данные. Пожалуйста, проверьте формат файла.")
else:
    st.info("Пожалуйста, загрузите Excel файл для анализа.")

# Добавляем информацию о приложении
with st.sidebar:
    st.title("О приложении")
    st.info("""
    Это приложение анализирует данные картинг-центра и предоставляет визуализацию.
    - Загрузите Excel файл с данными
    - Выберите режим анализа (отдельный месяц или все месяцы)
    - Просмотрите отчеты на разных вкладках
    """)

    st.header("Инструкция")
    st.markdown("""
    1. **Загрузите Excel-файл** с данными картинг-центра
    2. **Выберите режим анализа**:
       - Отдельный месяц - для анализа конкретного листа
       - Все месяцы - для суммарного анализа всех листов
    3. **Переключайтесь между вкладками** для просмотра различных отчетов
    4. На вкладке "Загруженность по часам" представлены графики для всех дней недели
    5. Вкладка "Тепловая карта загруженности" показывает все дни и часы на одном графике
    6. При анализе всех месяцев доступна дополнительная вкладка "Анализ по месяцам"
    7. Используйте опцию "Просмотр обработанных данных" для доступа к табличным данным
    """)