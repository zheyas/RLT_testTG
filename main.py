import subprocess
import sys
import pandas as pd
import numpy as np

# Установка openpyxl при необходимости
try:
    import openpyxl
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl

# -------------------------------------------------------------------
# 1. Загрузка данных из Excel-файла
# -------------------------------------------------------------------
input_file = "corporate_links_raw.xlsx"

# Читаем Excel, предполагая что все данные в первом столбце, разделённые запятыми
df_excel = pd.read_excel(input_file, header=None, engine='openpyxl')

# Первая строка содержит заголовки
header_row = df_excel.iloc[0, 0]
headers = [h.strip() for h in header_row.split(',')]

# Данные со второй строки
data_rows = []
for i in range(1, len(df_excel)):
    cell = df_excel.iloc[i, 0]
    if pd.notna(cell):
        row = [x.strip() for x in str(cell).split(',')]
        while len(row) < len(headers):
            row.append('')
        data_rows.append(row)

df = pd.DataFrame(data_rows, columns=headers)

# Переименуем колонки для удобства
df.rename(columns={
    'FIO_owner': 'fio',
    'Company_name': 'company',
    'INN_company': 'inn',
    'Ownership': 'ownership',
    'Region': 'region',
    'Source': 'source',
    'Ownership_date': 'date'
}, inplace=True)

# -------------------------------------------------------------------
# 2. Нормализация и очистка
# -------------------------------------------------------------------
def normalize_fio(name):
    if pd.isna(name) or name == '':
        return name
    name = str(name).strip()
    parts = name.split()
    if len(parts) == 0:
        return name
    last = parts[0]
    rest = ' '.join(parts[1:])
    initials = []
    if '.' in rest:
        for part in rest.split('.'):
            part = part.strip()
            if part and part[0].isalpha():
                initials.append(part[0].upper())
    else:
        for word in rest.split():
            if word and word[0].isalpha():
                initials.append(word[0].upper())
    while len(initials) < 2:
        initials.append('')
    result = last
    if initials[0]:
        result += f" {initials[0]}."
    if initials[1]:
        result += f"{initials[1]}."
    return result

df['fio_norm'] = df['fio'].apply(normalize_fio)

# Названия компаний: убираем кавычки, лишние пробелы
df['company_norm'] = df['company'].str.replace('"', '', regex=False).str.strip()

# ИНН как строка, пропуски -> NaN
df['inn_str'] = df['inn'].astype(str).str.strip()
df.loc[df['inn_str'].isin(['nan', '', 'None']), 'inn_str'] = np.nan

# Доля владения -> проценты (float)
def parse_ownership(val):
    if pd.isna(val) or str(val).strip() == '':
        return np.nan
    s = str(val).strip().replace('%', '').replace(' ', '').replace(',', '.')
    if s == '':
        return np.nan
    try:
        num = float(s)
        return num * 100 if num <= 1 else num
    except:
        return np.nan

df['ownership_pct'] = df['ownership'].apply(parse_ownership)

# Дата -> единый формат YYYY-MM-DD
def parse_date(val):
    if pd.isna(val) or str(val).strip() == '':
        return np.nan
    s = str(val).strip()
    for fmt in ['%d.%m.%Y', '%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y/%m/%d', '%Y.%m.%d']:
        try:
            return pd.to_datetime(s, format=fmt)
        except:
            continue
    try:
        return pd.to_datetime(s, dayfirst=True)
    except:
        return np.nan

df['date_norm'] = df['date'].apply(parse_date)
df['date_str'] = df['date_norm'].dt.strftime('%Y-%m-%d')

# Удаляем явные дубликаты (одинаковые ФИО + компания + дата + доля)
df.drop_duplicates(subset=['fio_norm', 'company_norm', 'date_str', 'ownership_pct'], inplace=True)

# Финальная чистая таблица
clean_df = df[['fio_norm', 'company_norm', 'inn_str', 'ownership_pct', 'region', 'source', 'date_str']].copy()
clean_df.rename(columns={
    'fio_norm': 'ФИО',
    'company_norm': 'Компания',
    'inn_str': 'ИНН',
    'ownership_pct': 'Доля_%',
    'region': 'Регион',
    'source': 'Источник',
    'date_str': 'Дата'
}, inplace=True)

# -------------------------------------------------------------------
# 3. Анализ данных
# -------------------------------------------------------------------
# 3.1. Записи с отсутствующим ИНН
missing_inn = clean_df[clean_df['ИНН'].isna()]

# 3.2. Компании с суммарной долей > 100% на конкретную дату
daily_sum = clean_df.groupby(['Компания', 'Дата'], as_index=False)['Доля_%'].sum()
over100 = daily_sum[daily_sum['Доля_%'] > 100].copy()
over100.rename(columns={'Доля_%': 'Суммарная_доля_%'}, inplace=True)

# 3.3. Анализ изменений в компаниях (более одной даты)
companies_with_changes = []
company_dates = clean_df.groupby('Компания')['Дата'].unique()
for company, dates in company_dates.items():
    dates = sorted([d for d in dates if pd.notna(d)])
    if len(dates) < 2:
        continue
    # Собираем снимки на каждую дату
    snapshots = []
    for d in dates:
        subset = clean_df[(clean_df['Компания'] == company) & (clean_df['Дата'] == d)]
        owners = []
        for _, row in subset.iterrows():
            owners.append(f"{row['ФИО']} {row['Доля_%']:.1f}%")
        snapshots.append({
            'date': d,
            'owners': owners,
            'total': subset['Доля_%'].sum()
        })
    # Формируем описание изменений
    description = []
    for i in range(len(snapshots)-1):
        d1 = snapshots[i]
        d2 = snapshots[i+1]
        set1 = set(d1['owners'])
        set2 = set(d2['owners'])
        # Кто вошёл
        entered = set2 - set1
        for o in entered:
            description.append(f" {o} (вошёл к {d2['date']})")
        # Кто вышел
        exited = set1 - set2
        for o in exited:
            description.append(f" {o} (вышел после {d1['date']})")
        # Изменение долей у тех, кто остался
        remained = set1 & set2
        for o in remained:
            # Найти долю в первом и втором периоде
            share1 = next(s for s in d1['owners'] if s.startswith(o.split()[0]))  # по фамилии
            share2 = next(s for s in d2['owners'] if s.startswith(o.split()[0]))
            if share1 != share2:
                description.append(f"✏"
                                   f"️ {share1} → {share2} (изменение доли)")
    if description:
        companies_with_changes.append({
            'Компания': company,
            'Первая дата': snapshots[0]['date'],
            'Состав на первую дату': ', '.join(snapshots[0]['owners']),
            'Последняя дата': snapshots[-1]['date'],
            'Состав на последнюю дату': ', '.join(snapshots[-1]['owners']),
            'Характер изменений': '; '.join(description)
        })

changes_df = pd.DataFrame(companies_with_changes)

# 3.4. Владельцы в нескольких компаниях
owner_counts = clean_df.groupby('ФИО')['Компания'].nunique().reset_index()
multi_owners = owner_counts[owner_counts['Компания'] > 1].copy()
multi_owners.rename(columns={'ФИО': 'Владелец', 'Компания': 'Количество_компаний'}, inplace=True)

# 3.5. Записи с отсутствующей долей
missing_ownership = clean_df[clean_df['Доля_%'].isna()]

# -------------------------------------------------------------------
# 4. Сохранение результатов в Excel (только нужные листы)
# -------------------------------------------------------------------
output_file = "corporate_analysis_result.xlsx"
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    clean_df.to_excel(writer, sheet_name='Очищенные_данные', index=False)
    missing_inn.to_excel(writer, sheet_name='Отсутствует_ИНН', index=False)
    over100.to_excel(writer, sheet_name='Сумма_долей_>100', index=False)
    if not changes_df.empty:
        changes_df.to_excel(writer, sheet_name='Изменения_в_компаниях', index=False)
    if not multi_owners.empty:
        multi_owners.to_excel(writer, sheet_name='Владельцы_>1_компании', index=False)
    missing_ownership.to_excel(writer, sheet_name='Отсутствует_доля', index=False)

# -------------------------------------------------------------------
# 5. Вывод в консоль (кратко)
# -------------------------------------------------------------------

print(f"Записей после очистки: {len(clean_df)}")
print(f"Отсутствует ИНН: {len(missing_inn)}")
print(f"Компаний с суммой долей >100% на дату: {len(over100)}")
print(f"Компаний с изменениями: {len(changes_df)}")
if not changes_df.empty:
    print("\nДетали изменений:")
    for _, row in changes_df.iterrows():
        print(f"\n{row['Компания']}:")
        print(f"  {row['Первая дата']} -> {row['Последняя дата']}")
        print(f"  {row['Характер изменений']}")
print(f"\nВладельцев в нескольких компаниях: {len(multi_owners)}")
if not multi_owners.empty:
    for _, row in multi_owners.iterrows():
        print(f"  {row['Владелец']} – {row['Количество_компаний']} компании")
print(f"Записей без доли: {len(missing_ownership)}")
print(f"Результаты сохранены в {output_file}")