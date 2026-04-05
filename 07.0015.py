import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import openpyxl

#добавление счетчиков
def analyze_data(spec_df, ref_df, manual_entries):
    total_elements = 0
    total_outputs = 0
    pin_outputs = 0
    planar_outputs = 0
    formable_elements = 0
    tinning_elements = 0
    adhesive_elements_pad = 0
    adhesive_elements_board = 0
    pmi_count = 0
    pmi_outputs_count = 0
    total_microchips = 0
    total_microchips_outputs = 0
    total_flux_places = 0
    microchips_with_planar_outputs = 0
    microchips_with_pin_outputs = 0
    formable_planar_elements = 0
    formable_pin_elements = 0
    microchips_pin_outputs = 0
    microchips_planar_outputs = 0
    connector_sh = 0
    connector_pl = 0
    outputs_connector_pl = 0
    outputs_connector_sh = 0

    elem_pl_3 = 0
    elem_pl_8 = 0
    elem_pl_9 = 0
    elem_pl_14 = 0
    elem_pl_16 = 0
    elem_pl_18 = 0
    elem_pl_24 = 0
    elem_pl_28 = 0
    elem_pl_42 = 0
    elem_pl_48 = 0
    elem_pl_64 = 0
    elem_pl_72 = 0
    elem_pl_108 = 0
    elem_pl_132 = 0
    elem_pl_208 = 0
    elem_pl_240 = 0
    elem_pl_256 = 0

    elem_pin_2 = 0
    elem_pin_3 = 0
    elem_pin_4 = 0
    elem_pin_5 = 0
    elem_pin_8 = 0
    elem_pin_12 = 0

    unique_pin_elements = set()
    unique_planar_elements = set()
    unique_pmi_elements = set()

    results = []

    found_elements_count = 0
    planar_element_count = 0
    pin_element_count = 0

    #подсчет и сверка с справочником
    for index, row in spec_df.iterrows():
        designation = row['Designator']
        results.append(f"Проверка обозначения {designation}.")

        ref_row = ref_df[ref_df['LibRef'].apply(lambda x: str(x) in row['Companent_key'])]

        if not ref_row.empty:
            found_elements_count += 1
            outputs_count = ref_row['количество выводов'].values[0]
            flux_count = ref_row['флюсование'].fillna(0).values[0]

            total_outputs += outputs_count

            total_flux_places += flux_count

            if ref_row['микросхема'].values[0] == 1:
                total_microchips += 1
                total_microchips_outputs += outputs_count

            if ref_row['формовка'].values[0] == 1:
                formable_elements += 1
                type_output = str(ref_row['тип выводов'].values[0]).lower()
                if 'пл' in type_output:
                    formable_planar_elements += 1
                elif 'ш' in type_output:
                    formable_pin_elements += 1

            if ref_row['лужение'].values[0] == 1:
                tinning_elements += 1

            if ref_row['приклеивание прокладки'].values[0] == 1:
                adhesive_elements_pad += 1

            if ref_row['приклеивание к плате'].values[0] == 1:
                adhesive_elements_board += 1

            type_output = str(ref_row['тип выводов'].values[0]).lower()

            results.append(f"Тип вывода для {designation}: {type_output}, Количество мест флюсования: {flux_count}")

            if 'пми' in type_output:
                pmi_count += 1
                pmi_outputs_count += outputs_count
                unique_pmi_elements.add(designation)

            if 'пл' in type_output:
                planar_outputs += outputs_count
                planar_element_count += 1
                unique_planar_elements.add(designation)

                if ref_row['СоК'].values[0] == 1:
                    connector_pl += 1
                    outputs_connector_pl += outputs_count

                if ref_row['микросхема'].values[0] == 1:
                    microchips_planar_outputs += outputs_count

                if outputs_count == 3 and ref_row['формовка'].values[0] == 1:
                    elem_pl_3 += 1
                if outputs_count == 8 and ref_row['формовка'].values[0] == 1:
                    elem_pl_8 += 1
                if outputs_count == 9 and ref_row['формовка'].values[0] == 1:
                    elem_pl_9 += 1
                if outputs_count == 14 and ref_row['формовка'].values[0] == 1:
                    elem_pl_14 += 1
                if outputs_count == 16 and ref_row['формовка'].values[0] == 1:
                    elem_pl_16 += 1
                if outputs_count == 18 and ref_row['формовка'].values[0] == 1:
                    elem_pl_18 += 1
                if outputs_count == 24 and ref_row['формовка'].values[0] == 1:
                    elem_pl_24 += 1
                if outputs_count == 28 and ref_row['формовка'].values[0] == 1:
                    elem_pl_28 += 1
                if outputs_count == 42 and ref_row['формовка'].values[0] == 1:
                    elem_pl_42 += 1
                if outputs_count == 48 and ref_row['формовка'].values[0] == 1:
                    elem_pl_48 += 1
                if outputs_count == 64 and ref_row['формовка'].values[0] == 1:
                    elem_pl_64 += 1
                if outputs_count == 72 and ref_row['формовка'].values[0] == 1:
                    elem_pl_72 += 1
                if outputs_count == 108 and ref_row['формовка'].values[0] == 1:
                    elem_pl_108 += 1
                if outputs_count == 132 and ref_row['формовка'].values[0] == 1:
                    elem_pl_132 += 1
                if outputs_count == 208 and ref_row['формовка'].values[0] == 1:
                    elem_pl_208 += 1
                if outputs_count == 240 and ref_row['формовка'].values[0] == 1:
                    elem_pl_240 += 1
                if outputs_count == 256 and ref_row['формовка'].values[0] == 1:
                    elem_pl_256 += 1

                microchips_with_planar_outputs += (ref_row['микросхема'].value[0] == 1)

            if 'ш' in type_output:
                pin_outputs += outputs_count
                pin_element_count += 1
                unique_pin_elements.add(designation)

                if ref_row['СоК'].values[0] == 1:
                    connector_sh += 1
                    outputs_connector_sh += outputs_count

                if ref_row['микросхема'].values[0] == 1:
                    microchips_pin_outputs += outputs_count

                if outputs_count == 2 and ref_row['формовка'].values[0] == 1:
                    elem_pin_2 += 1
                if outputs_count == 3 and ref_row['формовка'].values[0] == 1:
                    elem_pin_3 += 1
                if outputs_count == 4 and ref_row['формовка'].values[0] == 1:
                    elem_pin_4 += 1
                if outputs_count == 5 and ref_row['формовка'].values[0] == 1:
                    elem_pin_5 += 1
                if outputs_count == 8 and ref_row['формовка'].values[0] == 1:
                    elem_pin_8 += 1
                if outputs_count == 12 and ref_row['формовка'].values[0] == 1:
                    elem_pin_12 += 1

                microchips_with_pin_outputs += (ref_row['микросхема'].values[0] == 1)

        else:
            results.append(f"Не найдено совпадение в справочнике для шифра {row['Companent_key']}.")

    #анализ данных, введенных вручную
    for entry in manual_entries:
        designation, count_str = entry
        count = int(count_str)
        results.append(f"Проверка ручного ввода: {designation}, Количество: {count}")

        ref_row_manual_entry = ref_df[ref_df['LibRef'] == designation]

        if not ref_row_manual_entry.empty:
            outputs_count_manual_entry = ref_row_manual_entry['количество выводов'].values[0]

            total_outputs += outputs_count_manual_entry * count

            found_elements_count += count

            type_output_manual_entry = str(ref_row_manual_entry['тип выводов'].values[0]).lower()

            if ref_row['формовка'].values[0] == 1:
                formable_elements += 1

            if ref_row['лужение'].values[0] == 1:
                tinning_elements += 1

            if 'пми' in type_output_manual_entry:
                pmi_count += 1
                pmi_outputs_count += count * outputs_count_manual_entry

            if 'пл' in type_output_manual_entry:
                planar_element_count += 1
                planar_outputs += count * outputs_count_manual_entry

                if ref_row_manual_entry['СоК'].values[0] == 1:
                    connector_pl += 1
                    outputs_connector_pl += count * outputs_count_manual_entry

            if 'ш' in type_output_manual_entry:
                pin_element_count += 1
                pin_outputs += count * outputs_count_manual_entry

                if ref_row_manual_entry['СоК'].values[0] == 1:
                    connector_sh += 1
                    outputs_connector_sh += count * outputs_count_manual_entry

    #список результатов
    results.append(f"\nОбщее количество элементов: {found_elements_count}")
    results.append(f"Общее количество микрохсем: {total_microchips}")
    results.append(f"Количество выводов микросхем: {total_microchips_outputs}")
    results.append(f"Количество планарных выводов микросхем: {microchips_planar_outputs}")
    results.append(f"Количество штырьковых выводов микросхем: {microchips_pin_outputs}")
    results.append(f"Общее количество выводов: {total_outputs}")
    results.append(f"Общее количество мест флюсования: {total_flux_places}")
    results.append(f"Количество штырьковых выводов: {pin_outputs}")
    results.append(f"Количество планарных выводов: {planar_outputs}")
    results.append(f"Количество элементов с планарными выводами: {planar_element_count}")
    results.append(f"Количество элементов с штырьковыми выводами: {pin_element_count}")
    results.append(f"Количество микросхем с планарными выводами: {microchips_with_planar_outputs}")
    results.append(f"Количество микросхем с штырьковыми выводами: {microchips_pin_outputs}")
    results.append(f"Количество формуемых элементов: {formable_elements}")
    results.append(f"Количество формуемых элементов с планарными выводами: {formable_planar_elements}")
    results.append(f"Количество формуемых элементов с штырьковыми выводами: {formable_pin_elements}")
    results.append(f"Количество лудимых элементов: {tinning_elements}")
    results.append(f"Количество элементов с приклеиванием прокладки: {adhesive_elements_pad}")
    results.append(f"Количество элементов с приклеиванием к плате: {adhesive_elements_board}")

    results.append(f"\nКоличество ПМИ: {pmi_count}")
    results.append(f"Количество выводов ПМИ: {pmi_outputs_count}")

    #*******************************************************************************************************
    #запись результатов в ТТП
    ttp_file_path = ''
    workbook = openpyxl.load_workbook(ttp_file_path)

    sheet = workbook['5 Комплектование']
    sheet['F5'] = found_elements_count
    #sheet['F6'] =
    #sheet['F7'] =

    sheet = workbook['10 Контроль']
    sheet['F5'] = found_elements_count
    #sheet['F6'] =

    sheet = workbook['15 Подготовка ЭРЭ к монтажу']
    #ПРОВЕРКА НА ОТСУТ. МЕХ. ПОВ.
    sheet['G5'] = found_elements_count
    #ЗАЩИТА МАРКИРОВКИ ЭЛЕМЕНТОВ
    #sheet['G8'] =
    if sheet['G8'] is not None:
        sheet['G9'] = 1
    #sheet['G10'] =
    if sheet['G10'] is not None:
        sheet['G11'] = 1
    #ЗАЧИСТКА И ПРАВКА ВЫВОДОВ
    sheet['G12'] = total_flux_places
    sheet['G13'] = found_elements_count - pmi_count
    sheet['G14'] = found_elements_count - pmi_count
    #ПОДГОТОВКА К ЛУЖЕНИЮ ЖИЛЫ ПРОВОДОВ
    #sheet['G16'] =
    #sheet['G17'] =
    #sheet['G18'] =
    #sheet['G19'] =
    #sheet['G20'] =
    if sheet['G20'] is not None:
        sheet['G21'] = 1
    #ФОРМОВКА И ОБРЕЗКА ВЫВОДОВ МИКРОСХЕМ С ПЛАНАРНЫМИ ВЫВОДАМИ
    #sheet['G23'] =
    #sheet['G24'] =
    sheet['G25'] = formable_planar_elements
    sheet['G27'] = elem_pl_3
    sheet['G28'] = elem_pl_8
    sheet['G29'] = elem_pl_9
    sheet['G30'] = elem_pl_14
    sheet['G31'] = elem_pl_16
    sheet['G32'] = elem_pl_18
    sheet['G33'] = elem_pl_24
    sheet['G34'] = elem_pl_28
    sheet['G35'] = elem_pl_42
    sheet['G36'] = elem_pl_48
    sheet['G37'] = elem_pl_64
    sheet['G38'] = elem_pl_72
    sheet['G39'] = elem_pl_108
    sheet['G40'] = elem_pl_132
    sheet['G41'] = elem_pl_208
    sheet['G42'] = elem_pl_240
    sheet['G43'] = elem_pl_256
    sheet['G44'] = formable_planar_elements
    sheet['G45'] = formable_planar_elements
    #sheet['G46'] =
    #ЭРЭ, ФЛЮСОВАНИЕ И ЛУЖЕНИЕ
    sheet['G48'] = total_flux_places
    sheet['G49'] = total_flux_places
    if sheet['G49'] is not None:
        sheet['G50'] = 1
    #sheet['G51'] =
    sheet['G52'] = tinning_elements
    sheet['G53'] = total_flux_places
    sheet['G54'] = total_flux_places
    sheet['G55'] = total_flux_places
    #ПОДГОТОВКА ТРУБОК
    #sheet['G57'] =
    #sheet['G58'] =
    #sheet['G59'] =
    #sheet['G60'] =
    #sheet['G61'] =
    #sheet['G62'] =
    #sheet['G63'] =
    #sheet['G64'] =
    if sheet['G57'] is not None:
        sheet['G65'] = 1
        sheet['G66'] = 1
    #sheet['G67'] =
    #sheet['G68'] =
    if sheet['G57'] is not None:
        sheet['G69'] = 1
    #ЭРЭ С ШТЫРЬКОВЫМИ ВЫВОДАМИ, ФОРМОВКА
    #sheet['G71'] =
    sheet['G72'] = elem_pin_2
    sheet['G73'] = elem_pin_3
    sheet['G74'] = elem_pin_4
    sheet['G75'] = elem_pin_5
    sheet['G76'] = elem_pin_8
    sheet['G77'] = elem_pin_12
    sheet['G78'] = formable_pin_elements
    #ПРИКЛЕИВАНИЕ ПРОКЛАДОК ВК-9
    #sheet['G81'] =
    if sheet['G81'] is not None:
        sheet['G82'] = 1
    #sheet['G83'] =
    #sheet['G93'] =
    if sheet['G81'] is not None:
        sheet['G95'] = 1
    #sheet['G98'] =
    #ПРИКЛЕИВАНИЕ ПРОКЛАДОК АК-20
    #sheet['G100'] =
    if sheet['G100'] is not None:
        sheet['G101'] = 1
    #sheet['G102'] =
    if sheet['G100'] is not None:
        sheet['G112'] = 1
    #sheet['G113'] =
    if sheet['G100'] is not None:
        sheet['G123'] = 1
    #sheet['G124'] =
    if sheet['G100'] is not None:
        sheet['G125'] = 1
    #sheet['G126'] =
    #СОЕДИНИТЕЛЬ (Ш)
    sheet['G128'] = outputs_connector_sh
    sheet['G130'] = outputs_connector_sh
    sheet['G131'] = outputs_connector_sh
    if sheet['G128'] is not None:
        sheet['G132'] = 1
    #sheet['G133'] = soedinitel_sh
    sheet['G134'] = outputs_connector_sh
    sheet['G136'] = connector_sh
    #СОЕДИНИТЕЛЬ (П)
    sheet['G138'] = outputs_connector_pl
    sheet['G139'] = outputs_connector_pl
    if sheet['G138'] is not None:
        sheet['G140'] = 1
    #sheet['G141'] = soedinitel_sh * 2
    sheet['G143'] = outputs_connector_pl
    sheet['G144'] = outputs_connector_pl
    sheet['G146'] = outputs_connector_pl
    sheet['G147'] = 1

    sheet = workbook['G25 Подготовка ПП к монтажу']
    sheet['F5'] = 1
    sheet['F10'] = 1