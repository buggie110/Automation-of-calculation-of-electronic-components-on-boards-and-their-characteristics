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

    sheet = workbook['25 Подготовка ПП к монтажу']
    sheet['F5'] = 1
    sheet['F10'] = 1
    sheet['F11'] = 1
    sheet['F13'] = 1
    sheet['F14'] = 1
    sheet['F15'] = 1
    sheet['F16'] = 1
    sheet['F17'] = 1
    sheet['F18'] = 1
    sheet['F19'] = 1
    sheet['F20'] = 1
    sheet['F21'] = 1

    sheet = workbook['40 Установка и пайка ПМИ']
    sheet['F6'] = 1
    sheet['F7'] = 1
    sheet['F9'] = 1
    sheet['F10'] = 1
    sheet['F12'] = 1
    sheet['F13'] = 1
    sheet['F14'] = 1
    sheet['F15'] = pmi_count
    sheet['F16'] = pmi_count
    sheet['F17'] = pmi_outputs_count
    sheet['F18'] = pmi_count
    sheet['F19'] = pmi_count
    sheet['F21'] = pmi_count
    sheet['F22'] = pmi_count
    sheet['F25'] = 1
    sheet['F26'] = 1
    sheet['F27'] = 1
    sheet['F28'] = 1
    sheet['F29'] = 1
    sheet['F30'] = 1
    sheet['F31'] = 1
    sheet['F32'] = pmi_count
    sheet['F33'] = pmi_outputs_count
    sheet['F37'] = pmi_outputs_count / 100 * 10
    sheet['F38'] = 1
    sheet['F42'] = 1
    sheet['F47'] = pmi_outputs_count
    sheet['F48'] = 1

    sheet = workbook['50 Сборочно-монтажная 1']
    #УСТАНОВКА ЭРЭ С ШТЫРЬКОВЫМИ ВЫВОДАМИ
    sheet['G6'] = 1
    sheet['G8'] = pin_element_count - microchips_with_pin_outputs
    sheet['G9'] = pin_outputs - microchips_pin_outputs
    sheet['G10'] = pin_element_count - microchips_with_pin_outputs
    sheet['G11'] = pin_element_count - microchips_with_pin_outputs
    sheet['G12'] = pin_element_count - microchips_with_pin_outputs
    if sheet['G8'] is not None:
        sheet['G16'] = 1
    sheet['G19'] = pin_outputs - microchips_pin_outputs
    sheet['G20'] = pin_outputs - microchips_pin_outputs
    sheet['G22'] = pin_element_count - microchips_with_pin_outputs
    sheet['G23'] = pin_outputs - microchips_pin_outputs
    #УСТАНОВКА ЭРЭ В КОРПУСАХ С ПЛАНАРНЫМИ ВЫВОДАМИ
    sheet['G25'] = planar_element_count - microchips_with_planar_outputs
    #sheet['G26'] = planar_element_count - microchips_with_planar_outputs
    sheet['G27'] = planar_element_count - microchips_with_planar_outputs
    if sheet ['G25'] is not None:
        sheet['G30'] = 1
    sheet['G33'] = planar_outputs - microchips_planar_outputs
    #sheet['G34'] =
    sheet['G35'] = planar_element_count - microchips_with_planar_outputs
    sheet['G36'] = planar_element_count - microchips_with_planar_outputs
    sheet['G37'] = planar_outputs - microchips_planar_outputs
    #УСТАНОВКА МИКРОСХЕМ С ПЛАНАРНЫМИ ВЫВОДАМИ
    sheet['G39'] = microchips_with_planar_outputs
    sheet['G40'] = microchips_with_planar_outputs
    if sheet['G39'] is not None:
        sheet['G41'] = 1
    sheet['G42'] = microchips_with_planar_outputs
    sheet['G43'] = microchips_with_planar_outputs
    if sheet ['G39'] is not None:
        sheet['G45'] = 1
        sheet['G46'] = 1
    sheet['G50'] = microchips_with_planar_outputs
    sheet['G51'] = microchips_planar_outputs
    if sheet['G39'] is not None:
        sheet['G52'] = 1
    sheet['G54'] = microchips_with_planar_outputs
    sheet['G55'] = microchips_planar_outputs
    #УСТАНОВКА МИКРОСХЕМ С ШТЫРЬКОВЫМИ ВЫВОДАМИ
    sheet['G57'] = microchips_with_pin_outputs
    sheet['G58'] = microchips_pin_outputs
    sheet['G59'] = microchips_with_pin_outputs
    #sheet['G60'] = microchips_pin_outputs
    sheet['G61'] = microchips_pin_outputs
    sheet['G62'] = microchips_pin_outputs
    sheet['G63'] = microchips_with_pin_outputs
    sheet['G64'] = microchips_pin_outputs
    #УСТАНОВКА ОБЪЕМНЫХ ПЕРЕМЫЧЕК
    #sheet['G67'] =
    #sheet['G69'] =
    #sheet['G71'] =
    if sheet['G67'] is not None:
        sheet['G72'] = 1
    #sheet['G73'] =
    #sheet['G74'] =
    #sheet['G75'] =
    if sheet['G39'] is not None:
        sheet['G76'] = 1
    #sheet['G79'] =
    #sheet['G80'] =
    #УСТАНОВКА ЖГУТОВ
    #sheet['G105'] =
    #sheet['G106'] =
    #sheet['G107'] =
    #sheet['G108'] =
    #sheet['G109'] =
    #sheet['G110'] =
    #sheet['G111'] =
    #sheet['G112'] =
    #УСТАНОВКА ДРОССЕЛЕЙ
    #sheet['G114'] =
    if sheet['G114'] is not None:
        sheet['G115'] = 1
    #sheet['G116'] =
    #sheet['G117'] =
    if sheet['G114'] is not None:
        sheet['G119'] = 1
    #sheet['G122'] =
    #sheet['G124'] =
    #sheet['G125'] =
    if sheet['G114'] is not None:
        sheet['G132'] = 1
    #sheet['G135'] =
    #sheet['G136'] =
    #sheet['G137'] =
    if sheet['G114'] is not None:
        sheet['G135'] = 1
    #sheet['G139'] =

    sheet = workbook['55 Удаление остатков флюса']
    sheet['F5'] = 1
    sheet['F14'] = total_outputs - pmi_outputs_count
    sheet['F15'] = (total_outputs - pmi_outputs_count) / 100 * 5
    sheet['F16'] = (total_outputs - pmi_outputs_count) / 100 * 5
    sheet['F18'] = (total_outputs - pmi_outputs_count) / 100 * 5
    sheet['F19'] = 1
    sheet['F20'] = 1

    sheet - workbook['65 Проверка эл. параметров']
    if total_microchips_outputs == 0:
        sheet['F6'] = 1

    if total_microships_outputs > 0 and total_microchips_outputs < 29:
        sheet['F6'] = 0.5
        sheet['F7'] = 0.5
    elif total_microchips_outputs == 29:
        sheet['F7'] = 1

    if total_microships_outputs > 29 and total_microchips_outputs < 59:
        sheet['F6'] = 0.5
        sheet['F7'] = 0.5
    elif total_microchips_outputs == 59:
        sheet['F7'] = 1

    if total_microships_outputs > 59 and total_microchips_outputs < 88:
        sheet['F6'] = 0.5
        sheet['F7'] = 0.5
    elif total_microchips_outputs == 88:
        sheet['F7'] = 1

    if total_microships_outputs > 88 and total_microchips_outputs < 118:
        sheet['F6'] = 0.5
        sheet['F7'] = 0.5
    elif total_microchips_outputs == 118:
        sheet['F7'] = 1

    if total_microships_outputs > 118 and total_microchips_outputs < 147:
        sheet['F6'] = 0.5
        sheet['F7'] = 0.5
    elif total_microchips_outputs == 147:
        sheet['F7'] = 1

    if total_microships_outputs > 147 and total_microchips_outputs < 176:
        sheet['F6'] = 0.5
        sheet['F7'] = 0.5
    elif total_microchips_outputs == 176:
        sheet['F7'] = 1

    if total_microships_outputs > 176 and total_microchips_outputs < 206:
        sheet['F6'] = 0.5
        sheet['F7'] = 0.5
    elif total_microchips_outputs == 206:
        sheet['F7'] = 1

    if total_microships_outputs > 206 and total_microchips_outputs < 235:
        sheet['F6'] = 0.5
        sheet['F7'] = 0.5
    elif total_microchips_outputs == 235:
        sheet['F7'] = 1

    if total_microships_outputs > 235 and total_microchips_outputs < 265:
        sheet['F6'] = 0.5
        sheet['F7'] = 0.5
    elif total_microchips_outputs == 265:
        sheet['F7'] = 1

    if total_microships_outputs > 265 and total_microchips_outputs < 293:
        sheet['F6'] = 0.5
        sheet['F7'] = 0.5
    elif total_microchips_outputs == 293:
        sheet['F7'] = 1

    if total_microships_outputs > 293 and total_microchips_outputs < 553:
        sheet['F6'] = 0.5
        sheet['F7'] = 0.5
    elif total_microchips_outputs == 553:
        sheet['F7'] = 1

    if total_microships_outputs > 553 and total_microchips_outputs < 812:
        sheet['F6'] = 0.5
        sheet['F7'] = 0.5
    elif total_microchips_outputs == 812:
        sheet['F7'] = 1

    if total_microships_outputs > 812 and total_microchips_outputs < 1070:
        sheet['F6'] = 0.5
        sheet['F7'] = 0.5
    elif total_microchips_outputs == 1070:
        sheet['F7'] = 1

    if total_microships_outputs > 1070 and total_microchips_outputs < 1330:
        sheet['F6'] = 0.5
        sheet['F7'] = 0.5
    elif total_microchips_outputs == 1330:
        sheet['F7'] = 1

    if total_microships_outputs > 1330 and total_microchips_outputs < 1589:
        sheet['F6'] = 0.5
        sheet['F7'] = 0.5
    elif total_microchips_outputs == 1589:
        sheet['F7'] = 1

    if total_microships_outputs > 1589 and total_microchips_outputs < 1848:
        sheet['F6'] = 0.5
        sheet['F7'] = 0.5
    elif total_microchips_outputs == 1848:
        sheet['F7'] = 1

    if total_microships_outputs > 1848 and total_microchips_outputs < 2107:
        sheet['F6'] = 0.5
        sheet['F7'] = 0.5
    elif total_microchips_outputs == 2107:
        sheet['F7'] = 1

    if total_microships_outputs > 2107 and total_microchips_outputs < 2366:
        sheet['F6'] = 0.5
        sheet['F7'] = 0.5
    elif total_microchips_outputs == 2366:
        sheet['F7'] = 1

    if total_microships_outputs > 2366 and total_microchips_outputs < 2625:
        sheet['F6'] = 0.5
        sheet['F7'] = 0.5
    elif total_microchips_outputs == 2625:
        sheet['F7'] = 1

    if total_microships_outputs > 2625 and total_microchips_outputs < 2884:
        sheet['F6'] = 0.5
        sheet['F7'] = 0.5
    elif total_microchips_outputs == 2884:
        sheet['F7'] = 1

    workbook.save(ttp_file_path)

#**************************************************************************************

    return results

#анализ данных из файлов
def run_analysis():
    spec_path = filedialog.askopenfilemname(title="Выберите файл", filetypes=[("Excel files", ".xlsx")])

    ref_path = filedialog.askopenfilename(title="Выберите файл справочника", filetypes=[("Excel files", ".xslx")])

    spec_df = pd.read_excel(spec_path).dropna(subset=['Designator'])

    ref_df = pd.read_excel(ref_path)

    required_columns_ref = ['LibRef', 'количество выводов', 'тип выводов', 'формовка', 'лужение', 'приклеивание прокладки', 'приклеивание к плате', 'микросхема', 'соединитель']

    for col in required_columns_ref:
        if col not in ref_df.columns:
            raise ValueError(f"Отсутствует необходимый столбец '{col}' в файле справочника.")

    #отладочная информация загруженных данных
    print("Загруженные данные:")
    print("Спецификация:")
    print(spec_df)
    print("Справочник:")
    print(ref_df)

    manual_entries = []
    for entry in manual_entry_list.get(0, tk.END):
        designation, count_str = entry.split(',')
        manual_entries.append{(designation.strip(), int(count_str.strip()))}

    results = analyze_data(spec_df, ref_df, manual_entries)

    result_text.delete(1.0, tk.END)
    for line in results:
        result_text.insert(tk.END, line + "\n")

def delete_selected():
    selected_indies = manual_entry_list.curselection()
    for index in reversed(selected_indies):
        manual_entry_list.delete(index)

#создание основоного окна приложения
root = tk.Tk()

results_listbox = tk.Listbox(root)
results_listbox.pack()

def delete_selected():
    selected_indies = manual_entry_list.curselection()
    for index in reversed(selected_indies):
        manual_entry_list.delete(index)

def delete_selected_results():
    selected_indices = results_listbox.curselection()
    for index in reversed(selected_indies):
        results_listbox.delete(index)

frame = tk.Frame(root)
frame.pack()

result_text = tk.Text(frame, wrap=tk.WORD)
result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar = tk.Scrollbar(frame, command=result_text.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

result_text.config(yscrollcommand=scrollbar.set)

manual_entry_label = tk.Label(root, text="Введдимте шифр и количество элементов:")
manual_entry_label.pack()

entry_field = tk.Entry(root)
entry_field.pack()

add_button = tk.Button(root, text="Добавить", command=lambda: manual_entry_list.insert(tk.END, entry_field.get()))
add_button.pack()

delete_button = tk.Button(root, text="Удалить выбранное", command=delete_selected)
delete_button.pack()

sample_results = []
for res in sample_results:
    results_listbox.insert(tk.END, res)

#запуск анализа
run_analysis_button = tk.Button(root, text="Запустить анализ", command=run_analysis)
ruin_analysis_button.pack()

root.mainloop()
