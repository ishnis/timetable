<?php
require 'vendor/autoload.php'; // Для работы с Excel файлами

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;

function getNextClassDate() {
    try {
        $spreadsheet = IOFactory::load(__DIR__ . '/расписание.xlsx');
        $worksheet = $spreadsheet->getActiveSheet();

        // Читаем данные
        $startDateValue = $worksheet->getCell('G6')->getValue(); // Дата начала занятий
        $timeValue = $worksheet->getCell('F6')->getValue(); // Время (доля дня или строка)
        $address = $worksheet->getCell('H6')->getValue(); // Адрес

        // Проверяем, что время не пустое
        if (empty($timeValue) || $timeValue === '-') {
            return null;
        }

        // Получаем день недели начала занятий
        $startDate = null;
        if (is_numeric($startDateValue) && $startDateValue > 0) {
            $startDate = Date::excelToDateTimeObject($startDateValue);
        } else {
            try {
                $startDate = new DateTime($startDateValue);
            } catch (Exception $e) {
                return null;
            }
        }
        
        // Получаем номер дня недели начала занятий (1 = Пн, 7 = Вс)
        $classDayOfWeek = (int)$startDate->format('N');

        // Получаем текущую дату
        $today = new DateTime();
        
        // Находим следующий день занятий
        $nextClassDay = clone $today;
        $daysUntilNextClass = ($classDayOfWeek - $today->format('N') + 7) % 7;
        if ($daysUntilNextClass === 0) {
            $daysUntilNextClass = 7; // Если сегодня день занятий, берем следующий
        }
        $nextClassDay->modify("+{$daysUntilNextClass} days");

        // Убедимся, что адрес не пустой, иначе установим значение по умолчанию
        if (empty($address)) {
            $address = "Адрес не указан";
        }

        return [
            'date' => $nextClassDay, // Возвращаем дату следующего занятия
            'time' => $timeValue, // Возвращаем исходное значение времени из F6
            'address' => $address
        ];

    } catch (Exception $e) {
        return null;
    }
}

function generateCalendar() {
    $classInfo = getNextClassDate();

    if (!$classInfo) {
        return '<div class="calendar-error">Не удалось определить дату следующего занятия из файла. Убедитесь, что файл существует и данные в G6 (дата) и F6 (время) заполнены корректно.</div>';
    }

    $nextClassDateFromExcel = $classInfo['date']; // Объект даты-времени занятия из Excel
    $time = $classInfo['time']; // Исходная строка или число времени (для отображения)
    $address = $classInfo['address']; // Обработанный адрес

    // --- Определение начала недели от ТЕКУЩЕЙ даты ---
    $today = new DateTime();
    $today->setTime(0, 0, 0); // Устанавливаем время в начало дня
    $weekStart = clone $today;
    $weekStart->modify('monday this week');

    $calendar = '<div class="calendar-grid">';

    // Заголовок с днями недели
    $calendar .= '<div class="calendar-header">';
    $dayNames = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс'];
    for ($i = 0; $i < 7; $i++) {
        $date = clone $weekStart;
        $date->modify("+{$i} days");
        $calendar .= '<div class="calendar-day">';
        $calendar .= '<div class="day-name">' . $dayNames[$i] . '</div>';
        $calendar .= '<div class="day-number">' . $date->format('d') . '</div>';
        $calendar .= '</div>';
    }
    $calendar .= '</div>';

    // Ячейки с занятиями
    $calendar .= '<div class="calendar-body">';
    // Определяем день недели занятия из даты в Excel (1 = Пн, 7 = Вс)
    $classDayOfWeek = (int)$nextClassDateFromExcel->format('N');

    for ($i = 0; $i < 7; $i++) {
        $date = clone $weekStart; // Дата ячейки календаря (в отображаемой неделе)
        $date->modify("+{$i} days");
        $date->setTime(0, 0, 0); // Устанавливаем время в начало дня
        $currentDayOfWeek = (int)$date->format('N'); // День недели текущей ячейки (1 = Пн, 7 = Вс)

        $calendar .= '<div class="calendar-cell">'; // Ячейка создается всегда

        // Проверяем только, что это день занятий и дата не в прошлом
        if ($currentDayOfWeek === $classDayOfWeek && $date >= $today) {
            $calendar .= '<div class="class-dot"></div>';
            $calendar .= '<div class="class-name">Проектная деятельность</div>';
            // Отображаем время
            $displayTime = $time;
            if (is_numeric($time)) {
                try {
                    $excelTimeForDisplay = Date::excelToDateTimeObject($time);
                    $displayTime = $excelTimeForDisplay->format('H:i');
                } catch (Exception $e) {
                    $displayTime = 'Неверное время';
                }
            }
            $calendar .= '<div class="class-time">' . $displayTime . '</div>';
        }

        $calendar .= '</div>'; // Закрываем тег ячейки
    }
    $calendar .= '</div>'; // Закрываем тег calendar-body

    $calendar .= '</div>'; // Закрываем тег calendar-grid

    // Добавляем адрес под календарем
    $calendar .= '<div class="class-address">Адрес: ' . $address . '</div>';

    return $calendar;
}
?>