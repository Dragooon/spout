<?php

namespace Box\Spout\Writer\Style;

/**
 * Class NumberFormat
 * This class provides constants and functions to work with NumberFormatting
 *
 * @package Box\Spout\Writer\Style
 */
class NumberFormat
{
    const FORMAT_GENERAL = 'General';
    const FORMAT_TEXT = '@';
    const FORMAT_NUMBER = '0';
    const FORMAT_NUMBER_00 = '0.00';
    const FORMAT_NUMBER_COMMA_SEPARATED1 = '#,##0.00';
    const FORMAT_NUMBER_COMMA_SEPARATED2 = '#,##0.00_-';
    const FORMAT_PERCENTAGE = '0%';
    const FORMAT_PERCENTAGE_00 = '0.00%';
    const FORMAT_DATE_YYYYMMDD2 = 'yyyy-mm-dd';
    const FORMAT_DATE_YYYYMMDD = 'yy-mm-dd';
    const FORMAT_DATE_DDMMYYYY = 'dd/mm/yy';
    const FORMAT_DATE_DMYSLASH = 'd/m/yy';
    const FORMAT_DATE_DMYMINUS = 'd-m-yy';
    const FORMAT_DATE_DMMINUS = 'd-m';
    const FORMAT_DATE_MYMINUS = 'm-yy';
    const FORMAT_DATE_XLSX14 = 'mm-dd-yy';
    const FORMAT_DATE_XLSX15 = 'd-mmm-yy';
    const FORMAT_DATE_XLSX16 = 'd-mmm';
    const FORMAT_DATE_XLSX17 = 'mmm-yy';
    const FORMAT_DATE_XLSX22 = 'm/d/yy h:mm';
    const FORMAT_DATE_DATETIME = 'd/m/yy h:mm';
    const FORMAT_DATE_TIME1 = 'h:mm AM/PM';
    const FORMAT_DATE_TIME2 = 'h:mm:ss AM/PM';
    const FORMAT_DATE_TIME3 = 'h:mm';
    const FORMAT_DATE_TIME4 = 'h:mm:ss';
    const FORMAT_DATE_TIME5 = 'mm:ss';
    const FORMAT_DATE_TIME6 = 'h:mm:ss';
    const FORMAT_DATE_TIME7 = 'i:s.S';
    const FORMAT_DATE_TIME8 = 'h:mm:ss;@';
    const FORMAT_DATE_YYYYMMDDSLASH = 'yy/mm/dd;@';
    const FORMAT_CURRENCY_USD_SIMPLE = '"$"#,##0.00_-';
    const FORMAT_CURRENCY_USD = '$#,##0_-';
    const FORMAT_CURRENCY_EUR_SIMPLE = '#,##0.00_-"€"';
    const FORMAT_CURRENCY_EUR = '#,##0_-"€"';


    /**
     * Convert a PHP DateTime object to an MS Excel serialized date/time value.
     *
     * @param \DateTimeInterface $dateValue PHP DateTime object
     *
     * @return float MS Excel serialized date/time value
     */
    public static function dateTimeToExcel(\DateTimeInterface $dateValue)
    {
        return self::formattedPHPToExcel(
            $dateValue->format('Y'),
            $dateValue->format('m'),
            $dateValue->format('d'),
            $dateValue->format('H'),
            $dateValue->format('i'),
            $dateValue->format('s')
        );
    }

    /**
     * formattedPHPToExcel.
     *
     * Code borrowed from PhpOffice/PhpSpreadsheet
     *
     * @param int $year
     * @param int $month
     * @param int $day
     * @param int $hours
     * @param int $minutes
     * @param int $seconds
     *
     * @return float Excel date/time value
     */
    public static function formattedPHPToExcel($year, $month, $day, $hours = 0, $minutes = 0, $seconds = 0)
    {
        //
        //    Fudge factor for the erroneous fact that the year 1900 is treated as a Leap Year in MS Excel
        //    This affects every date following 28th February 1900
        //
        $excel1900isLeapYear = true;
        if (($year == 1900) && ($month <= 2)) {
            $excel1900isLeapYear = false;
        }
        $myexcelBaseDate = 2415020;

        //    Julian base date Adjustment
        if ($month > 2) {
            $month -= 3;
        } else {
            $month += 9;
            --$year;
        }

        //    Calculate the Julian Date, then subtract the Excel base date (JD 2415020 = 31-Dec-1899 Giving Excel Date of 0)
        $century = substr($year, 0, 2);
        $decade = substr($year, 2, 2);
        $excelDate = floor((146097 * $century) / 4) + floor((1461 * $decade) / 4) + floor((153 * $month + 2) / 5) + $day + 1721119 - $myexcelBaseDate + $excel1900isLeapYear;

        $excelTime = (($hours * 3600) + ($minutes * 60) + $seconds) / 86400;

        return (float) $excelDate + $excelTime;
    }
}
