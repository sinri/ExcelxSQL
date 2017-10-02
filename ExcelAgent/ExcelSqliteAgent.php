<?php
/**
 * Created by PhpStorm.
 * User: Sinri
 * Date: 2017/10/1
 * Time: 22:33
 */

namespace sinri\excelxsql\ExcelAgent;

use Box\Spout\Common\Helper\GlobalFunctionsHelper;
use Box\Spout\Reader\XLSX\Reader;
use Box\Spout\Reader\XLSX\Sheet;
use Box\Spout\Writer\XLSX\Writer;
use DateTime;
use sinri\excelxsql\SQLite\LibSqlite3;
use SQLite3Result;

class ExcelSqliteAgent
{
    const FIELD_TYPE_TEXT = "TEXT";
    const FIELD_TYPE_NUMERIC = "NUMERIC";
    const FIELD_TYPE_INTEGER = "INTEGER";
    const FIELD_TYPE_REAL = "REAL";
    const FIELD_TYPE_NONE = "NONE";

    protected $db;

    /**
     * @return LibSqlite3
     */
    public function getDb()
    {
        return $this->db;
    }

    /**
     * ExcelSqliteAgent constructor.
     * @param string $dbLink 默认使用内存中的临时数据库。也可以指定数据库文件路径。
     */
    public function __construct($dbLink = ":memory:")
    {
        $this->db = new LibSqlite3($dbLink);
    }

    /**
     * 以全自动方式加载一个xlsx，其中的每个范围内的工作薄将转化为$prefix_$sheetName为名的表，表中字段以首行内容为名。
     * @param string $prefix 表前缀，用于提示来源xlsx
     * @param string $xlsxFilePath xlsx文件路径
     * @param null|int|int[]|string|string[] $sheetRange 默认为null，即加载全部工作薄。可以为一个数字（自0开始）作为工作薄索引或者一个字符串作为工作薄名称。也可以为一个数组，组合前述条目。
     * @throws \Exception 当无法执行SQL是抛出异常
     */
    public function loadXlsx($prefix, $xlsxFilePath, $sheetRange = null)
    {
        $reader = new Reader();
        $reader->setGlobalFunctionsHelper(new GlobalFunctionsHelper());
        $reader->setShouldFormatDates(false);
        $reader->open($xlsxFilePath);

        foreach ($reader->getSheetIterator() as $sheet) {
            $sheet = $this->getSheet($sheet);
            if ($sheetRange !== null) {
                if (is_integer($sheetRange)) {
                    if ($sheet->getIndex() != $sheetRange && $sheet->getName() != $sheetRange) {
                        continue;
                    }
                } elseif (is_array($sheetRange)) {
                    if (!in_array($sheet->getName(), $sheetRange) && !in_array($sheet->getIndex(), $sheetRange)) {
                        continue;
                    }
                }
            }
            $fields = [];
            foreach ($sheet->getRowIterator() as $row_index => $row) {
                //echo "[{$row_index}]" . PHP_EOL;
                $values = [];
                if ($row_index == 1) {
                    // title_row, build the virtual table
                    $fields_with_type = [];
                    foreach ($row as $item) {
                        $fields[] = $item;
                        $fields_with_type[] = $item . " " . self::FIELD_TYPE_TEXT;
                    }
                    $sql = "create table {$prefix}_{$sheet->getName()} (";
                    $sql .= implode(",", $fields_with_type);
                    $sql .= ")";
                } else {
                    $sql = "insert into {$prefix}_{$sheet->getName()} (";
                    $sql .= implode(",", $fields);
                    $sql .= ")values(";
                    $x = [];
                    for ($i = 0; $i < count($fields); $i++) {
                        $value_key = ":" . $fields[$i];
                        $x[] = $value_key;
                        $this_value = isset($row[$i]) ? $row[$i] : '';
                        if (is_a($this_value, DateTime::class)) {
                            $this_value = $this_value->format('Y-m-d H:i:s');
                        }
                        $values[$value_key] = $this_value;
                    }
                    $sql .= implode(',', $x);
                    $sql .= ")";
                }
                //echo $sql . PHP_EOL;
                $done = $this->db->safeExecute($sql, $values);
                if (!$done) {
                    throw new \Exception("cannot run sql: " . $sql);
                }
            }
        }
    }

    /**
     * 以全自动方式加载一个xlsx，其中的每个范围内的工作薄将转化为$prefix_sheet_$sheetIndex为名的表，表中字段以col_$colIndex为名。
     * @param string $prefix 表前缀，用于提示来源xlsx
     * @param string $xlsxFilePath xlsx文件路径
     * @param null|int|int[]|string|string[] $sheetRange 默认为null，即加载全部工作薄。可以为一个数字（自0开始）作为工作薄索引或者一个字符串作为工作薄名称。也可以为一个数组，组合前述条目。
     * @throws \Exception 当无法执行SQL是抛出异常
     */
    public function loadXlsxWithIndexFormat($prefix, $xlsxFilePath, $sheetRange = null)
    {
        $reader = new Reader();
        $reader->setGlobalFunctionsHelper(new
        GlobalFunctionsHelper());
        $reader->setShouldFormatDates(false);
        $reader->open($xlsxFilePath);

        foreach ($reader->getSheetIterator() as $sheet) {
            $sheet = $this->getSheet($sheet);
            if ($sheetRange !== null) {
                if (is_integer($sheetRange)) {
                    if ($sheet->getIndex() != $sheetRange && $sheet->getName() != $sheetRange) {
                        continue;
                    }
                } elseif (is_array($sheetRange)) {
                    if (!in_array($sheet->getName(), $sheetRange) && !in_array($sheet->getIndex(), $sheetRange)) {
                        continue;
                    }
                }
            }
            $fields = [];
            foreach ($sheet->getRowIterator() as $row_index => $row) {
                //echo "[{$row_index}]" . PHP_EOL;
                $values = [];
                if ($row_index == 1) {
                    // title_row, build the virtual table
                    $fields_with_type = [];
                    foreach ($row as $field_index => $item) {
                        $fields[] = "col_" . $field_index;
                        $fields_with_type[] = "col_" . $field_index . " " . self::FIELD_TYPE_TEXT;
                    }
                    $sql = "create table {$prefix}_sheet_{$sheet->getIndex()} (";
                    $sql .= implode(",", $fields_with_type);
                    $sql .= ")";
                } else {
                    $sql = "insert into {$prefix}_sheet_{$sheet->getIndex()} (";
                    $sql .= implode(",", $fields);
                    $sql .= ")values(";
                    $x = [];
                    for ($i = 0; $i < count($fields); $i++) {
                        $value_key = ":" . $fields[$i];
                        $x[] = $value_key;
                        $this_value = isset($row[$i]) ? $row[$i] : '';
                        if (is_a($this_value, DateTime::class)) {
                            $this_value = $this_value->format('Y-m-d H:i:s');
                        }
                        $values[$value_key] = $this_value;
                    }
                    $sql .= implode(',', $x);
                    $sql .= ")";
                }
                //echo $sql . PHP_EOL;
                $done = $this->db->safeExecute($sql, $values);
                if (!$done) {
                    throw new \Exception("cannot run sql: " . $sql);
                }
            }
        }
    }

    /**
     * 以指定列内容类型方式加载一个xlsx中的一个工作薄，转化为表。
     * @param string $prefix 表前缀
     * @param string $xlsxFilePath 文件路径
     * @param int $sheetIndex 工作薄索引（自0开始）
     * @param array $fieldDefinition 如{first_row_col_value:type}。默认为空数组。未定义的默认类型为TEXT。
     * @param bool $useIndexFieldName 是否使用索引作为字段名，默认为true。
     * @param bool $useIndexTableName 是否使用索引作为表名，默认为true。
     * @throws \Exception 无法执行SQL时抛出异常
     */
    public function loadSheetWithTypeDefinition($prefix, $xlsxFilePath, $sheetIndex, $fieldDefinition = [], $useIndexFieldName = true, $useIndexTableName = true)
    {
        $reader = new Reader();
        $reader->setGlobalFunctionsHelper(new
        GlobalFunctionsHelper());
        $reader->setShouldFormatDates(false);
        $reader->open($xlsxFilePath);

        foreach ($reader->getSheetIterator() as $sheet) {
            $sheet = $this->getSheet($sheet);
            if ($sheet->getIndex() != $sheetIndex) {
                continue;
            }

            $tableName = ($useIndexTableName ? "{$prefix}_sheet_{$sheet->getIndex()}" : "{$prefix}_{$sheet->getName()}");

            $fields = [];
            $field_type_dict = [];
            foreach ($sheet->getRowIterator() as $row_index => $row) {
                //echo "[{$row_index}]" . PHP_EOL;
                $values = [];
                if ($row_index == 1) {
                    // title_row, build the virtual table
                    $fields_with_type = [];
                    foreach ($row as $field_index => $item) {
                        $field_name = $useIndexFieldName ? ("col_" . $field_index) : $item;
                        $fields[] = $field_name;
                        $field_type = isset($fieldDefinition[$item]) ? $fieldDefinition[$item] : self::FIELD_TYPE_TEXT;
                        $field_type_dict[] = $field_type;
                        $fields_with_type[] = $field_name . " " . $field_type;
                    }
                    $sql = "create table {$tableName} (";
                    $sql .= implode(",", $fields_with_type);
                    $sql .= ")";
                } else {
                    $sql = "insert into {$tableName} (";
                    $sql .= implode(",", $fields);
                    $sql .= ")values(";
                    $x = [];
                    for ($i = 0; $i < count($fields); $i++) {
                        $value_key = ":" . $fields[$i];
                        $x[] = $value_key;
                        $this_value = isset($row[$i]) ? $row[$i] : '';
                        if (is_a($this_value, DateTime::class)) {
                            $this_value = $this_value->format('Y-m-d H:i:s');
                        } elseif (in_array($field_type_dict[$i], [
                            self::FIELD_TYPE_INTEGER, self::FIELD_TYPE_NUMERIC, self::FIELD_TYPE_REAL
                        ])) {
                            $old_value = $this_value;
                            if (empty($this_value)) {
                                $this_value = 0;
                            } else {
                                $this_value = str_replace(',', '', $this_value);
                                //echo " fix: " . $old_value . " -> " . $this_value . PHP_EOL;
                            }
                        }
                        $values[$value_key] = $this_value;
                    }
                    $sql .= implode(',', $x);
                    $sql .= ")";
                }
                //echo $sql . PHP_EOL;
                $done = $this->db->safeExecute($sql, $values);
                if (!$done) {
                    throw new \Exception("cannot run sql: " . $sql);
                }
            }
        }
    }

    /**
     * 将某个SQLite3查询结果导出xlsx
     * @param SQLite3Result $result
     * @param string $sheetName
     * @param string $outputXlsxFile
     * @param bool $forDownload 直接输出到浏览器下载，默认false
     * @internal param bool $isAppendSheet
     */
    public function writeResultMatrixToXlsx($result, $sheetName = 'Excel-SQL-Result', $outputXlsxFile = 'output.xlsx', $forDownload = false)
    {
        $writer = new Writer();
        $writer->setGlobalFunctionsHelper(new GlobalFunctionsHelper());
        $writer->setShouldUseInlineStrings(false);

        if ($forDownload) {
            $writer->openToBrowser($outputXlsxFile); // stream data directly to the browser
        } else {
            $writer->openToFile($outputXlsxFile); // write data to a file or to a PHP stream
        }

        $sheet = $writer->getCurrentSheet();
        $sheet->setName($sheetName);

        $isFirstRow = true;
        while ($row = $result->fetchArray(SQLITE3_ASSOC)) {
            if ($isFirstRow) {
                $values = [];
                foreach ($row as $filed_name => $field_value) {
                    $values[] = $filed_name;
                }
                $writer->addRow($values);
                $isFirstRow = false;
            }
            $values = [];
            foreach ($row as $filed_name => $field_value) {
                $values[] = $field_value;
            }
            $writer->addRow($values);
        }

        $writer->close();
    }

    /**
     * @param $sheet
     * @return Sheet
     */
    private function getSheet(&$sheet)
    {
        return $sheet;
    }
}