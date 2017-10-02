<?php
/**
 * Created by PhpStorm.
 * User: Sinri
 * Date: 2017/10/2
 * Time: 13:36
 */

require_once __DIR__ . '/../autoload.php';

$dbLink = __DIR__ . '/io/sample_sqlite.db';
@unlink($dbLink);

$ESA = new \sinri\excelxsql\ExcelAgent\ExcelSqliteAgent($dbLink);


$ESA->loadSheetWithTypeDefinition(
    "test2",
    __DIR__ . '/io/sample_input.xlsx',
    0,
    [
        "rec_id" => \sinri\excelxsql\ExcelAgent\ExcelSqliteAgent::FIELD_TYPE_INTEGER,
        "name" => \sinri\excelxsql\ExcelAgent\ExcelSqliteAgent::FIELD_TYPE_TEXT,
        "value" => \sinri\excelxsql\ExcelAgent\ExcelSqliteAgent::FIELD_TYPE_NUMERIC,
        "date" => \sinri\excelxsql\ExcelAgent\ExcelSqliteAgent::FIELD_TYPE_TEXT,
    ],
    false,
    false
);
$ESA->loadSheetWithTypeDefinition(
    "test2",
    __DIR__ . '/io/sample_input.xlsx',
    1,
    [
        "address_id" => \sinri\excelxsql\ExcelAgent\ExcelSqliteAgent::FIELD_TYPE_INTEGER,
        "name" => \sinri\excelxsql\ExcelAgent\ExcelSqliteAgent::FIELD_TYPE_TEXT,
        "address" => \sinri\excelxsql\ExcelAgent\ExcelSqliteAgent::FIELD_TYPE_TEXT,
    ],
    false,
    false
);

$sql = "SELECT name,value,date FROM test2_Sheet1 WHERE name=:name";
$values = [":name" => "abe"];
$result = $ESA->getDb()->safeReadRow($sql, $values);
var_dump($result);

$sql = "SELECT address FROM test2_Sheet2 WHERE name=:name";
$values = [":name" => "abe"];
$result = $ESA->getDb()->safeReadCol($sql, $values);
var_dump($result);

$sql = "SELECT test2_Sheet1.name,test2_Sheet1.value,test2_Sheet1.date,test2_Sheet2.address FROM test2_Sheet1 INNER JOIN test2_Sheet2 ON test2_Sheet1.name=test2_Sheet2.name WHERE test2_Sheet1.value<:limit";
$values = [":limit" => 10];
$result = $ESA->getDb()->safeReadAll($sql, $values);
var_dump($result);

echo "------" . PHP_EOL;

$sql = "SELECT test2_Sheet1.name AS 姓名,test2_Sheet1.value AS 值,test2_Sheet1.date AS 日期,group_concat(test2_Sheet2.address) AS 地址 
FROM test2_Sheet1 INNER JOIN test2_Sheet2 ON test2_Sheet1.name=test2_Sheet2.name 
WHERE test2_Sheet1.value<:limit GROUP BY test2_Sheet1.name
";
$values = [":limit" => 50];
$result = $ESA->getDb()->safeReadAll($sql, $values);
var_dump($result);
$result = $ESA->getDb()->safeExecute($sql, $values);
$ESA->writeResultMatrixToXlsx($result, "group_concat", __DIR__ . '/io/sample_output.xlsx', false);