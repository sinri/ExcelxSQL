<?php
/**
 * Created by PhpStorm.
 * User: Sinri
 * Date: 2017/10/1
 * Time: 23:39
 */

require_once __DIR__ . '/../autoload.php';

$dbLink = __DIR__ . '/io/sample_sqlite.db';
@unlink($dbLink);

$ESA = new \sinri\excelxsql\ExcelAgent\ExcelSqliteAgent($dbLink);

$ESA->loadXlsx("test2", __DIR__ . '/io/sample_input.xlsx');

$sql = "SELECT name,value,date FROM test2_Sheet1 WHERE name=:name";
$values = [":name" => "abe"];
$result = $ESA->getDb()->safeReadRow($sql, $values);
var_dump($result);

$sql = "SELECT address FROM test2_Sheet2 WHERE name=:name";
$values = [":name" => "abe"];
$result = $ESA->getDb()->safeReadCol($sql, $values);
var_dump($result);

$sql = "SELECT test2_Sheet1.name,test2_Sheet1.value,test2_Sheet1.date,test2_Sheet2.address FROM test2_Sheet1 INNER JOIN test2_Sheet2 ON test2_Sheet1.name=test2_Sheet2.name WHERE cast(test2_Sheet1.value AS INT)<:limit";
$values = [":limit" => 10];
$result = $ESA->getDb()->safeReadAll($sql, $values);
var_dump($result);