<?php
/**
 * Created by PhpStorm.
 * User: Sinri
 * Date: 2017/10/2
 * Time: 12:50
 */

require_once __DIR__ . '/../autoload.php';

$dbLink = __DIR__ . '/io/sample_sqlite.db';
@unlink($dbLink);

$ESA = new \sinri\excelxsql\ExcelAgent\ExcelSqliteAgent($dbLink);

$ESA->loadXlsxWithIndexFormat("test2", __DIR__ . '/io/sample_input.xlsx');

$sql = "SELECT col_1,col_2,col_3 FROM test2_sheet_0 WHERE col_1=:name";
$values = [":name" => "abe"];
$result = $ESA->getDb()->safeReadRow($sql, $values);
var_dump($result);

$sql = "SELECT col_2 FROM test2_sheet_1 WHERE col_1=:name";
$values = [":name" => "abe"];
$result = $ESA->getDb()->safeReadCol($sql, $values);
var_dump($result);

$sql = "SELECT test2_sheet_0.col_1,test2_sheet_0.col_2,test2_sheet_0.col_3,test2_sheet_1.col_2 
  FROM test2_sheet_0 INNER JOIN test2_sheet_1 ON test2_sheet_0.col_1=test2_sheet_1.col_1 
  WHERE cast(test2_sheet_0.col_2 AS INT)<:limit
";
$values = [":limit" => 10];
$result = $ESA->getDb()->safeReadAll($sql, $values);
var_dump($result);