<?php
require_once 'config.php';
function db(){
 static $pdo;
 if(!$pdo){
  $pdo=new PDO("mysql:host=".DB_HOST.";dbname=".DB_NAME.";charset=utf8",DB_USER,DB_PASS);
 }
 return $pdo;
}
?>