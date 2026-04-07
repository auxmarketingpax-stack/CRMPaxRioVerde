<?php
require_once '../db.php';
$data=json_decode(file_get_contents("php://input"),true);

$pass=password_hash($data['password'],PASSWORD_DEFAULT);
db()->prepare("INSERT INTO users(name,email,password) VALUES(?,?,?)")
->execute([$data['name'],$data['email'],$pass]);

echo json_encode(["ok"=>true]);
?>