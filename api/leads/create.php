<?php
require_once '../db.php';
require_once '../security.php';
auth();

$data=json_decode(file_get_contents("php://input"),true);

db()->prepare("INSERT INTO leads(name,contact,owner,value) VALUES(?,?,?,?)")
->execute([$data['name'],$data['contact'],$data['owner'],$data['value']]);

echo json_encode(["ok"=>true]);
?>