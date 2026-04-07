<?php
require_once '../db.php';
session_start();
$data=json_decode(file_get_contents("php://input"),true);

$stmt=db()->prepare("SELECT * FROM users WHERE email=?");
$stmt->execute([$data['email']]);
$user=$stmt->fetch(PDO::FETCH_ASSOC);

if($user && password_verify($data['password'],$user['password'])){
 $_SESSION['user_id']=$user['id'];
 echo json_encode(["ok"=>true]);
}else{
 echo json_encode(["error"=>"login inválido"]);
}
?>