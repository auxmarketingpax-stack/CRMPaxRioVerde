<?php
session_start();
function auth(){
 if(!isset($_SESSION['user_id'])){
  http_response_code(401);
  echo json_encode(["error"=>"unauthorized"]);
  exit;
 }
}
?>