<?php
require_once '../db.php';
require_once '../security.php';
auth();

echo json_encode(db()->query("SELECT * FROM leads")->fetchAll(PDO::FETCH_ASSOC));
?>