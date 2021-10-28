<?php
session_start();
$_SESSION['tipo'] = null;
header('location:login.php');