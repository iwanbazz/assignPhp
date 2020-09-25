<?php

if (!function_exists('file_type')) {
  function file_type($value)
  {
    return include './file_types/Type_' . $value . '.php';
  }
}

if (!function_exists('diedump')) {
  function diedump($value)
  {
    die(var_dump($value));
  }
}
