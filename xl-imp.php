<?php
/**
* Plugin Name: Excel Import
* Plugin URI:
* Description: Import to post from excel.
* Version: 0.1
* Author: Mr. Bruno
* Author URI: Author's website
*/

define('PLUGIN_PATH', __FILE__);
define('PLUGIN_DIR_PATH', __DIR__);

include_once(PLUGIN_DIR_PATH . '/classes/XLIMP.class.php');
new XLIMP();
