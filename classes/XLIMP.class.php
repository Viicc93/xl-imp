<?php

class XLIMP {

  public function __construct (){
    // Add admin page for plugin
    add_action('admin_menu', array(&$this, 'admin_menu'));

  }

  // Add a page to admin menu add call  function
  public function admin_menu() {
    // Add page to admin menu
    // add_menu_page( page title, menu title, capability, menu slug, callble function, icon, position )
    add_menu_page('Excel Import', 'Excel Import', 'edit_theme_options', 'excel_import', array($this, 'admin_page'), 'dashicons-upload' );
  }

  // Admin page
  public function admin_page(){
    // upload file and insert to post
    $this->upload();
    ?>
    <!-- Page form -->
    <h2>Excel File Uploader</h2>
    <form enctype="multipart/form-data" action="#" method="post" >
      <label class="form-label span3" for="file">File</label>
      <input type="file" name="file" id="file" required />
      <br><br>
      <input class="button button-primary button-large" type="submit" value="<?php _e('Submit'); ?>" />
    </form>
  <?php
  }



  public function upload (){
    /** Set default timezone (will throw a notice otherwise) */
    date_default_timezone_set('Europe/Berlin');
    // Include library for excel reader
    include PLUGIN_DIR_PATH . '/libraries/PHPExcel/IOFactory.php';

    $error_count = 0; // count errors of insert to post

    /*
    *
    * CUSTOMIZE ERROR MESSAGES
    *
    */
    // Error Messages
    $error_message_01 = __('Wrong File, please use XLSX or CSV', 'importer');
    $error_message_02 = __('Error loading file', 'importer' );
    $error_message_03 = __('Error insert fish nr', 'importer' );

    // Success Message
    $success_message_start = __( 'Succeed to import', 'importer' );
    $success_message_end = __( 'fishes', 'importer' );



    if(isset($_FILES['file']['name'])){

      $file_name = $_FILES['file']['name'];
      $ext = pathinfo($file_name, PATHINFO_EXTENSION);

      //Checking the file extension
      if($ext == "xlsx" || $ext == "csv" ){

        $file_name = $_FILES['file']['tmp_name'];
        $inputFileName = $file_name;

        /**********************PHPExcel Script to Read Excel File**********************/
        //  Read your Excel workbook
        try {
          $inputFileType = PHPExcel_IOFactory::identify($inputFileName); //Identify the file
          $objReader = PHPExcel_IOFactory::createReader($inputFileType); //Creating the reader
          $objPHPExcel = $objReader->load($inputFileName); //Loading the file
        }
        catch (Exception $e) {
          die('<div class="error notice"><p style="color:red;">' . $error_message_02 . ' "' . pathinfo($inputFileName, PATHINFO_BASENAME)
          . '"</p></div><br><br><div class="error notice">' . $e->getMessage() . '</p></div>');
        }

          //  Get worksheet dimensions
          $sheet = $objPHPExcel->getSheet(0);     //Selecting sheet 0
          $highestRow = $sheet->getHighestRow();     //Getting number of rows
          $highestColumn = $sheet->getHighestColumn();     //Getting number of columns

          //  Loop through each row of the worksheet in turn
          for ($row = 1; $row <= $highestRow; $row++) {

            //  Read a row of data into an array
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, NULL, TRUE, FALSE);

            /*
            *
            * CHANGES HERE FOR CUSTOUMIZE INSERT VARIABLES
            *
            */

            // Values from flie
            $art = $rowData[0][2];
            $color = $rowData[0][3];

            // Post Content can't be null
            if ( $rowData[0][6] != null ) {
              $content = $rowData[0][6];
            }
            else {
              $content = ' ';
            }

            // Custom Post Metas
            $meta1 = $rowData[0][0];
            $meta2 = $rowData[0][4];
            $meta3 = $rowData[0][5];

            // Post Title
            $title = $art . ' ' . $color;

            /*
            * CHANGE HERE TO INSERT INTO POST
            */
            $postarr = array(
              'post_title' => $title,
              'post_content' => $content,
              'post_status' => 'publish',
              'post_type' => 'post',
              'meta_input'   => array(
                'meta_1' => $meta1,
                'meta_2' => $meta2,
                'meta_3' => $meta3,
              ),
            );

            $post = wp_insert_post( $postarr, true );

            // Error and success messages
            if (is_wp_error( $post )) {
              $error_count++;
              echo '<div class="error notice"><p>' . $error_message_03 . ' ' . $row . '</p></div>';
            }
            if ($row == $highestRow) {
              $success_count = $row - $error_count;
              echo '<div class="updated notice"><p>' . $success_message_start . ' ' . $success_count . '/' . $row . ' ' . $success_message_end . '</p></div>';
            }

          }

      }
      else {
        echo '<div class="error notice"><p style="color:red;">' . $error_message_01 . '</p></div>';
      }

    }

  }



}
