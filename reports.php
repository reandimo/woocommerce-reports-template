<?php

/*
Plugin Name: Woocommerce Orders Report
Plugin URI: https://nativo.team
Version: 1.0
Author: Renan Diaz 
*/

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Automattic\WooCommerce;
use \WC_Order_Query;
use \WC_Order;

class Orders_Report
{

    /**
     * Run the actions
     */
    public function run()
    {

        add_action('admin_menu', array($this, 'addSubmenuPages'));

        // Reporte
        // Add action hook only if action=downloadOrdersExcelExport
        if (isset($_POST['action']) && $_POST['action'] == 'downloadOrdersExcelExport') {
            // Handle CSV Export
            add_action('admin_init', array($this, 'ordersExcelExport'));
        } 

    }

    function addSubmenuPages()
    {

        add_submenu_page(
            'woocommerce',
            'Woocommerce Orders Report',
            'Woocommerce Orders Report',
            'manage_options',
            'orders-excel-reports',
            array($this, 'ordersExcelExportForm')
        );
    }

    function ordersExcelExportForm()
    {

        echo '<div class="wrap"><div id="icon-tools" class="icon32"></div>';
        echo '<h2>Woocommerce Orders Report</h2>';
        echo '</div>';

        ?>
                <form method="post" action="<?= admin_url('tools.php') ?>">
                    <input type="hidden" name="action" value="downloadOrdersExcelExport">
                    <input type="hidden" name="_wpnonce" value="<?= wp_create_nonce('downloadOrdersExcelExport') ?>">
                    <br>
                    <label for="">Desde</label>
                    <input type="date" name="start" required="">
                    <br>
                    <br>
                    <label for="">Hasta</label>
                    <input type="date" name="end" required="">
                    <br>
                    <br>
                    <label for="">Status</label>
                    <select name="status" required="">
                        <option value="all">All</option>
                        <option value="pending">Payment Pending</option>
                        <option value="processing">Processing</option>
                        <option value="completed">Completed</option>
                        <option value="refunded">Refunded</option>
                        <option value="failed">Failed</option>
                        <option value="cancelled">Cancelled</option>
                    </select>
                    <br>
                    <br>
                    <button class="button">Download</button>
                </form>
        <?php


    }

    function ordersExcelExport()
    {

        // Check for current user privileges 
        if (!current_user_can('manage_options')) {
            return false;
        }
        // Check if we are in WP-Admin
        if (!is_admin()) {
            return false;
        }
        // Nonce Check
        $nonce = isset($_POST['_wpnonce']) ? $_POST['_wpnonce'] : '';
        if (!wp_verify_nonce($nonce, 'downloadOrdersExcelExport')) {
            die('Security check error');
        }

        $args = array(
            'limit' => -1,
            'return' => 'ids',
            // 'meta_key' => '_billing_state', // Postmeta key field
            // 'meta_value' => 'CA', // Postmeta value field
        );

        if ($_POST['status'] == 'completed') {
            $args['date_completed'] = $_POST['start'] . '...' . $_POST['end'];
        } else {
            $args['date_created'] = $_POST['start'] . '...' . $_POST['end'];
        }

        // check status
        if ($_POST['status'] !== 'all') {
            $args['status'] = $_POST['status'];
        }

        $query = new \WC_Order_Query($args);
        $customer_orders = $query->get_orders();

        $excel = array();

        foreach ($customer_orders as $id_order) {

            $order = new \WC_Order($id_order);

            if (count($order->get_items()) == 0) {
                continue;
            }

            // Get and Loop Over Order Items
            foreach ($order->get_items() as $item_id => $item) { 

                $product_id = $item->get_product_id();
                $variation_id = $item->get_variation_id();
                $product = $item->get_product(); // see link above to get $product info
                $product_name = $item->get_name();
                $quantity = $item->get_quantity();
                $subtotal = $item->get_subtotal();
                $total = $item->get_total();
                $tax = $item->get_subtotal_tax();
                $tax_class = $item->get_tax_class();
                $tax_status = $item->get_tax_status();
                $allmeta = $item->get_meta_data();
                $somemeta = $item->get_meta( '_whatever', true );
                $item_type = $item->get_type(); // e.g. "line_item", "fee"

                $excel[] =  array(
                    $order->get_date_created()->format('m/d/Y'),
                    $order->get_status(),
                    $order->get_id(),
                    $order->get_billing_first_name() . ' ' . $order->get_billing_last_name(),
                    $order->get_billing_phone(),
                    $order->get_billing_email(),
                    $order->get_billing_address_1() . ' ' . $order->get_billing_address_2(),
                    $order->get_billing_state(),
                    $order->get_billing_city(),
                    $item->get_name(),
                    $item->get_quantity(),
                    $item->get_total(),
                );
            }
        }

        $header = array(
            'Order Date',
            'Status',
            'Order ID',
            'Customer',
            'Phone',
            'Email',
            'Billing Address',
            'Billing State',
            'Billing City',
            'Item',
            'Quantity',
            'Cost',
        );

        array_unshift($excel, $header); 

        // Excel
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $spreadsheet->getProperties()
            ->setCreator('Woocommerce Report Web')
            ->setLastModifiedBy('Woocommerce Report Web')
            ->setTitle("Orden Woocommerce Report " . $_POST['status'] . ' ' . date('Y-m-d') . ".xlsx")
            ->setSubject("Orden Woocommerce Report")
            ->setKeywords("office 2007  openxml php");

        $spreadsheet->setActiveSheetIndex(0); 

        // DATA
        $spreadsheet->getActiveSheet()
            ->fromArray(
                $excel,  // The data to set
                NULL,        // Array values with this value will not be set
                'A1'         // Top left coordinate of the worksheet range where
                //    we want to set these values (default is A1)
            );

        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('J')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('K')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('L')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('M')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('N')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('O')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('P')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('Q')->setAutoSize(true);

        /* Here there will be some code where you create $spreadsheet */

        // redirect output to client browser
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="Reporte Woocommerce - ' . $_POST['status'] . ' ' . date('Y-m-d') . '.xlsx"');
        header('Cache-Control: max-age=0');

        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');

        $writer->save('php://output', 'Xlsx');

        die();

    } 
 
}

$_reports = new Orders_Report();
$_reports->run();
