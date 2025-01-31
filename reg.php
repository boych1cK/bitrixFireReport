<?
    require_once($_SERVER['DOCUMENT_ROOT'] . "/bitrix/modules/main/include/prolog_before.php");
    \Bitrix\Main\Loader::includeModule('crm');
    require_once($_SERVER['DOCUMENT_ROOT'] . '/xlsxPhpLib/autoload.php'); 
    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
    function getUFValue($row){
        global $DB;
        $query = 'SELECT * FROM `b_user_field_enum` WHERE ID = \''.$row.'\'';
        $results = $DB->Query($query);
        $val = $results->Fetch();
        return $val['VALUE'];
    }
    if(isset($_POST['startDate']) && isset($_POST['endDate'])){
        $arFilter = array(
            "CATEGORY_ID"=>7, //выбираем определенную сделку по ID
            "CHECK_PERMISSIONS"=>"N",
            ">=BEGINDATE" => \Bitrix\Main\Type\DateTime::createFromTimestamp(strtotime($_POST['startDate'])),
            "<=BEGINDATE" => \Bitrix\Main\Type\DateTime::createFromTimestamp(strtotime($_POST['endDate']))
        );
        $data = array();
        //$fields = [ 1041,1042, 1043,1044, 1047, 1048, 1049, 1050, 1051, 1052, 1053];
        $filter = ['CONTACT_ID','UF_CRM_1698915872','UF_CRM_1698916009','UF_CRM_1698916207','UF_CRM_1698916263','UF_CRM_1698916300','UF_CRM_1698916306','UF_CRM_1698917391','UF_CRM_1698917607','UF_CRM_1698917672','UF_CRM_1698917752','UF_CRM_1698917826','UF_CRM_1698917868','UF_CRM_1698917903','UF_CRM_1698917923','UF_CRM_1698917957','UF_CRM_1698917431','UF_CRM_1698917460','BEGINDATE'];
        $header = ['ФИО','Должность', 'Дата','Причина ухода','Оплата труда','График работы','Уровень нагрузки','Отн. в коллективе',
        'Неравномерная нагрузка','Нашли работу?','Рабочее место','Отн. руководителя','Система оплаты труда','Карьерный рост','Баланс р и лж','Забота о сотр.','Соцпакет','Сколько работал','Своя причина ухода', 'Своя оценка оплаты труда'];
        $res = CCrmDeal::GetListEx(Array(), $arFilter,false,false, $filter);
        $dataPost = array();
        while($row = $res->Fetch()){
            for($i=1; $i<count($filter)-4; $i++){
            global $DB;
                if(!is_array($row[$filter[$i]])){
                    $row[$filter[$i]]=getUFValue($row[$filter[$i]]);
                }else{
                    $string = '';
                    for($j=0; $j<count($row[$filter[$i]]); $j++){
        
                        $query = 'SELECT * FROM `b_user_field_enum` WHERE ID = \''.$row[$filter[$i]][$j].'\'';
                        $results = $DB->Query($query);
                        $val = $results->Fetch();
                        $string .=" ".$val['VALUE']."\r\n";
                    }
                    $row[$filter[$i]]=$string;
                    
                }
            }
            unset($row['ID']);
            $cont = CCrmContact::GetListEx(Array(),array('ID'=>$row["CONTACT_ID"]), false,false,array('*', 'UF_*'));
            $cont = $cont->Fetch();
            $name = $cont['LAST_NAME'].' '.$cont['NAME'].' '.$cont['SECOND_NAME'];
            $job ='';
            if(!$cont['UF_CRM_655195758DF18'] == null)
            {
                $job = $cont['UF_CRM_655195758DF18'];
            }else{
                $job .= getUFValue($cont['UF_CRM_6551ACDDB960C']);
            }
            unset($row['CONTACT_ID']);
            array_unshift($row , $job);
            array_unshift($row , $name);
            array_push($data, $row);
        }
            $spreadsheet = new Spreadsheet();
            $sheet = $spreadsheet->getActiveSheet();
            $sheet->fromArray($header, NULL, 'A1');
            $sheet->fromArray($data, NULL, 'A2');
            $sheet->getStyle('A1:T1')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FFA07A');
            foreach ($sheet->getColumnIterator() as $column) {
                $sheet->getColumnDimension($column->getColumnIndex())->setAutoSize(true);
            }
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment;filename="Отчет увольнение.xlsx"');
            header('Cache-Control: max-age=0');
        
        
            $writer = new Xlsx($spreadsheet);
          $writer->save('php://output');
            die;
    }
    
?>