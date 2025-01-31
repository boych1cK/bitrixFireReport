<? 
require_once($_SERVER['DOCUMENT_ROOT'] . "/bitrix/modules/main/include/prolog_before.php");
$today = date("Y-m-d");
use Bitrix\UI\Buttons\Button;
\Bitrix\Main\Loader::includeModule('crm');
require_once($_SERVER['DOCUMENT_ROOT'] . '/xlsxPhpLib/autoload.php'); 
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Bitrix\Main\UserFieldTable;

?>
<!DOCTYPE html>
<html style="height: 100%" lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Сводный отчет</title>
</head>
<body style="height: 100%">
    <div style="display: flex; width: 100%; height: 100%; justify-content: center; align-items: center;">
            <div>
            <h1>Выберите временной промежуток</h1>
            <form action="reg.php" method="POST">
                <input style="display: block; margin: 0 auto;" type="date" id="start" name="startDate" value="<?=$today?>" max="<?=$today?>" />
                <p style="display: block; margin: 0 auto; width: fit-content;">-</p>
                <input style="display: block; margin: 0 auto;" type="date" id="end" name="endDate" value="<?=$today?>" max="<?=$today?>" />
                <input style="display: block; margin: 0 auto; margin-top: 20px" type="submit" value="Сформировать">
            </form>
        </div>
    </div>
</body>
</html>