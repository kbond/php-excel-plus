<style>
  table { border-bottom: 2px solid black; }
  thead { vertical-align: bottom; font-weight: bold; background: black; color: white; }
  td, th {padding: .2em; }
</style>

<?php
require_once __DIR__.'/../vendor/autoload.php';

$excel = new \PHPExcel_Plus;

$excel->load('list.xls');

$table = $excel->convertToComplexArray();
?>

<table>
  <?php foreach ($table as $section => $rows): ?>
  <?php if ($section == 'head'): ?>
      <thead>
    <?php else: ?>
      <tbody>
    <?php endif; ?>

    <?php foreach ($rows as $row): ?>
      <tr>
      <?php foreach ($row as $cell): ?>
        <<?php echo $cell['bold'] ? 'th' : 'td' ?><?php echo array_key_exists('rows',$cell) ? ' rowspan="'.$cell['rows'].'"' : ''?><?php echo array_key_exists('cols',$cell) ? ' colspan="'.$cell['cols'].'"' : ''?>>
          <?php echo $cell['value'] ?>
        </<?php echo $cell['bold'] ? 'th' : 'td' ?>>
      <?php endforeach; ?>
      </tr>
    <?php endforeach; ?>

    <?php if ($section == 'head'): ?>
      </thead>
    <?php else: ?>
      </tbody>
  <?php endif; ?>
  <?php endforeach; ?>
</table>