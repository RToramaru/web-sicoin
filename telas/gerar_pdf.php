<?php
require "../banco_dados/conecta.php";
require "../vendor/autoload.php";


use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\Settings;

$templateProcessor = new \PhpOffice\PhpWord\TemplateProcessor('modelo.docx');
 
$tipo = $_POST['tipo'];
$codigo = $_POST['codigo'];
$produto = $_POST['produto'];
$setor = $_POST['setor'];
$inicio = $_POST['inicio'];
$fim = $_POST['fim'];

$tipo_text = $tipo == 'entrada' ? "entrada_produto" : "saida_produto";
$cod_text = $codigo == '' ? "1 = 1" : "codigo = '$codigo'";
$prod_text = $produto == '' ? "2 = 2" : "produto = '$produto'";
$setor_text = $setor == '' ? "3 = 3" : "destino = '$setor'";
$inicio_text = $inicio == '' ? "4 = 4" : "emissao >= '$inicio'";
$fim_text = $fim == '' ? "5 = 5" : "emissao <= '$fim'";

$sql = "SELECT * FROM $tipo_text, $tipo WHERE $cod_text and $prod_text and $setor_text and $inicio_text and $fim_text and $tipo_text.codigo = $tipo.codigo";

if($codigo != ''){
  $sql = "SELECT * FROM $tipo, $tipo_text WHERE $tipo_text.codigo = '$codigo'";
}

$resultado = mysqli_query($conn, $sql);
$total = 0;
$row = mysqli_num_rows($resultado);
$templateProcessor->cloneRow('codigo', $row);

$count = 1;
while($row = mysqli_fetch_assoc($resultado)){
   
  $templateProcessor->setValue('codigo#'.$count, $row['codigo']);
  $templateProcessor->setValue('origem#'.$count, $row['destino']);
  $templateProcessor->setValue('emissao#'.$count, $row['emissao']);
  $templateProcessor->setValue('produto#'.$count, $row['produto']);
  $templateProcessor->setValue('preco#'.$count, $row['valor']);
  $templateProcessor->setValue('quantidade#'.$count, $row['quantidade']);
  $templateProcessor->setValue('total#'.$count, $row['valor']*$row['quantidade']);
  
  $count += 1;
  $total += ($row['valor']*$row['quantidade']);
}

$templateProcessor->setValue('valor_total', $total);
$templateProcessor->saveAs('modelo_.docx');


$phpWord = \PhpOffice\PhpWord\IOFactory::load('modelo_.docx','Word2007');
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'HTML');
$objWriter->save('hello.html');
$dompdf = new \Dompdf\Dompdf();
$dompdf->load_html(file_get_contents('hello.html'));
$dompdf->setPaper('A4', 'landscape');
$dompdf->render();
$pdf_string = $dompdf->output();
$dompdf->stream(
  "form.pdf", 
  array(
    "Attachment" => false
  )
);
file_put_contents('result2.pdf', $pdf_string);


//header('location:relatorio.php');

