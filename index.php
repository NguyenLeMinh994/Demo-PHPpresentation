<?php

require_once './PhpPresentation/src/PhpPresentation/Autoloader.php';
\PhpOffice\PhpPresentation\Autoloader::register();


// with Composer
require_once './PHPPresentation/vendor/autoload.php';


use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Color;
// use PhpOffice\PhpPresentation\Slide\Background\Color as ColorBackground;

use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Slide\Background\Image;
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;







$objPHPPowerPoint = new PhpPresentation();
$objPHPPowerPoint->getLayout()->setDocumentLayout(['cx' => 1280, 'cy' => 700], true)
      ->setCX(25, DocumentLayout::UNIT_INCH)
      ->setCY(14, DocumentLayout::UNIT_INCH);
$currentSlide = $objPHPPowerPoint->getActiveSlide();


// Create a shape (text)
// $objPHPPresentation->createSlide();
echo "=============  </br>";
echo "Slide 1  </br>";
$shape = $currentSlide->createRichTextShape()
      ->setHeight(282.24)
      ->setWidth(1611.84)
      ->setOffsetX(175.68)
      ->setOffsetY(1010.88);
$textRun = $shape->createTextRun('DC to Azure Migration Workshop');
$textRun->getFont()->setBold(true)
      ->setSize(82)
      ->setName('Calibri Light (Headings)')
      ->setColor(new Color('FFFFFF'));
$shape->createParagraph()->setLineSpacing(200);

$textRun = $shape->createTextRun('Dik van Brummen and Herb Prooy, the CloudLab on behalf of Microsoft');
$textRun->getFont()->setBold(true)
      ->setSize(40)
      ->setName('Calibri Light (Headings)')
      ->setColor(new Color('FFFFFF'));

echo "Image Background </br>";
$oBkgImage = new Image();
$oBkgImage->setPath('./images/p1.png');
$currentSlide->setBackground($oBkgImage);

echo "============= </br>";

echo "============= </br>";
echo "Slide 2 </br>";
$currentSlide = $objPHPPowerPoint->createSlide();

$shape = $currentSlide->createRichTextShape()->setRotation(270);
$shape->setHeight(106.56)
      ->setWidth(591.36)
      ->setOffsetX(-201.6)
      ->setOffsetY(488.64);
$textRun = $shape->createTextRun("About us");
$textRun->getFont()
      ->setSize(60)
      ->setName('Gotham Medium')
      ->setColor(new Color('FF4672A8'));

$shape = $currentSlide->createLineShape(0, 0, 0, 0);
$shape->setHeight(1132.8)
      ->setOffsetX(147.84)
      ->setOffsetY(211.2);

$shape = $currentSlide->createLineShape(0, 0, 0, 0);
$shape->setWidth(823.68)
      ->setOffsetX(147.84)
      ->setOffsetY(474.24);


$shape = $currentSlide->createRichTextShape();
$shape->setHeight(527.04)
      ->setWidth(724.8)
      ->setOffsetX(287.04)
      ->setOffsetY(587.52);

$textRun = $shape->createTextRun("The CloudLab is founded by industry veterans and serial entrepreneurs in the ASP, SaaS and Cloud business, strongly connected with Microsoft. As an experienced team we have a strong track record of building successful cloud companies and creating profitable exits. Our approach is direct, entrepreneurial and very much focused on sustainable results.");
$textRun->getFont()
      ->setSize(24)
      ->setName('Gotham Medium')
      ->setColor(new Color('FF000000'));
// Border

$shape = $currentSlide->createRichTextShape();
$shape->setHeight(417.6)
      ->setWidth(504)
      ->setOffsetX(971.52)
      ->setOffsetY(198.72)
      ->getBorder()->setColor(new Color('FF4672A8'))->setDashStyle(Border::DASH_SOLID)->setLineStyle(Border::LINE_SINGLE);

$shape = $currentSlide->createRichTextShape();
$shape->setHeight(417.6)
      ->setWidth(364.8)
      ->setOffsetX(1160.64)
      ->setOffsetY(887.04)
      ->getBorder()->setColor(new Color('FF4672A8'))->setDashStyle(Border::DASH_SOLID)->setLineStyle(Border::LINE_SINGLE);

$shape = $currentSlide->createRichTextShape();
$shape->setHeight(599.04)
      ->setWidth(432.96)
      ->setOffsetX(1761.6)
      ->setOffsetY(288)
      ->getBorder()->setColor(new Color('FF4672A8'))->setDashStyle(Border::DASH_SOLID)->setLineStyle(Border::LINE_SINGLE);

//Text

$shape = $currentSlide->createRichTextShape();
$shape->setHeight(64.32)
      ->setWidth(360.96)
      ->setOffsetX(1117.44)
      ->setOffsetY(532.8);

$textRun = $shape->createTextRun("Dik van Brummen");
$textRun->getFont()
      ->setSize(34)
      ->setName('Gotham Medium')
      ->setColor(new Color('FF4672A8'));


$shape = $currentSlide->createRichTextShape();
$shape->setHeight(64.32)
      ->setWidth(552.96)
      ->setOffsetX(1117.44)
      ->setOffsetY(1206.72);
$textRun = $shape->createTextRun("        Rainer Strassner");
$textRun->getFont()
      ->setSize(34)
      ->setName('Gotham Medium')
      ->setColor(new Color('FF4672A8'));


$shape = $currentSlide->createRichTextShape();
$shape->setHeight(64.32)
      ->setWidth(360.96)
      ->setOffsetX(1988.16)
      ->setOffsetY(985.92);
$textRun = $shape->createTextRun("Herb Prooy");
$textRun->getFont()
      ->setSize(34)
      ->setName('Gotham Medium')
      ->setColor(new Color('FF4672A8'));

// Image Person


$shape = $currentSlide->createDrawingShape();
$shape->setPath("./images/person1.png")
      ->setHeight(493.44)
      ->setWidth(494.4)
      ->setOffsetX(1060.8)
      ->setOffsetY(44.16);

$shape = $currentSlide->createDrawingShape();
$shape->setPath("./images/person2.png")
      ->setHeight(491.52)
      ->setWidth(310.08)
      ->setOffsetX(1258.56)
      ->setOffsetY(717.12);

$shape = $currentSlide->createDrawingShape();
$shape->setPath("./images/person3.png")
      ->setHeight(619.2)
      ->setWidth(498.24)
      ->setOffsetX(1842.24)
      ->setOffsetY(362.88);

echo "============= </br>";

echo "Slide 3 </br></br>";
$currentSlide = $objPHPPowerPoint->createSlide();

echo "Shape 1 </br></br>";
$shape = $currentSlide->createRichTextShape();
$shape->setHeight(1147.2)
      ->setWidth(809.28)
      ->setOffsetX(343.68)
      ->setOffsetY(70.08)
      ->getBorder()->setColor(new Color('FF4672A8'))->setDashStyle(Border::DASH_SOLID)->setLineStyle(Border::LINE_SINGLE);

//Text left 
$shape = $currentSlide->createRichTextShape();
$shape->setWidth(476.16)
      ->setHeight(1102.08)
      ->setOffsetX(439.68)
      ->setOffsetY(94.08);

$arrTextLeft = ['Amdahl Computer', 'Apollo Computers', 'Apricot Computers', 'Atari ', 'Burroughs', 'Bull', 'Commodore', 'Compaq', 'CDC', 'Convex Computer', 'Data General', 'Digital Equipment', 'Honeywell', 'Next', 'Nixdorf Computers', 'Packard Bell'];
foreach ($arrTextLeft as $val) {
      $textRun = $shape->createTextRun($val);
      $textRun->getFont()
            ->setSize(40)
            ->setName('Calibri')
            ->setColor(new Color('FF4672A8'));

      $shape->createParagraph();
}

echo "Shape 2 </br></br>";

//Text Right 
$shape = $currentSlide->createRichTextShape();
$shape->setHeight(1147.2)
      ->setWidth(809.28)
      ->setOffsetX(1247.04)
      ->setOffsetY(70.08)
      ->getBorder()->setColor(new Color('FF4672A8'))
      ->setDashStyle(Border::DASH_SOLID)
      ->setLineStyle(Border::LINE_SINGLE);


$shape = $currentSlide->createRichTextShape();
$shape->setWidth(637.44)
      ->setHeight(1108.8)
      ->setOffsetX(1341.12)
      ->setOffsetY(80.64);
$arrTextRight = ['Philips', 'Prime Computer', 'Quantex Microsystems', 'Remmington', 'Sequent Systems', 'Siemens', 'Silicon Graphics', 'Sperry  Univac', 'Stratus Computer', 'Sun Microsystems', 'Tandon Corporation', 'Tandy Corporation', 'Texas Instruments', 'Tulip Computers', 'Wang Laboratories', 'Xerox', 'Zenith Data Systems'];
foreach ($arrTextRight as $val) {
      $textRun = $shape->createTextRun($val);
      $textRun->getFont()
            ->setSize(40)
            ->setName('Calibri')
            ->setColor(new Color('FF4672A8'));

      $shape->createParagraph();
}
echo "============= </br>";

echo "Slide 4 </br>";
$currentSlide = $objPHPPowerPoint->createSlide();
$shape = $currentSlide->createDrawingShape();
$shape->setPath('./images/p2.png')
      ->setHeight(934.08)
      ->setWidth(1951.68)
      ->setOffsetX(223.68)
      ->setOffsetY(196.8);

echo "============= </br>";

echo "Slide 5 </br>";
$currentSlide = $objPHPPowerPoint->createSlide();

// Shape left

//===========================================
$shape = $currentSlide->createRichTextShape()
      ->setHeight(547.2)
      ->setWidth(445.44)
      ->setOffsetX(192.96)
      ->setOffsetY(229.44);

$shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER)->setVertical(Alignment::VERTICAL_CENTER);

$shape->getFill()->setFillType(Fill::FILL_SOLID)->setRotation(90)->setStartColor(new Color('FF0F85F1'))->setEndColor(new Color('FF0F85F1'));
$textRun = $shape->createTextRun('Differentiation');
$textRun->getFont()->setBold(true)
      ->setSize(34)
      ->setName('Calibri')
      ->setColor(new Color('FFFFFFFF'));
//===========================================

//===========================================

$shape = $currentSlide->createRichTextShape()
      ->setHeight(547.2)
      ->setWidth(445.44)
      ->setOffsetX(664.32)
      ->setOffsetY(229.44);

$shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER)->setVertical(Alignment::VERTICAL_CENTER);

$shape->getFill()->setFillType(Fill::FILL_SOLID)->setRotation(90)->setStartColor(new Color('FFC00000'))->setEndColor(new Color('FFC00000'));
$textRun = $shape->createTextRun('Overall Cost Leadership');
$textRun->getFont()->setBold(true)
      ->setSize(34)
      ->setName('Calibri')
      ->setColor(new Color('FFFFFFFF'));
//===========================================

//===========================================
$shape = $currentSlide->createRichTextShape()
      ->setHeight(284.16)
      ->setWidth(915.84)
      ->setOffsetX(192.96)
      ->setOffsetY(809.28);

$shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER)->setVertical(Alignment::VERTICAL_CENTER);

$shape->getFill()->setFillType(Fill::FILL_SOLID)->setRotation(90)->setStartColor(new Color('FFFF9900'))->setEndColor(new Color('FFFF9900'));
$textRun = $shape->createTextRun('Focus');
$textRun->getFont()->setBold(true)
      ->setSize(34)
      ->setName('Calibri')
      ->setColor(new Color('FFFFFFFF'));
//===========================================

//===========================================
$shape = $currentSlide->createRichTextShape()
      ->setHeight(246.72)
      ->setWidth(695.04)
      ->setOffsetX(304.32)
      ->setOffsetY(672);

// $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER)->setVertical(Alignment::VERTICAL_CENTER);

$shape->getFill()->setFillType(Fill::FILL_SOLID)->setRotation(90)->setStartColor(new Color('FFFFFFFF'))->setEndColor(new Color('FFFFFFFF'));
// $textRun = $shape->createTextRun('Focus');
// $textRun->getFont()->setBold(true)
//       ->setSize(34)
//       ->setName('Calibri')
//       ->setColor(new Color('FFFFFFFF'));
//===========================================


//===========================================
$shape = $currentSlide->createRichTextShape()
      ->setHeight(170.88)
      ->setWidth(569.28)
      ->setOffsetX(366.72)
      ->setOffsetY(707.52);

$shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER)->setVertical(Alignment::VERTICAL_CENTER);

$shape->getFill()->setFillType(Fill::FILL_SOLID)->setRotation(90)->setStartColor(new Color('FF00B050'))->setEndColor(new Color('FF00B050'));
$textRun = $shape->createTextRun('Stuck in the middle');
$textRun->getFont()->setBold(true)
      ->setSize(34)
      ->setName('Calibri')
      ->setColor(new Color('FFFFFFFF'));
//===========================================


//===========================================
$shape = $currentSlide->createRichTextShape()
      ->setHeight(864)
      ->setWidth(959.04)
      ->setOffsetX(1248)
      ->setOffsetY(229.44);

$shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

$shape->getFill()->setFillType(Fill::FILL_SOLID)->setRotation(90)->setStartColor(new Color('FF0F85F1'))->setEndColor(new Color('FF0F85F1'));
$textRun = $shape->createTextRun('Porter’s Generic Competitive Strategies');
$textRun->getFont()->setBold(true)
      ->setSize(36)
      ->setName('Calibri')
      ->setColor(new Color('FF000000'));

$shape->createParagraph();
$shape->createParagraph();
$shape->createParagraph();
$shape->createParagraph();


$textRun = $shape->createTextRun("“ The firm failing to develop its strategy in at least one of the three directions –a firm that is “");
$textRun->getFont()
      ->setSize(44)
      ->setName('Calibri')
      ->setColor(new Color('FFFFFFFF'));

$textRun = $shape->createTextRun('stuck in the middle');
$textRun->getFont()
      ->setBold(true)
      ->setSize(44)
      ->setName('Calibri')
      ->setColor(new Color('FFFFFFFF'));

$textRun = $shape->createTextRun('” – is in an extremely poor strategic direction”');
$textRun->getFont()
      ->setSize(44)
      ->setName('Calibri')
      ->setColor(new Color('FFFFFFFF'));

$shape->createParagraph();
$shape->createParagraph();


$textRun = $shape->createTextRun('-Michael E. Porter, 1980');
$textRun->getFont()
      ->setSize(34)
      ->setName('Calibri')
      ->setColor(new Color('FF000000'));
//===========================================

// Shape right
echo "============= </br>";

$oWriterPPTX = IOFactory::createWriter($objPHPPowerPoint, 'PowerPoint2007');
$oWriterPPTX->save(__DIR__ . "/result/sample.pptx");
echo "Done";
echo "<a href='./result/sample.pptx'>PPTX</a>";
