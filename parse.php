<?php

$urls = [
	'http://alexavto52.ru/g6301047-zapasnye-chasti-dlya',
	'http://alexavto52.ru/g6301047-zapasnye-chasti-dlya/page_2',
	'http://alexavto52.ru/g6301047-zapasnye-chasti-dlya/page_3',
];

$pages = [];

$data = [];

if (!file_exists('./cache'))
	mkdir('./cache');

/**
 * @param string $url
 * @return string content
 * @throws Exception
 */
function getContent($url)
{
	if (!$url)
		throw new \Exception('Invalid url');

	$key = md5($url);
	$cache_file = "./cache/$key";

	if (file_exists($cache_file))
		return file_get_contents($cache_file);

	$content = file_get_contents($url);
	file_put_contents($cache_file, $content);

	return $content;
}

foreach ($urls as $url)
{
	$content = getContent($url);

	$domDoc  = new \DOMDocument('1.0', 'utf-8');
	$domDoc->strictErrorChecking = false;
	$success = @$domDoc->loadHTML($content);

	$errors = libxml_get_errors();
	if ($errors)
		throw new \Exception('Invalid html document');
	if (!$success)
		throw new \Exception('Error parsing document');

	$xpath = new \DOMXPath($domDoc);
	$nodeList = $xpath->query("descendant-or-self::*[contains(concat(' ', normalize-space(@class), ' '), ' b-layout__clear ')]/descendant::a[contains(concat(' ', normalize-space(@class), ' '), ' b-centered-image ')]");

	/**
	 * @var \DOMElement $a
	 */
	foreach ($nodeList as $a)
	{
		$href = $a->getAttribute('href');

		if (!$href)
			continue;

		$pages[] = $href;
	}
}

foreach ($pages as $url)
{
	$content = getContent($url);

	$domDoc  = new \DOMDocument('1.0', 'utf-8');
	$domDoc->strictErrorChecking = false;
	$success = @$domDoc->loadHTML($content);

	$errors = libxml_get_errors();
	if ($errors)
		throw new \Exception('Invalid html document');
	if (!$success)
		throw new \Exception('Error parsing document');

	$xpath = new \DOMXPath($domDoc);
	$nodeTitle = $xpath->query("descendant-or-self::h1[contains(concat(' ', normalize-space(@class), ' '), ' b-product__name ')]")->item(0);

	if (!$nodeTitle)
		throw new \Exception('title not found');

	$title = trim($nodeTitle->textContent);

	$nodeCode = $xpath->query("descendant-or-self::*[contains(concat(' ', normalize-space(@class), ' '), ' b-product__info-holder ')]/descendant::*[contains(concat(' ', normalize-space(@class), ' '), ' b-product__sku ')]")->item(0);

	if (!$nodeCode)
		throw new \Exception('code not found');

	$code = trim($nodeCode->getAttribute('title'));

	$nodePrice = $xpath->query("descendant-or-self::*[contains(concat(' ', normalize-space(@class), ' '), ' b-product__info-holder ')]/descendant::*[contains(concat(' ', normalize-space(@class), ' '), ' b-product__price ')]")->item(0);

	if (!$nodePrice)
		throw new \Exception('price not found');

	$price = trim($nodePrice->textContent);

	$price = explode(' ', $price);
	$currency = $price[1];
	$price = $price[0];

	if (!$price || !$currency)
		throw new \Exception('Invalid price');

	$nodeDesc = $xpath->query("descendant-or-self::*[contains(concat(' ', normalize-space(@class), ' '), ' b-layout__content ')]/descendant::*[contains(concat(' ', normalize-space(@class), ' '), ' b-user-content ')]")->item(0);

	if (!$nodeDesc)
		throw new \Exception('description not found');

	$desc = trim($nodeDesc->textContent);

	$nodesSpec = $xpath->query("descendant-or-self::*[contains(concat(' ', normalize-space(@class), ' '), ' b-layout__content ')]/descendant::table[contains(concat(' ', normalize-space(@class), ' '), ' b-product-info ')]/descendant::td");

	if (!$nodesSpec || !$nodesSpec->length)
		throw new \Exception('Spec not found');

	$spec = [];
	$last_key = null;

	/**
	 * @var \DOMElement $s
	 */
	foreach ($nodesSpec as $s)
	{
		$s = trim($s->textContent);

		if ($last_key)
		{
			$spec[$last_key] = $s;
			$last_key = null;
		}
		else
			$last_key = $s;
	}

	$nodeImg = $xpath->query("descendant-or-self::*[contains(concat(' ', normalize-space(@class), ' '), ' b-product__container ')]/descendant::*[contains(concat(' ', normalize-space(@class), ' '), ' b-product__image ')]/descendant::img")->item(0);

	if (!$nodeImg)
		throw new \Exception('image not found');

	$img = trim($nodeImg->getAttribute('src'));

	$file_name = end(explode('/', $img));
	$img_src = getContent($img);
	file_put_contents($file_name, $img_src);

	$data[] = [
		'title' => $title,
		'code' => $code,
		'price' => $price,
		'currency' => $currency,
		'description' => $desc,
		'spec' => $spec,
		'img' => $file_name,
	];
}

if (!$data)
	throw new \Exception('Empty data');

$table = '';

foreach ($data as $d)
{
	$row = '<Row>';

	foreach ($d as $k => $v)
	{
		if (is_array($v))
		{
			$a = [];
			foreach ($v as $ks => $vs)
			{
				$a[] = $ks . ': ' . $vs;
			}

			$v = implode("\r\n", $a);
		}

		$row .= <<<TABLE
<Cell><Data ss:Type="String">{$v}</Data></Cell>
TABLE;
	}

	$row .= '</Row>';
	$table .= $row;
}


$tmpl = <<<TMPL
<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40">
	<Styles>
  		<Style ss:ID="bold">
			<Font ss:Bold="1"/>
		</Style>
 	</Styles>
	<Worksheet ss:Name="WorksheetName">
		<Table>
			$table
		</Table>
	</Worksheet>
</Workbook>
TMPL;

file_put_contents('export.xls', $tmpl);